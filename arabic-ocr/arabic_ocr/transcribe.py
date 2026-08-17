"""Pass 1 — Surya transcription and layout / reading-order handling.

Surya transcribes what is physically on the page. This is the authoritative
text and is never overwritten automatically. Its errors look like garbage and
are therefore catchable by a human, which is exactly what this design wants.

Text handling rules enforced here:
  * Tashkīl (vocalization marks) is preserved exactly as Surya recognizes it.
    Nothing is stripped, normalized, or added.
  * No arabic-reshaper and no bidi algorithm is applied to the stored text.
    Reshaping substitutes presentation forms for the real characters and would
    corrupt the file for every downstream consumer, including the translation
    tool.
  * No RLM/LRM directional marks are inserted.

The module sets TORCH_DEVICE from an autodetect (defaulting to CPU) before the
Surya models are imported, and processes images in batches — Surya is
substantially faster batched than one image per call.
"""

from __future__ import annotations

import os
from dataclasses import dataclass

from PIL import Image

# Surya is substantially faster batched. Tune this if memory is tight.
BATCH_SIZE = 8

# Surya layout region labels that belong to the editor's footnote apparatus.
FOOTNOTE_LABELS = {"Footnote"}


@dataclass
class PageTranscription:
    """Result of transcribing one page."""

    body: str = ""
    notes: str = ""
    regions_detected: int = 0
    layout_uncertain: bool = False
    ok: bool = True


def autodetect_device() -> str:
    """Return "cuda" if a CUDA GPU is usable, else "cpu"."""
    try:
        import torch

        return "cuda" if torch.cuda.is_available() else "cpu"
    except Exception:
        return "cpu"


def _line_boxes(ocr_result) -> list[tuple[str, tuple[float, float, float, float]]]:
    """Extract (text, bbox) pairs from a Surya OCR result, defensively."""
    lines = []
    for line in getattr(ocr_result, "text_lines", []) or []:
        text = (getattr(line, "text", "") or "").rstrip("\n")
        if not text.strip():
            continue
        bbox = getattr(line, "bbox", None)
        if not bbox or len(bbox) < 4:
            # No geometry — place it at the top so it still lands in the body.
            bbox = (0.0, 0.0, 0.0, 0.0)
        lines.append((text, tuple(float(v) for v in bbox[:4])))
    return lines


def _layout_regions(layout_result) -> list[tuple[str, tuple[float, float, float, float]]]:
    """Extract (label, bbox) pairs from a Surya layout result, defensively."""
    regions = []
    for region in getattr(layout_result, "bboxes", []) or []:
        label = getattr(region, "label", "") or ""
        bbox = getattr(region, "bbox", None)
        if not bbox or len(bbox) < 4:
            continue
        regions.append((label, tuple(float(v) for v in bbox[:4])))
    return regions


def _overlap_area(a: tuple[float, ...], b: tuple[float, ...]) -> float:
    ix0 = max(a[0], b[0])
    iy0 = max(a[1], b[1])
    ix1 = min(a[2], b[2])
    iy1 = min(a[3], b[3])
    if ix1 <= ix0 or iy1 <= iy0:
        return 0.0
    return (ix1 - ix0) * (iy1 - iy0)


def _is_footnote_line(
    line_bbox: tuple[float, ...],
    regions: list[tuple[str, tuple[float, float, float, float]]],
) -> bool:
    """Return True when the line's best-overlapping region is a footnote region."""
    best_label = None
    best_area = 0.0
    for label, rbbox in regions:
        area = _overlap_area(line_bbox, rbbox)
        if area > best_area:
            best_area = area
            best_label = label
    if best_area <= 0.0 or best_label is None:
        # No region contains this line — keep it in the body rather than drop it.
        return False
    return best_label in FOOTNOTE_LABELS


class SuryaEngine:
    """Wraps Surya's OCR + layout predictors and produces PageTranscriptions."""

    def __init__(self) -> None:
        self.device = autodetect_device()
        os.environ.setdefault("TORCH_DEVICE", self.device)

        # Import only after TORCH_DEVICE is set — Surya reads it at import time.
        self._rec, self._det, self._layout = _build_predictors()

    # ------------------------------------------------------------------ #
    # Public API
    # ------------------------------------------------------------------ #
    def transcribe_batch(
        self, images: list[Image.Image], separate_footnotes: bool
    ) -> list[PageTranscription]:
        """Transcribe a batch of page images.

        On a batch-level failure the batch is retried one image at a time so a
        single bad scan cannot lose the rest of the batch.
        """
        try:
            ocr_results = self._run_ocr(images)
            layout_results = (
                self._run_layout(images) if separate_footnotes else [None] * len(images)
            )
        except Exception:
            return [
                self._transcribe_single(img, separate_footnotes) for img in images
            ]

        out = []
        for ocr_result, layout_result in zip(ocr_results, layout_results):
            out.append(
                self._assemble(ocr_result, layout_result, separate_footnotes)
            )
        return out

    # ------------------------------------------------------------------ #
    # Internals
    # ------------------------------------------------------------------ #
    def _transcribe_single(
        self, image: Image.Image, separate_footnotes: bool
    ) -> PageTranscription:
        try:
            ocr_result = self._run_ocr([image])[0]
            layout_result = (
                self._run_layout([image])[0] if separate_footnotes else None
            )
        except Exception:
            return PageTranscription(ok=False)
        return self._assemble(ocr_result, layout_result, separate_footnotes)

    def _assemble(
        self, ocr_result, layout_result, separate_footnotes: bool
    ) -> PageTranscription:
        lines = _line_boxes(ocr_result)

        if not separate_footnotes:
            ordered = sorted(lines, key=lambda lb: (lb[1][1], lb[1][0]))
            body = "\n".join(text for text, _ in ordered)
            return PageTranscription(body=body, notes="", regions_detected=0)

        regions = _layout_regions(layout_result) if layout_result is not None else []

        if not regions:
            # Layout analysis could not confidently identify a split. Do not
            # guess — everything goes to the body and the page is flagged.
            ordered = sorted(lines, key=lambda lb: (lb[1][1], lb[1][0]))
            body = "\n".join(text for text, _ in ordered)
            return PageTranscription(
                body=body, notes="", regions_detected=0, layout_uncertain=True
            )

        body_lines = []
        note_lines = []
        for text, bbox in lines:
            if _is_footnote_line(bbox, regions):
                note_lines.append((text, bbox))
            else:
                body_lines.append((text, bbox))

        body_lines.sort(key=lambda lb: (lb[1][1], lb[1][0]))
        note_lines.sort(key=lambda lb: (lb[1][1], lb[1][0]))

        return PageTranscription(
            body="\n".join(t for t, _ in body_lines),
            notes="\n".join(t for t, _ in note_lines),
            regions_detected=len(regions),
        )

    def _run_ocr(self, images: list[Image.Image]):
        # Newer Surya auto-detects language and takes det_predictor as a kwarg.
        # Older Surya requires an explicit per-image language list.
        try:
            return self._rec(images, det_predictor=self._det)
        except TypeError:
            langs = [["ar"]] * len(images)
            return self._rec(images, langs, self._det)

    def _run_layout(self, images: list[Image.Image]):
        return self._layout(images)


def _build_predictors():
    """Construct Surya predictors, accommodating the two current API shapes."""
    try:
        # Surya >= 0.13 routes recognition and layout through a shared
        # FoundationPredictor.
        from surya.foundation import FoundationPredictor
        from surya.recognition import RecognitionPredictor
        from surya.detection import DetectionPredictor
        from surya.layout import LayoutPredictor

        foundation = FoundationPredictor()
        rec = RecognitionPredictor(foundation)
        det = DetectionPredictor()
        try:
            layout = LayoutPredictor(foundation)
        except TypeError:
            layout = LayoutPredictor()
        return rec, det, layout
    except ImportError:
        # Older Surya: predictors are standalone.
        from surya.recognition import RecognitionPredictor
        from surya.detection import DetectionPredictor
        from surya.layout import LayoutPredictor

        return RecognitionPredictor(), DetectionPredictor(), LayoutPredictor()

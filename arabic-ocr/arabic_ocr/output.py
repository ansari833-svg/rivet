"""Output writing — per-page text files, full concatenations, review CSV, run log.

All text is written with an explicit UTF-8 encoding. The review CSV uses
utf-8-sig (the BOM is required or Excel on Windows renders the Arabic columns as
mojibake) and newline="" (which prevents doubled blank rows).

Every write goes through _safe_write, which catches PermissionError — typically
"the file is open in Excel" — and offers a retry rather than losing completed
work.
"""

from __future__ import annotations

import csv
from dataclasses import dataclass
from pathlib import Path

# A page marker that is trivially strippable and cannot occur in Arabic text.
PAGE_MARKER = "[[page {n}]]"

_CONFIDENCE_ORDER = {"high": 0, "medium": 1, "low": 2}


@dataclass
class RunRecord:
    page: int
    filename: str
    surya_ok: bool
    regions_detected: int
    layout_uncertain: bool
    verification_status: str
    discrepancy_count: int
    seconds: float


@dataclass
class ReviewRow:
    page: int
    category: str
    context: str
    transcribed: str
    image_shows: str
    confidence: str


def _safe_write(path: Path, write_fn) -> None:
    """Run write_fn(path), retrying on PermissionError after user confirmation."""
    while True:
        try:
            write_fn(path)
            return
        except PermissionError:
            print(
                f"Cannot write {path.name} — the file may be open in another "
                "program. Close it and press Enter to retry.",
                flush=True,
            )
            input()


class OutputWriter:
    def __init__(
        self, output_dir: Path, total_pages: int, separate_footnotes: bool
    ) -> None:
        self.output_dir = output_dir
        self.separate_footnotes = separate_footnotes
        # Zero-pad page numbers to the width of the total page count.
        self.pad = max(2, len(str(total_pages)))

    # ------------------------------------------------------------------ #
    # Per-page text
    # ------------------------------------------------------------------ #
    def page_stub(self, page_number: int) -> str:
        return f"page_{page_number:0{self.pad}d}"

    def write_page_text(
        self, page_number: int, body: str, notes: str, apparatus_detected: bool
    ) -> None:
        stub = self.page_stub(page_number)
        body_path = self.output_dir / f"{stub}_body.txt"
        _safe_write(body_path, lambda p: p.write_text(body, encoding="utf-8"))

        # A notes file is written only when footnote separation is on and
        # apparatus was actually detected on this page.
        if self.separate_footnotes and apparatus_detected:
            notes_path = self.output_dir / f"{stub}_notes.txt"
            _safe_write(notes_path, lambda p: p.write_text(notes, encoding="utf-8"))

    # ------------------------------------------------------------------ #
    # Concatenations
    # ------------------------------------------------------------------ #
    def write_full_body(self, pages: list[tuple[int, str]]) -> None:
        path = self.output_dir / "_full_body.txt"
        _safe_write(path, lambda p: p.write_text(self._concat(pages), encoding="utf-8"))

    def write_full_notes(self, pages: list[tuple[int, str]]) -> None:
        path = self.output_dir / "_full_notes.txt"
        _safe_write(path, lambda p: p.write_text(self._concat(pages), encoding="utf-8"))

    @staticmethod
    def _concat(pages: list[tuple[int, str]]) -> str:
        chunks = []
        for page_number, text in pages:
            chunks.append(PAGE_MARKER.format(n=page_number))
            chunks.append(text)
        return "\n".join(chunks) + "\n"

    # ------------------------------------------------------------------ #
    # Review CSV
    # ------------------------------------------------------------------ #
    def write_review_csv(self, rows: list[ReviewRow]) -> None:
        ordered = sorted(
            rows,
            key=lambda r: (r.page, _CONFIDENCE_ORDER.get(r.confidence, 99)),
        )
        path = self.output_dir / "_review.csv"

        def _write(p: Path) -> None:
            with p.open("w", encoding="utf-8-sig", newline="") as fh:
                writer = csv.writer(fh)
                writer.writerow(
                    [
                        "page",
                        "category",
                        "context",
                        "transcribed",
                        "image_shows",
                        "confidence",
                        "resolved",
                    ]
                )
                for r in ordered:
                    writer.writerow(
                        [
                            r.page,
                            r.category,
                            r.context,
                            r.transcribed,
                            r.image_shows,
                            r.confidence,
                            "",  # resolved — left blank for the user to fill in
                        ]
                    )

        _safe_write(path, _write)

    # ------------------------------------------------------------------ #
    # Run log
    # ------------------------------------------------------------------ #
    def write_run_log(self, records: list[RunRecord]) -> None:
        ordered = sorted(records, key=lambda r: r.page)
        path = self.output_dir / "_run_log.csv"

        def _write(p: Path) -> None:
            with p.open("w", encoding="utf-8-sig", newline="") as fh:
                writer = csv.writer(fh)
                writer.writerow(
                    [
                        "page",
                        "filename",
                        "surya_ok",
                        "regions_detected",
                        "layout_uncertain",
                        "verification_status",
                        "discrepancy_count",
                        "seconds",
                    ]
                )
                for r in ordered:
                    writer.writerow(
                        [
                            r.page,
                            r.filename,
                            r.surya_ok,
                            r.regions_detected,
                            r.layout_uncertain,
                            r.verification_status,
                            r.discrepancy_count,
                            f"{r.seconds:.1f}",
                        ]
                    )

        _safe_write(path, _write)

"""Resume support.

The checkpoint records everything needed to continue an interrupted run without
re-doing completed work: which source folder is being processed, the page list,
which passes have completed for each page, and the transcribed text itself.

It is written with an explicit UTF-8 encoding on every write. Windows defaults
to cp1252 and will raise UnicodeEncodeError the moment Arabic content is written
without this.

Unlike a translation run, an OCR run produces source material a user may want to
re-verify later, so the checkpoint is never auto-deleted on success.
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

CHECKPOINT_NAME = "_checkpoint.json"


class Checkpoint:
    """In-memory view of the run's progress, mirrored to _checkpoint.json."""

    def __init__(self, path: Path, data: dict[str, Any]) -> None:
        self.path = path
        self.data = data

    # ------------------------------------------------------------------ #
    # Construction / loading
    # ------------------------------------------------------------------ #
    @classmethod
    def create(
        cls,
        output_dir: Path,
        source_folder: Path,
        page_numbers: list[int],
        separate_footnotes: bool,
        verify: bool,
    ) -> "Checkpoint":
        path = output_dir / CHECKPOINT_NAME
        data: dict[str, Any] = {
            "source_folder": str(source_folder),
            "page_numbers": page_numbers,
            "separate_footnotes": separate_footnotes,
            "verify": verify,
            # Keyed by page number (as a string, because JSON object keys are
            # always strings). Each value tracks completion of each pass and
            # the transcribed text.
            "pages": {},
        }
        cp = cls(path, data)
        cp.save()
        return cp

    @classmethod
    def load(cls, output_dir: Path) -> "Checkpoint | None":
        path = output_dir / CHECKPOINT_NAME
        if not path.exists():
            return None
        try:
            with path.open("r", encoding="utf-8") as fh:
                data = json.load(fh)
        except (json.JSONDecodeError, OSError):
            return None
        return cls(path, data)

    # ------------------------------------------------------------------ #
    # Persistence
    # ------------------------------------------------------------------ #
    def save(self) -> None:
        # Write to a temp file then replace, so an interruption mid-write does
        # not leave a truncated checkpoint behind.
        tmp = self.path.with_suffix(".json.tmp")
        with tmp.open("w", encoding="utf-8") as fh:
            json.dump(self.data, fh, ensure_ascii=False, indent=2)
        tmp.replace(self.path)

    # ------------------------------------------------------------------ #
    # Per-page state
    # ------------------------------------------------------------------ #
    def _page(self, page_number: int) -> dict[str, Any]:
        key = str(page_number)
        page = self.data["pages"].get(key)
        if page is None:
            page = {
                "surya_done": False,
                "verify_done": False,
                "body": "",
                "notes": "",
                "filename": "",
                "regions_detected": 0,
                "layout_uncertain": False,
                "surya_ok": True,
                "verification_status": "skipped",
                "discrepancy_count": 0,
                "discrepancies": [],
                "seconds_surya": 0.0,
                "seconds_verify": 0.0,
            }
            self.data["pages"][key] = page
        return page

    def record_surya(
        self,
        page_number: int,
        filename: str,
        body: str,
        notes: str,
        regions_detected: int,
        layout_uncertain: bool,
        surya_ok: bool,
        seconds: float,
    ) -> None:
        page = self._page(page_number)
        page["surya_done"] = True
        page["filename"] = filename
        page["body"] = body
        page["notes"] = notes
        page["regions_detected"] = regions_detected
        page["layout_uncertain"] = layout_uncertain
        page["surya_ok"] = surya_ok
        page["seconds_surya"] = seconds
        self.save()

    def record_verify(
        self,
        page_number: int,
        verification_status: str,
        discrepancies: list[dict[str, Any]],
        seconds: float,
    ) -> None:
        page = self._page(page_number)
        page["verify_done"] = True
        page["verification_status"] = verification_status
        page["discrepancies"] = discrepancies
        page["discrepancy_count"] = len(discrepancies)
        page["seconds_verify"] = seconds
        self.save()

    def all_pages(self) -> list[dict[str, Any]]:
        """Return per-page records with the page number attached, in page order."""
        out = []
        for key, page in self.data["pages"].items():
            record = dict(page)
            record["page"] = int(key)
            out.append(record)
        out.sort(key=lambda r: r["page"])
        return out

    def surya_done(self, page_number: int) -> bool:
        page = self.data["pages"].get(str(page_number))
        return bool(page and page.get("surya_done"))

    def verify_done(self, page_number: int) -> bool:
        page = self.data["pages"].get(str(page_number))
        return bool(page and page.get("verify_done"))

    def get_text(self, page_number: int) -> tuple[str, str]:
        page = self.data["pages"].get(str(page_number), {})
        return page.get("body", ""), page.get("notes", "")

    def completed_count(self) -> int:
        """Pages whose required passes are all done."""
        verify = self.data.get("verify", False)
        count = 0
        for page in self.data["pages"].values():
            if not page.get("surya_done"):
                continue
            if verify and not page.get("verify_done"):
                continue
            count += 1
        return count

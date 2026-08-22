"""CLI entry point — prompts, pre-flight, and orchestration.

Two-pass design: Surya (Pass 1) writes the authoritative text; Claude (Pass 2)
only flags discrepancies. A human decides. See the package docstring.

This tool targets Windows only. It reconfigures the console to UTF-8 and enables
ANSI/VT processing before printing anything, because Windows terminals default
to a legacy code page that mangles box-drawing characters, the ✓ symbol, and
Arabic text.
"""

from __future__ import annotations

import argparse
import sys
import time
from pathlib import Path

import colorama
from colorama import Fore, Style
from natsort import natsorted

from . import checkpoint as checkpoint_mod
from .output import OutputWriter, ReviewRow, RunRecord
from .verify import MODEL as VERIFY_MODEL

IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png"}
OUTPUT_DIR_NAME = "ocr_output"


# --------------------------------------------------------------------------- #
# Console helpers
# --------------------------------------------------------------------------- #
def _box(title: str, lines: list[str] | None = None) -> str:
    body = [title] + (lines or [])
    width = max(len(line) for line in body)
    width = max(width, 45)
    top = "┌" + "─" * (width + 2) + "┐"
    bottom = "└" + "─" * (width + 2) + "┘"
    rows = [top, f"│ {title.ljust(width)} │"]
    for line in lines or []:
        rows.append(f"│ {line.ljust(width)} │")
    rows.append(bottom)
    return "\n".join(rows)


def _header() -> None:
    print(
        Fore.CYAN
        + _box("arabic-ocr", ["Arabic page transcription with review"])
        + Style.RESET_ALL
    )


def _prompt(text: str) -> str:
    return input(text).strip()


def _prompt_yes_no(text: str, default_yes: bool = True) -> bool:
    answer = _prompt(text).lower()
    if not answer:
        return default_yes
    return answer.startswith("y")


# --------------------------------------------------------------------------- #
# Input selection
# --------------------------------------------------------------------------- #
def _pick_folder() -> Path | None:
    import tkinter as tk
    from tkinter import filedialog

    root = tk.Tk()
    root.withdraw()
    # Required on Windows — without it the dialog opens behind the terminal and
    # the program looks like it has hung.
    root.attributes("-topmost", True)
    folder = filedialog.askdirectory(title="Select folder of page images")
    root.destroy()
    if not folder:
        return None
    return Path(folder)


def _scan_images(folder: Path) -> list[Path]:
    found = [
        p
        for p in folder.iterdir()
        if p.is_file() and p.suffix.lower() in IMAGE_EXTENSIONS
    ]
    # Natural sort: page_2 before page_10, which plain alphabetical gets wrong.
    return natsorted(found, key=lambda p: p.name)


def _parse_page_selection(raw: str, total: int) -> list[int] | None:
    """Parse "" (all), "5" (single), or "3-12" (inclusive range). None = invalid."""
    raw = raw.strip()
    if not raw:
        return list(range(1, total + 1))
    if "-" in raw:
        parts = raw.split("-", 1)
        try:
            start = int(parts[0].strip())
            end = int(parts[1].strip())
        except ValueError:
            return None
        if start < 1 or end > total or start > end:
            return None
        return list(range(start, end + 1))
    try:
        page = int(raw)
    except ValueError:
        return None
    if page < 1 or page > total:
        return None
    return [page]


# --------------------------------------------------------------------------- #
# Main
# --------------------------------------------------------------------------- #
def main() -> int:
    # Fix the Windows console BEFORE printing anything.
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
    colorama.just_fix_windows_console()

    parser = argparse.ArgumentParser(
        prog="arabic-ocr",
        description="Two-pass Arabic transcription: Surya writes, Claude verifies.",
    )
    parser.add_argument(
        "--offline",
        action="store_true",
        help="Skip the Claude verification pass entirely (Surya only, no network).",
    )
    parser.add_argument(
        "--restart",
        action="store_true",
        help="Ignore any existing checkpoint and start fresh.",
    )
    parser.add_argument(
        "--no-footnotes",
        action="store_true",
        help="Treat the whole page as body text (no footnote separation).",
    )
    args = parser.parse_args()

    _header()

    folder = _pick_folder()
    if folder is None:
        print("No folder selected — exiting.")
        return 0

    images = _scan_images(folder)
    if not images:
        print(f"No .jpg/.jpeg/.png images found in {folder} — exiting.")
        return 0

    total = len(images)
    print(f"\nFolder: {folder}")
    print(
        f"Found: {total} images ({images[0].name} … {images[-1].name})\n"
    )

    # --- Prompts ---------------------------------------------------------- #
    while True:
        selection = _parse_page_selection(
            _prompt("Pages to process (press Enter for all, or e.g. 1-20): "), total
        )
        if selection is not None:
            break
        print(f"  Invalid selection. Enter a page 1-{total}, a range, or Enter for all.")

    if args.no_footnotes:
        separate_footnotes = False
    else:
        separate_footnotes = _prompt_yes_no(
            "Separate footnote apparatus from body text? (Y/n): ", default_yes=True
        )

    if args.offline:
        verify = False
    else:
        verify = _prompt_yes_no(
            "Run Claude verification pass? (Y/n): ", default_yes=True
        )

    output_dir = folder / OUTPUT_DIR_NAME
    output_dir.mkdir(exist_ok=True)

    # --- Resume / checkpoint --------------------------------------------- #
    cp = None
    if not args.restart:
        cp = checkpoint_mod.Checkpoint.load(output_dir)
        if cp is not None:
            done = cp.completed_count()
            print(
                Fore.YELLOW
                + f"\nResuming — {done} of {len(selection)} pages already complete."
                + Style.RESET_ALL
            )
    if cp is None:
        cp = checkpoint_mod.Checkpoint.create(
            output_dir, folder, selection, separate_footnotes, verify
        )

    # --- Pre-flight summary ---------------------------------------------- #
    lo, hi = selection[0], selection[-1]
    page_span = f"{lo}" if lo == hi else f"{lo}–{hi} (inclusive)"
    from .transcribe import autodetect_device

    device = autodetect_device().upper()
    verify_line = (
        f"Verify: Claude ({VERIFY_MODEL})" if verify else "Verify: off (offline)"
    )
    footnote_line = (
        "Footnotes: separated" if separate_footnotes else "Footnotes: body only"
    )
    print(
        "\n"
        + _box(
            "Ready",
            [
                "",
                f"Pages: {page_span}",
                f"Engine: Surya ({device})",
                verify_line,
                footnote_line,
                f"Output: {output_dir}",
            ],
        )
    )
    confirm = _prompt("\nStart? (press Enter to confirm, Q to quit): ").lower()
    if confirm == "q":
        print("Quit — nothing processed.")
        return 0

    # --- Run -------------------------------------------------------------- #
    page_paths = {i + 1: images[i] for i in range(total)}
    _run(
        selection=selection,
        page_paths=page_paths,
        total=total,
        separate_footnotes=separate_footnotes,
        verify=verify,
        cp=cp,
        output_dir=output_dir,
    )
    return 0


def _run(
    *,
    selection: list[int],
    page_paths: dict[int, Path],
    total: int,
    separate_footnotes: bool,
    verify: bool,
    cp,
    output_dir: Path,
) -> None:
    from PIL import Image

    from .transcribe import BATCH_SIZE, PageTranscription, SuryaEngine

    print("\nLoading Surya models (first run downloads weights from Hugging Face)…")
    engine = SuryaEngine()

    verifier = None
    if verify:
        from .verify import ClaudeVerifier

        try:
            verifier = ClaudeVerifier()
        except Exception as exc:  # noqa: BLE001 - surface config errors clearly
            print(
                Fore.RED
                + f"Could not initialize Claude verifier ({exc}). "
                "Continuing with Surya only."
                + Style.RESET_ALL
            )
            verify = False

    writer = OutputWriter(output_dir, total, separate_footnotes)
    num_selected = len(selection)
    processed = 0

    for batch_start in range(0, num_selected, BATCH_SIZE):
        batch = selection[batch_start : batch_start + BATCH_SIZE]

        to_transcribe = [pn for pn in batch if not cp.surya_done(pn)]
        results: dict[int, PageTranscription] = {}
        if to_transcribe:
            batch_images = [
                Image.open(page_paths[pn]).convert("RGB") for pn in to_transcribe
            ]
            t0 = time.time()
            transes = engine.transcribe_batch(batch_images, separate_footnotes)
            per_page_seconds = (time.time() - t0) / max(1, len(to_transcribe))
            for pn, tr in zip(to_transcribe, transes):
                results[pn] = tr
                cp.record_surya(
                    pn,
                    page_paths[pn].name,
                    tr.body,
                    tr.notes,
                    tr.regions_detected,
                    tr.layout_uncertain,
                    tr.ok,
                    per_page_seconds,
                )

        for pn in batch:
            processed += 1
            if pn in results:
                tr = results[pn]
            else:
                body, notes = cp.get_text(pn)
                tr = PageTranscription(body=body, notes=notes)

            apparatus_detected = separate_footnotes and bool(tr.notes.strip())
            writer.write_page_text(pn, tr.body, tr.notes, apparatus_detected)

            flags = 0
            if verify and not cp.verify_done(pn):
                v0 = time.time()
                result = verifier.verify_page(page_paths[pn], tr.body)
                discreps = [
                    {
                        "category": d.category,
                        "context": d.context,
                        "transcribed": d.transcribed,
                        "image_shows": d.image_shows,
                        "confidence": d.confidence,
                    }
                    for d in result.discrepancies
                ]
                cp.record_verify(
                    pn, result.status, discreps, time.time() - v0
                )
                flags = len(discreps)
            elif verify:
                # Already verified in a prior run — reflect its flag count.
                page = cp.data["pages"].get(str(pn), {})
                flags = int(page.get("discrepancy_count", 0))

            body_words = len(tr.body.split())
            notes_words = len(tr.notes.split())
            status = Fore.GREEN + "✓" + Style.RESET_ALL if tr.ok else Fore.RED + "✗" + Style.RESET_ALL
            note_frag = f", notes {notes_words} words" if separate_footnotes else ""
            flag_frag = f", {flags} flags" if verify else ""
            print(
                f"{status} Page {processed}/{num_selected} — "
                f"body {body_words} words{note_frag}{flag_frag}"
            )

    _finalize(cp, writer, separate_footnotes, verify, output_dir)


def _finalize(cp, writer: OutputWriter, separate_footnotes: bool, verify: bool, output_dir: Path) -> None:
    pages = cp.all_pages()

    writer.write_full_body([(p["page"], p.get("body", "")) for p in pages])
    if separate_footnotes:
        writer.write_full_notes([(p["page"], p.get("notes", "")) for p in pages])

    run_records = [
        RunRecord(
            page=p["page"],
            filename=p.get("filename", ""),
            surya_ok=p.get("surya_ok", True),
            regions_detected=p.get("regions_detected", 0),
            layout_uncertain=p.get("layout_uncertain", False),
            verification_status=p.get("verification_status", "skipped"),
            discrepancy_count=p.get("discrepancy_count", 0),
            seconds=p.get("seconds_surya", 0.0) + p.get("seconds_verify", 0.0),
        )
        for p in pages
    ]
    writer.write_run_log(run_records)

    review_rows: list[ReviewRow] = []
    high_conf = 0
    pages_with_flags = set()
    if verify:
        for p in pages:
            for d in p.get("discrepancies", []):
                review_rows.append(
                    ReviewRow(
                        page=p["page"],
                        category=d.get("category", ""),
                        context=d.get("context", ""),
                        transcribed=d.get("transcribed", ""),
                        image_shows=d.get("image_shows", ""),
                        confidence=d.get("confidence", ""),
                    )
                )
                pages_with_flags.add(p["page"])
                if d.get("confidence", "").lower() == "high":
                    high_conf += 1
        writer.write_review_csv(review_rows)

    _summary(
        output_dir,
        num_pages=len(pages),
        separate_footnotes=separate_footnotes,
        verify=verify,
        total_flags=len(review_rows),
        pages_with_flags=len(pages_with_flags),
        high_conf=high_conf,
        cp_path=cp.path,
    )


def _summary(
    output_dir: Path,
    *,
    num_pages: int,
    separate_footnotes: bool,
    verify: bool,
    total_flags: int,
    pages_with_flags: int,
    high_conf: int,
    cp_path: Path,
) -> None:
    print()
    print(Fore.GREEN + f"✓ Done — {num_pages} pages transcribed" + Style.RESET_ALL)
    print()
    print(f"  Body text:      {output_dir / '_full_body.txt'}")
    if separate_footnotes:
        print(f"  Footnotes:      {output_dir / '_full_notes.txt'}")
    if verify:
        print(
            f"  Review queue:   {output_dir / '_review.csv'}  "
            f"({total_flags} flags across {pages_with_flags} pages)"
        )
    print(f"  Run log:        {output_dir / '_run_log.csv'}")
    print(f"  Checkpoint:     {cp_path}  (kept for re-verification)")
    if verify and high_conf:
        print()
        print(
            f"  {high_conf} flags are high-confidence and worth checking first."
        )


if __name__ == "__main__":
    raise SystemExit(main())

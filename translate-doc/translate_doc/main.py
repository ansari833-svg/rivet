"""CLI entry point: startup menu, prompts, and flow orchestration.

Windows-only. See README.md for setup and usage.
"""

from __future__ import annotations

import argparse
import os
import sys
from datetime import date
from pathlib import Path
from typing import Callable, Optional

import colorama

from . import builder, formatter, logger, merger, pdf_export
from .parser import ParsedDocument, parse
from .translator import Checkpoint, Translator, TranslationError

APP_TITLE = "translate-doc"


# ---------------------------------------------------------------------------
# Terminal setup
# ---------------------------------------------------------------------------


def _configure_console() -> None:
    # Windows terminals default to a legacy code page that mangles box-drawing
    # characters, the check mark, and Arabic text. Force UTF-8 first.
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except (AttributeError, ValueError):
        pass
    # Enable ANSI/VT processing on legacy cmd.exe and Windows Terminal.
    colorama.just_fix_windows_console()


# ---------------------------------------------------------------------------
# Generic prompt helpers
# ---------------------------------------------------------------------------


def _retry_on_permission_error(action: Callable[[], None], filename: str) -> None:
    """Run ``action``; on PermissionError, ask the user to close Word and retry."""
    while True:
        try:
            action()
            return
        except PermissionError:
            print(
                f"Cannot write {filename} — the file is open in Word. "
                "Close it and press Enter to retry."
            )
            input()


def _require_api_key() -> str:
    key = os.environ.get("ANTHROPIC_API_KEY")
    if not key:
        print("ANTHROPIC_API_KEY is not set.\n")
        print("Set it in one of these ways, then open a NEW terminal:\n")
        print('  1. From a terminal:  setx ANTHROPIC_API_KEY "sk-ant-..."')
        print(
            "  2. System Properties → Advanced → Environment Variables → New,\n"
            "     under User variables:  name ANTHROPIC_API_KEY, value sk-ant-...\n"
        )
        print("The tool reads the key from the environment; it is never stored.")
        sys.exit(1)
    return key


# ---------------------------------------------------------------------------
# Startup menu
# ---------------------------------------------------------------------------


def _print_main_menu() -> None:
    print(
        "\n"
        "┌─────────────────────────────────────────┐\n"
        "│  translate-doc                          │\n"
        "│                                         │\n"
        "│  What would you like to do?             │\n"
        "│  [T] Translate a document               │\n"
        "│  [M] Merge existing documents           │\n"
        "│  [F] Format a translated document       │\n"
        "│                                         │\n"
        "│  Enter T, M, or F:                      │\n"
        "└─────────────────────────────────────────┘"
    )


def _main_menu_choice() -> str:
    while True:
        _print_main_menu()
        choice = input("> ").strip().upper()
        if choice in ("T", "M", "F"):
            return choice
        print("Please enter T, M, or F")


# ---------------------------------------------------------------------------
# File selection (native Windows picker)
# ---------------------------------------------------------------------------


def _pick_file(title: str = "Select an Arabic .docx file") -> Optional[Path]:
    import tkinter as tk
    from tkinter import filedialog

    root = tk.Tk()
    root.withdraw()
    # Required on Windows or the dialog opens behind the terminal.
    root.attributes("-topmost", True)
    file_path = filedialog.askopenfilename(
        title=title,
        filetypes=[("Word documents", "*.docx")],
    )
    root.destroy()
    if not file_path:
        return None
    return Path(file_path)


# ---------------------------------------------------------------------------
# Page selection
# ---------------------------------------------------------------------------


def _prompt_page_selection(
    filename: str, page_count: int
) -> tuple[int, int, str]:
    """Return (start_page, end_page, range_suffix)."""
    print(
        "\n"
        "┌─────────────────────────────────────────────┐\n"
        f"│  Document: {filename:<33}│\n"
        f"│  Detected length: ~{page_count} pages{' ' * max(0, 20 - len(str(page_count)))}│\n"
        "│                                             │\n"
        "│  How much would you like to translate?      │\n"
        "│  [A] All pages                              │\n"
        "│  [R] A page range                           │\n"
        "│                                             │\n"
        "│  Enter A or R:                              │\n"
        "└─────────────────────────────────────────────┘"
    )
    while True:
        choice = input("> ").strip().upper()
        if choice == "A":
            return 1, page_count, ""
        if choice == "R":
            return _prompt_range(page_count)
        print("Please enter A or R")


def _prompt_range(page_count: int) -> tuple[int, int, str]:
    while True:
        raw = input("Enter page range (e.g. 1-10): ").strip()
        parts = raw.replace(" ", "").split("-")
        if len(parts) != 2 or not all(p.isdigit() for p in parts):
            print("Invalid range. Use the form start-end, e.g. 1-10.")
            continue
        start, end = int(parts[0]), int(parts[1])
        if start < 1:
            print("Start page must be >= 1.")
            continue
        if end > page_count:
            print(f"End page must be <= {page_count} (the detected page count).")
            continue
        if start > end:
            print("Start page must be <= end page.")
            continue
        return start, end, f"_p{start}-{end}"


# ---------------------------------------------------------------------------
# Pre-flight
# ---------------------------------------------------------------------------


def _preflight_confirm(
    filename: str,
    start: int,
    end: int,
    page_count: int,
    output_dir: Path,
) -> bool:
    pages_line = (
        "All pages"
        if start == 1 and end == page_count
        else f"{start}–{end} of ~{page_count}"
    )
    print(
        "\n"
        "┌─────────────────────────────────────────────┐\n"
        "│  Ready                                      │\n"
        "│                                             │\n"
        f"│  Document: {filename}\n"
        f"│  Pages: {pages_line}\n"
        "│  Model: claude-sonnet-5\n"
        f"│  Outputs: {output_dir}\n"
        "└─────────────────────────────────────────────┘"
    )
    answer = input("Translate? (press Enter to confirm, Q to quit): ").strip().upper()
    return answer != "Q"


# ---------------------------------------------------------------------------
# Translate flow
# ---------------------------------------------------------------------------


def run_translate_flow(input_path: Path, restart: bool) -> None:
    if input_path.suffix.lower() != ".docx":
        print(f"Not a .docx file: {input_path}")
        sys.exit(1)
    if not input_path.exists():
        print(f"File does not exist: {input_path}")
        sys.exit(1)

    api_key = _require_api_key()
    input_path = input_path.resolve()
    print(f"Selected file: {input_path}")

    # Handle an existing checkpoint before parsing so --restart is honoured.
    checkpoint_path = Checkpoint.checkpoint_path_for(input_path)
    if restart and checkpoint_path.exists():
        checkpoint_path.unlink()
        print("Ignoring previous checkpoint (--restart).")

    print("Reading document...")
    parsed: ParsedDocument = parse(input_path)
    page_count = parsed.page_count

    start, end, range_suffix = _prompt_page_selection(input_path.name, page_count)
    selected_ids = set(parsed.block_ids_for_pages(start, end))
    if not selected_ids:
        print("No content found in the selected page range — exiting.")
        return

    outputs = builder.output_paths(input_path, range_suffix)

    if not _preflight_confirm(
        input_path.name, start, end, page_count, input_path.parent
    ):
        print("Cancelled.")
        return

    # Resume from checkpoint if present.
    existing = None if restart else Checkpoint.load(input_path)
    if existing is not None:
        summary = existing.remaining_summary()
        if summary:
            print(f"Resuming previous translation ({summary})...")
        checkpoint = existing
    else:
        checkpoint = Checkpoint.start(input_path)

    translator = Translator(client=_build_client(api_key))

    def _progress(current: int, total: int) -> None:
        print(f"Translating... chunk {current} of {total}")

    try:
        result = translator.translate(
            sorted(selected_ids), parsed.blocks, checkpoint, progress=_progress
        )
    except TranslationError as exc:
        print(f"\nTranslation failed: {exc}")
        print(
            "Your progress has been checkpointed. Re-run the tool to resume, "
            "or lower MAX_PARAGRAPHS_PER_CHUNK in translator.py if responses "
            "are being truncated."
        )
        sys.exit(1)

    for bid in result.missing_ids:
        print(f"Warning: block {bid} missing from response — keeping original.")

    translations = checkpoint.all_translations()

    # Build outputs, retrying on the common "file open in Word" failure.
    print("Building translated document...")
    _retry_on_permission_error(
        lambda: builder.build_translated_docx(
            parsed, translations, selected_ids, outputs["translated"]
        ),
        outputs["translated"].name,
    )

    print("Building bilingual parallel-text document...")
    _retry_on_permission_error(
        lambda: builder.build_bilingual_docx(
            parsed,
            translations,
            selected_ids,
            outputs["bilingual_docx"],
            title=input_path.name,
            translation_date=date.today().isoformat(),
        ),
        outputs["bilingual_docx"].name,
    )

    print("Writing translation log...")
    flagged = 0

    def _write_log() -> None:
        nonlocal flagged
        flagged = logger.write_log(
            parsed, translations, selected_ids, outputs["log"]
        )

    _retry_on_permission_error(_write_log, outputs["log"].name)

    print("Exporting PDF...")
    pdf_ok = pdf_export.export_pdf(
        outputs["bilingual_docx"], outputs["bilingual_pdf"]
    )

    # Checkpoint no longer needed once all outputs are written.
    checkpoint.delete()

    _print_summary(input_path, outputs, selected_ids, flagged, pdf_ok, range_suffix)


def _build_client(api_key: str):
    from anthropic import Anthropic

    return Anthropic(api_key=api_key)


def _print_summary(
    input_path: Path,
    outputs: dict,
    selected_ids: set,
    flagged: int,
    pdf_ok: bool,
    range_suffix: str,
) -> None:
    name = input_path.stem + range_suffix
    translated_count = len(selected_ids)
    print("\n" + "=" * 50)
    print("Done.")
    print(f"  Paragraphs/blocks translated: {translated_count}")
    print(f"  Difficult sections flagged:   {flagged}")
    print("=" * 50)
    print(f"✓ {name}_translated.docx — English translation (editable)")
    print(f"✓ {name}_bilingual.docx — Arabic/English parallel text (editable)")
    if pdf_ok:
        print(f"✓ {name}_bilingual.pdf — Parallel text PDF (print-ready)")
    else:
        print(f"  {name}_bilingual.pdf — not created (see message above)")
    print(f"✓ {name}_translation_log.csv — Review log ({flagged} items flagged)")
    print(f"\nAll outputs are in: {input_path.parent}")


# ---------------------------------------------------------------------------
# Merge flow
# ---------------------------------------------------------------------------


def _parse_index_list(raw: str, count: int) -> Optional[list[int]]:
    parts = [p.strip() for p in raw.split(",") if p.strip()]
    if not parts or not all(p.isdigit() for p in parts):
        return None
    indices = [int(p) for p in parts]
    if any(i < 1 or i > count for i in indices):
        return None
    return indices


def run_merge_flow() -> None:
    raw_folder = input(
        "Enter folder path to scan (or press Enter for current directory): "
    ).strip()
    folder = Path(raw_folder) if raw_folder else Path.cwd()
    folder = folder.resolve()
    if not folder.is_dir():
        print(f"Not a folder: {folder}")
        return

    files = merger.scan_folder(folder)
    if not files:
        print(f"No .docx or .txt files found in {folder}")
        return

    print(f"\nFiles found in {folder}:")
    for i, path in enumerate(files, start=1):
        print(f"  [{i}] {path.name}")

    # Initial selection.
    while True:
        raw = input(
            "\nEnter file numbers in order, separated by commas (e.g. 1,2,3): "
        )
        indices = _parse_index_list(raw, len(files))
        if indices is None:
            print("Invalid selection — use numbers from the list above.")
            continue
        if len(indices) < 2:
            print("Select at least 2 files to merge.")
            continue
        break

    selected = [files[i - 1] for i in indices]

    # Confirm / correct order.
    while True:
        print("\nMerge order:")
        for i, path in enumerate(selected, start=1):
            print(f"  [{i}] {path.name}")
        raw = input(
            "Press Enter to use this order, or type a new order "
            "(e.g. 1,2,4,3): "
        ).strip()
        if not raw:
            break
        indices = _parse_index_list(raw, len(files))
        if indices is None or len(indices) < 2:
            print("Invalid order — use file numbers from the original list.")
            continue
        selected = [files[i - 1] for i in indices]

    # Warn on mixing translated and bilingual files.
    if merger.has_translated_and_bilingual(selected):
        answer = input(
            "Warning: you are merging translated and bilingual files — "
            "are you sure? (y/n) "
        ).strip().lower()
        if answer != "y":
            print("Cancelled.")
            return

    # Warn on mixing .txt and .docx.
    if merger.is_mixed_types(selected):
        answer = input(
            "Warning: mixing .txt and .docx files — output will be saved as "
            ".rtf. Continue? (y/n) "
        ).strip().lower()
        if answer != "y":
            print("Cancelled.")
            return

    fmt = merger.determine_output_format(selected)
    default_name = f"merged.{'docx' if fmt == 'docx' else 'rtf'}"
    out_name = input(
        f"Output filename (press Enter for '{default_name}'): "
    ).strip() or default_name
    output_path = folder / out_name

    print("Merging...")

    def _do_merge() -> None:
        merger.merge(selected, output_path)

    _retry_on_permission_error(_do_merge, output_path.name)

    print(f"\n✓ Merged {len(selected)} files → {output_path.name}")
    print(f"Saved to: {output_path}")


# ---------------------------------------------------------------------------
# Format flow
# ---------------------------------------------------------------------------


def run_format_flow(input_path: Path, restart: bool) -> None:
    if input_path.suffix.lower() != ".docx":
        print(f"Not a .docx file: {input_path}")
        sys.exit(1)
    if not input_path.exists():
        print(f"File does not exist: {input_path}")
        sys.exit(1)

    # Pandoc is required; fail fast (before any API spend) so the checkpoint,
    # if one exists from a prior run, is left untouched.
    if not formatter.pandoc_available():
        print(
            "Format mode needs Pandoc. Install it with: "
            "winget install --id JohnMacFarlane.Pandoc"
        )
        sys.exit(1)

    input_path = input_path.resolve()
    print(f"Selected file: {input_path}")

    # Intelligent paragraphing needs the API key; without it we still format.
    client = None
    if os.environ.get("ANTHROPIC_API_KEY"):
        client = _build_client(os.environ["ANTHROPIC_API_KEY"])
    else:
        print(
            "ANTHROPIC_API_KEY is not set — skipping intelligent paragraphing.\n"
            "The document will still be formatted using page-based paragraphs."
        )

    def _progress(current: int, total: int) -> None:
        print(f"Paragraphing... chunk {current} of {total}")

    print("Reading and cleaning the document...")

    result_holder: dict[str, formatter.FormatResult] = {}

    def _do_format() -> None:
        result_holder["result"] = formatter.format_document(
            input_path, client=client, restart=restart, progress=_progress
        )

    try:
        _retry_on_permission_error(_do_format, f"{input_path.stem}_formatted.docx")
    except FileNotFoundError:
        # build_reference_doc / run_pandoc invoke the pandoc binary.
        print(
            "Format mode needs Pandoc. Install it with: "
            "winget install --id JohnMacFarlane.Pandoc"
        )
        sys.exit(1)

    result = result_holder["result"]
    _print_format_summary(input_path, result)


def _print_format_summary(input_path: Path, result: "formatter.FormatResult") -> None:
    skipped = (
        ", ".join(str(n) for n in result.skipped_pages)
        if result.skipped_pages
        else "none"
    )
    print("\n" + "=" * 50)
    print("Done.")
    print(f"  Input:                 {input_path.name}")
    print(f"  Output:                {result.output_path.name}")
    print(f"  Pages formatted:       {result.pages_formatted}")
    print(f"  Footnotes created:     {result.footnotes_created}")
    print(f"  Asterisk flags:        {result.asterisks}")
    print(f"  Cover pages skipped:   {skipped}")
    print(
        "  Intelligent paragraphing: "
        + ("ran" if result.paragraphing_ran else "skipped")
    )
    print("=" * 50)
    print(f"✓ {result.output_path.name} — clean, footnoted Word document")
    print(f"\nSaved to: {result.output_path}")


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------


def main() -> None:
    _configure_console()

    parser_ = argparse.ArgumentParser(
        prog="translate-doc",
        description="Translate Arabic .docx files into English via Claude.",
    )
    parser_.add_argument("file", nargs="?", help="Path to an Arabic .docx file.")
    parser_.add_argument(
        "--restart",
        action="store_true",
        help="Ignore any existing checkpoint and start fresh.",
    )
    parser_.add_argument(
        "--merge",
        action="store_true",
        help="Go straight to merge mode.",
    )
    parser_.add_argument(
        "--format",
        dest="format_mode",
        action="store_true",
        help="Format a translated document into a clean, footnoted .docx.",
    )
    args = parser_.parse_args()

    if args.merge:
        run_merge_flow()
        return

    if args.format_mode:
        _dispatch_format(args)
        return

    # Direct file shortcut: skip the menu.
    if args.file:
        run_translate_flow(Path(args.file), restart=args.restart)
        return

    choice = _main_menu_choice()
    if choice == "M":
        run_merge_flow()
        return
    if choice == "F":
        _dispatch_format(args)
        return

    # Translate flow with the native file picker.
    selected = _pick_file()
    if selected is None:
        print("No file selected — exiting.")
        return
    run_translate_flow(selected, restart=args.restart)


def _dispatch_format(args) -> None:
    if args.file:
        run_format_flow(Path(args.file), restart=args.restart)
        return
    selected = _pick_file(title="Select a translated .docx file")
    if selected is None:
        print("No file selected — exiting.")
        return
    run_format_flow(selected, restart=args.restart)


if __name__ == "__main__":
    main()

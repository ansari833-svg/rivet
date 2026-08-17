"""CLI entry point for pdf-screenshots.

A simple prompt-based terminal UI (no frameworks) that converts each page of a
PDF into a JPEG image. Windows-only.
"""

import sys
from pathlib import Path

import colorama

from pdf_screenshots.render import get_page_count, render_pdf_to_jpegs

DEFAULT_WIDTH = 900
DEFAULT_HEIGHT = 1300
DEFAULT_QUALITY = 85


def _box(lines):
    """Render a list of text lines inside a single ASCII/box-drawing frame."""
    inner_width = max(len(line) for line in lines)
    top = "┌" + "─" * (inner_width + 2) + "┐"
    bottom = "└" + "─" * (inner_width + 2) + "┘"
    body = ["│ " + line.ljust(inner_width) + " │" for line in lines]
    return "\n".join([top, *body, bottom])


def _print_header():
    print(
        _box(
            [
                "pdf-screenshots",
                "Convert PDF pages to JPEG images",
            ]
        )
    )
    print()


def _select_pdf():
    """Open the native Windows file picker and return the chosen Path, or None."""
    import tkinter as tk
    from tkinter import filedialog

    root = tk.Tk()
    root.withdraw()
    # Required on Windows, otherwise the dialog often opens behind the terminal
    # and the program appears to hang.
    root.attributes("-topmost", True)
    file_path = filedialog.askopenfilename(
        title="Select a PDF file",
        filetypes=[("PDF files", "*.pdf")],
    )
    root.destroy()

    if not file_path:
        return None
    return Path(file_path)


def _prompt_output_dir(pdf_path: Path) -> Path:
    """Prompt for an output folder; Enter accepts the auto-resolved default."""
    default = pdf_path.parent / pdf_path.stem
    raw = input("Output folder (press Enter for auto): ").strip()
    if not raw:
        return default
    return Path(raw)


def _prompt_int(prompt: str, default: int, minimum=None, maximum=None) -> int:
    """Prompt for an integer, re-prompting on invalid input."""
    while True:
        raw = input(prompt).strip()
        if not raw:
            return default
        try:
            value = int(raw)
        except ValueError:
            print("  Please enter a whole number.")
            continue
        if minimum is not None and value < minimum:
            print(f"  Please enter a number >= {minimum}.")
            continue
        if maximum is not None and value > maximum:
            print(f"  Please enter a number <= {maximum}.")
            continue
        return value


def _confirm_overwrite(output_dir: Path) -> bool:
    """If prior page_*.jpg files exist, ask before overwriting. Returns proceed."""
    existing = list(output_dir.glob("page_*.jpg"))
    if not existing:
        return True
    while True:
        raw = input(
            f"Output folder already contains {len(existing)} images — "
            f"overwrite? (y/n) "
        ).strip().lower()
        if raw in ("y", "yes"):
            return True
        if raw in ("n", "no"):
            return False
        print("  Please answer y or n.")


def main():
    # Windows terminals default to a legacy code page that mangles box-drawing
    # characters and the check mark. Do this before printing anything.
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
    # Enable ANSI/VT processing in both Windows Terminal and legacy cmd.exe.
    colorama.just_fix_windows_console()

    _print_header()

    pdf_path = _select_pdf()
    if pdf_path is None:
        print("No file selected — exiting.")
        return

    print(f"Selected: {pdf_path}")
    print()

    # Validate the selected file before going further.
    if not pdf_path.exists():
        print(f"File does not exist: {pdf_path}")
        return
    if pdf_path.suffix.lower() != ".pdf":
        print(f"Not a PDF file: {pdf_path}")
        return

    output_dir = _prompt_output_dir(pdf_path)
    width = _prompt_int(
        f"Width in pixels (press Enter for {DEFAULT_WIDTH}): ",
        DEFAULT_WIDTH,
        minimum=1,
    )
    height = _prompt_int(
        f"Height in pixels (press Enter for {DEFAULT_HEIGHT}): ",
        DEFAULT_HEIGHT,
        minimum=1,
    )
    quality = _prompt_int(
        f"JPEG quality 1-95 (press Enter for {DEFAULT_QUALITY}): ",
        DEFAULT_QUALITY,
        minimum=1,
        maximum=95,
    )

    try:
        n_pages = get_page_count(pdf_path)
    except Exception as exc:  # noqa: BLE001 - surface a friendly message
        print(f"Could not read the PDF: {exc}")
        return

    print()
    print(
        _box(
            [
                "Ready",
                "",
                f"PDF:     {pdf_path.name}",
                f"Pages:   {n_pages}",
                f"Output:  {output_dir}",
                f"Size:    {width} × {height} px",
                f"Quality: {quality}",
            ]
        )
    )
    print()

    answer = input("Convert? (press Enter to confirm, Q to quit): ").strip().lower()
    if answer in ("q", "quit"):
        print("Cancelled.")
        return

    output_dir.mkdir(parents=True, exist_ok=True)
    if not _confirm_overwrite(output_dir):
        print("Cancelled.")
        return

    print()
    print(f"Converting {n_pages} pages...")

    def _progress(page_number, total, filename):
        print(f"  ✓ Saved page {page_number}/{total} → {filename}")

    render_pdf_to_jpegs(
        pdf_path=pdf_path,
        output_dir=output_dir,
        width=width,
        height=height,
        quality=quality,
        progress_callback=_progress,
    )

    print()
    print(f"✓ Done — {n_pages} images saved to {output_dir}")


if __name__ == "__main__":
    main()

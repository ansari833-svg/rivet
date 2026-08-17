# pdf-screenshots

A small **Windows** command-line tool that converts each page of a PDF into a
JPEG image at a consistent, screenshot-like size (around 900×1300 px by
default). It uses a simple prompt-based terminal UI with ASCII box styling — no
frameworks, just clean terminal output.

## Why pypdfium2 (and not pdf2image / PyMuPDF)?

- **pdf2image** depends on [poppler](https://poppler.freedesktop.org/), which
  on Windows is *not* pip-installable — you have to manually download binaries,
  unzip them to a fixed location, and pass an explicit `poppler_path`.
- **pypdfium2** ships as a self-contained pip wheel for Windows with no
  external dependencies, and is permissively licensed (Apache 2.0 / BSD-3),
  which matters if this ends up on a work machine.
- **PyMuPDF** is avoided for licensing reasons — it is AGPL unless separately
  licensed.

## Requirements

- Windows
- Python (from [python.org](https://www.python.org/downloads/windows/) —
  the standard installer bundles `tkinter`, which the file picker uses)

## Setup

From the `pdf-screenshots` folder, run:

```powershell
powershell -ExecutionPolicy Bypass -File setup.ps1
```

> The `-ExecutionPolicy Bypass` flag is required — the default execution policy
> will otherwise block the script.

`setup.ps1` will:

1. Create a virtual environment at `.venv`.
2. Install dependencies into it via `.venv\Scripts\pip install -e .`.
3. Create a launcher batch file at
   `%LOCALAPPDATA%\Programs\pdf-screenshots\pdf-screenshots.bat` that invokes
   the venv's Python directly, so the tool works without activating the venv
   first.
4. Add that folder to your **user** `PATH`.

After setup, **open a new terminal window** for the `PATH` change to take
effect.

## Usage

Once set up, just run:

```powershell
pdf-screenshots
```

Then:

1. A native Windows file picker opens — choose a PDF.
2. Answer the prompts (press Enter to accept each default):
   - **Output folder** — defaults to a folder named after the PDF, next to it.
   - **Width** in pixels — default `900`.
   - **Height** in pixels — default `1300`.
   - **JPEG quality** (1–95) — default `85`.
3. Review the pre-flight summary and press Enter to confirm (or `Q` to quit).
4. Each page is rendered and saved as `page_01.jpg`, `page_02.jpg`, … (zero
   padded based on the total page count).

Example session:

```
┌─────────────────────────────────────────┐
│ pdf-screenshots                          │
│ Convert PDF pages to JPEG images         │
└─────────────────────────────────────────┘

Selected: C:\Users\hansari\Documents\document.pdf

Output folder (press Enter for auto):
Width in pixels (press Enter for 900):
Height in pixels (press Enter for 1300):
JPEG quality 1-95 (press Enter for 85):

┌─────────────────────────────────────────┐
│ Ready                                    │
│                                          │
│ PDF:     document.pdf                    │
│ Pages:   12                              │
│ Output:  C:\Users\hansari\...\document   │
│ Size:    900 × 1300 px                   │
│ Quality: 85                              │
└─────────────────────────────────────────┘

Convert? (press Enter to confirm, Q to quit):

Converting 12 pages...
  ✓ Saved page 1/12 → page_01.jpg
  ✓ Saved page 2/12 → page_02.jpg
  ...

✓ Done — 12 images saved to C:\Users\hansari\Documents\document\
```

## Sizing & aspect ratio

By default the output is stretched to exactly the requested width × height.
This slightly distorts a standard Letter/A4 page, but produces a consistent,
screenshot-like frame for every page. To keep the native aspect ratio instead
(fitting each page inside the frame with white padding), set `PRESERVE_ASPECT =
True` at the top of `pdf_screenshots/render.py`.

Pages are always rendered at a scale large enough to cover the requested output
before being downsampled with Pillow's LANCZOS filter, so text stays sharp.

## Running without the launcher

If you'd rather not add anything to `PATH`, you can run the tool directly from
the venv:

```powershell
.venv\Scripts\python -m pdf_screenshots.main
```

## Project structure

```
pdf-screenshots\
├── setup.ps1
├── pyproject.toml
├── README.md
└── pdf_screenshots\
    ├── __init__.py
    ├── main.py        # CLI entry point, prompts, pre-flight
    └── render.py      # pypdfium2 rendering and JPEG output
```

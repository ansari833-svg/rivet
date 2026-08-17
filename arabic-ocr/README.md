# arabic-ocr

A reusable Windows CLI tool that transcribes Arabic text from a folder of
scanned page images (JPEG/PNG). It is built for classical Arabic printed books
(fiqh, hadith, tafsir, balāgha) in modern critical editions, where transcription
fidelity matters far more than speed. The output is exact plain-text UTF-8, ready
to feed into a separate translation tool.

> **Windows only.** This tool assumes a Windows environment and does not include
> macOS or Linux compatibility code.

## How it works — two passes, and why

Every page goes through two passes that are then reconciled:

- **Pass 1 — Surya (local, deterministic).** Surya transcribes what is
  physically on the page. **This is the authoritative text. It is never
  overwritten automatically.**
- **Pass 2 — Claude (verification only).** Claude receives the page image *and*
  Surya's transcription and reports only where they disagree. Claude does **not**
  produce its own transcription and does **not** rewrite Surya's output.

This split is the whole point of the design. A language model reading classical
Arabic will silently normalize archaic spellings, regularize unusual readings
toward common ones, and complete damaged words from memory rather than from the
page. Those errors are invisible in the output. Surya's errors, by contrast,
look like garbage and are therefore catchable. **So Surya writes, Claude flags,
and a human decides.** The tool never uses Claude's reading as the text when the
two disagree — disagreements are logged to `_review.csv` for a human to resolve.

## Installation

From the `arabic-ocr` folder, run:

```
powershell -ExecutionPolicy Bypass -File setup.ps1
```

This will:

1. Create a virtual environment at `.venv`.
2. Install the **CPU-only** PyTorch wheel explicitly (so surya-ocr does not pull
   several gigabytes of CUDA libraries on a machine without an NVIDIA GPU),
   then install the tool with `pip install -e .`.
3. If an NVIDIA GPU is detected, print a note that the CUDA build of PyTorch
   would give a 5–10x speedup, along with the install command. It is **not**
   installed automatically.
4. Create a launcher at `%LOCALAPPDATA%\Programs\arabic-ocr\arabic-ocr.bat` and
   add that folder to your **user** PATH.

**Open a new terminal window after setup** for the PATH change to take effect.
Then run:

```
arabic-ocr
```

## Configuration

Set your Anthropic API key (required for the Claude verification pass):

```
setx ANTHROPIC_API_KEY "sk-ant-..."
```

`setx` writes to the persistent user environment, so **open a new terminal
window** afterward for it to be visible.

If you do not want to use Claude at all, run with `--offline` (see Flags) — no
key or network is needed for Surya-only transcription.

## First run and offline use

- On its **first run**, Surya downloads its model weights from Hugging Face and
  therefore needs network access once. **Every run after that is fully offline**
  for the transcription pass.
- The `--offline` flag skips the Claude verification pass entirely: Surya output
  only, no network, and no `_review.csv` is produced.

## Usage

Run `arabic-ocr` and follow the prompts:

1. A native Windows folder picker opens (it is forced to the foreground). Select
   the folder of page images. Cancelling exits cleanly.
2. The tool scans for `.jpg`, `.jpeg`, and `.png` files, sorts them with a
   natural sort (so `image-2` comes before `image-10`), and shows the count plus
   the first and last filenames so you can confirm the order.
3. You are asked which pages to process, whether to separate footnotes, and
   whether to run the Claude verification pass, followed by a pre-flight summary
   you confirm before anything runs.

### Page selection

At the "Pages to process" prompt:

- Press **Enter** for all pages.
- Enter a single page number, e.g. `5`, to process page 5 only.
- Enter a range, e.g. `3-12`, to process pages 3 through 12 **inclusive**.

### Footnote separation

Modern critical editions (muḥaqqaq editions) often devote a third to half of the
page to the editor's footnote apparatus, separated from the body by a horizontal
rule. Naive OCR interleaves the two and the result is unusable. With footnote
separation on (the default), the tool uses Surya's layout regions to write body
text and footnote apparatus as **separate streams**. If layout analysis cannot
confidently identify a split, the tool does not guess — it writes everything to
the body stream and flags the page as `layout_uncertain` in the run log.

## Flags

- `--offline` — skip the Claude verification pass entirely (Surya only, no
  network, no `_review.csv`).
- `--restart` — ignore any existing checkpoint and start fresh.
- `--no-footnotes` — treat the whole page as body text.

## Resuming

Progress is checkpointed to `_checkpoint.json` after every page. If a run is
interrupted, the next run detects the checkpoint and offers to resume
(`Resuming — 34 of 58 pages already complete.`). Use `--restart` to ignore it.

The checkpoint is **not** auto-deleted on success. Unlike a translation run, an
OCR run is source material you may want to re-verify, so it is kept and mentioned
in the final summary.

## Output

A folder named `ocr_output` is created inside the source folder, containing:

- **Per-page text files** — `page_01_body.txt` and (when footnote separation is
  on and apparatus was detected) `page_01_notes.txt`. Page numbers are
  zero-padded to match the total page count. Plain UTF-8 text.
- **`_full_body.txt`** — all body text concatenated in page order, with a
  `[[page 12]]` marker line between pages (trivially strippable; cannot occur in
  Arabic text).
- **`_full_notes.txt`** — the same for footnote apparatus, when separation is on.
- **`_review.csv`** — the discrepancy log from the Claude pass. Columns: `page`,
  `category`, `context`, `transcribed`, `image_shows`, `confidence`, `resolved`.
  Written with a UTF-8 BOM so Excel renders the Arabic columns correctly. Sorted
  by page, then confidence (high first). The `resolved` column is left blank for
  you to fill in.
- **`_run_log.csv`** — one row per page: `page`, `filename`, `surya_ok`,
  `regions_detected`, `layout_uncertain`, `verification_status`,
  `discrepancy_count`, `seconds`. This is how you find the pages that need
  attention without opening every file.

## Text fidelity guarantees

- **Tashkīl** (vocalization marks) is preserved exactly as recognized. It is
  never stripped, normalized, or added.
- **No reshaping and no bidi algorithm** is applied to the stored text.
  Reshaping is for rendering glyphs to a display; applying it to stored text
  substitutes presentation forms for the real characters and corrupts the file
  for every downstream consumer.
- **No RLM/LRM directional marks** are inserted.
- All text files are written with UTF-8 encoding.

Note: the console deliberately prints only short excerpts of Arabic, never full
pages. Terminal bidirectional rendering is unreliable and would make correct
text look wrong.

## Project layout

```
arabic-ocr\
├── setup.ps1
├── pyproject.toml
├── README.md
└── arabic_ocr\
    ├── __init__.py
    ├── main.py          # CLI entry point, prompts, pre-flight, orchestration
    ├── transcribe.py    # Surya OCR and layout/reading-order handling
    ├── verify.py        # Claude verification pass and JSON parsing
    ├── checkpoint.py    # Resume support
    └── output.py        # Text files, review CSV, run log
```

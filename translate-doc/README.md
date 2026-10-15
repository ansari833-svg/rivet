# translate-doc

A Windows command-line tool that translates Arabic `.docx` files into English
using the Anthropic Claude API.

For each document it produces:

1. **`{name}_translated.docx`** — an English translation with formatting
   preserved as closely as possible.
2. **`{name}_bilingual.docx`** and **`{name}_bilingual.pdf`** — an
   Arabic/English parallel-text document with footnotes for difficult passages.
3. **`{name}_translation_log.csv`** — a review log flagging every difficult or
   untranslated section.

It also has a **merge mode** for combining several `.docx` or `.txt` files into
one document, and a **format mode** that turns a raw `_translated.docx` into a
clean, properly footnoted Word document (`{name}_formatted.docx`).

> This tool is Windows-only.

---

## Requirements

- Windows
- Python 3.10+ (install from [python.org](https://www.python.org/downloads/);
  the standard installer bundles `tkinter`, which the file picker needs)
- An Anthropic API key
- Optional, for PDF export: Microsoft Word **or**
  [LibreOffice](https://www.libreoffice.org/)
- Required for **format mode**: [Pandoc](https://pandoc.org/) (install with
  `winget install --id JohnMacFarlane.Pandoc`)

---

## Installation

From the `translate-doc` folder, run:

```powershell
powershell -ExecutionPolicy Bypass -File setup.ps1
```

The `-ExecutionPolicy Bypass` is required — Windows' default execution policy
will otherwise block the script from running.

`setup.ps1` will:

- create a virtual environment at `.venv`
- install the tool and its dependencies into it
- create a launcher at `%LOCALAPPDATA%\Programs\translate-doc\translate-doc.bat`
- add that folder to your **user** PATH

**Open a new terminal window after setup** so the PATH change takes effect.
Then you can run `translate-doc` from anywhere.

---

## Setting your API key

The tool reads your key from the `ANTHROPIC_API_KEY` environment variable and
never stores it. Set it in either of these ways:

**Option 1 — from a terminal** (takes effect in *new* terminals only):

```cmd
setx ANTHROPIC_API_KEY "sk-ant-..."
```

**Option 2 — via the GUI:**

System Properties → Advanced → Environment Variables → **New** under
*User variables*, with name `ANTHROPIC_API_KEY` and value `sk-ant-...`.

If the key is not set, the tool prints these instructions and exits cleanly
(no traceback).

---

## Usage

### Interactive

Just run:

```cmd
translate-doc
```

You'll see a menu:

```
┌─────────────────────────────────────────┐
│  translate-doc                          │
│                                         │
│  [T] Translate a document               │
│  [M] Merge existing documents           │
│  [F] Format a translated document       │
│                                         │
│  Enter T, M, or F:                      │
└─────────────────────────────────────────┘
```

- **[T]** opens a native file picker, asks how many pages to translate, shows a
  pre-flight summary, then translates.
- **[M]** scans a folder and merges the files you choose.
- **[F]** cleans up and reformats a `_translated.docx` (see
  [Format mode](#format-mode) below).

### Shortcuts

Translate a specific file directly (skips the menu):

```cmd
translate-doc "C:\Users\you\Documents\contract.docx"
```

Ignore a saved checkpoint and start fresh:

```cmd
translate-doc "contract.docx" --restart
```

Go straight to merge mode:

```cmd
translate-doc --merge
```

Go straight to format mode (with or without a file argument):

```cmd
translate-doc --format
translate-doc --format "contract_translated.docx"
```

---

## Page selection

Before translating, the tool estimates the document's page count and asks
whether to translate all pages or a range (e.g. `1-10`). The selected range is
reflected in the output filenames, e.g. `contract_p1-10_translated.docx`.

## Resume / checkpoints

Translation happens in chunks. After each chunk a checkpoint file
(`{name}_checkpoint.json`) is written next to the input. If the run is
interrupted, just run the tool again on the same file — it resumes from where it
left off. Use `--restart` to ignore the checkpoint. The checkpoint is deleted
automatically once all outputs are written.

## Difficult passages

Claude flags ambiguous, culturally specific, or hard-to-translate phrases. In
the English-only document these are replaced with a highlighted marker; in the
bilingual document they become numbered footnotes; and every one of them is
recorded in the CSV log along with any passages that were missing from the API
response.

## PDF export

The bilingual PDF is produced with Microsoft Word (via COM) if available, then
LibreOffice headless as a fallback. If neither is installed, the `.docx` files
are still created and the tool tells you how to export the PDF manually from
Word (Save As → PDF).

---

## Format mode

A raw translated book reads as walls of shredded text: cover/front matter at the
top, a running header repeated on every page, source-page markers and difficulty
flags inline, and one short paragraph per OCR line. Format mode cleans all of
that and produces **`{name}_formatted.docx`**: a clean, legible, properly
footnoted Word document.

Run it from the menu (**[F]**), or with `translate-doc --format`. It opens a
picker for the `_translated.docx` (or takes a path argument) and writes the
formatted file into the same folder.

**It is a Pandoc round-trip.** python-docx cannot create or cleanly read real
Word footnotes, so Format mode never rebuilds the document by hand:

1. **Export** — `pandoc input.docx -t markdown --wrap=none`. Real Word footnotes
   come out as `[^N]` references inline and `[^N]:` definitions.
2. **Clean the Markdown** (below) — every footnote `[^N]` ref and `[^N]:`
   definition is left untouched; definitions are set aside verbatim and
   re-appended at the end.
3. **Rebuild** — `pandoc work.md -o output.docx --reference-doc=styled.docx`.
   The footnotes come back as genuine, auto-numbered Word footnotes anchored at
   the bottom of the page.

The cleaning pass:

- **Drops cover / front matter.** It starts the body at the first real content —
  the basmala, or the first clean-prose page — discarding everything before it
  (segments carrying ISBN, "All Rights Reserved", a URL, "Edition", etc., or a
  high ratio of non-dictionary tokens).
- **Removes source-page markers** (`^[N]^` superscripts) while leaving `[^N]`
  footnote refs alone.
- **Strips running headers by recurrence, not keywords.** The leading phrase of
  each page is normalized and counted; any phrase appearing on ≥ max(8, 5% of
  pages) is a running header and is removed throughout, along with a trailing
  printed page number and any adjacent flag.
- **Promotes section titles to headings.** A short, title-like leading phrase
  that occurs only a few times becomes a Markdown `## heading` at its first
  occurrence; later repeats are stripped.
- **Strips junk and fixes spacing:** `<br>`, `***`, `XXX`, `{.mark}` spans and
  leaked tags are removed; each difficulty flag becomes a superscript asterisk
  (`^\*^`); a space is restored where OCR glued sentence punctuation to the next
  word; whitespace is collapsed.
- **Paragraphs the body.** When `ANTHROPIC_API_KEY` is set, the **preferred**
  path sends the text to Claude (`claude-sonnet-5`) in chunks, which inserts
  paragraph breaks only where the argument shifts (it handles the long run-on
  sentences the mechanical pass can't). The model is **only** allowed to add
  breaks: every chunk is verified and discarded if a word or marker changed.
  Progress is checkpointed to `{name}_format_checkpoint.json` (resume by
  re-running, or `--restart` to ignore). Without a key, a **mechanical fallback**
  splits the body into sentences and groups them at discourse markers (*As for*,
  *The first*, *Then*, *He said:*, *Ibn …*, *Al-…*, …) and sentence-length
  thresholds — it never produces one paragraph per page. A post-pass glues any
  orphaned footnote ref or asterisk onto the previous paragraph.
- **Styles with a generated reference doc:** US Letter, one-inch margins, 12pt
  Times New Roman body — justified, 1.3 line spacing, 0.3" first-line indent,
  ~8pt space after — ~9.5pt footnotes, section titles as Heading 2, and centered
  page numbers in the footer.

Format mode requires **Pandoc**. If it is not installed the tool prints the
install command and stops without discarding any checkpoint.

The styling comes from a reference document bundled **inside the package** at
`translate_doc/assets/styled.docx` (declared as package data in
`pyproject.toml` and loaded with `importlib.resources`), so it is found from any
install — `setup.ps1`, `pip install`, or `uv tool install .`. If that file is
missing the tool stops with an error rather than silently producing unstyled
output, and the run summary prints the exact path it loaded the reference doc
from. To change the styling, edit `build_reference_doc()` in `formatter.py` and
regenerate the asset:

```cmd
python -c "from pathlib import Path; from translate_doc.formatter import build_reference_doc; build_reference_doc(Path('translate_doc/assets/styled.docx'))"
```

The end-of-run summary reports: cover pages dropped, running headers removed,
section headings promoted, paragraphs produced, footnotes preserved, and whether
the Claude pass ran or the mechanical fallback was used.

### Notes and honest limits

- **Running headers are found by recurrence.** The tool keys off the phrase that
  repeats at the top of each page. A book whose pages don't share a consistent
  header line, or whose header is OCR'd differently each time, may keep some
  header text; a one-off line that happens to repeat could be trimmed.
- **The paragraph breaks are Claude's judgment** (or the mechanical rules'), not
  the original author's. The source's real paragraphing is not recoverable from
  the flattened text, so the breaks reflect where the *argument* appears to
  shift — a best-effort reconstruction, not the author's own layout.
- **The asterisks are pointers, not the notes themselves.** They mark passages
  the translator flagged as difficult during Translate mode; the full text of
  those notes lives in that run's `{name}_translation_log.csv`.

---

## Tuning

Constants you may want to adjust live near the top of the source modules:

- `translate_doc/translator.py`
  - `MAX_PARAGRAPHS_PER_CHUNK` — lower this if responses hit the output token
    limit.
  - `MODEL` — the Claude model string (`claude-sonnet-5`).
- `translate_doc/parser.py`
  - `PARAGRAPHS_PER_PAGE` — the ratio used to estimate page count.
- `translate_doc/formatter.py`
  - `WORDS_PER_CHUNK` — chunk size for the Claude paragraphing calls.
  - `_PARA_*` thresholds — tune where the mechanical fallback starts new
    paragraphs (discourse-marker and sentence-length limits).
  - `FRONT_MATTER_MARKERS` / `DISCOURSE_RE` — the cover-detection signals and the
    discourse markers that trigger a paragraph break.

---

## Project structure

```
translate-doc\
├── setup.ps1
├── pyproject.toml
├── README.md
└── translate_doc\
    ├── __init__.py
    ├── main.py          # CLI entry point, startup menu, prompts
    ├── parser.py        # .docx parsing and structure extraction
    ├── translator.py    # Claude API calls and chunking logic
    ├── builder.py       # Reconstructing the output .docx files
    ├── pdf_export.py    # Word COM / LibreOffice PDF conversion
    ├── merger.py        # Merge mode
    ├── formatter.py     # Format mode (clean + real footnotes via Pandoc)
    └── logger.py        # CSV log generation
```

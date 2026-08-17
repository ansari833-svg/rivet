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
one document.

> This tool is Windows-only.

---

## Requirements

- Windows
- Python 3.10+ (install from [python.org](https://www.python.org/downloads/);
  the standard installer bundles `tkinter`, which the file picker needs)
- An Anthropic API key
- Optional, for PDF export: Microsoft Word **or**
  [LibreOffice](https://www.libreoffice.org/)

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
│  What would you like to do?             │
│  [T] Translate a document               │
│  [M] Merge existing documents           │
│                                         │
│  Enter T or M:                          │
└─────────────────────────────────────────┘
```

- **[T]** opens a native file picker, asks how many pages to translate, shows a
  pre-flight summary, then translates.
- **[M]** scans a folder and merges the files you choose.

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

## Tuning

Constants you may want to adjust live near the top of the source modules:

- `translate_doc/translator.py`
  - `MAX_PARAGRAPHS_PER_CHUNK` — lower this if responses hit the output token
    limit.
  - `MODEL` — the Claude model string (`claude-sonnet-5`).
- `translate_doc/parser.py`
  - `PARAGRAPHS_PER_PAGE` — the ratio used to estimate page count.

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
    └── logger.py        # CSV log generation
```

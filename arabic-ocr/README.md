# arabic-ocr

Two-pass Arabic page transcription for classical printed books: Surya (local)
transcribes and is the authoritative text; Claude verifies against the page
image and flags discrepancies for a human to resolve.

## Install (one line)

Open **PowerShell** and paste this. It installs uv if you don't have it,
installs the tool, and puts it on your PATH:

    powershell -ExecutionPolicy Bypass -c "irm https://astral.sh/uv/install.ps1 | iex; uv tool install 'git+https://github.com/ansari833-svg/rivet.git@claude/arabic-ocr-cli-2qn0cr#subdirectory=arabic-ocr'; uv tool update-shell"

Then **open a new terminal** and run:

    arabic-ocr

## Install from a downloaded copy (alternative)

If you have the project folder instead, run this from inside it:

    uv tool install .

## Anthropic API key (Claude verification pass)

The verification pass needs an API key in the `ANTHROPIC_API_KEY` environment
variable. Set it in **PowerShell** for the current session:

    $env:ANTHROPIC_API_KEY = "sk-ant-..."

or persistently via **Windows Settings -> Edit environment variables for your
account**. You can skip the key entirely and run Surya only with `--offline`.

## First run

Surya downloads its OCR model weights once (needs network); every run after
that is offline for transcription. Flags: `--offline` (skip the Claude pass),
`--restart` (ignore a saved checkpoint), `--no-footnotes` (whole page as body).

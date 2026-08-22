# pdf-screenshots

A Windows command-line tool that converts each page of a PDF into a
consistently sized JPEG image.

## Install

If you don't have `uv`, install it first (Windows PowerShell):

```powershell
powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
```

Then install the tool from this folder:

```powershell
uv tool install .
```

If `pdf-screenshots` is not on your PATH afterward, run:

```powershell
uv tool update-shell
```

## Run

```powershell
pdf-screenshots
```

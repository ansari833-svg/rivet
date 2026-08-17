"""Export the bilingual .docx to PDF.

Tries Microsoft Word via COM first (best fidelity for RTL text, footnotes and
headers), then falls back to LibreOffice headless. If neither is available the
user is told how to export manually.
"""

from __future__ import annotations

import shutil
import subprocess
from pathlib import Path

# Standard install locations for LibreOffice on Windows (not on PATH by default).
_LIBREOFFICE_PATHS = (
    Path(r"C:\Program Files\LibreOffice\program\soffice.exe"),
    Path(r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"),
)

_MANUAL_MESSAGE = (
    "PDF export requires Microsoft Word or LibreOffice. The .docx files were "
    "created successfully — open the bilingual .docx in Word and use "
    "Save As → PDF to produce the PDF manually."
)


def _find_libreoffice() -> Path | None:
    for candidate in _LIBREOFFICE_PATHS:
        if candidate.exists():
            return candidate
    found = shutil.which("soffice")
    if found:
        return Path(found)
    return None


def _convert_with_word(docx_path: Path, pdf_path: Path) -> bool:
    """Method 1 — Microsoft Word via COM automation (docx2pdf)."""
    try:
        from docx2pdf import convert  # imported lazily; pulls in pywin32/COM
    except Exception:
        return False
    try:
        convert(str(docx_path), str(pdf_path))
    except Exception:
        # Word not installed, or the file is already open in a Word process.
        return False
    return pdf_path.exists()


def _convert_with_libreoffice(docx_path: Path, pdf_path: Path) -> bool:
    """Method 2 — LibreOffice headless fallback."""
    soffice = _find_libreoffice()
    if soffice is None:
        return False
    try:
        subprocess.run(
            [
                str(soffice),
                "--headless",
                "--convert-to",
                "pdf",
                "--outdir",
                str(pdf_path.parent),
                str(docx_path),
            ],
            check=True,
            capture_output=True,
        )
    except (subprocess.CalledProcessError, OSError):
        return False
    # LibreOffice names the output after the input stem in the outdir.
    produced = pdf_path.parent / f"{docx_path.stem}.pdf"
    if produced != pdf_path and produced.exists():
        produced.replace(pdf_path)
    return pdf_path.exists()


def export_pdf(docx_path: Path, pdf_path: Path) -> bool:
    """Attempt PDF conversion, returning True on success.

    Prints the manual-export message if no converter is available.
    """
    if _convert_with_word(docx_path, pdf_path):
        return True
    if _convert_with_libreoffice(docx_path, pdf_path):
        return True
    print(_MANUAL_MESSAGE)
    return False

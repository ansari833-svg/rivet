"""CSV translation log.

Produces ``{name}_translation_log.csv`` — a complete record of everything in
the translated portion that a human should review: every ``[DIFFICULT: ...]``
flag plus every block that was missing from the API response.

The file is written with ``utf-8-sig`` (BOM) so Excel on Windows renders the
Arabic column correctly, and ``newline=""`` so Windows does not double the
row spacing.
"""

from __future__ import annotations

import csv
from dataclasses import dataclass
from pathlib import Path

from .parser import ParsedDocument
from .translator import extract_difficult_flags, strip_difficult_tags

COLUMNS = [
    "section_number",
    "paragraph_or_location",
    "original_arabic",
    "translated_english",
    "difficulty_reason",
    "translator_note",
]

MISSING_REASON = "MISSING — not returned by API"


@dataclass
class LogRow:
    block_id: int  # used for document-order sorting only
    paragraph_or_location: str
    original_arabic: str
    translated_english: str
    difficulty_reason: str
    translator_note: str


def _location_label(parsed: ParsedDocument, block_id: int) -> str:
    """Human-friendly location, e.g. 'Paragraph 12' or 'Table cell 12'."""
    for el in parsed.elements:
        if el.kind == "para" and el.block_id == block_id:
            return f"Paragraph {block_id}"
        if el.kind == "table" and el.grid is not None:
            for r, row in enumerate(el.grid):
                for c, bid in enumerate(row):
                    if bid == block_id:
                        return f"Table cell (row {r + 1}, col {c + 1})"
    return f"Block {block_id}"


def build_rows(
    parsed: ParsedDocument,
    translations: dict[int, str],
    selected_ids: set[int],
) -> list[LogRow]:
    rows: list[LogRow] = []
    for bid in sorted(selected_ids):
        block = parsed.blocks[bid]
        if not block.text.strip():
            continue
        english = translations.get(bid, "")
        location = _location_label(parsed, bid)

        # Missing block: kept original Arabic, no translation returned.
        if not english.strip():
            rows.append(
                LogRow(
                    block_id=bid,
                    paragraph_or_location=location,
                    original_arabic=block.text,
                    translated_english="",
                    difficulty_reason=MISSING_REASON,
                    translator_note=(
                        "This passage was not translated in the automated run "
                        "and requires manual review."
                    ),
                )
            )
            continue

        # One row per difficulty flag in the block.
        for flag in extract_difficult_flags(bid, english):
            rows.append(
                LogRow(
                    block_id=bid,
                    paragraph_or_location=location,
                    original_arabic=flag.original_arabic,
                    translated_english=strip_difficult_tags(english),
                    difficulty_reason=flag.reason,
                    translator_note=(
                        f"Translator's note: {flag.original_arabic} — {flag.reason}"
                    ),
                )
            )
    return rows


def write_log(
    parsed: ParsedDocument,
    translations: dict[int, str],
    selected_ids: set[int],
    output_path: Path,
) -> int:
    """Write the CSV log; returns the number of flagged rows."""
    rows = build_rows(parsed, translations, selected_ids)
    rows.sort(key=lambda r: r.block_id)  # document order

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with open(output_path, "w", encoding="utf-8-sig", newline="") as fh:
        writer = csv.writer(fh)
        writer.writerow(COLUMNS)
        for i, row in enumerate(rows, start=1):
            writer.writerow(
                [
                    i,
                    row.paragraph_or_location,
                    row.original_arabic,
                    row.translated_english,
                    row.difficulty_reason,
                    row.translator_note,
                ]
            )
    return len(rows)

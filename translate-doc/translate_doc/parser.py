"""Parse an Arabic .docx file into an ordered structural model.

The model is intentionally flat: every translatable unit of text becomes a
:class:`Block` with a stable integer id (document order). The overall document
structure is captured as an ordered list of :class:`Element` entries that either
reference a single paragraph block or lay out a table as a grid of block ids.

This keeps chunking, checkpointing and "missing block" handling simple: the
translator works on block ids, and the builder rebuilds the document from the
same ids.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

from docx import Document as _open_document
from docx.document import Document as _DocxDocument
from docx.oxml.ns import qn
from docx.table import Table as _DocxTable
from docx.text.paragraph import Paragraph as _DocxParagraph

# Approximate number of paragraphs that make up one printed page. Tunable.
PARAGRAPHS_PER_PAGE = 25


@dataclass
class RunFmt:
    """Character-run formatting captured from the source document."""

    text: str
    bold: Optional[bool] = None
    italic: Optional[bool] = None
    underline: Optional[bool] = None
    size_pt: Optional[float] = None
    font_name: Optional[str] = None


@dataclass
class Block:
    """A single translatable paragraph-like unit."""

    id: int
    text: str
    style_name: Optional[str] = None
    heading_level: Optional[int] = None
    alignment: Optional[str] = None  # WD_ALIGN_PARAGRAPH member name, e.g. "LEFT"
    rtl: bool = False
    runs: list[RunFmt] = field(default_factory=list)
    page: int = 1

    def dominant_bold(self) -> bool:
        real = [r for r in self.runs if r.text.strip()]
        return bool(real) and all(r.bold for r in real)

    def dominant_italic(self) -> bool:
        real = [r for r in self.runs if r.text.strip()]
        return bool(real) and all(r.italic for r in real)

    def dominant_underline(self) -> bool:
        real = [r for r in self.runs if r.text.strip()]
        return bool(real) and all(r.underline for r in real)

    def first_size_pt(self) -> Optional[float]:
        for r in self.runs:
            if r.size_pt:
                return r.size_pt
        return None

    def first_font_name(self) -> Optional[str]:
        for r in self.runs:
            if r.font_name:
                return r.font_name
        return None


@dataclass
class Element:
    """A structural element: either a paragraph or a table.

    - kind == "para": ``block_id`` refers to the paragraph block.
    - kind == "table": ``grid`` is a list of rows, each a list of block ids.
    """

    kind: str
    block_id: Optional[int] = None
    grid: Optional[list[list[int]]] = None


@dataclass
class ParsedDocument:
    blocks: dict[int, Block]
    elements: list[Element]
    page_count: int

    def ordered_block_ids(self) -> list[int]:
        return sorted(self.blocks.keys())

    def block_ids_for_pages(self, start: int, end: int) -> list[int]:
        return [
            bid
            for bid in self.ordered_block_ids()
            if start <= self.blocks[bid].page <= end
        ]


def _iter_body_items(document: _DocxDocument):
    """Yield paragraphs and tables from the document body in true order."""
    body = document.element.body
    for child in body.iterchildren():
        if child.tag == qn("w:p"):
            yield _DocxParagraph(child, document)
        elif child.tag == qn("w:tbl"):
            yield _DocxTable(child, document)


def _paragraph_is_rtl(paragraph: _DocxParagraph) -> bool:
    p_pr = paragraph._p.find(qn("w:pPr"))
    if p_pr is not None and p_pr.find(qn("w:bidi")) is not None:
        return True
    # Fall back to run-level rtl marker.
    for run in paragraph.runs:
        r_pr = run._element.find(qn("w:rPr"))
        if r_pr is not None and r_pr.find(qn("w:rtl")) is not None:
            return True
    return False


def _paragraph_has_page_break(paragraph: _DocxParagraph) -> bool:
    for br in paragraph._p.iter(qn("w:br")):
        if br.get(qn("w:type")) == "page":
            return True
    # A rendered page-break element (Word 2013+).
    if list(paragraph._p.iter(qn("w:lastRenderedPageBreak"))):
        return True
    # Section break introduces a new page in most layouts.
    p_pr = paragraph._p.find(qn("w:pPr"))
    if p_pr is not None and p_pr.find(qn("w:sectPr")) is not None:
        return True
    return False


def _extract_runs(paragraph: _DocxParagraph) -> list[RunFmt]:
    runs: list[RunFmt] = []
    for run in paragraph.runs:
        size_pt = None
        if run.font is not None and run.font.size is not None:
            size_pt = run.font.size.pt
        runs.append(
            RunFmt(
                text=run.text,
                bold=run.bold,
                italic=run.italic,
                underline=bool(run.underline) if run.underline is not None else None,
                size_pt=size_pt,
                font_name=run.font.name if run.font is not None else None,
            )
        )
    return runs


def _heading_level(style_name: Optional[str]) -> Optional[int]:
    if not style_name:
        return None
    name = style_name.strip().lower()
    if name.startswith("heading"):
        tail = name.replace("heading", "").strip()
        if tail.isdigit():
            return int(tail)
        return 1
    if name in ("title",):
        return 0
    return None


def parse(path: Path) -> ParsedDocument:
    """Parse a .docx file into a :class:`ParsedDocument`."""
    document = _open_document(str(path))

    blocks: dict[int, Block] = {}
    elements: list[Element] = []
    next_id = 0

    # Page tracking as we walk the document in order.
    current_page = 1
    paras_on_page = 0

    def advance_page(paragraph: _DocxParagraph) -> None:
        nonlocal current_page, paras_on_page
        paras_on_page += 1
        if _paragraph_has_page_break(paragraph):
            current_page += 1
            paras_on_page = 0
        elif paras_on_page >= PARAGRAPHS_PER_PAGE:
            current_page += 1
            paras_on_page = 0

    def make_paragraph_block(paragraph: _DocxParagraph) -> int:
        nonlocal next_id
        style_name = paragraph.style.name if paragraph.style is not None else None
        alignment = None
        if paragraph.paragraph_format.alignment is not None:
            try:
                alignment = paragraph.paragraph_format.alignment._member_name  # type: ignore[attr-defined]
            except AttributeError:
                alignment = str(paragraph.paragraph_format.alignment)
        block = Block(
            id=next_id,
            text=paragraph.text,
            style_name=style_name,
            heading_level=_heading_level(style_name),
            alignment=alignment,
            rtl=_paragraph_is_rtl(paragraph),
            runs=_extract_runs(paragraph),
            page=current_page,
        )
        blocks[block.id] = block
        next_id += 1
        return block.id

    for item in _iter_body_items(document):
        if isinstance(item, _DocxParagraph):
            advance_page(item)
            bid = make_paragraph_block(item)
            elements.append(Element(kind="para", block_id=bid))
        elif isinstance(item, _DocxTable):
            grid: list[list[int]] = []
            for row in item.rows:
                row_ids: list[int] = []
                for cell in row.cells:
                    # Join the cell's paragraphs into a single block. Cell text is
                    # usually short; this keeps the grid rectangular and simple.
                    cell_para = cell.paragraphs[0] if cell.paragraphs else None
                    if cell_para is not None:
                        runs: list[RunFmt] = []
                        text_parts: list[str] = []
                        for p in cell.paragraphs:
                            runs.extend(_extract_runs(p))
                            text_parts.append(p.text)
                        block = Block(
                            id=next_id,
                            text="\n".join(text_parts),
                            style_name=(
                                cell_para.style.name
                                if cell_para.style is not None
                                else None
                            ),
                            heading_level=None,
                            rtl=_paragraph_is_rtl(cell_para),
                            runs=runs,
                            page=current_page,
                        )
                    else:
                        block = Block(id=next_id, text="", page=current_page)
                    blocks[block.id] = block
                    row_ids.append(block.id)
                    next_id += 1
                grid.append(row_ids)
            elements.append(Element(kind="table", grid=grid))

    page_count = max(current_page, 1)
    return ParsedDocument(blocks=blocks, elements=elements, page_count=page_count)


def estimate_page_count(path: Path) -> int:
    """Cheap page-count estimate without a full parse (used for the prompt)."""
    return parse(path).page_count

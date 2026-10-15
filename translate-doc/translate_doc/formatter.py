"""Format mode: a pandoc round-trip pipeline.

python-docx cannot create or cleanly read real Word footnotes; pandoc does
both. So Format mode exports the input ``.docx`` to Markdown with pandoc,
performs every text manipulation on that Markdown (never touching ``[^N]``
footnote refs or ``[^N]:`` definitions), then rebuilds a styled ``.docx`` with
pandoc — the footnotes come back as genuine, auto-numbered Word footnotes.

Pipeline:
    1. pandoc input.docx -t markdown --wrap=none            (export)
    2. clean + reparagraph the Markdown                     (this module)
    3. pandoc work_clean.md -o out.docx --reference-doc=... (rebuild)
"""

from __future__ import annotations

import json
import math
import re
import shutil
import subprocess
import time
from collections import Counter
from importlib import resources
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

MODEL = "claude-sonnet-5"
WORDS_PER_CHUNK = 2500

# ---- Markdown tokens we must recognise or preserve ------------------------

# Footnote reference (no leading caret) and definition line.
FOOTNOTE_REF_RE = re.compile(r"\[\^\d+\]")
FOOTNOTE_DEF_RE = re.compile(r"^\[\^\d+\]:")
# Source-page marker: superscript [N], brackets possibly escaped by pandoc.
PAGE_MARKER_RE = re.compile(r"\^\\?\[\d+\\?\]\^")
# The asterisk flag we emit (escaped star inside a pandoc superscript).
ASTERISK_MD = "^\\*^"

# Junk / OCR noise.
SEE_LOG_RE = re.compile(r"\\?\[\s*(?:†|\\u2020)?\s*see\s*log\s*\\?\]", re.IGNORECASE)
LEAKED_TAG_RE = re.compile(r"</?(?:sup|sub|u)>", re.IGNORECASE)
GLUED_PUNCT_RE = re.compile(r"([.!?؟])([\"“A-Z])")

# Cover / front-matter signals.
FRONT_MATTER_MARKERS = (
    "isbn",
    "all rights reserved",
    "http",
    "www.",
    "@",
    "p.o. box",
    "edition",
    "<br>",
    "xxx",
)

# Discourse markers that should begin a new paragraph (mechanical fallback).
DISCOURSE_RE = re.compile(
    r"^(?:as for\b|the first\b|the second\b|the third\b|then\b|accordingly\b|"
    r"therefore\b|thus\b|in summary\b|know that\b|he said:|and allah\b|"
    r"allah(?: the exalted| the almighty| almighty)? said|ibn\b|al-)",
    re.IGNORECASE,
)

# Mechanical-paragraphing thresholds.
_PARA_MARKER_MIN_CHARS = 160
_PARA_LEN_MIN_SENTENCES = 2
_PARA_LEN_MIN_CHARS = 380
_PARA_MAX_CHARS = 1100
_PARA_SPLIT_MIN_CHARS = 600

_SENTENCE_END = set(".!?؟…")

PARAGRAPHING_SYSTEM_PROMPT = """\
You insert paragraph breaks into an already-translated passage whose \
paragraphing has been flattened into running text.

Your ONLY permitted action is to insert paragraph breaks (blank lines) between \
sentences, and only where the argument genuinely shifts — a new point, a new \
hadith, a move from a ruling to its evidence, or the start of a new section.

Absolute rules:
1. Do NOT change, add, remove, reorder, correct, or translate any words, \
punctuation, or characters. The text between breaks must be identical to the \
input apart from the whitespace you turn into paragraph breaks.
2. Preserve every marker exactly and in place. Footnote references look like \
[^12]; the flag marker looks like ^\\*^. Never alter, drop, duplicate, move, or \
translate them.
3. Only break BETWEEN sentences, never inside one. If unsure, do not break.
4. Return ONLY the re-paragraphed text — no commentary, no headings, no markup \
beyond the blank lines separating paragraphs.
"""


# ---------------------------------------------------------------------------
# Result + availability
# ---------------------------------------------------------------------------


@dataclass
class FormatResult:
    output_path: Path
    cover_pages_dropped: int
    running_headers_removed: int
    section_headings: int
    paragraphs_produced: int
    footnotes_preserved: int
    paragraphing_ran: bool
    reference_doc: str = ""


def pandoc_available() -> bool:
    return shutil.which("pandoc") is not None


PANDOC_INSTALL_HINT = (
    "Format mode needs Pandoc. Install it with: "
    "winget install --id JohnMacFarlane.Pandoc"
)


class PandocMissingError(RuntimeError):
    """Pandoc is not on PATH."""


class ReferenceDocMissingError(RuntimeError):
    """The bundled styled.docx reference doc could not be found."""


def _require_pandoc() -> None:
    if not pandoc_available():
        raise PandocMissingError(PANDOC_INSTALL_HINT)


def bundled_reference_doc():
    """Context manager yielding a real filesystem path to the bundled
    ``translate_doc/assets/styled.docx``.

    The file ships as package data and is located with importlib.resources, so
    it is found whether the tool runs from the project folder or from an
    installed copy. If it is missing, raise rather than fall back to pandoc's
    default (unstyled) reference doc.
    """
    resource = resources.files("translate_doc").joinpath("assets").joinpath("styled.docx")
    if not resource.is_file():
        raise ReferenceDocMissingError(
            "Format mode's bundled reference doc (translate_doc/assets/styled.docx) "
            f"was not found in the installed package (looked for: {resource}). "
            "The install is incomplete — reinstall translate-doc. Refusing to "
            "continue with default styling."
        )
    return resources.as_file(resource)


# ---------------------------------------------------------------------------
# Step 0 — pandoc export
# ---------------------------------------------------------------------------


def export_markdown(input_docx: Path) -> str:
    result = subprocess.run(
        ["pandoc", str(input_docx), "-t", "markdown", "--wrap=none"],
        capture_output=True,
        text=True,
        check=True,
    )
    return result.stdout


# ---------------------------------------------------------------------------
# Step 1 — set footnote definitions aside
# ---------------------------------------------------------------------------


def extract_footnote_defs(md: str) -> tuple[str, list[str]]:
    """Pull every ``[^N]: ...`` definition (plus indented continuations) out of
    the Markdown, returning (body, definitions). Definitions are kept verbatim.
    """
    lines = md.split("\n")
    body: list[str] = []
    defs: list[str] = []
    i, n = 0, len(lines)
    while i < n:
        if FOOTNOTE_DEF_RE.match(lines[i]):
            block = [lines[i]]
            i += 1
            # Capture multi-line / multi-paragraph continuations (indented).
            while i < n:
                if lines[i].strip() == "":
                    j = i
                    while j < n and lines[j].strip() == "":
                        j += 1
                    if j < n and (lines[j].startswith("    ") or lines[j].startswith("\t")):
                        block.extend(lines[i : j + 1])
                        i = j + 1
                        continue
                    break
                if lines[i].startswith("    ") or lines[i].startswith("\t"):
                    block.append(lines[i])
                    i += 1
                    continue
                break
            defs.append("\n".join(block).rstrip())
        else:
            body.append(lines[i])
            i += 1
    return "\n".join(body), defs


# ---------------------------------------------------------------------------
# Step 2 — segmentation + cover / front-matter detection
# ---------------------------------------------------------------------------


@dataclass
class Segment:
    text: str
    has_page_number: bool


def split_into_segments(body: str) -> list[Segment]:
    """Split the body at page markers. Each piece is one source page-segment."""
    pieces = PAGE_MARKER_RE.split(body)
    segments = [Segment(text=pieces[0], has_page_number=False)]
    for piece in pieces[1:]:
        segments.append(Segment(text=piece, has_page_number=True))
    return [s for s in segments if s.text.strip()]


def _non_dictionary_ratio(text: str) -> float:
    tokens = re.findall(r"\S+", text)
    if not tokens:
        return 1.0
    bad = 0
    for tok in tokens:
        core = re.sub(r"[^A-Za-z]", "", tok)
        if not core or (len(core) == 1 and core.lower() not in ("a", "i")):
            bad += 1
        elif any(ch.isdigit() for ch in tok):
            bad += 1
    return bad / len(tokens)


def _is_front_matter(text: str) -> bool:
    low = text.lower()
    if any(marker in low for marker in FRONT_MATTER_MARKERS):
        return True
    return _non_dictionary_ratio(text) > 0.35


def _has_basmala(text: str) -> bool:
    return "in the name of allah" in text.lower()


def _is_clean_prose(text: str) -> bool:
    words = re.findall(r"[A-Za-z]+", text)
    if len(words) < 12:
        return False
    if _non_dictionary_ratio(text) > 0.3:
        return False
    return bool(re.search(r"[.!?؟]", text))


def find_content_start(segments: list[Segment]) -> int:
    """Index of the first real-content segment; everything before is cover."""
    for idx, seg in enumerate(segments):
        if _has_basmala(seg.text):
            return idx
        if _is_front_matter(seg.text):
            continue
        if _is_clean_prose(seg.text):
            return idx
    return 0


# ---------------------------------------------------------------------------
# Step 4 / 5 — leading phrases, running headers, section titles
# ---------------------------------------------------------------------------


def _leading_phrase(text: str) -> str:
    # The running header / title is the first line of the page-segment. Take that
    # line, then truncate to the first sentence-end or the first ~10 words.
    first_line = ""
    for line in text.split("\n"):
        if line.strip():
            first_line = line.strip()
            break
    t = re.sub(r'^(?:\[\^\d+\]|\^\\\*\^|["“\'(]|\s)+', "", first_line).strip()
    m = re.search(r"[.!?؟…]", t)
    if m and m.start() <= 120:
        phrase = t[: m.start()]
    else:
        phrase = " ".join(t.split()[:10])
    return phrase.strip()


def _normalize_phrase(phrase: str) -> str:
    low = re.sub(r"[^a-z\s]", " ", phrase.lower())
    return re.sub(r"\s+", " ", low).strip()


def _is_title_like(phrase: str) -> bool:
    if phrase.endswith("?"):
        return True
    words = [w for w in re.findall(r"[A-Za-z]+", phrase)]
    if not words:
        return False
    capped = sum(1 for w in words if w[0].isupper())
    return capped / len(words) >= 0.6


def _header_removal_regex(normalized: str) -> re.Pattern:
    words = normalized.split()
    core = r"[\W\s]+".join(re.escape(w) for w in words)
    # Optional trailing page number and/or asterisk flag.
    trailer = r"(?:\s*(?:\^\\?\*\^|\d{1,4}\b))*"
    return re.compile(r"\b" + core + r"\b" + trailer, re.IGNORECASE)


# ---------------------------------------------------------------------------
# Step 6 — per-segment cleanup
# ---------------------------------------------------------------------------


def _clean_segment_text(text: str) -> str:
    text = PAGE_MARKER_RE.sub(" ", text)
    text = re.sub(r"<br\s*/?>", " ", text, flags=re.IGNORECASE)
    text = text.replace("***", " ").replace("XXX", " ")
    text = re.sub(r"\[\]\{[^}]*\}", "", text)  # empty spans
    text = re.sub(r"\{\.[^}]*\}", "", text)     # {.mark} / {.underline}
    text = LEAKED_TAG_RE.sub("", text)
    # Flag markers -> superscript asterisk.
    text = SEE_LOG_RE.sub(f" {ASTERISK_MD} ", text)
    text = text.replace("†", f" {ASTERISK_MD} ")
    # Un-glue sentence punctuation stuck to the next word.
    text = GLUED_PUNCT_RE.sub(r"\1 \2", text)
    # Reflow to prose: newlines -> spaces, collapse whitespace.
    text = re.sub(r"\s*\n\s*", " ", text)
    text = re.sub(r"[ \t]+", " ", text)
    return re.sub(r"\s{2,}", " ", text).strip()


# ---------------------------------------------------------------------------
# Step 7 — paragraphing (mechanical fallback)
# ---------------------------------------------------------------------------


def _split_sentences(text: str) -> list[str]:
    sentences: list[str] = []
    start, i, L = 0, 0, len(text)
    while i < L:
        if text[i] in _SENTENCE_END:
            j = i + 1
            # Consume trailing closers and markers so they stay with the sentence.
            while j < L:
                if text[j] in "\"”»')]":
                    j += 1
                elif text.startswith(ASTERISK_MD, j):
                    j += len(ASTERISK_MD)
                else:
                    m = FOOTNOTE_REF_RE.match(text, j)
                    if m:
                        j = m.end()
                    else:
                        break
            k = j
            while k < L and text[k] in " \t":
                k += 1
            if k > j and k < L and (text[k].isupper() or text[k] in "\"“(["):
                sentences.append(text[start:j].strip())
                start = i = k
                continue
        i += 1
    tail = text[start:].strip()
    if tail:
        sentences.append(tail)
    return [s for s in sentences if s]


def _starts_discourse(sentence: str) -> bool:
    s = FOOTNOTE_REF_RE.sub("", sentence).replace(ASTERISK_MD, "")
    s = s.lstrip(" \"“'‘([")
    return bool(DISCOURSE_RE.match(s))


def _split_long_paragraph(para: str) -> list[str]:
    if len(para) <= _PARA_MAX_CHARS:
        return [para]
    for m in re.finditer(r"[;:]", para):
        if m.start() >= _PARA_SPLIT_MIN_CHARS:
            head = para[: m.end()].strip()
            tail = para[m.end() :].strip()
            return [head] + _split_long_paragraph(tail) if tail else [head]
    return [para]


def _only_markers(text: str) -> bool:
    stripped = FOOTNOTE_REF_RE.sub("", text).replace(ASTERISK_MD, "")
    return not re.search(r"[A-Za-z0-9]", stripped)


def _glue_orphans(paragraphs: list[str]) -> list[str]:
    """Never let a paragraph start with a bare marker or be marker-only."""
    out: list[str] = []
    for para in paragraphs:
        p = para.strip()
        if not p:
            continue
        starts_marker = bool(FOOTNOTE_REF_RE.match(p)) or p.startswith(ASTERISK_MD)
        if out and (starts_marker or _only_markers(p)):
            out[-1] = f"{out[-1]} {p}".strip()
        else:
            out.append(p)
    return out


def mechanical_paragraphs(text: str) -> list[str]:
    body = re.sub(r"\s+", " ", text).strip()
    body = GLUED_PUNCT_RE.sub(r"\1 \2", body)
    if not body:
        return []
    sentences = _split_sentences(body)

    paragraphs: list[str] = []
    cur: list[str] = []
    cur_len = 0
    for s in sentences:
        if (
            cur
            and _starts_discourse(s)
            and len(cur) >= 1
            and cur_len >= _PARA_MARKER_MIN_CHARS
        ):
            paragraphs.append(" ".join(cur))
            cur, cur_len = [], 0
        cur.append(s)
        cur_len += len(s) + 1
        if len(cur) >= _PARA_LEN_MIN_SENTENCES and cur_len >= _PARA_LEN_MIN_CHARS:
            paragraphs.append(" ".join(cur))
            cur, cur_len = [], 0
    if cur:
        paragraphs.append(" ".join(cur))

    result: list[str] = []
    for para in paragraphs:
        result.extend(_split_long_paragraph(para))
    return _glue_orphans(result)


# ---------------------------------------------------------------------------
# Step 7 — LLM paragraphing (preferred when a key is present)
# ---------------------------------------------------------------------------


class FormatCheckpoint:
    """Checkpoint for the LLM paragraphing step."""

    def __init__(self, path: Path) -> None:
        self.path = path
        self.total: Optional[int] = None
        self.outputs: dict[int, str] = {}

    @classmethod
    def path_for(cls, input_path: Path) -> Path:
        return input_path.parent / f"{input_path.stem}_format_checkpoint.json"

    @classmethod
    def load(cls, input_path: Path) -> Optional["FormatCheckpoint"]:
        path = cls.path_for(input_path)
        if not path.exists():
            return None
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
        except (json.JSONDecodeError, OSError):
            return None
        cp = cls(path)
        cp.total = data.get("total")
        cp.outputs = {int(k): v for k, v in data.get("outputs", {}).items()}
        return cp

    @classmethod
    def start(cls, input_path: Path) -> "FormatCheckpoint":
        return cls(cls.path_for(input_path))

    def _save(self) -> None:
        payload = {
            "total": self.total,
            "outputs": {str(k): v for k, v in self.outputs.items()},
        }
        self.path.write_text(
            json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8"
        )

    def ensure_total(self, total: int) -> None:
        if self.total is None:
            self.total = total
            self._save()

    def record(self, index: int, text: str) -> None:
        self.outputs[index] = text
        self._save()

    def delete(self) -> None:
        try:
            self.path.unlink()
        except FileNotFoundError:
            pass


def _normalize(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


def _chunk_paragraphs(paragraphs: list[str], target_words: int = WORDS_PER_CHUNK):
    chunks: list[list[str]] = []
    current: list[str] = []
    words = 0
    for para in paragraphs:
        w = len(para.split())
        if current and words + w > target_words:
            chunks.append(current)
            current, words = [], 0
        current.append(para)
        words += w
    if current:
        chunks.append(current)
    return chunks


def _paragraph_chunk(client, chunk_text: str) -> str:
    response = client.messages.create(
        model=MODEL,
        max_tokens=8000,
        system=PARAGRAPHING_SYSTEM_PROMPT,
        messages=[{"role": "user", "content": chunk_text}],
    )
    return "".join(
        part.text for part in response.content if getattr(part, "type", "") == "text"
    )


def _call_with_retry(client, chunk_text: str) -> Optional[str]:
    try:
        return _paragraph_chunk(client, chunk_text)
    except Exception:  # noqa: BLE001
        time.sleep(2)
        try:
            return _paragraph_chunk(client, chunk_text)
        except Exception:  # noqa: BLE001
            return None


def _split_into_paragraphs(text: str) -> list[str]:
    out: list[str] = []
    for block in re.split(r"\n\s*\n+", text):
        para = re.sub(r"\s*\n\s*", " ", block)
        para = re.sub(r"[ \t]+", " ", para).strip()
        if para:
            out.append(para)
    return out


def llm_paragraph_blocks(
    body_blocks: list[str],
    client,
    checkpoint: FormatCheckpoint,
    progress=None,
) -> list[list[str]]:
    """Re-paragraph each body block via the API, verifying each chunk.

    A chunk whose text changed (beyond whitespace) is discarded in favour of the
    mechanical paragraphing of that chunk, so the step can only re-break text.
    """
    per_block_chunks = [_chunk_paragraphs(mechanical_paragraphs(b)) for b in body_blocks]
    total = sum(len(c) for c in per_block_chunks)
    checkpoint.ensure_total(total)

    results: list[list[str]] = []
    idx = 0
    for chunks in per_block_chunks:
        processed: list[str] = []
        for chunk in chunks:
            chunk_text = "\n\n".join(chunk)
            if progress is not None:
                progress(idx + 1, total)
            if idx in checkpoint.outputs:
                processed.append(checkpoint.outputs[idx])
            else:
                returned = _call_with_retry(client, chunk_text)
                if returned is not None and _normalize(returned) == _normalize(chunk_text):
                    out = returned.strip()
                else:
                    out = chunk_text
                checkpoint.record(idx, out)
                processed.append(out)
            idx += 1
        combined = "\n\n".join(processed)
        results.append(_glue_orphans(_split_into_paragraphs(combined)))
    return results


# ---------------------------------------------------------------------------
# Step 8 — the styled reference document
# ---------------------------------------------------------------------------


def _add_page_number_field(paragraph) -> None:
    run = paragraph.add_run()
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = "PAGE"
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    run._r.append(begin)
    run._r.append(instr)
    run._r.append(end)


def _set_doc_default_font(document, name: str, size_pt: float) -> None:
    styles_el = document.styles.element
    path = "/".join((qn("w:docDefaults"), qn("w:rPrDefault"), qn("w:rPr")))
    rpr = styles_el.find(path)
    if rpr is None:
        return
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.insert(0, rfonts)
    for attr in ("w:ascii", "w:hAnsi", "w:cs"):
        rfonts.set(qn(attr), name)
    half = str(int(size_pt * 2))
    for tag in ("w:sz", "w:szCs"):
        el = rpr.find(qn(tag))
        if el is None:
            el = OxmlElement(tag)
            rpr.append(el)
        el.set(qn("w:val"), half)


def _style_named(document, name: str):
    """Find a style by its exact name. (python-docx's ``styles[name]`` maps
    "Heading N" to the internal "heading N", which misses pandoc's
    capitalised "Heading N" styles.)"""
    for style in document.styles:
        if style.name == name:
            return style
    raise KeyError(name)


def _clear_theme_overrides(style) -> None:
    """Drop theme font/colour attributes, which Word applies in preference to the
    explicit font name and RGB colour we set."""
    rpr = style.element.find(qn("w:rPr"))
    if rpr is None:
        return
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is not None:
        for attr in ("w:asciiTheme", "w:hAnsiTheme", "w:cstheme", "w:eastAsiaTheme"):
            rfonts.attrib.pop(qn(attr), None)
    color = rpr.find(qn("w:color"))
    if color is not None:
        for attr in ("w:themeColor", "w:themeShade", "w:themeTint"):
            color.attrib.pop(qn(attr), None)


def build_reference_doc(ref_path: Path) -> None:
    """Maintainer helper: regenerate ``translate_doc/assets/styled.docx``.

    Not used at runtime — Format mode always loads the bundled copy. Run:
        python -c "from pathlib import Path; from translate_doc.formatter import build_reference_doc; build_reference_doc(Path('translate_doc/assets/styled.docx'))"
    """
    with open(ref_path, "wb") as fh:
        subprocess.run(
            ["pandoc", "--print-default-data-file", "reference.docx"],
            stdout=fh,
            check=True,
        )
    document = Document(str(ref_path))
    _set_doc_default_font(document, "Times New Roman", 12)

    # Body text: 12pt TNR, justified, line 312 (1.3) auto, after 160, indent 432.
    for style_name in ("Normal", "Body Text", "First Paragraph", "Compact"):
        try:
            style = document.styles[style_name]
        except KeyError:
            continue
        style.font.name = "Times New Roman"
        style.font.size = Pt(12)
        pf = style.paragraph_format
        pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        pf.line_spacing = 1.3
        pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
        pf.space_after = Pt(8)
        pf.first_line_indent = Inches(0.3)

    # Footnotes: ~9.5pt (sz 19), compact.
    for style_name in ("Footnote Text", "Footnote"):
        try:
            style = document.styles[style_name]
        except KeyError:
            continue
        style.font.name = "Times New Roman"
        style.font.size = Pt(9.5)
        pf = style.paragraph_format
        pf.space_after = Pt(0)
        pf.line_spacing = 1.0

    # Title (Heading 1) and section titles (Heading 2): serif, not the theme font.
    try:
        h1 = _style_named(document, "Heading 1")
        h1.font.name = "Times New Roman"
        _clear_theme_overrides(h1)
    except KeyError:
        pass
    try:
        h2 = _style_named(document, "Heading 2")
        h2.font.name = "Times New Roman"
        h2.font.size = Pt(14)
        h2.font.bold = True
        h2.font.color.rgb = RGBColor(0x1F, 0x3A, 0x5F)
        h2.paragraph_format.space_before = Pt(12)
        h2.paragraph_format.space_after = Pt(6)
        h2.paragraph_format.keep_with_next = True
        _clear_theme_overrides(h2)
    except KeyError:
        pass

    for section in document.sections:
        section.page_width = Inches(8.5)
        section.page_height = Inches(11)
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)
        footer_para = section.footer.paragraphs[0]
        footer_para.text = ""
        footer_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _add_page_number_field(footer_para)

    document.save(str(ref_path))


def run_pandoc(md_path: Path, output_path: Path, ref_path: Path) -> None:
    subprocess.run(
        [
            "pandoc",
            str(md_path),
            "-f",
            "markdown",
            "-o",
            str(output_path),
            f"--reference-doc={ref_path}",
        ],
        check=True,
    )


# ---------------------------------------------------------------------------
# Orchestration
# ---------------------------------------------------------------------------


def _derive_title(input_path: Path) -> str:
    stem = re.sub(r"_translated$", "", input_path.stem, flags=re.IGNORECASE)
    stem = stem.replace("_", " ").replace("-", " ")
    return re.sub(r"\s+", " ", stem).strip() or input_path.stem


@dataclass
class _Block:
    kind: str  # "heading" or "body"
    text: str = ""
    paragraphs: Optional[list[str]] = None


def _assemble_blocks(
    segments: list[Segment],
    header_norms: set[str],
    section_norms: set[str],
) -> tuple[list[_Block], int, int]:
    """Turn segments into an ordered block list, promoting section titles to
    headings and merging consecutive body prose. Returns (blocks, headers_hit,
    headings_emitted)."""
    blocks: list[_Block] = []
    seen_sections: set[str] = set()
    headers_hit = 0
    headings_emitted = 0

    for seg in segments:
        text = seg.text
        phrase = _leading_phrase(text)
        norm = _normalize_phrase(phrase)

        if norm and norm in header_norms:
            # Strip the running header (and its trailing page no. / flag) at the start.
            text = _header_removal_regex(norm).sub(" ", text, count=1)
            headers_hit += 1
        elif norm and norm in section_norms:
            if norm not in seen_sections:
                seen_sections.add(norm)
                blocks.append(_Block(kind="heading", text=phrase.rstrip(":.")))
                headings_emitted += 1
                # Remove the title text from the start of the body.
                text = _header_removal_regex(norm).sub(" ", text, count=1)
            else:
                text = _header_removal_regex(norm).sub(" ", text, count=1)

        # Remove any remaining header occurrences anywhere in this segment.
        for hnorm in header_norms:
            text = _header_removal_regex(hnorm).sub(" ", text)

        cleaned = _clean_segment_text(text)
        if not cleaned:
            continue
        if blocks and blocks[-1].kind == "body":
            blocks[-1].text = f"{blocks[-1].text} {cleaned}".strip()
        else:
            blocks.append(_Block(kind="body", text=cleaned))

    return blocks, headers_hit, headings_emitted


def _build_markdown(title: str, blocks: list[_Block], defs: list[str]) -> str:
    parts = [f"# {title}", "", "*Translated from Arabic*", "", ""]
    for block in blocks:
        if block.kind == "heading":
            parts.extend([f"## {block.text}", ""])
        else:
            for para in block.paragraphs or []:
                parts.extend([para, ""])
    parts.append("")
    if defs:
        parts.append("\n\n".join(defs))
        parts.append("")
    return "\n".join(parts)


def format_document(
    input_path: Path,
    client=None,
    restart: bool = False,
    progress=None,
) -> FormatResult:
    """Full Format-mode pipeline (pandoc round-trip). Pandoc availability is
    assumed to have been checked by the caller."""
    _require_pandoc()
    input_path = input_path.resolve()
    output_path = input_path.parent / f"{input_path.stem}_formatted.docx"
    if restart:
        FormatCheckpoint.path_for(input_path).unlink(missing_ok=True)

    # 0/1 — export and set footnote defs aside.
    md = export_markdown(input_path)
    body, defs = extract_footnote_defs(md)

    # 2 — segment and drop cover / front matter.
    segments = split_into_segments(body)
    start = find_content_start(segments)
    cover_dropped = sum(1 for s in segments[:start])
    segments = segments[start:]

    # 4/5 — running headers (by recurrence) and section titles.
    lead_phrases = [_leading_phrase(s.text) for s in segments]
    norms = [_normalize_phrase(p) for p in lead_phrases]
    freq = Counter(n for n in norms if n)
    representative: dict[str, str] = {}
    for phrase, norm in zip(lead_phrases, norms):
        representative.setdefault(norm, phrase)
    threshold = max(8, math.ceil(0.05 * len(segments))) if segments else 8
    header_norms = {n for n, c in freq.items() if c >= threshold}
    section_norms = {
        n
        for n, c in freq.items()
        if n not in header_norms
        and c <= 3
        and len(n.split()) <= 10
        and _is_title_like(representative[n])  # test the ORIGINAL-cased phrase
    }

    # 3/6 — assemble blocks (strips headers, promotes titles, cleans prose).
    blocks, headers_hit, headings_emitted = _assemble_blocks(
        segments, header_norms, section_norms
    )

    # 7 — paragraphing. LLM preferred; mechanical fallback otherwise.
    body_blocks = [b for b in blocks if b.kind == "body"]
    paragraphing_ran = False
    if client is not None and body_blocks:
        checkpoint = None if restart else FormatCheckpoint.load(input_path)
        if checkpoint is None:
            checkpoint = FormatCheckpoint.start(input_path)
        paragraphed = llm_paragraph_blocks(
            [b.text for b in body_blocks], client, checkpoint, progress=progress
        )
        checkpoint.delete()
        paragraphing_ran = True
    else:
        paragraphed = [mechanical_paragraphs(b.text) for b in body_blocks]

    it = iter(paragraphed)
    for b in blocks:
        if b.kind == "body":
            b.paragraphs = next(it)

    total_paragraphs = sum(len(b.paragraphs or []) for b in blocks if b.kind == "body")

    # 3b/8 — rebuild with pandoc against the styled reference doc.
    title = _derive_title(input_path)
    work_md = _build_markdown(title, blocks, defs)
    md_path = input_path.parent / f"{input_path.stem}_formatted.md"
    md_path.write_text(work_md, encoding="utf-8")
    try:
        with bundled_reference_doc() as ref_path:
            reference_doc = str(Path(ref_path).resolve())
            run_pandoc(md_path, output_path, Path(ref_path))
    finally:
        md_path.unlink(missing_ok=True)

    return FormatResult(
        output_path=output_path,
        cover_pages_dropped=cover_dropped,
        running_headers_removed=len(header_norms),
        section_headings=headings_emitted,
        paragraphs_produced=total_paragraphs,
        footnotes_preserved=len(defs),
        paragraphing_ran=paragraphing_ran,
        reference_doc=reference_doc,
    )

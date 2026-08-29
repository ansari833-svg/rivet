"""Claude API translation with chunking, retry and checkpoint support.

The translator works on :class:`~translate_doc.parser.Block` objects. Blocks are
grouped into chunks; each chunk is one API call. After every successful chunk a
checkpoint file is written so an interrupted run can resume.

Difficult passages are flagged inline by the model using the tag

    [DIFFICULT: original Arabic | reason]

which is preserved verbatim in the returned translations. Downstream code
(builder / logger) is responsible for extracting and rendering those tags.
"""

from __future__ import annotations

import json
import re
import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Optional

from anthropic import Anthropic

from .parser import Block

# Use this exact model string.
MODEL = "claude-sonnet-5"

# Maximum number of blocks sent in a single API call. Reduce this if responses
# hit the output token limit.
MAX_PARAGRAPHS_PER_CHUNK = 40

# Generous output cap; Arabic->English rarely exceeds the source length by much.
MAX_OUTPUT_TOKENS = 16000

# Marker used to delimit blocks in both directions.
_BLOCK_RE = re.compile(r"\[\[BLOCK\s+(\d+)\]\]", re.IGNORECASE)

# Matches an inline difficulty flag: [DIFFICULT: original | reason]
DIFFICULT_RE = re.compile(r"\[DIFFICULT:\s*(.*?)\s*\|\s*(.*?)\]", re.DOTALL)

SYSTEM_PROMPT = """\
You are an expert Arabic-to-English translator specializing in formal, legal, \
medical, and governmental documents. You produce authentic, literal \
translations that prioritize fidelity to the source over readability.

Follow these rules exactly:

1. Produce a literal, faithful translation from Arabic to English. Do NOT \
paraphrase, summarize, simplify, or "improve" the text. Preserve the author's \
sentence and paragraph structure as closely as English grammar allows.

2. Preserve formal and technical terminology as distinct from everyday meaning. \
When a word carries a specific legal, medical, administrative, or governmental \
sense, translate it with that technical meaning. If the technical term and the \
literal meaning diverge in a way a reader should know about, flag it (see rule 4).

3. Do not omit anything. Every sentence in the source must appear in the \
translation.

4. Flag any phrase, idiom, or term that is ambiguous, culturally specific, or \
difficult to translate literally by inserting an inline tag exactly in this form:

   [DIFFICULT: <original Arabic text> | <short reason / alternative reading>]

   Place the tag immediately after the English rendering of the difficult \
phrase. Keep the original Arabic inside the tag. Use the tag sparingly, only \
for genuinely difficult or ambiguous material.

5. You will be given the source text split into numbered blocks, each \
introduced by a marker line of the form [[BLOCK n]]. Translate each block \
independently and return your output using the SAME markers, in the SAME order, \
one marker per block, followed by the English translation of that block. \
Return the marker even for blocks that are empty or contain only numbers or \
punctuation (echo the content as-is in that case).

Return ONLY the block markers and their translations. Do not add commentary, \
preamble, or a summary.
"""


@dataclass
class TranslationResult:
    """Result of translating a set of blocks."""

    translations: dict[int, str] = field(default_factory=dict)
    missing_ids: list[int] = field(default_factory=list)


class TranslationError(RuntimeError):
    pass


def chunk_blocks(block_ids: list[int], size: int = MAX_PARAGRAPHS_PER_CHUNK) -> list[list[int]]:
    """Split an ordered list of block ids into chunks of at most ``size``."""
    return [block_ids[i : i + size] for i in range(0, len(block_ids), size)]


def _build_user_message(chunk_ids: list[int], blocks: dict[int, Block]) -> str:
    parts: list[str] = [
        "Translate the following Arabic blocks into English following the "
        "system instructions. Preserve the [[BLOCK n]] markers exactly.\n",
    ]
    for bid in chunk_ids:
        text = blocks[bid].text
        parts.append(f"[[BLOCK {bid}]]\n{text}")
    return "\n\n".join(parts)


def _parse_response(text: str) -> dict[int, str]:
    """Split a model response back into a {block_id: translated_text} map."""
    out: dict[int, str] = {}
    matches = list(_BLOCK_RE.finditer(text))
    for i, m in enumerate(matches):
        bid = int(m.group(1))
        start = m.end()
        end = matches[i + 1].start() if i + 1 < len(matches) else len(text)
        out[bid] = text[start:end].strip()
    return out


class Translator:
    def __init__(
        self,
        client: Optional[Anthropic] = None,
        model: str = MODEL,
        max_paragraphs_per_chunk: int = MAX_PARAGRAPHS_PER_CHUNK,
    ) -> None:
        self.client = client or Anthropic()
        self.model = model
        self.max_paragraphs_per_chunk = max_paragraphs_per_chunk

    def _call_once(self, user_message: str) -> str:
        response = self.client.messages.create(
            model=self.model,
            max_tokens=MAX_OUTPUT_TOKENS,
            system=SYSTEM_PROMPT,
            messages=[{"role": "user", "content": user_message}],
        )
        return "".join(
            part.text for part in response.content if getattr(part, "type", "") == "text"
        )

    def _call_with_retry(self, user_message: str) -> str:
        """Call the API, retrying once with exponential backoff on failure."""
        try:
            return self._call_once(user_message)
        except Exception as first_error:  # noqa: BLE001 - surface after retry
            time.sleep(2)
            try:
                return self._call_once(user_message)
            except Exception as second_error:  # noqa: BLE001
                raise TranslationError(
                    f"API call failed twice: {first_error!r} / {second_error!r}"
                ) from second_error

    def translate(
        self,
        block_ids: list[int],
        blocks: dict[int, Block],
        checkpoint: "Checkpoint",
        progress: Optional[Callable[[int, int], None]] = None,
    ) -> TranslationResult:
        """Translate the given block ids, chunk by chunk, checkpointing as we go."""
        chunks = chunk_blocks(block_ids, self.max_paragraphs_per_chunk)
        checkpoint.ensure_total(len(chunks))

        result = TranslationResult()

        for chunk_index, chunk_ids in enumerate(chunks):
            if progress is not None:
                progress(chunk_index + 1, len(chunks))

            if checkpoint.is_done(chunk_index):
                result.translations.update(checkpoint.chunk_translations(chunk_index))
                continue

            # Skip blocks that are entirely empty of translatable content.
            translatable = [bid for bid in chunk_ids if blocks[bid].text.strip()]
            empty = [bid for bid in chunk_ids if not blocks[bid].text.strip()]

            chunk_out: dict[int, str] = {bid: "" for bid in empty}

            if translatable:
                user_message = _build_user_message(translatable, blocks)
                raw = self._call_with_retry(user_message)
                parsed = _parse_response(raw)
                for bid in translatable:
                    if bid in parsed:
                        chunk_out[bid] = parsed[bid]
                    else:
                        result.missing_ids.append(bid)

            result.translations.update(chunk_out)
            checkpoint.record_chunk(chunk_index, chunk_out)

        return result


class Checkpoint:
    """JSON checkpoint stored next to the input file.

    Layout::

        {
          "original_filename": "contract.docx",
          "total_chunks": 4,
          "completed": [0, 1],
          "translations": {"0": {"12": "..."}, "1": {"40": "..."}}
        }
    """

    def __init__(self, path: Path, original_filename: str) -> None:
        self.path = path
        self.original_filename = original_filename
        self.total_chunks: Optional[int] = None
        self.completed: set[int] = set()
        self._translations: dict[int, dict[int, str]] = {}

    # ---- persistence -------------------------------------------------

    @classmethod
    def checkpoint_path_for(cls, input_path: Path) -> Path:
        return input_path.parent / f"{input_path.stem}_checkpoint.json"

    @classmethod
    def load(cls, input_path: Path) -> Optional["Checkpoint"]:
        path = cls.checkpoint_path_for(input_path)
        if not path.exists():
            return None
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
        except (json.JSONDecodeError, OSError):
            return None
        cp = cls(path, data.get("original_filename", input_path.name))
        cp.total_chunks = data.get("total_chunks")
        cp.completed = set(data.get("completed", []))
        cp._translations = {
            int(k): {int(bk): bv for bk, bv in v.items()}
            for k, v in data.get("translations", {}).items()
        }
        return cp

    @classmethod
    def start(cls, input_path: Path) -> "Checkpoint":
        return cls(cls.checkpoint_path_for(input_path), input_path.name)

    def _save(self) -> None:
        payload = {
            "original_filename": self.original_filename,
            "total_chunks": self.total_chunks,
            "completed": sorted(self.completed),
            "translations": {
                str(cidx): {str(bid): text for bid, text in blocks.items()}
                for cidx, blocks in self._translations.items()
            },
        }
        # Explicit utf-8 so Arabic content survives on Windows (cp1252 default).
        self.path.write_text(
            json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8"
        )

    def delete(self) -> None:
        try:
            self.path.unlink()
        except FileNotFoundError:
            pass

    # ---- state -------------------------------------------------------

    def ensure_total(self, total: int) -> None:
        if self.total_chunks is None:
            self.total_chunks = total
            self._save()

    def is_done(self, chunk_index: int) -> bool:
        return chunk_index in self.completed

    def chunk_translations(self, chunk_index: int) -> dict[int, str]:
        return dict(self._translations.get(chunk_index, {}))

    def record_chunk(self, chunk_index: int, translations: dict[int, str]) -> None:
        self._translations[chunk_index] = dict(translations)
        self.completed.add(chunk_index)
        self._save()

    def all_translations(self) -> dict[int, str]:
        merged: dict[int, str] = {}
        for blocks in self._translations.values():
            merged.update(blocks)
        return merged

    def remaining_summary(self) -> Optional[str]:
        """Human-readable 'chunks a-b remaining' string, or None if nothing done."""
        if not self.completed or self.total_chunks is None:
            return None
        remaining = [i for i in range(self.total_chunks) if i not in self.completed]
        if not remaining:
            return None
        # Present as 1-based inclusive range for the user.
        return f"chunks {remaining[0] + 1}-{remaining[-1] + 1} remaining"


# ---- difficulty extraction ------------------------------------------


@dataclass
class DifficultFlag:
    block_id: int
    original_arabic: str
    reason: str


def extract_difficult_flags(block_id: int, translated: str) -> list[DifficultFlag]:
    """Pull all [DIFFICULT: ...] flags out of a translated block."""
    flags: list[DifficultFlag] = []
    for m in DIFFICULT_RE.finditer(translated):
        flags.append(
            DifficultFlag(
                block_id=block_id,
                original_arabic=m.group(1).strip(),
                reason=m.group(2).strip(),
            )
        )
    return flags


def strip_difficult_tags(translated: str) -> str:
    """Remove [DIFFICULT: ...] tags, keeping the surrounding English text."""
    return DIFFICULT_RE.sub("", translated).replace("  ", " ").strip()

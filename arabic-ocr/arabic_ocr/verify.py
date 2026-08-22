"""Pass 2 — Claude verification.

Claude receives the page image AND Surya's transcription and reports ONLY where
they disagree. Claude does not produce its own transcription and does not
rewrite Surya's output. A language model reading classical Arabic will silently
normalize archaic spellings, regularize unusual readings, and complete damaged
words from memory rather than from the page — those errors are invisible in the
output. So Surya writes, Claude flags, and a human decides. This inversion must
never be reversed: Claude output is never used as the text when the two disagree.

A failed verification pass must never lose the Surya transcription. On any
failure the page is recorded as ``verification_failed`` and the run continues.
"""

from __future__ import annotations

import base64
import json
import time
from dataclasses import dataclass, field
from pathlib import Path

MODEL = "claude-sonnet-5"
MAX_TOKENS = 4096
RETRY_BACKOFF_SECONDS = 2.0

SYSTEM_PROMPT = """\
You are a transcription verifier for scanned pages of classical Arabic printed \
books (fiqh, hadith, tafsir, balāgha) in modern critical editions. You are given \
a page image and an existing transcription of that page produced by an OCR \
engine. Your ONLY job is to compare the transcription against the image and \
report where they disagree.

Absolute rules:
- Do NOT produce your own transcription. Do NOT restate text that is correct.
- Report ONLY discrepancies between the supplied transcription and the image.
- Never suggest a reading based on what the text "should" say, on your \
familiarity with the work, or on grammatical expectation. Report ONLY what is \
visibly on the page.
- If a word is damaged, faded, or unclear, say so explicitly. Do NOT supply the \
expected reading.
- If the image (or a region of it) is too degraded to judge, state that plainly \
instead of guessing.
- Treat tashkīl (vocalization marks) differences as their own discrepancy \
category. They are easy to overlook and matter for this material.

For each discrepancy provide:
- context: a short snippet of surrounding text so the location is findable
- transcribed: what the supplied transcription says at that point
- image_shows: what the image appears to show (or a statement that it is \
unclear/damaged/illegible)
- category: one of "text", "tashkil", "missing", "extra", "illegible"
- confidence: one of "high", "medium", "low"

Respond with a single JSON object and nothing else. No preamble, no explanation, \
no markdown code fences. The object must have exactly this shape:

{"discrepancies": [{"context": "...", "transcribed": "...", "image_shows": "...", \
"category": "...", "confidence": "..."}]}

If there are no discrepancies, return {"discrepancies": []}.\
"""

USER_INSTRUCTION = (
    "Here is the page image and the OCR transcription to verify. Compare them "
    "character by character and report only the discrepancies as JSON."
)


@dataclass
class Discrepancy:
    context: str = ""
    transcribed: str = ""
    image_shows: str = ""
    category: str = ""
    confidence: str = ""


@dataclass
class VerificationResult:
    status: str = "verified"  # "verified" | "verification_failed"
    discrepancies: list[Discrepancy] = field(default_factory=list)


_MEDIA_TYPES = {
    ".jpg": "image/jpeg",
    ".jpeg": "image/jpeg",
    ".png": "image/png",
}


class ClaudeVerifier:
    """Wraps a single-page Claude verification call."""

    def __init__(self) -> None:
        # Imported lazily so --offline runs never require the SDK or a key.
        import anthropic

        self._client = anthropic.Anthropic()

    def verify_page(self, image_path: Path, transcription: str) -> VerificationResult:
        """Verify one page. Retries once with backoff; never raises."""
        media_type = _MEDIA_TYPES.get(image_path.suffix.lower(), "image/jpeg")
        image_b64 = base64.standard_b64encode(image_path.read_bytes()).decode("ascii")

        transcription = transcription if transcription.strip() else "(empty transcription)"

        for attempt in range(2):
            try:
                response = self._client.messages.create(
                    model=MODEL,
                    max_tokens=MAX_TOKENS,
                    system=SYSTEM_PROMPT,
                    messages=[
                        {
                            "role": "user",
                            "content": [
                                {
                                    "type": "image",
                                    "source": {
                                        "type": "base64",
                                        "media_type": media_type,
                                        "data": image_b64,
                                    },
                                },
                                {
                                    "type": "text",
                                    "text": (
                                        USER_INSTRUCTION
                                        + "\n\n=== TRANSCRIPTION START ===\n"
                                        + transcription
                                        + "\n=== TRANSCRIPTION END ==="
                                    ),
                                },
                            ],
                        }
                    ],
                )
            except Exception:
                if attempt == 0:
                    time.sleep(RETRY_BACKOFF_SECONDS)
                    continue
                return VerificationResult(status="verification_failed")

            text = _response_text(response)
            parsed = _parse_discrepancies(text)
            if parsed is None:
                # Parse failure. Record the page as failed and move on rather
                # than aborting the run.
                return VerificationResult(status="verification_failed")
            return VerificationResult(status="verified", discrepancies=parsed)

        return VerificationResult(status="verification_failed")


def _response_text(response) -> str:
    parts = []
    for block in getattr(response, "content", []) or []:
        if getattr(block, "type", None) == "text":
            parts.append(getattr(block, "text", ""))
    return "".join(parts)


def _strip_fences(text: str) -> str:
    """Remove stray markdown code fences the model may emit despite instructions."""
    text = text.strip()
    if text.startswith("```"):
        # Drop the opening fence line (```json or ```) and the closing fence.
        newline = text.find("\n")
        if newline != -1:
            text = text[newline + 1 :]
        if text.rstrip().endswith("```"):
            text = text.rstrip()[: -len("```")]
    return text.strip()


def _parse_discrepancies(text: str) -> list[Discrepancy] | None:
    """Defensively parse the model's JSON. Returns None on unrecoverable failure."""
    cleaned = _strip_fences(text)
    if not cleaned:
        return None
    try:
        data = json.loads(cleaned)
    except json.JSONDecodeError:
        # Last resort: try to isolate the outermost JSON object.
        start = cleaned.find("{")
        end = cleaned.rfind("}")
        if start == -1 or end == -1 or end <= start:
            return None
        try:
            data = json.loads(cleaned[start : end + 1])
        except json.JSONDecodeError:
            return None

    if not isinstance(data, dict):
        return None
    raw = data.get("discrepancies", [])
    if not isinstance(raw, list):
        return None

    out = []
    for item in raw:
        if not isinstance(item, dict):
            continue
        out.append(
            Discrepancy(
                context=str(item.get("context", "")),
                transcribed=str(item.get("transcribed", "")),
                image_shows=str(item.get("image_shows", "")),
                category=str(item.get("category", "")),
                confidence=str(item.get("confidence", "")).lower(),
            )
        )
    return out

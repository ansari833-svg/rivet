"""arabic-ocr: two-pass Arabic page transcription with human-in-the-loop review.

Pass 1 (Surya, local, deterministic) is the authoritative text.
Pass 2 (Claude) only flags discrepancies; it never rewrites Surya's output.
"""

__version__ = "0.1.0"

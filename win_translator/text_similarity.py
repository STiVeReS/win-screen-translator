from __future__ import annotations

import re
from difflib import SequenceMatcher


_WS_RE = re.compile(r"\s+", re.UNICODE)


def normalize_text(s: str) -> str:
    """Нормалізація для порівняння: нижній регістр, схлопування пробілів, прибрати зайве."""
    if not s:
        return ""

    s2 = s.strip().lower()
    s2 = _WS_RE.sub(" ", s2)

    # трохи прибираємо типові OCR-сміття, але обережно
    s2 = s2.replace("\u200b", "")  # zero-width
    s2 = s2.strip()

    return s2


def similarity_ratio(a: str, b: str) -> float:
    """0..1"""
    a2 = normalize_text(a)
    b2 = normalize_text(b)

    if not a2 and not b2:
        return 1.0
    if not a2 or not b2:
        return 0.0
    if a2 == b2:
        return 1.0

    return float(SequenceMatcher(None, a2, b2).ratio())


def is_same_or_similar(new_text: str, prev_text: str, threshold: float = 0.9) -> bool:
    th = float(threshold or 0.0)
    if th < 0.0:
        th = 0.0
    if th > 1.0:
        th = 1.0

    r = similarity_ratio(new_text, prev_text)
    return r >= th
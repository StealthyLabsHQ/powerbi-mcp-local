"""Accent inference from measure-name semantics.

The styling tool picks an accent colour for each card from the preset's
``accentMap`` based on the bound measure's name. The rules below are
heuristic — matched case-insensitively against substrings of the
measure name. ``custom_spec`` can override by passing
``accent_overrides: {measure_name: accent_key}``.
"""

from __future__ import annotations

import re

# Substrings → accent key. The first matching rule wins.
ACCENT_RULES: list[tuple[re.Pattern[str], str]] = [
    (
        re.compile(
            r"croissance|growth|marge brute|gross\s*margin|ebe|ebit|marge nette|net\s*margin|profit", re.IGNORECASE
        ),
        "positive",
    ),
    (re.compile(r"endettement|debt|leverage|bfr|wcr|charge|expense|frais|cost", re.IGNORECASE), "warning"),
    (re.compile(r"\bvar\b|variance|geo|atelier|workshop|store", re.IGNORECASE), "info"),
]


def infer_accent_key(measure_name: str | None) -> str:
    """Return one of ``"positive" / "warning" / "info" / "neutral"``.

    Defaults to ``"neutral"`` when the name is missing or matches none
    of the rules.
    """
    if not measure_name:
        return "neutral"
    text = str(measure_name)
    for regex, accent_key in ACCENT_RULES:
        if regex.search(text):
            return accent_key
    return "neutral"


def pick_accent(measure_name: str | None, accent_map: dict[str, str]) -> str:
    """Return the hex colour from ``accent_map`` for the inferred accent
    key. Falls back through ``neutral`` → ``info`` → first value.
    """
    key = infer_accent_key(measure_name)
    if key in accent_map:
        return accent_map[key]
    if "neutral" in accent_map:
        return accent_map["neutral"]
    if "info" in accent_map:
        return accent_map["info"]
    return next(iter(accent_map.values()))


__all__ = ["ACCENT_RULES", "infer_accent_key", "pick_accent"]

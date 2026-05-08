"""Power BI visual literal + format-property encoders.

Builds the JSON shapes that Power BI's report engine expects for typed
literals (text, decimal, int, bool, color) used in visual ``objects``
property bags.
"""

from __future__ import annotations

import json
from typing import Any

from pbi_connection import PowerBIValidationError

from ._base import HEX_COLOR_RE

_VISUAL_FORMAT_TYPES = frozenset({"auto", "text", "bool", "int", "decimal", "color", "raw"})


def _literal_value(value: Any) -> dict[str, Any]:
    return {"expr": {"Literal": {"Value": json.dumps(value)}}}


def _decimal_literal(value: float) -> dict[str, Any]:
    """Power BI numeric literal (Decimal). Uses 'D' suffix expected by the report engine."""
    return {"expr": {"Literal": {"Value": f"{float(value)}D"}}}


def _int_literal(value: int) -> dict[str, Any]:
    return {"expr": {"Literal": {"Value": f"{int(value)}L"}}}


def _text_literal(value: str) -> dict[str, Any]:
    """Power BI canonical text literal: 'value' with embedded quotes doubled.

    PBI's Literal.Value field for text uses single-quoted form (the older
    json.dumps-derived '"…"' style is silently ignored by some visual
    serializers, which is why titles set that way never render).
    """
    escaped = str(value).replace("'", "''")
    return {"expr": {"Literal": {"Value": f"'{escaped}'"}}}


def _solid_color(color: str) -> dict[str, Any]:
    if not HEX_COLOR_RE.match(color):
        raise PowerBIValidationError(
            "color must match '#RRGGBB'.",
            details={"value": color},
        )
    return {"solid": {"color": {"expr": {"Literal": {"Value": f"'{color}'"}}}}}


def _gauge_axis_objects(
    min_value: float | None, max_value: float | None, target_value: float | None
) -> list[dict[str, Any]]:
    properties: dict[str, Any] = {}
    if min_value is not None:
        properties["min"] = _decimal_literal(min_value)
    if max_value is not None:
        properties["max"] = _decimal_literal(max_value)
    if target_value is not None:
        properties["target"] = _decimal_literal(target_value)
    if not properties:
        return []
    return [{"properties": properties}]


def _encode_visual_format_value(value: Any, hint: str | None = None) -> Any:
    """Encode a Python value as a Power BI visual format property.

    ``hint`` (optional, one of ``text``, ``bool``, ``int``, ``decimal``,
    ``color``, ``raw``) selects the literal form. ``auto`` (default) infers
    from the Python type. ``raw`` returns the value untouched so callers can
    pass an already-shaped dict (e.g. a measure binding).
    """
    if hint is not None and hint not in _VISUAL_FORMAT_TYPES:
        raise PowerBIValidationError(
            f"unknown property type hint '{hint}'.",
            details={"hint": hint, "allowed": sorted(_VISUAL_FORMAT_TYPES)},
        )
    if hint == "raw":
        return value
    if hint == "color":
        if not isinstance(value, str):
            raise PowerBIValidationError("color values must be strings.", details={"value": repr(value)})
        return _solid_color(value)
    if hint == "text":
        return _text_literal(str(value))
    if hint == "bool":
        return _literal_value(bool(value))
    if hint == "int":
        return _int_literal(int(value))
    if hint == "decimal":
        return _decimal_literal(float(value))
    # auto
    if isinstance(value, bool):
        return _literal_value(value)
    if isinstance(value, int):
        return _int_literal(value)
    if isinstance(value, float):
        return _decimal_literal(value)
    if isinstance(value, str):
        if HEX_COLOR_RE.match(value):
            return _solid_color(value)
        return _text_literal(value)
    if isinstance(value, dict):
        # Allow callers to pass already-shaped expr/Measure/etc. payloads
        return value
    raise PowerBIValidationError(
        f"cannot encode value of type {type(value).__name__} for visual format property.",
        details={"value": repr(value)},
    )


def _datapoint_fill_objects(fill_color: str | None, target_color: str | None) -> list[dict[str, Any]]:
    properties: dict[str, Any] = {}
    if fill_color is not None:
        properties["fill"] = _solid_color(fill_color)
    if target_color is not None:
        properties["targetFill"] = _solid_color(target_color)
    if not properties:
        return []
    return [{"properties": properties}]


def _title_objects(title: str) -> dict[str, Any]:
    return {
        "title": [
            {
                "properties": {
                    "show": _literal_value(True),
                    "text": _literal_value(title),
                }
            }
        ]
    }

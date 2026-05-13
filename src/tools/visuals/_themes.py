"""Power BI report-theme JSON validation.

Used by ``pbi_apply_theme_tool``, ``pbi_validate_theme_tool``, and
``pbi_export_active_theme_tool``. v0.13 hardening: user-supplied theme
JSON is checked against an explicit schema before it is copied into the
report's ``StaticResources/Themes`` folder. Unknown top-level keys,
oversized payloads, malformed colours, and obvious URL-bearing values
are rejected up-front so a crafted theme cannot be used as a smuggling
vector for non-theme content.
"""

from __future__ import annotations

import re
from typing import Any

from pbi_connection import PowerBIValidationError

MAX_THEME_BYTES = 256 * 1024

THEME_ALLOWED_TOP_LEVEL_KEYS: set[str] = {
    "name",
    "dataColors",
    "foreground",
    "foregroundNeutralSecondary",
    "foregroundNeutralTertiary",
    "background",
    "backgroundLight",
    "backgroundNeutral",
    "tableAccent",
    "good",
    "neutral",
    "bad",
    "maximum",
    "center",
    "minimum",
    "null",
    "hyperlink",
    "visitedHyperlink",
    "textClasses",
    "visualStyles",
}

_HEX6_RE = re.compile(r"^#[0-9A-Fa-f]{6}$")
_HEX8_RE = re.compile(r"^#[0-9A-Fa-f]{8}$")
# Strings that resemble URLs or external references that have no business
# living inside a theme JSON. Themes describe colours and typography, not
# behaviour; if a string value looks like a URL we treat it as smuggling.
_FORBIDDEN_VALUE_PATTERNS = [
    re.compile(r"^\s*javascript:", re.IGNORECASE),
    re.compile(r"^\s*data:", re.IGNORECASE),
    re.compile(r"^\s*vbscript:", re.IGNORECASE),
    re.compile(r"^\s*file://", re.IGNORECASE),
    re.compile(r"^\s*https?://", re.IGNORECASE),
]

_COLOUR_FIELD_HINTS = {
    "color",
    "foreground",
    "background",
    "backgroundLight",
    "backgroundNeutral",
    "foregroundNeutralSecondary",
    "foregroundNeutralTertiary",
    "tableAccent",
    "good",
    "neutral",
    "bad",
    "maximum",
    "center",
    "minimum",
    "hyperlink",
    "visitedHyperlink",
    "fontColor",
    "borderColor",
    "lineColor",
    "backColor",
    "altBackColor",
}


class ThemeValidationError(PowerBIValidationError):
    code = "theme_validation_error"


def _is_hex_colour(value: str) -> bool:
    return bool(_HEX6_RE.match(value) or _HEX8_RE.match(value))


def _check_value(value: Any, path: str, issues: list[dict[str, Any]]) -> None:
    if isinstance(value, str):
        for pattern in _FORBIDDEN_VALUE_PATTERNS:
            if pattern.search(value):
                issues.append(
                    {
                        "level": "error",
                        "path": path,
                        "message": "Theme string contains a forbidden URL-like pattern.",
                        "matched": pattern.pattern,
                    }
                )
                return
        # ``#``-prefixed strings in colour-shaped fields must be hex.
        leaf = path.rsplit(".", 1)[-1]
        if value.startswith("#") and leaf in _COLOUR_FIELD_HINTS and not _is_hex_colour(value):
            issues.append(
                {
                    "level": "error",
                    "path": path,
                    "message": "Colour value must be '#RRGGBB' or '#RRGGBBAA'.",
                    "value": value,
                }
            )
        return
    if isinstance(value, list):
        for index, item in enumerate(value):
            _check_value(item, f"{path}[{index}]", issues)
        return
    if isinstance(value, dict):
        for key, item in value.items():
            _check_value(item, f"{path}.{key}", issues)
        return
    # primitives (int/float/bool/None) are accepted as-is.


def validate_theme_payload(
    payload: Any,
    *,
    strict_top_level: bool = True,
) -> list[dict[str, Any]]:
    """Return a list of validation issues for ``payload``.

    ``strict_top_level`` rejects any unknown top-level key. Deep keys
    inside ``visualStyles`` are intentionally not enumerated — Power BI
    visuals carry arbitrary extension keys there.
    """
    if not isinstance(payload, dict):
        raise ThemeValidationError(
            "Theme payload must be a JSON object at the top level.",
            details={"top_level_type": type(payload).__name__},
        )

    issues: list[dict[str, Any]] = []

    if strict_top_level:
        unknown = sorted(set(payload.keys()) - THEME_ALLOWED_TOP_LEVEL_KEYS)
        for key in unknown:
            issues.append(
                {
                    "level": "error",
                    "path": key,
                    "message": "Unknown top-level theme key.",
                }
            )

    name = payload.get("name")
    if name is not None and not isinstance(name, str):
        issues.append({"level": "error", "path": "name", "message": "Theme 'name' must be a string."})
    if isinstance(name, str) and not name.strip():
        issues.append({"level": "error", "path": "name", "message": "Theme 'name' cannot be blank."})

    data_colors = payload.get("dataColors")
    if data_colors is not None:
        if not isinstance(data_colors, list):
            issues.append(
                {"level": "error", "path": "dataColors", "message": "'dataColors' must be a list of hex strings."}
            )
        else:
            for index, item in enumerate(data_colors):
                if not isinstance(item, str) or not _is_hex_colour(item):
                    issues.append(
                        {
                            "level": "error",
                            "path": f"dataColors[{index}]",
                            "message": "dataColors entries must be '#RRGGBB' or '#RRGGBBAA'.",
                            "value": item,
                        }
                    )

    _check_value(payload, "$", issues)
    return issues


def assert_theme_within_size_limit(raw_bytes: int) -> None:
    if raw_bytes > MAX_THEME_BYTES:
        raise ThemeValidationError(
            f"Theme JSON exceeds the {MAX_THEME_BYTES} byte limit.",
            details={"size_bytes": raw_bytes, "limit_bytes": MAX_THEME_BYTES},
        )


__all__ = [
    "MAX_THEME_BYTES",
    "THEME_ALLOWED_TOP_LEVEL_KEYS",
    "ThemeValidationError",
    "validate_theme_payload",
    "assert_theme_within_size_limit",
]

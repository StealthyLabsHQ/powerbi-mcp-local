"""Format-string preset library for Power BI measures.

Centralises the format strings most commonly applied to measures so an LLM
can pick a preset by name (``currency_eur_k``, ``percent_2dp``, …) rather than
crafting a Power BI numeric format expression by hand.

Tools exposed:
- ``pbi_list_format_presets_tool`` — return the catalogue
- ``pbi_apply_format_preset_tool`` — apply a preset to one or more measures /
  columns by delegating to ``pbi_set_format_tool``
"""

from __future__ import annotations

from typing import Any

from pbi_connection import PowerBIValidationError, ok

from .measures import pbi_set_format_tool

# Each preset entry: format string + a short description.
# Format strings follow the Power BI / Excel custom number format grammar.
# The escaping (``\K``, ``\€``) prevents PBI's auto-K/M scaling from kicking in
# on cards because the format is no longer "general".
PRESETS: dict[str, dict[str, str]] = {
    # --- Currency ---
    "currency_eur": {
        "format_string": "#,##0.00 \\€",
        "description": "Euro currency, 2 decimals (e.g. 1 234,56 €).",
    },
    "currency_eur_k": {
        "format_string": "#,##0,\\K \\€",
        "description": "Euro currency in thousands (e.g. 1 234K €).",
    },
    "currency_eur_m": {
        "format_string": "#,##0.0,,\\M \\€",
        "description": "Euro currency in millions (e.g. 12.3M €).",
    },
    "currency_usd": {
        "format_string": "$#,##0.00",
        "description": "US dollar currency, 2 decimals.",
    },
    "currency_usd_k": {
        "format_string": "$#,##0,\\K",
        "description": "US dollar currency in thousands.",
    },
    "currency_usd_m": {
        "format_string": "$#,##0.0,,\\M",
        "description": "US dollar currency in millions.",
    },
    # --- Percent ---
    "percent": {
        "format_string": "0.00%",
        "description": "Percentage with 2 decimals (default).",
    },
    "percent_0dp": {
        "format_string": "0%",
        "description": "Percentage rounded to integer.",
    },
    "percent_1dp": {
        "format_string": "0.0%",
        "description": "Percentage with 1 decimal.",
    },
    "percent_2dp": {
        "format_string": "0.00%",
        "description": "Percentage with 2 decimals.",
    },
    "percent_4dp": {
        "format_string": "0.0000%",
        "description": "Percentage with 4 decimals (e.g. tax/commission rates).",
    },
    # --- Generic numeric ---
    "thousands": {
        "format_string": "#,##0,\\K",
        "description": "Number in thousands (e.g. 1 234K).",
    },
    "millions": {
        "format_string": "#,##0.0,,\\M",
        "description": "Number in millions (e.g. 12.3M).",
    },
    "decimal_2": {
        "format_string": "#,##0.00",
        "description": "Number with thousand separator and 2 decimals.",
    },
    "integer": {
        "format_string": "#,##0",
        "description": "Integer with thousand separator (e.g. 1 234).",
    },
    "integer_no_sep": {
        "format_string": "0",
        "description": "Integer without thousand separator (e.g. 1234).",
    },
    # --- Dates ---
    "date_iso": {
        "format_string": "yyyy-MM-dd",
        "description": "ISO 8601 short date (e.g. 2026-05-07).",
    },
    "date_short_fr": {
        "format_string": "dd/MM/yyyy",
        "description": "French short date (e.g. 07/05/2026).",
    },
    "date_short_us": {
        "format_string": "MM/dd/yyyy",
        "description": "US short date (e.g. 05/07/2026).",
    },
    "date_long_fr": {
        "format_string": "dddd d MMMM yyyy",
        "description": "French long date (e.g. mercredi 7 mai 2026).",
    },
    "datetime_iso": {
        "format_string": "yyyy-MM-dd HH:mm:ss",
        "description": "ISO date + 24h time.",
    },
}


def _resolve_preset(preset: str) -> str:
    if not preset or not str(preset).strip():
        raise PowerBIValidationError("preset must be a non-empty string.")
    key = str(preset).strip()
    spec = PRESETS.get(key)
    if spec is None:
        raise PowerBIValidationError(
            f"Unknown format preset '{preset}'.",
            details={"preset": preset, "available": sorted(PRESETS)},
        )
    return spec["format_string"]


def pbi_list_format_presets_tool(*, filter_substring: str | None = None) -> dict[str, Any]:
    """Return the catalogue of format-string presets.

    Optional ``filter_substring`` is matched (case-insensitive) against each
    preset's name to narrow the result — useful when an LLM only needs the
    "percent_*" or "currency_*" subset.
    """
    if filter_substring is None:
        items = PRESETS
    else:
        token = str(filter_substring).strip().casefold()
        items = {k: v for k, v in PRESETS.items() if token in k.casefold()}
    return ok(
        f"Returning {len(items)} format preset(s).",
        presets={
            name: {"format_string": data["format_string"], "description": data["description"]}
            for name, data in sorted(items.items())
        },
        count=len(items),
    )


def pbi_apply_format_preset_tool(
    manager: Any,
    *,
    table: str,
    names: list[str],
    preset: str,
    object_type: str = "measure",
) -> dict[str, Any]:
    """Apply a named format preset to a list of measures or columns.

    Wraps ``pbi_set_format_tool`` so callers don't need to memorise the raw
    format string. Supported preset families: currency_eur(_k|_m), currency_usd,
    percent (0dp/1dp/2dp/4dp), thousands, millions, decimal_2, integer,
    integer_no_sep, date_iso, date_short_fr, date_short_us, date_long_fr,
    datetime_iso. Use ``pbi_list_format_presets_tool`` to inspect them.
    """
    format_string = _resolve_preset(preset)
    response = pbi_set_format_tool(
        manager,
        table=table,
        names=names,
        format_string=format_string,
        object_type=object_type,
    )
    response.setdefault("preset", preset)
    response.setdefault("format_string", format_string)
    return response


__all__ = [
    "PRESETS",
    "pbi_apply_format_preset_tool",
    "pbi_list_format_presets_tool",
]

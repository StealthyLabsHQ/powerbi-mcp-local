"""Field reference normalisation: ``Table.Column`` / ``Table[Column]`` /
``'Table'[Column]`` / bare measure name."""

from __future__ import annotations

import re

from pbi_connection import PowerBIValidationError

# Match the standard Power BI bracket forms: Table[Column] or 'Table With Spaces'[Column].
_BRACKET_REF_RE = re.compile(r"^\s*'?(?P<table>[^'\[\]]+?)'?\s*\[\s*(?P<column>[^\[\]]+?)\s*\]\s*$")


def _normalize_reference(reference: str) -> str:
    """Normalise a user-supplied field reference into ``"Table.Column"`` form.

    Accepts (case-insensitive on whitespace; surrounding quotes optional):
    - ``"Table.Column"`` (existing canonical form, returned as-is)
    - ``"Table[Column]"``
    - ``"'Table With Spaces'[Column]"``
    - ``"BareMeasureName"`` (measure references stay unchanged)
    """
    if not isinstance(reference, str):
        return reference  # type: ignore[return-value]
    raw = reference.strip()
    if not raw:
        return raw
    match = _BRACKET_REF_RE.match(raw)
    if match:
        table = match.group("table").strip()
        column = match.group("column").strip()
        return f"{table}.{column}"
    return raw


def _split_column_ref(reference: str) -> tuple[str, str]:
    normalized = _normalize_reference(reference)
    if "." not in normalized:
        raise PowerBIValidationError(
            "Column references must use 'TableName.ColumnName', 'TableName[ColumnName]', "
            "or '\\'Table Name\\'[Column Name]' format.",
            details={"reference": reference},
        )
    table, column = normalized.rsplit(".", 1)
    if not table.strip() or not column.strip():
        raise PowerBIValidationError(
            "Column references must include both a table and a column name.",
            details={"reference": reference},
        )
    return table.strip(), column.strip()


def _query_ref(reference: str) -> str:
    """Return the short queryRef name (column part only, without table prefix).

    Accepts the same flexible reference forms as :func:`_normalize_reference`
    (``Table.Column``, ``Table[Column]``, ``'Table'[Column]``, bare measure).
    """
    normalized = _normalize_reference(reference)
    return normalized.split(".", 1)[1] if "." in normalized else normalized

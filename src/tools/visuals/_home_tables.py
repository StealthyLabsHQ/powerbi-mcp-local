"""Measure → home table resolution.

Visual bindings need each measure's owning table to fill the ``Entity``
field of the prototype query. Two sources, on-disk PBIP metadata + the
live AS model, are merged so writes don't trigger a post-hoc
``measure_home_table_needs_repair`` issue.
"""

from __future__ import annotations

import logging
import re
from pathlib import Path
from typing import Any

from ._base import MODEL_TABLES_RELATIVE_DIR
from ._paths import _resolve_extract_folder

logger = logging.getLogger("tools.visuals._home_tables")

_DAX_LINE_COMMENT_RE = re.compile(r"//[^\n]*")
_DAX_BLOCK_COMMENT_RE = re.compile(r"/\*.*?\*/", re.DOTALL)
_DAX_STRING_LITERAL_RE = re.compile(r'"(?:[^"]|"")*"')
_DAX_BLANK_ONLY_RE = re.compile(r"^\s*BLANK\s*\(\s*\)\s*$", re.IGNORECASE)


def _is_likely_constant_dax(expression: str) -> tuple[bool, str | None]:
    """Best-effort heuristic for "DAX expression that resolves to a scalar
    independent of filter context".

    Returns ``(is_constant, hint)``. Strips comments and string literals
    before checking so commented-out references and quoted hints don't
    cause false negatives.

    Catches:
    - No ``[...]`` reference at all (pure literal arithmetic, hard-coded
      numeric / boolean / null values).
    - ``BLANK()`` as the entire expression body.

    Misses dynamic-but-still-constant expressions that contain a column
    reference but always evaluate to the same scalar (e.g.
    ``CALCULATE(SUM(Sales[Amount]), Sales[Amount] = 0)``). Detecting those
    requires a runtime probe.
    """
    raw = (expression or "").strip()
    if not raw:
        return False, None
    cleaned = _DAX_LINE_COMMENT_RE.sub("", raw)
    cleaned = _DAX_BLOCK_COMMENT_RE.sub("", cleaned)
    cleaned = _DAX_STRING_LITERAL_RE.sub('""', cleaned)
    if "[" not in cleaned:
        return True, "no column or measure reference (looks like a scalar literal)"
    if _DAX_BLANK_ONLY_RE.match(cleaned):
        return True, "expression body is BLANK() — every Y value collapses to BLANK"
    return False, None


def _scan_measure_home_tables(extract_folder: Path) -> dict[str, str]:
    """Map measure name → home table from extract metadata folders."""
    table_root = extract_folder / MODEL_TABLES_RELATIVE_DIR
    if not table_root.is_dir():
        return {}

    measure_home_map: dict[str, str] = {}
    for table_dir in table_root.iterdir():
        if not table_dir.is_dir():
            continue
        measures_dir = table_dir / "measures"
        if not measures_dir.is_dir():
            continue
        for dax_file in measures_dir.glob("*.dax"):
            measure_name = dax_file.stem.strip()
            if not measure_name:
                continue
            existing = measure_home_map.get(measure_name)
            if existing and existing != table_dir.name:
                logger.warning(
                    "Measure '%s' found in multiple tables ('%s', '%s'); keeping first.",
                    measure_name,
                    existing,
                    table_dir.name,
                )
                continue
            measure_home_map[measure_name] = table_dir.name
    return measure_home_map


def _resolve_measure_home_map(
    extract_folder: str,
    manager: Any | None = None,
) -> dict[str, str]:
    """Build a measure → home table map combining on-disk PBIP metadata and the
    live model (when ``manager`` is supplied).

    Use at the top of every ``pbi_add_*_tool`` so the visual write carries the
    correct ``Entity`` reference and callers don't get the post-hoc
    ``measure_home_table_needs_repair`` validation issue.
    """
    home_map = _scan_measure_home_tables(_resolve_extract_folder(extract_folder, must_exist=True))
    return _augment_measure_home_map_with_live(home_map, manager)


def _augment_measure_home_map_with_live(
    measure_home_map: dict[str, str],
    manager: Any | None,
    *,
    include_hidden: bool = False,
) -> dict[str, str]:
    """Fill in missing measure → table mappings from the live model.

    The on-disk PBIP extract metadata is the canonical source, but it isn't
    always present (e.g. a layout-only extract from a closed PBIX). When a
    connection manager is supplied, we pull the same information from the
    live TOM so visual writes carry the correct ``Entity`` reference and
    callers don't see ``measure_home_table_needs_repair`` after the write.

    Returns the (possibly augmented) map. Existing entries take priority so
    the on-disk metadata always wins on conflict.
    """
    if manager is None:
        return measure_home_map
    # Resolve through the visuals package re-export so unittest.mock patches
    # against ``tools.visuals.pbi_model_info_tool`` keep working.
    from . import pbi_model_info_tool

    try:
        model = pbi_model_info_tool(manager, include_hidden=include_hidden, include_row_counts=False)
    except Exception:
        return measure_home_map
    if not model.get("ok"):
        return measure_home_map
    existing_lower = {key.casefold() for key in measure_home_map}
    for measure in model.get("measures", []) or []:
        name = str(measure.get("name", ""))
        table_name = str(measure.get("table", ""))
        if not name or not table_name:
            continue
        if name.casefold() in existing_lower:
            continue
        measure_home_map[name] = table_name
        existing_lower.add(name.casefold())
    return measure_home_map


def _runtime_probe_measure_constancy(
    manager: Any | None,
    axis_ref: str | None,
    measure: str,
    *,
    sample_count: int = 3,
) -> tuple[bool, dict[str, Any] | None]:
    """Live-engine probe for "this measure returns the same scalar across
    every axis value".

    Catches the dynamic-but-still-constant cases the static
    :func:`_is_likely_constant_dax` heuristic misses (e.g.
    ``CALCULATE(SUM(Sales[Amount]), Sales[Amount] = 0)``).

    Returns ``(is_constant, probe_data)``:
    - ``manager`` ``None`` or ``axis_ref`` not in ``Table.Column`` form →
      ``(False, None)`` — no probe attempted.
    - Engine call fails or returns < 2 distinct axis values →
      ``(False, None)`` — inconclusive.
    - All ``sample_count`` measure values match → ``(True, {…samples…})``.
    """
    if manager is None or not axis_ref or "." not in axis_ref:
        return False, None
    table, column = axis_ref.split(".", 1)
    table = table.strip()
    column = column.strip()
    if not table or not column:
        return False, None
    sample_count = max(2, min(int(sample_count), 10))
    table_q = f"'{table}'" if any(ch in table for ch in " -.") else table
    column_q = f"[{column}]"
    measure_q = f"[{measure}]"
    query = f'EVALUATE TOPN({sample_count}, ADDCOLUMNS(VALUES({table_q}{column_q}), "__probe_v", {measure_q}))'
    try:
        from ..query import pbi_execute_dax_tool
    except ImportError:
        return False, None
    try:
        result = pbi_execute_dax_tool(manager, query=query, max_rows=sample_count)
    except Exception:
        return False, None
    if not isinstance(result, dict) or not result.get("ok"):
        return False, None
    rows = result.get("rows") or []
    if len(rows) < 2:
        # Need at least 2 distinct axis values to call something constant.
        return False, None
    values: list[Any] = []
    for row in rows:
        if not isinstance(row, dict):
            continue
        # ADOMD returns column names as "[__probe_v]" — try both forms.
        v = row.get("[__probe_v]")
        if v is None:
            v = row.get("__probe_v")
        values.append(v)
    if len(values) < 2:
        return False, None
    distinct = {repr(v) for v in values}
    is_constant = len(distinct) == 1
    return is_constant, {
        "axis": axis_ref,
        "measure": measure,
        "samples": values,
        "distinct_count": len(distinct),
        "rows_probed": len(rows),
    }


def _inspect_value_measures(
    value_measures: list[str],
    measure_home_map: dict[str, str],
    manager: Any | None,
    *,
    axis_ref: str | None = None,
) -> list[dict[str, Any]]:
    """Diagnostic warnings for line/area/combo Y measures.

    Detects two failure modes that PBI Desktop reports as opaque rendering errors:

    - ``measure_home_unresolved``: the measure exists in the live model but no
      home table could be resolved (extract metadata absent and live lookup
      failed). The binding falls back to the synthetic ``$Measures`` entity
      which PBI refuses to plot in cartesian charts.
    - ``constant_measure``: the DAX expression has no column or measure
      reference, e.g. ``Ratio = 0.92``. Line/combo charts trigger an internal
      error on some PBI builds when every Y value collapses to the same scalar
      with no axis dependency. Wrap the value in ``VAR`` or persist the model
      to bind it through a real table entity first.
    """
    from ..measures import pbi_list_measures_tool  # local: avoid circular imports

    warnings: list[dict[str, Any]] = []
    expressions: dict[str, str] = {}
    if manager is not None:
        try:
            listed = pbi_list_measures_tool(manager, include_hidden=True)
            for entry in listed.get("measures", []) or []:
                expressions[str(entry.get("name", ""))] = str(entry.get("expression", ""))
        except Exception:  # pragma: no cover — manager might be detached
            pass

    for measure in value_measures:
        if measure not in measure_home_map:
            warnings.append(
                {
                    "measure": measure,
                    "issue": "measure_home_unresolved",
                    "hint": (
                        "Home table not found on disk or live. Binding will fall "
                        "back to '$Measures' which PBI refuses to render. Save the "
                        ".pbix (Ctrl+S) to persist the measure before adding it to "
                        "a cartesian chart."
                    ),
                }
            )
        expr = expressions.get(measure, "")
        is_constant, hint = _is_likely_constant_dax(expr)
        if is_constant:
            warnings.append(
                {
                    "measure": measure,
                    "issue": "constant_measure",
                    "hint": (
                        f"{hint}. Line/combo/waterfall charts may error out on render "
                        "— wrap the value via CALCULATE/VAR with a real filter, or "
                        "use a card visual instead."
                    ),
                    "expression_preview": expr[:120],
                }
            )
            continue  # static check fired — no need for the runtime probe
        # Runtime probe (only when axis ref + manager available): catches
        # dynamic-but-still-constant DAX (e.g. ``CALCULATE(... , <filter>)``
        # that always evaluates to the same scalar) which static parsing
        # cannot see.
        is_runtime_constant, probe = _runtime_probe_measure_constancy(manager, axis_ref, measure)
        if is_runtime_constant and probe is not None:
            warnings.append(
                {
                    "measure": measure,
                    "issue": "runtime_constant_measure",
                    "hint": (
                        f"Runtime probe: measure returned the same value across "
                        f"{probe.get('rows_probed', '?')} distinct axis points "
                        f"({probe.get('axis')}). Line/combo/waterfall charts may "
                        "render as a flat baseline or trigger an internal error. "
                        "Add filter context that varies along the axis."
                    ),
                    "probe": probe,
                    "expression_preview": expr[:120] if expr else None,
                }
            )
    return warnings


def _persistence_risks(issues: list[dict[str, Any]]) -> list[dict[str, Any]]:
    return [item for item in issues if item.get("source") == "live_model" and item.get("extract_metadata") == "missing"]

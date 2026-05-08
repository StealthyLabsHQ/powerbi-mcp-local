"""Measure → home table resolution.

Visual bindings need each measure's owning table to fill the ``Entity``
field of the prototype query. Two sources, on-disk PBIP metadata + the
live AS model, are merged so writes don't trigger a post-hoc
``measure_home_table_needs_repair`` issue.
"""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Any

from ._base import MODEL_TABLES_RELATIVE_DIR
from ._paths import _resolve_extract_folder


logger = logging.getLogger("tools.visuals._home_tables")


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


def _inspect_value_measures(
    value_measures: list[str],
    measure_home_map: dict[str, str],
    manager: Any | None,
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
            warnings.append({
                "measure": measure,
                "issue": "measure_home_unresolved",
                "hint": (
                    "Home table not found on disk or live. Binding will fall "
                    "back to '$Measures' which PBI refuses to render. Save the "
                    ".pbix (Ctrl+S) to persist the measure before adding it to "
                    "a cartesian chart."
                ),
            })
        expr = expressions.get(measure, "")
        if expr and "[" not in expr:
            warnings.append({
                "measure": measure,
                "issue": "constant_measure",
                "hint": (
                    "DAX expression has no column or measure reference (looks "
                    "like a scalar constant). Line/combo charts may error out "
                    "on render — wrap the value via CALCULATE/VAR with a "
                    "harmless filter, or use a card visual instead."
                ),
                "expression_preview": expr[:120],
            })
    return warnings


def _persistence_risks(issues: list[dict[str, Any]]) -> list[dict[str, Any]]:
    return [
        item for item in issues
        if item.get("source") == "live_model"
        and item.get("extract_metadata") == "missing"
    ]

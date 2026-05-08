"""Structure-style visual tools: table, slicer, matrix, map."""

from __future__ import annotations

from typing import Any

from pbi_connection import PowerBIValidationError

from ._bindings import _validate_field_references_live
from ._containers import _append_visual, _create_chart_container
from ._formatting import _int_literal, _literal_value, _text_literal
from ._home_tables import _resolve_measure_home_map
from ._refs import _query_ref


def pbi_add_table_visual_tool(extract_folder: str, page: str, columns: list[str], x: int, y: int, width: int = 520, height: int = 320, title: str = "", *, manager: Any | None = None) -> dict[str, Any]:
    if not columns:
        raise PowerBIValidationError("columns must contain at least one field or measure.")
    _validate_field_references_live(manager, list(columns))
    measure_home_map = _resolve_measure_home_map(extract_folder, manager=manager)
    return _append_visual(
        extract_folder,
        page,
        lambda section, home_map: _create_chart_container(
            section,
            visual_type="tableEx",
            x=x,
            y=y,
            width=width,
            height=height,
            title=title or None,
            projections={"Values": [{"queryRef": _query_ref(item)} for item in columns]},
            references=list(columns),
            measure_home_map=home_map,
        ),
        measure_home_map,
    )


def pbi_add_slicer_tool(extract_folder: str, page: str, column: str, x: int, y: int, width: int = 220, height: int = 120, slicer_type: str = "dropdown", *, manager: Any | None = None) -> dict[str, Any]:
    slicer_kind = slicer_type.strip().casefold()
    if slicer_kind not in {"dropdown", "list", "range", "tile"}:
        raise PowerBIValidationError("slicer_type must be one of: dropdown, list, range, tile.", details={"slicer_type": slicer_type})
    _validate_field_references_live(manager, [column], expected_kinds={column: "column"})
    measure_home_map = _resolve_measure_home_map(extract_folder, manager=manager)
    if slicer_kind == "tile":
        emitted_kind = "list"
        extra: dict[str, Any] = {
            "slicerType": emitted_kind,
            "objects": {
                "general": [
                    {"properties": {"orientation": _int_literal(1)}}
                ]
            },
        }
    else:
        extra = {"slicerType": slicer_kind}
    return _append_visual(
        extract_folder,
        page,
        lambda section, home_map: _create_chart_container(
            section,
            visual_type="slicer",
            x=x,
            y=y,
            width=width,
            height=height,
            title=None,
            projections={"Values": [{"queryRef": _query_ref(column)}]},
            references=[column],
            measure_home_map=home_map,
            extra_single_visual=extra,
        ),
        measure_home_map,
    )


def pbi_add_matrix_tool(
    extract_folder: str,
    page: str,
    rows: list[str],
    values: list[str],
    x: int,
    y: int,
    columns: list[str] | None = None,
    width: int = 540,
    height: int = 360,
    title: str = "",
    subtotals: bool = True,
    column_layout: str = "stepped",
    *,
    manager: Any | None = None,
) -> dict[str, Any]:
    """Add a matrix / pivot-table visual.

    Roles: Rows (column list, required), Columns (column list, optional),
    Values (measure list, required). ``column_layout``: ``"stepped"`` or
    ``"tabular"``.
    """
    if not rows:
        raise PowerBIValidationError("rows must contain at least one column.")
    if not values:
        raise PowerBIValidationError("values must contain at least one measure.")
    layout_token = column_layout.strip().casefold()
    if layout_token not in {"stepped", "tabular"}:
        raise PowerBIValidationError(
            "column_layout must be 'stepped' or 'tabular'.",
            details={"column_layout": column_layout},
        )
    projections: dict[str, list[dict[str, str]]] = {
        "Rows": [{"queryRef": _query_ref(item)} for item in rows],
        "Values": [{"queryRef": _query_ref(item)} for item in values],
    }
    references = [*rows, *values]
    expected_kinds = {r: "column" for r in rows}
    expected_kinds.update({v: "measure" for v in values})
    if columns:
        projections["Columns"] = [{"queryRef": _query_ref(item)} for item in columns]
        references.extend(columns)
        for c in columns:
            expected_kinds[c] = "column"
    _validate_field_references_live(manager, references, expected_kinds=expected_kinds)
    extra_single_visual = {
        "objects": {
            "subTotals": [
                {"properties": {"rowSubtotals": _literal_value(bool(subtotals))}}
            ],
            "general": [
                {"properties": {"layout": _text_literal("Stepped" if layout_token == "stepped" else "Tabular")}}
            ],
        }
    }
    measure_home_map = _resolve_measure_home_map(extract_folder, manager=manager)
    return _append_visual(
        extract_folder,
        page,
        lambda section, home_map: _create_chart_container(
            section,
            visual_type="pivotTable",
            x=x,
            y=y,
            width=width,
            height=height,
            title=title or None,
            projections=projections,
            references=references,
            measure_home_map=home_map,
            extra_single_visual=extra_single_visual,
            manager=manager,
        ),
        measure_home_map,
    )


def pbi_add_map_tool(
    extract_folder: str,
    page: str,
    location_column: str,
    value_measure: str | None,
    x: int,
    y: int,
    width: int = 420,
    height: int = 320,
    title: str = "",
    *,
    manager: Any | None = None,
) -> dict[str, Any]:
    """Add a bubble/map visual.

    Roles: ``Category`` (location column — country, city, Lat/Long…) and
    optional ``Y`` (measure that drives bubble size).
    """
    projections: dict[str, list[dict[str, str]]] = {
        "Category": [{"queryRef": _query_ref(location_column)}]
    }
    references = [location_column]
    expected_kinds = {location_column: "column"}
    if value_measure:
        projections["Y"] = [{"queryRef": _query_ref(value_measure)}]
        references.append(value_measure)
        expected_kinds[value_measure] = "measure"
    _validate_field_references_live(manager, references, expected_kinds=expected_kinds)
    measure_home_map = _resolve_measure_home_map(extract_folder, manager=manager)
    return _append_visual(
        extract_folder,
        page,
        lambda section, home_map: _create_chart_container(
            section,
            visual_type="map",
            x=x,
            y=y,
            width=width,
            height=height,
            title=title or None,
            projections=projections,
            references=references,
            measure_home_map=home_map,
            manager=manager,
        ),
        measure_home_map,
    )

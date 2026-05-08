"""Cartesian chart tools: bar, line, donut, waterfall, scatter, combo."""

from __future__ import annotations

from typing import Any

from pbi_connection import PowerBIValidationError

from ._bindings import _validate_field_references_live
from ._containers import _append_visual, _create_chart_container
from ._home_tables import _inspect_value_measures, _resolve_measure_home_map
from ._refs import _query_ref


def pbi_add_bar_chart_tool(
    extract_folder: str,
    page: str,
    category_column: str,
    value_measure: str,
    x: int,
    y: int,
    width: int = 400,
    height: int = 300,
    title: str = "",
    legend_column: str | None = None,
    *,
    manager: Any | None = None,
) -> dict[str, Any]:
    measure_home_map = _resolve_measure_home_map(extract_folder, manager=manager)
    projections = {
        "Category": [{"queryRef": _query_ref(category_column)}],
        "Y": [{"queryRef": _query_ref(value_measure)}],
    }
    references = [category_column, value_measure]
    expected_kinds = {category_column: "column", value_measure: "measure"}
    if legend_column:
        projections["Series"] = [{"queryRef": _query_ref(legend_column)}]
        references.append(legend_column)
        expected_kinds[legend_column] = "column"
    _validate_field_references_live(manager, references, expected_kinds=expected_kinds)
    return _append_visual(
        extract_folder,
        page,
        lambda section, home_map: _create_chart_container(
            section,
            visual_type="clusteredBarChart",
            x=x,
            y=y,
            width=width,
            height=height,
            title=title or None,
            projections=projections,
            references=references,
            measure_home_map=home_map,
        ),
        measure_home_map,
    )


def pbi_add_line_chart_tool(
    extract_folder: str,
    page: str,
    axis_column: str,
    value_measures: list[str],
    x: int,
    y: int,
    width: int = 420,
    height: int = 300,
    title: str = "",
    *,
    manager: Any | None = None,
) -> dict[str, Any]:
    if not value_measures:
        raise PowerBIValidationError("value_measures must contain at least one measure.")
    expected_kinds = {axis_column: "column"}
    for m in value_measures:
        expected_kinds[m] = "measure"
    _validate_field_references_live(manager, [axis_column, *value_measures], expected_kinds=expected_kinds)
    measure_home_map = _resolve_measure_home_map(extract_folder, manager=manager)
    diagnostics = _inspect_value_measures(value_measures, measure_home_map, manager, axis_ref=axis_column)
    result = _append_visual(
        extract_folder,
        page,
        lambda section, home_map: _create_chart_container(
            section,
            visual_type="lineChart",
            x=x,
            y=y,
            width=width,
            height=height,
            title=title or None,
            projections={
                "Category": [{"queryRef": _query_ref(axis_column)}],
                "Y": [{"queryRef": _query_ref(measure)} for measure in value_measures],
            },
            references=[axis_column, *value_measures],
            measure_home_map=home_map,
        ),
        measure_home_map,
    )
    if diagnostics:
        result["warnings"] = diagnostics
    return result


def pbi_add_donut_chart_tool(
    extract_folder: str,
    page: str,
    category_column: str,
    value_measure: str,
    x: int,
    y: int,
    width: int = 320,
    height: int = 280,
    title: str = "",
    *,
    manager: Any | None = None,
) -> dict[str, Any]:
    _validate_field_references_live(
        manager, [category_column, value_measure], expected_kinds={category_column: "column", value_measure: "measure"}
    )
    measure_home_map = _resolve_measure_home_map(extract_folder, manager=manager)
    return _append_visual(
        extract_folder,
        page,
        lambda section, home_map: _create_chart_container(
            section,
            visual_type="donutChart",
            x=x,
            y=y,
            width=width,
            height=height,
            title=title or None,
            projections={
                "Category": [{"queryRef": _query_ref(category_column)}],
                "Y": [{"queryRef": _query_ref(value_measure)}],
            },
            references=[category_column, value_measure],
            measure_home_map=home_map,
        ),
        measure_home_map,
    )


def pbi_add_waterfall_tool(
    extract_folder: str,
    page: str,
    category_column: str,
    value_measure: str,
    x: int,
    y: int,
    width: int = 420,
    height: int = 300,
    title: str = "",
    *,
    manager: Any | None = None,
) -> dict[str, Any]:
    _validate_field_references_live(
        manager, [category_column, value_measure], expected_kinds={category_column: "column", value_measure: "measure"}
    )
    measure_home_map = _resolve_measure_home_map(extract_folder, manager=manager)
    diagnostics = _inspect_value_measures([value_measure], measure_home_map, manager, axis_ref=category_column)
    result = _append_visual(
        extract_folder,
        page,
        lambda section, home_map: _create_chart_container(
            section,
            visual_type="waterfallChart",
            x=x,
            y=y,
            width=width,
            height=height,
            title=title or None,
            projections={
                "Category": [{"queryRef": _query_ref(category_column)}],
                "Y": [{"queryRef": _query_ref(value_measure)}],
            },
            references=[category_column, value_measure],
            measure_home_map=home_map,
        ),
        measure_home_map,
    )
    if diagnostics:
        result["warnings"] = diagnostics
    return result


def pbi_add_scatter_chart_tool(
    extract_folder: str,
    page: str,
    category_column: str,
    x_measure: str,
    y_measure: str,
    x: int,
    y: int,
    width: int = 420,
    height: int = 320,
    title: str = "",
    size_measure: str | None = None,
    legend_column: str | None = None,
    *,
    manager: Any | None = None,
) -> dict[str, Any]:
    """Add a scatter chart visual.

    Roles: Category (column), X (measure), Y (measure), Size (measure,
    optional), Series (column, optional). Use for correlation analysis
    between two measures grouped by a dimension.
    """
    projections = {
        "Category": [{"queryRef": _query_ref(category_column)}],
        "X": [{"queryRef": _query_ref(x_measure)}],
        "Y": [{"queryRef": _query_ref(y_measure)}],
    }
    references = [category_column, x_measure, y_measure]
    expected_kinds = {category_column: "column", x_measure: "measure", y_measure: "measure"}
    if size_measure:
        projections["Size"] = [{"queryRef": _query_ref(size_measure)}]
        references.append(size_measure)
        expected_kinds[size_measure] = "measure"
    if legend_column:
        projections["Series"] = [{"queryRef": _query_ref(legend_column)}]
        references.append(legend_column)
        expected_kinds[legend_column] = "column"
    _validate_field_references_live(manager, references, expected_kinds=expected_kinds)
    measure_home_map = _resolve_measure_home_map(extract_folder, manager=manager)
    return _append_visual(
        extract_folder,
        page,
        lambda section, home_map: _create_chart_container(
            section,
            visual_type="scatterChart",
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


def pbi_add_combo_chart_tool(
    extract_folder: str,
    page: str,
    category_column: str,
    bar_measures: list[str],
    line_measures: list[str],
    x: int,
    y: int,
    width: int = 480,
    height: int = 320,
    title: str = "",
    legend_column: str | None = None,
    *,
    manager: Any | None = None,
) -> dict[str, Any]:
    """Add a combo chart (column bars + line overlay).

    Roles: Category (column), Y (bar measures), Y2 (line measures).
    """
    if not bar_measures:
        raise PowerBIValidationError("bar_measures must contain at least one measure.")
    if not line_measures:
        raise PowerBIValidationError("line_measures must contain at least one measure.")
    projections: dict[str, list[dict[str, str]]] = {
        "Category": [{"queryRef": _query_ref(category_column)}],
        "Y": [{"queryRef": _query_ref(item)} for item in bar_measures],
        "Y2": [{"queryRef": _query_ref(item)} for item in line_measures],
    }
    references = [category_column, *bar_measures, *line_measures]
    expected_kinds = {category_column: "column"}
    for m in bar_measures:
        expected_kinds[m] = "measure"
    for m in line_measures:
        expected_kinds[m] = "measure"
    if legend_column:
        projections["Series"] = [{"queryRef": _query_ref(legend_column)}]
        references.append(legend_column)
        expected_kinds[legend_column] = "column"
    _validate_field_references_live(manager, references, expected_kinds=expected_kinds)
    measure_home_map = _resolve_measure_home_map(extract_folder, manager=manager)
    diagnostics = _inspect_value_measures(
        [*bar_measures, *line_measures], measure_home_map, manager, axis_ref=category_column
    )
    result = _append_visual(
        extract_folder,
        page,
        lambda section, home_map: _create_chart_container(
            section,
            visual_type="lineClusteredColumnComboChart",
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
    if diagnostics:
        result["warnings"] = diagnostics
    return result

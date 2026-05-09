"""Cartesian chart tools: bar, line, donut, waterfall, scatter, combo."""

from __future__ import annotations

from typing import Any

from pbi_connection import PowerBIValidationError

from ._bindings import _validate_field_references_live
from ._containers import _append_visual, _create_chart_container
from ._home_tables import _inspect_value_measures, _resolve_measure_home_map
from ._refs import _projection


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
        "Category": [_projection(category_column)],
        "Y": [_projection(value_measure)],
    }
    references = [category_column, value_measure]
    expected_kinds = {category_column: "column", value_measure: "measure"}
    if legend_column:
        projections["Series"] = [_projection(legend_column)]
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
                "Category": [_projection(axis_column)],
                "Y": [_projection(measure) for measure in value_measures],
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
                "Category": [_projection(category_column)],
                "Y": [_projection(value_measure)],
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
                "Category": [_projection(category_column)],
                "Y": [_projection(value_measure)],
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
        "Category": [_projection(category_column)],
        "X": [_projection(x_measure)],
        "Y": [_projection(y_measure)],
    }
    references = [category_column, x_measure, y_measure]
    expected_kinds = {category_column: "column", x_measure: "measure", y_measure: "measure"}
    if size_measure:
        projections["Size"] = [_projection(size_measure)]
        references.append(size_measure)
        expected_kinds[size_measure] = "measure"
    if legend_column:
        projections["Series"] = [_projection(legend_column)]
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


# ─────────────────────────────────────────────────────────────────────
# v0.12.6 chart pack — pieChart + stacked / 100% bar+column variants +
# area family + ribbon + treemap + funnel + multiRowCard. The new tools
# share the same projection shape as the existing pbi_add_bar_chart /
# pbi_add_line_chart / pbi_add_donut_chart helpers; only the
# ``visual_type`` differs in the underlying Power BI JSON.
# ─────────────────────────────────────────────────────────────────────


def _add_categorical_chart(
    *,
    visual_type: str,
    extract_folder: str,
    page: str,
    category_column: str,
    value_measure: str,
    x: int,
    y: int,
    width: int,
    height: int,
    title: str,
    legend_column: str | None,
    manager: Any | None,
) -> dict[str, Any]:
    """Shared backbone for bar / column / ribbon / pie / treemap variants."""
    measure_home_map = _resolve_measure_home_map(extract_folder, manager=manager)
    projections: dict[str, list[dict[str, str]]] = {
        "Category": [_projection(category_column)],
        "Y": [_projection(value_measure)],
    }
    references = [category_column, value_measure]
    expected_kinds = {category_column: "column", value_measure: "measure"}
    if legend_column:
        projections["Series"] = [_projection(legend_column)]
        references.append(legend_column)
        expected_kinds[legend_column] = "column"
    _validate_field_references_live(manager, references, expected_kinds=expected_kinds)
    return _append_visual(
        extract_folder,
        page,
        lambda section, home_map: _create_chart_container(
            section,
            visual_type=visual_type,
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


def _add_axis_chart(
    *,
    visual_type: str,
    extract_folder: str,
    page: str,
    axis_column: str,
    value_measures: list[str],
    x: int,
    y: int,
    width: int,
    height: int,
    title: str,
    legend_column: str | None,
    manager: Any | None,
) -> dict[str, Any]:
    """Shared backbone for line / area variants — accepts multiple Y measures."""
    if not value_measures:
        raise PowerBIValidationError("value_measures must contain at least one measure.")
    expected_kinds: dict[str, str] = {axis_column: "column"}
    for m in value_measures:
        expected_kinds[m] = "measure"
    references = [axis_column, *value_measures]
    if legend_column:
        expected_kinds[legend_column] = "column"
        references.append(legend_column)
    _validate_field_references_live(manager, references, expected_kinds=expected_kinds)
    measure_home_map = _resolve_measure_home_map(extract_folder, manager=manager)
    diagnostics = _inspect_value_measures(value_measures, measure_home_map, manager, axis_ref=axis_column)
    projections: dict[str, list[dict[str, str]]] = {
        "Category": [_projection(axis_column)],
        "Y": [_projection(measure) for measure in value_measures],
    }
    if legend_column:
        projections["Series"] = [_projection(legend_column)]
    result = _append_visual(
        extract_folder,
        page,
        lambda section, home_map: _create_chart_container(
            section,
            visual_type=visual_type,
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
    if diagnostics:
        result["warnings"] = diagnostics
    return result


def pbi_add_pie_chart_tool(
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
    """Pie chart — proportional slices for a single measure broken down by a category."""
    return _add_categorical_chart(
        visual_type="pieChart",
        extract_folder=extract_folder,
        page=page,
        category_column=category_column,
        value_measure=value_measure,
        x=x,
        y=y,
        width=width,
        height=height,
        title=title,
        legend_column=None,
        manager=manager,
    )


def pbi_add_stacked_bar_chart_tool(
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
    """Stacked bar chart — horizontal bars with stacked series."""
    return _add_categorical_chart(
        visual_type="stackedBarChart",
        extract_folder=extract_folder,
        page=page,
        category_column=category_column,
        value_measure=value_measure,
        x=x,
        y=y,
        width=width,
        height=height,
        title=title,
        legend_column=legend_column,
        manager=manager,
    )


def pbi_add_stacked_column_chart_tool(
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
    """Stacked column chart — vertical bars with stacked series."""
    return _add_categorical_chart(
        visual_type="stackedColumnChart",
        extract_folder=extract_folder,
        page=page,
        category_column=category_column,
        value_measure=value_measure,
        x=x,
        y=y,
        width=width,
        height=height,
        title=title,
        legend_column=legend_column,
        manager=manager,
    )


def pbi_add_clustered_column_chart_tool(
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
    """Clustered column chart — vertical bars side-by-side per series."""
    return _add_categorical_chart(
        visual_type="clusteredColumnChart",
        extract_folder=extract_folder,
        page=page,
        category_column=category_column,
        value_measure=value_measure,
        x=x,
        y=y,
        width=width,
        height=height,
        title=title,
        legend_column=legend_column,
        manager=manager,
    )


def pbi_add_hundred_percent_stacked_bar_chart_tool(
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
    """100% stacked bar chart — every bar normalised to 100% (proportions)."""
    return _add_categorical_chart(
        visual_type="hundredPercentStackedBarChart",
        extract_folder=extract_folder,
        page=page,
        category_column=category_column,
        value_measure=value_measure,
        x=x,
        y=y,
        width=width,
        height=height,
        title=title,
        legend_column=legend_column,
        manager=manager,
    )


def pbi_add_hundred_percent_stacked_column_chart_tool(
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
    """100% stacked column chart — every column normalised to 100% (proportions)."""
    return _add_categorical_chart(
        visual_type="hundredPercentStackedColumnChart",
        extract_folder=extract_folder,
        page=page,
        category_column=category_column,
        value_measure=value_measure,
        x=x,
        y=y,
        width=width,
        height=height,
        title=title,
        legend_column=legend_column,
        manager=manager,
    )


def pbi_add_ribbon_chart_tool(
    extract_folder: str,
    page: str,
    category_column: str,
    value_measure: str,
    x: int,
    y: int,
    width: int = 480,
    height: int = 320,
    title: str = "",
    legend_column: str | None = None,
    *,
    manager: Any | None = None,
) -> dict[str, Any]:
    """Ribbon chart — series ranked per category with continuity ribbons.

    Most useful when the rank ordering of series shifts across the
    category axis (top-N over time / regions).
    """
    return _add_categorical_chart(
        visual_type="ribbonChart",
        extract_folder=extract_folder,
        page=page,
        category_column=category_column,
        value_measure=value_measure,
        x=x,
        y=y,
        width=width,
        height=height,
        title=title,
        legend_column=legend_column,
        manager=manager,
    )


def pbi_add_treemap_tool(
    extract_folder: str,
    page: str,
    category_column: str,
    value_measure: str,
    x: int,
    y: int,
    width: int = 360,
    height: int = 300,
    title: str = "",
    details_column: str | None = None,
    *,
    manager: Any | None = None,
) -> dict[str, Any]:
    """Treemap — nested rectangles sized by the measure, grouped by category.

    Power BI's treemap uses the ``Category``, ``Details``, and ``Values``
    projection roles — *not* the cartesian ``Y`` role. Earlier builds of
    this tool emitted ``Y`` and the visual rendered empty (PBI Desktop's
    data-shape pass dropped the unrecognised role for the treemap
    visualType). The ``details_column`` parameter is optional — pass a
    second-level grouping column to drill the treemap into nested
    rectangles per Details value.
    """
    measure_home_map = _resolve_measure_home_map(extract_folder, manager=manager)
    projections: dict[str, list[dict[str, Any]]] = {
        "Category": [_projection(category_column)],
        "Values": [_projection(value_measure)],
    }
    references = [category_column, value_measure]
    expected_kinds: dict[str, str] = {category_column: "column", value_measure: "measure"}
    if details_column:
        projections["Details"] = [_projection(details_column)]
        references.append(details_column)
        expected_kinds[details_column] = "column"
    _validate_field_references_live(manager, references, expected_kinds=expected_kinds)
    return _append_visual(
        extract_folder,
        page,
        lambda section, home_map: _create_chart_container(
            section,
            visual_type="treemap",
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


def pbi_add_funnel_tool(
    extract_folder: str,
    page: str,
    group_column: str,
    value_measure: str,
    x: int,
    y: int,
    width: int = 360,
    height: int = 320,
    title: str = "",
    *,
    manager: Any | None = None,
) -> dict[str, Any]:
    """Funnel chart — sequential stages sorted by value (sales pipeline / conversion).

    Power BI's funnel uses the ``Group`` and ``Values`` projection roles
    rather than the canonical ``Category`` / ``Y`` of cartesian charts;
    this tool maps them transparently.
    """
    measure_home_map = _resolve_measure_home_map(extract_folder, manager=manager)
    projections = {
        "Group": [_projection(group_column)],
        "Values": [_projection(value_measure)],
    }
    references = [group_column, value_measure]
    expected_kinds = {group_column: "column", value_measure: "measure"}
    _validate_field_references_live(manager, references, expected_kinds=expected_kinds)
    return _append_visual(
        extract_folder,
        page,
        lambda section, home_map: _create_chart_container(
            section,
            visual_type="funnel",
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


def pbi_add_area_chart_tool(
    extract_folder: str,
    page: str,
    axis_column: str,
    value_measures: list[str],
    x: int,
    y: int,
    width: int = 420,
    height: int = 300,
    title: str = "",
    legend_column: str | None = None,
    *,
    manager: Any | None = None,
) -> dict[str, Any]:
    """Area chart — line chart with the area under each series filled."""
    return _add_axis_chart(
        visual_type="areaChart",
        extract_folder=extract_folder,
        page=page,
        axis_column=axis_column,
        value_measures=value_measures,
        x=x,
        y=y,
        width=width,
        height=height,
        title=title,
        legend_column=legend_column,
        manager=manager,
    )


def pbi_add_stacked_area_chart_tool(
    extract_folder: str,
    page: str,
    axis_column: str,
    value_measures: list[str],
    x: int,
    y: int,
    width: int = 420,
    height: int = 300,
    title: str = "",
    legend_column: str | None = None,
    *,
    manager: Any | None = None,
) -> dict[str, Any]:
    """Stacked area chart — series stacked on top of each other (cumulative trend)."""
    return _add_axis_chart(
        visual_type="stackedAreaChart",
        extract_folder=extract_folder,
        page=page,
        axis_column=axis_column,
        value_measures=value_measures,
        x=x,
        y=y,
        width=width,
        height=height,
        title=title,
        legend_column=legend_column,
        manager=manager,
    )


def pbi_add_hundred_percent_stacked_area_chart_tool(
    extract_folder: str,
    page: str,
    axis_column: str,
    value_measures: list[str],
    x: int,
    y: int,
    width: int = 420,
    height: int = 300,
    title: str = "",
    legend_column: str | None = None,
    *,
    manager: Any | None = None,
) -> dict[str, Any]:
    """100% stacked area chart — proportions of each series over time, normalised to 100%."""
    return _add_axis_chart(
        visual_type="hundredPercentStackedAreaChart",
        extract_folder=extract_folder,
        page=page,
        axis_column=axis_column,
        value_measures=value_measures,
        x=x,
        y=y,
        width=width,
        height=height,
        title=title,
        legend_column=legend_column,
        manager=manager,
    )


def pbi_add_multi_row_card_tool(
    extract_folder: str,
    page: str,
    measures: list[str],
    x: int,
    y: int,
    width: int = 320,
    height: int = 320,
    title: str = "",
    category_column: str | None = None,
    *,
    manager: Any | None = None,
) -> dict[str, Any]:
    """Multi-row card — vertical stack of KPI rows (one row per category value).

    Common executive-dashboard element. With ``category_column`` set, the
    card renders one row per category value; without it, the card shows
    one row per measure (totals view).
    """
    if not measures:
        raise PowerBIValidationError("measures must contain at least one measure.")
    expected_kinds: dict[str, str] = {m: "measure" for m in measures}
    references = list(measures)
    if category_column:
        expected_kinds[category_column] = "column"
        references.append(category_column)
    _validate_field_references_live(manager, references, expected_kinds=expected_kinds)
    measure_home_map = _resolve_measure_home_map(extract_folder, manager=manager)
    projections: dict[str, list[dict[str, str]]] = {
        "Values": [_projection(item) for item in measures],
    }
    if category_column:
        projections["Category"] = [_projection(category_column)]
    return _append_visual(
        extract_folder,
        page,
        lambda section, home_map: _create_chart_container(
            section,
            visual_type="multiRowCard",
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
        "Category": [_projection(category_column)],
        "Y": [_projection(item) for item in bar_measures],
        "Y2": [_projection(item) for item in line_measures],
    }
    references = [category_column, *bar_measures, *line_measures]
    expected_kinds = {category_column: "column"}
    for m in bar_measures:
        expected_kinds[m] = "measure"
    for m in line_measures:
        expected_kinds[m] = "measure"
    if legend_column:
        projections["Series"] = [_projection(legend_column)]
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

"""MCP wrappers — domain: visuals."""

from __future__ import annotations

from typing import Any

from mcp_core import CONNECTION_MANAGER, _run, mcp
from tools import (
    pbi_add_bar_chart_tool,
    pbi_add_card_tool,
    pbi_add_combo_chart_tool,
    pbi_add_donut_chart_tool,
    pbi_add_gauge_tool,
    pbi_add_kpi_tool,
    pbi_add_labelled_card_tool,
    pbi_add_line_chart_tool,
    pbi_add_map_tool,
    pbi_add_matrix_tool,
    pbi_add_scatter_chart_tool,
    pbi_add_slicer_tool,
    pbi_add_table_visual_tool,
    pbi_add_text_box_tool,
    pbi_add_visual_tool,
    pbi_add_waterfall_tool,
    pbi_apply_design_tool,
    pbi_apply_theme_tool,
    pbi_auto_grid_layout_tool,
    pbi_build_dashboard_tool,
    pbi_compile_report_tool,
    pbi_convert_visual_type_tool,
    pbi_create_page_tool,
    pbi_delete_page_tool,
    pbi_describe_page_tool,
    pbi_disable_card_autoscale_tool,
    pbi_extract_report_tool,
    pbi_get_page_tool,
    pbi_list_pages_tool,
    pbi_move_visual_tool,
    pbi_patch_layout_tool,
    pbi_remove_visual_tool,
    pbi_repair_report_fields_tool,
    pbi_set_page_size_tool,
    pbi_set_visual_format_property_tool,
    pbi_validate_report_fields_tool,
)


@mcp.tool()
def pbi_extract_report(pbix_path: str, extract_folder: str | None = None) -> dict[str, Any]:
    """Extract a .pbix report into a pbi-tools folder structure."""
    return _run(
        "pbi_extract_report",
        pbi_extract_report_tool,
        pbix_path=pbix_path,
        extract_folder=extract_folder,
    )


@mcp.tool()
def pbi_compile_report(extract_folder: str, output_path: str, force: bool = False) -> dict[str, Any]:
    """Compile an extracted report folder back into a .pbix."""
    return _run(
        "pbi_compile_report",
        pbi_compile_report_tool,
        extract_folder=extract_folder,
        output_path=output_path,
        force=force,
    )


@mcp.tool()
def pbi_patch_layout(
    extract_folder: str,
    pbix_path: str,
    force: bool = False,
    fail_on_persistence_risk: bool = True,
    include_hidden: bool = False,
) -> dict[str, Any]:
    """Patch Report/Layout directly into an existing .pbix archive."""
    return _run(
        "pbi_patch_layout",
        pbi_patch_layout_tool,
        extract_folder=extract_folder,
        pbix_path=pbix_path,
        force=force,
        fail_on_persistence_risk=fail_on_persistence_risk,
        manager=CONNECTION_MANAGER if fail_on_persistence_risk else None,
        include_hidden=include_hidden,
    )


@mcp.tool()
def pbi_list_pages(extract_folder: str) -> dict[str, Any]:
    """List pages in an extracted report."""
    return _run("pbi_list_pages", pbi_list_pages_tool, extract_folder=extract_folder)


@mcp.tool()
def pbi_validate_report_fields(extract_folder: str, page: str | None = None, include_hidden: bool = False) -> dict[str, Any]:
    """Validate report visual field bindings for broken Power BI visuals."""
    return _run(
        "pbi_validate_report_fields",
        pbi_validate_report_fields_tool,
        extract_folder=extract_folder,
        page=page,
        include_hidden=include_hidden,
        manager=CONNECTION_MANAGER,
    )


@mcp.tool()
def pbi_repair_report_fields(extract_folder: str, page: str | None = None, apply: bool = False) -> dict[str, Any]:
    """Plan or apply deterministic repairs for broken report visual field bindings."""
    return _run(
        "pbi_repair_report_fields",
        pbi_repair_report_fields_tool,
        extract_folder=extract_folder,
        page=page,
        apply=apply,
        manager=CONNECTION_MANAGER,
    )


@mcp.tool()
def pbi_get_page(extract_folder: str, page: str) -> dict[str, Any]:
    """Get page details and visual metadata from an extracted report."""
    return _run("pbi_get_page", pbi_get_page_tool, extract_folder=extract_folder, page=page)


@mcp.tool()
def pbi_convert_visual_type(
    extract_folder: str,
    page: str,
    visual_id: str,
    new_type: str,
) -> dict[str, Any]:
    """Migrate an existing visual to a different type while preserving compatible
    bindings.

    Compatible groups:
    - card ↔ kpi
    - clusteredBarChart ↔ clusteredColumnChart ↔ lineChart ↔
      lineClusteredColumnComboChart
    - donutChart ↔ treemap

    Incompatible source/target combinations are rejected with a structured
    error so an LLM can fall back to delete + recreate explicitly.
    """
    return _run(
        "pbi_convert_visual_type",
        pbi_convert_visual_type_tool,
        extract_folder=extract_folder,
        page=page,
        visual_id=visual_id,
        new_type=new_type,
    )


@mcp.tool()
def pbi_auto_grid_layout(
    specs: list[dict[str, Any]],
    cols: int = 4,
    gap: int = 16,
    start_x: int = 20,
    start_y: int = 60,
    cell_width: int | None = None,
    cell_height: int | None = None,
    page_width: int = 1280,
    page_height: int = 720,
) -> dict[str, Any]:
    """Auto-position a list of visual specs on an N-column grid.

    Each input spec is returned annotated with x/y/width/height so the caller
    can pass it through ``pbi_add_visual`` / ``pbi_build_dashboard`` without
    doing arithmetic. Specs may set ``col_span`` / ``row_span`` to grow over
    neighbouring cells. Pure utility — no live model touch.
    """
    return _run(
        "pbi_auto_grid_layout",
        pbi_auto_grid_layout_tool,
        specs=specs,
        cols=cols,
        gap=gap,
        start_x=start_x,
        start_y=start_y,
        cell_width=cell_width,
        cell_height=cell_height,
        page_width=page_width,
        page_height=page_height,
    )


@mcp.tool()
def pbi_describe_page(extract_folder: str, page: str) -> dict[str, Any]:
    """LLM-friendly snapshot of a page: per-visual id/type/position/bindings/
    formatting (title, axis titles, label_display_units), plus a
    ``binding_health`` rollup. When a Power BI Desktop is connected, bindings
    are checked against the live model so missing fields surface in the same
    response without parsing layout JSON.
    """
    return _run(
        "pbi_describe_page",
        pbi_describe_page_tool,
        extract_folder=extract_folder,
        page=page,
        manager=CONNECTION_MANAGER,
    )


@mcp.tool()
def pbi_create_page(
    extract_folder: str,
    display_name: str,
    width: int = 1280,
    height: int = 720,
) -> dict[str, Any]:
    """Create a new report page."""
    return _run(
        "pbi_create_page",
        pbi_create_page_tool,
        extract_folder=extract_folder,
        display_name=display_name,
        width=width,
        height=height,
    )


@mcp.tool()
def pbi_delete_page(extract_folder: str, page: str) -> dict[str, Any]:
    """Delete a report page."""
    return _run("pbi_delete_page", pbi_delete_page_tool, extract_folder=extract_folder, page=page)


@mcp.tool()
def pbi_set_page_size(extract_folder: str, page: str, width: int, height: int) -> dict[str, Any]:
    """Resize a report page."""
    return _run(
        "pbi_set_page_size",
        pbi_set_page_size_tool,
        extract_folder=extract_folder,
        page=page,
        width=width,
        height=height,
    )


@mcp.tool()
def pbi_add_card(
    extract_folder: str,
    page: str,
    measure: str,
    x: int,
    y: int,
    width: int = 200,
    height: int = 120,
    title: str = "",
) -> dict[str, Any]:
    """Add a card visual to a report page.

    The referenced measure is validated against the live model when a Power BI
    Desktop instance is connected; a missing measure fails fast with a
    structured ``validation_error`` instead of silently producing a broken
    visual.
    """
    return _run(
        "pbi_add_card",
        pbi_add_card_tool,
        extract_folder=extract_folder,
        page=page,
        measure=measure,
        x=x,
        y=y,
        width=width,
        height=height,
        title=title,
        manager=CONNECTION_MANAGER,
    )


@mcp.tool()
def pbi_add_bar_chart(
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
) -> dict[str, Any]:
    """Add a clustered bar chart visual.

    Columns must be qualified: ``Table[Column]``, ``'Table'[Column]``, or
    ``Table.Column``. Bare names like ``Year`` are rejected.
    """
    return _run(
        "pbi_add_bar_chart",
        pbi_add_bar_chart_tool,
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
        manager=CONNECTION_MANAGER,
    )


@mcp.tool()
def pbi_add_line_chart(
    extract_folder: str,
    page: str,
    axis_column: str,
    value_measures: list[str],
    x: int,
    y: int,
    width: int = 420,
    height: int = 300,
    title: str = "",
) -> dict[str, Any]:
    """Add a line chart visual.

    Columns must be qualified with the table: ``Table[Column]``, ``'Table'[Column]``,
    or ``Table.Column``. Bare names like ``Year`` are rejected — use ``Date.Year``.
    """
    return _run(
        "pbi_add_line_chart",
        pbi_add_line_chart_tool,
        extract_folder=extract_folder,
        page=page,
        axis_column=axis_column,
        value_measures=value_measures,
        x=x,
        y=y,
        width=width,
        height=height,
        title=title,
        manager=CONNECTION_MANAGER,
    )


@mcp.tool()
def pbi_add_donut_chart(
    extract_folder: str,
    page: str,
    category_column: str,
    value_measure: str,
    x: int,
    y: int,
    width: int = 320,
    height: int = 280,
    title: str = "",
) -> dict[str, Any]:
    """Add a donut chart visual.

    Columns must be qualified: ``Table[Column]``, ``'Table'[Column]``, or ``Table.Column``.
    """
    return _run(
        "pbi_add_donut_chart",
        pbi_add_donut_chart_tool,
        extract_folder=extract_folder,
        page=page,
        category_column=category_column,
        value_measure=value_measure,
        x=x,
        y=y,
        width=width,
        height=height,
        title=title,
        manager=CONNECTION_MANAGER,
    )


@mcp.tool()
def pbi_add_table_visual(
    extract_folder: str,
    page: str,
    columns: list[str],
    x: int,
    y: int,
    width: int = 520,
    height: int = 320,
    title: str = "",
) -> dict[str, Any]:
    """Add a table visual.

    Each column must be qualified: ``Table[Column]``, ``'Table'[Column]``, or
    ``Table.Column``. Bare measure names are accepted.
    """
    return _run(
        "pbi_add_table_visual",
        pbi_add_table_visual_tool,
        extract_folder=extract_folder,
        page=page,
        columns=columns,
        x=x,
        y=y,
        width=width,
        height=height,
        title=title,
        manager=CONNECTION_MANAGER,
    )


@mcp.tool()
def pbi_add_scatter_chart(
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
) -> dict[str, Any]:
    """Add a scatter chart. Roles: Category (column), X (measure), Y (measure),
    Size (measure, optional), Series (column, optional). Use for correlation
    analysis between two measures grouped by a dimension.

    Columns must be qualified: ``Table[Column]`` / ``'Table'[Column]`` / ``Table.Column``.
    """
    return _run(
        "pbi_add_scatter_chart",
        pbi_add_scatter_chart_tool,
        extract_folder=extract_folder,
        page=page,
        category_column=category_column,
        x_measure=x_measure,
        y_measure=y_measure,
        x=x,
        y=y,
        width=width,
        height=height,
        title=title,
        size_measure=size_measure,
        legend_column=legend_column,
        manager=CONNECTION_MANAGER,
    )


@mcp.tool()
def pbi_add_combo_chart(
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
) -> dict[str, Any]:
    """Add a combo chart (bars + line overlay). Roles: Category (column),
    Y (bar measures, list), Y2 (line measures, list). Use for actual vs target
    comparisons where the target is rendered as a line over actual bars."""
    return _run(
        "pbi_add_combo_chart",
        pbi_add_combo_chart_tool,
        extract_folder=extract_folder,
        page=page,
        category_column=category_column,
        bar_measures=bar_measures,
        line_measures=line_measures,
        x=x,
        y=y,
        width=width,
        height=height,
        title=title,
        legend_column=legend_column,
        manager=CONNECTION_MANAGER,
    )


@mcp.tool()
def pbi_add_kpi(
    extract_folder: str,
    page: str,
    indicator_measure: str,
    trend_axis_column: str,
    x: int,
    y: int,
    width: int = 240,
    height: int = 160,
    title: str = "",
    goal_measure: str | None = None,
    direction: str = "high_is_good",
) -> dict[str, Any]:
    """Add a native KPI visual. Roles: Indicator (measure), TrendLine (column,
    typically Date), Goal (measure, optional). ``direction`` controls status
    colour: ``"high_is_good"`` (green when actual > goal) or ``"low_is_good"``."""
    return _run(
        "pbi_add_kpi",
        pbi_add_kpi_tool,
        extract_folder=extract_folder,
        page=page,
        indicator_measure=indicator_measure,
        trend_axis_column=trend_axis_column,
        x=x,
        y=y,
        width=width,
        height=height,
        title=title,
        goal_measure=goal_measure,
        direction=direction,
        manager=CONNECTION_MANAGER,
    )


@mcp.tool()
def pbi_add_map(
    extract_folder: str,
    page: str,
    location_column: str,
    x: int,
    y: int,
    value_measure: str | None = None,
    width: int = 420,
    height: int = 320,
    title: str = "",
) -> dict[str, Any]:
    """Add a map / bubble visual. Roles: location_column (geographic field —
    country, city, lat/long…) and optional value_measure (drives bubble size).

    References accept ``Table.Column``, ``Table[Column]``, or
    ``'Table With Spaces'[Column]``.
    """
    return _run(
        "pbi_add_map",
        pbi_add_map_tool,
        extract_folder=extract_folder,
        page=page,
        location_column=location_column,
        value_measure=value_measure,
        x=x,
        y=y,
        width=width,
        height=height,
        title=title,
        manager=CONNECTION_MANAGER,
    )


@mcp.tool()
def pbi_add_matrix(
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
) -> dict[str, Any]:
    """Add a matrix / pivot-table visual. Roles: Rows (column list — required),
    Columns (column list — optional), Values (measure list — required).
    ``column_layout``: ``"stepped"`` (compact) or ``"tabular"`` (one column per
    row level). Matches the docx-style multi-dim table common in business
    reports."""
    return _run(
        "pbi_add_matrix",
        pbi_add_matrix_tool,
        extract_folder=extract_folder,
        page=page,
        rows=rows,
        values=values,
        x=x,
        y=y,
        columns=columns,
        width=width,
        height=height,
        title=title,
        subtotals=subtotals,
        column_layout=column_layout,
        manager=CONNECTION_MANAGER,
    )


@mcp.tool()
def pbi_add_waterfall(
    extract_folder: str,
    page: str,
    category_column: str,
    value_measure: str,
    x: int,
    y: int,
    width: int = 420,
    height: int = 300,
    title: str = "",
) -> dict[str, Any]:
    """Add a waterfall chart visual.

    Columns must be qualified: ``Table[Column]``, ``'Table'[Column]``, or ``Table.Column``.
    """
    return _run(
        "pbi_add_waterfall",
        pbi_add_waterfall_tool,
        extract_folder=extract_folder,
        page=page,
        category_column=category_column,
        value_measure=value_measure,
        x=x,
        y=y,
        width=width,
        height=height,
        title=title,
        manager=CONNECTION_MANAGER,
    )


@mcp.tool()
def pbi_add_slicer(
    extract_folder: str,
    page: str,
    column: str,
    x: int,
    y: int,
    width: int = 220,
    height: int = 120,
    slicer_type: str = "dropdown",
) -> dict[str, Any]:
    """Add a slicer visual. slicer_type: dropdown | list | range | tile (horizontal list).

    The slicer ``column`` must be qualified: ``Table[Column]``, ``'Table'[Column]``,
    or ``Table.Column``.
    """
    return _run(
        "pbi_add_slicer",
        pbi_add_slicer_tool,
        extract_folder=extract_folder,
        page=page,
        column=column,
        x=x,
        y=y,
        width=width,
        height=height,
        slicer_type=slicer_type,
        manager=CONNECTION_MANAGER,
    )


@mcp.tool()
def pbi_add_gauge(
    extract_folder: str,
    page: str,
    measure: str,
    x: int,
    y: int,
    width: int = 280,
    height: int = 220,
    title: str = "",
    target_measure: str | None = None,
    min_value: float | None = None,
    max_value: float | None = None,
    target_value: float | None = None,
    fill_color: str | None = None,
    target_color: str | None = None,
    fill_color_measure: str | None = None,
) -> dict[str, Any]:
    """Add a gauge visual.

    Optional kwargs override the default 0-100 axis and blue fill:
    - min_value / max_value: gauge range bounds (constants)
    - target_value: marker on the arc (constant; alternative to target_measure)
    - fill_color / target_color: '#RRGGBB' arc + target colors
    - fill_color_measure: name of a DAX measure returning '#RRGGBB' for
      conditional formatting (overrides fill_color, reacts to slicer context)
    """
    return _run(
        "pbi_add_gauge",
        pbi_add_gauge_tool,
        extract_folder=extract_folder,
        page=page,
        measure=measure,
        x=x,
        y=y,
        width=width,
        height=height,
        title=title,
        target_measure=target_measure,
        min_value=min_value,
        max_value=max_value,
        target_value=target_value,
        fill_color=fill_color,
        target_color=target_color,
        fill_color_measure=fill_color_measure,
        manager=CONNECTION_MANAGER,
    )


@mcp.tool()
def pbi_add_labelled_card(
    extract_folder: str,
    page: str,
    measure: str,
    label: str,
    x: int,
    y: int,
    width: int = 220,
    height: int = 110,
    label_height: int = 28,
    label_font_size: int = 11,
    label_bold: bool = True,
    label_color: str = "#1F2937",
) -> dict[str, Any]:
    """Add a docx-style card: text label on top, value card underneath.

    Returns both visual ids under visuals.label and visuals.value.
    """
    return _run(
        "pbi_add_labelled_card",
        pbi_add_labelled_card_tool,
        extract_folder=extract_folder,
        page=page,
        measure=measure,
        label=label,
        x=x,
        y=y,
        width=width,
        height=height,
        label_height=label_height,
        label_font_size=label_font_size,
        label_bold=label_bold,
        label_color=label_color,
        manager=CONNECTION_MANAGER,
    )


@mcp.tool()
def pbi_add_text_box(
    extract_folder: str,
    page: str,
    text: str,
    x: int,
    y: int,
    width: int = 280,
    height: int = 80,
    font_size: int = 16,
    bold: bool = False,
    color: str = "#222222",
) -> dict[str, Any]:
    """Add a text box visual."""
    return _run(
        "pbi_add_text_box",
        pbi_add_text_box_tool,
        extract_folder=extract_folder,
        page=page,
        text=text,
        x=x,
        y=y,
        width=width,
        height=height,
        font_size=font_size,
        bold=bold,
        color=color,
    )


@mcp.tool()
def pbi_remove_visual(extract_folder: str, page: str, visual_id: str) -> dict[str, Any]:
    """Remove a visual from a report page."""
    return _run(
        "pbi_remove_visual",
        pbi_remove_visual_tool,
        extract_folder=extract_folder,
        page=page,
        visual_id=visual_id,
    )


@mcp.tool()
def pbi_move_visual(
    extract_folder: str,
    page: str,
    visual_id: str,
    x: int,
    y: int,
    width: int | None = None,
    height: int | None = None,
) -> dict[str, Any]:
    """Move or resize a visual."""
    return _run(
        "pbi_move_visual",
        pbi_move_visual_tool,
        extract_folder=extract_folder,
        page=page,
        visual_id=visual_id,
        x=x,
        y=y,
        width=width,
        height=height,
    )


@mcp.tool()
def pbi_set_visual_format_property(
    extract_folder: str,
    page: str,
    visual_id: str,
    object_name: str,
    properties: dict[str, Any],
    property_types: dict[str, str] | None = None,
) -> dict[str, Any]:
    """Override format properties on an existing visual.

    object_name examples: ``title``, ``categoryAxis``, ``valueAxis``,
    ``labels``, ``dataPoint``, ``general``. Properties are merged into
    ``singleVisual.objects[object_name][0].properties`` and encoded to PBI's
    canonical literal forms (single-quoted text, ``L`` int suffix, ``D``
    decimal suffix, ``#RRGGBB`` solid color, ``true``/``false`` bool).

    Pass property_types to force an encoding ("text", "bool", "int",
    "decimal", "color", or "raw" for pre-shaped expr dicts).
    """
    return _run(
        "pbi_set_visual_format_property",
        pbi_set_visual_format_property_tool,
        extract_folder=extract_folder,
        page=page,
        visual_id=visual_id,
        object_name=object_name,
        properties=properties,
        property_types=property_types,
    )


@mcp.tool()
def pbi_disable_card_autoscale(
    extract_folder: str,
    page: str | None = None,
    visual_ids: list[str] | None = None,
    label_precision: int = 0,
) -> dict[str, Any]:
    """Disable Power BI's auto K/M unit-scaling on card visuals.

    Sets ``labelDisplayUnits=1`` (None) and an explicit ``labelPrecision``
    on every card. Use to fix the "119K K €" double-suffix bug where a
    measure already pre-divides by 1000 with a ``K €`` format string and
    PBI auto-scales on top.
    """
    return _run(
        "pbi_disable_card_autoscale",
        pbi_disable_card_autoscale_tool,
        extract_folder=extract_folder,
        page=page,
        visual_ids=visual_ids,
        label_precision=label_precision,
    )


@mcp.tool()
def pbi_apply_theme(extract_folder: str, theme_json_path: str) -> dict[str, Any]:
    """Apply a theme JSON to an extracted report."""
    return _run(
        "pbi_apply_theme",
        pbi_apply_theme_tool,
        extract_folder=extract_folder,
        theme_json_path=theme_json_path,
    )


@mcp.tool()
def pbi_apply_design(
    extract_folder: str,
    preset: str = "powerbi-navy-pro",
    page_background: str | None = "#F0F4FB",
    style_cards: bool = True,
) -> dict[str, Any]:
    """Apply a complete visual design preset (theme + page background + card styling)."""
    return _run(
        "pbi_apply_design",
        pbi_apply_design_tool,
        extract_folder=extract_folder,
        preset=preset,
        page_background=page_background,
        style_cards=style_cards,
    )


@mcp.tool()
def pbi_add_visual(
    extract_folder: str,
    page: str,
    visual_type: str,
    x: int,
    y: int,
    width: int | None = None,
    height: int | None = None,
    title: str = "",
    config: dict[str, Any] | None = None,
    dry_run: bool = False,
) -> dict[str, Any]:
    """Add any visual via a generic dispatcher.

    visual_type: card | bar_chart | line_chart | donut | table | waterfall |
    slicer | gauge | kpi | scatter_chart | combo_chart | matrix | map |
    text_box | labelled_card.

    Set ``dry_run=True`` to validate and resolve bindings without writing the
    layout to disk — the response carries ``dry_run=True`` and a ``write_log``
    summarising what would have been persisted. Use it to preview a change
    before committing.

    The active connection manager is forwarded to every dispatcher so live
    field validation and home-table resolution happen — no
    ``measure_home_table_needs_repair`` after the write.
    """
    return _run(
        "pbi_add_visual",
        pbi_add_visual_tool,
        extract_folder=extract_folder,
        page=page,
        visual_type=visual_type,
        x=x,
        y=y,
        width=width,
        height=height,
        title=title,
        config=config,
        manager=CONNECTION_MANAGER,
        dry_run=dry_run,
    )


@mcp.tool()
def pbi_build_dashboard(extract_folder: str, page: str, layout: list[dict[str, Any]]) -> dict[str, Any]:
    """Build a dashboard page from a bulk layout specification."""
    return _run(
        "pbi_build_dashboard",
        pbi_build_dashboard_tool,
        extract_folder=extract_folder,
        page=page,
        layout=layout,
    )

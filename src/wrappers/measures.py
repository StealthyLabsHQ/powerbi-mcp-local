"""MCP wrappers — domain: measures."""

from __future__ import annotations

from typing import Any

from mcp_core import CONNECTION_MANAGER, _run, mcp
from tools import (
    pbi_apply_format_preset_tool,
    pbi_create_contribution_measure_tool,
    pbi_create_measure_tool,
    pbi_create_measures_tool,
    pbi_create_mtd_measure_tool,
    pbi_create_rolling_average_measure_tool,
    pbi_create_spy_measure_tool,
    pbi_create_time_intelligence_pack_tool,
    pbi_create_topn_measure_tool,
    pbi_create_variance_measure_tool,
    pbi_create_yoy_measure_tool,
    pbi_create_ytd_measure_tool,
    pbi_delete_measure_tool,
    pbi_import_dax_file_tool,
    pbi_list_format_presets_tool,
    pbi_list_measures_tool,
    pbi_measure_dependencies_tool,
    pbi_rename_measure_tool,
    pbi_set_format_tool,
)


@mcp.tool()
def pbi_list_measures(include_hidden: bool = False) -> dict[str, Any]:
    """List DAX measures in the active Power BI model."""
    return _run(
        "pbi_list_measures",
        pbi_list_measures_tool,
        CONNECTION_MANAGER,
        include_hidden=include_hidden,
    )


@mcp.tool()
def pbi_create_measure(
    table: str,
    name: str,
    expression: str,
    format_string: str = "",
    description: str = "",
    display_folder: str = "",
    is_hidden: bool = False,
    overwrite: bool = True,
) -> dict[str, Any]:
    """Create or update a DAX measure."""
    return _run(
        "pbi_create_measure",
        pbi_create_measure_tool,
        CONNECTION_MANAGER,
        table=table,
        name=name,
        expression=expression,
        format_string=format_string,
        description=description,
        display_folder=display_folder,
        is_hidden=is_hidden,
        overwrite=overwrite,
    )


@mcp.tool()
def pbi_create_measures(
    table: str,
    measures: list[dict[str, Any]],
    overwrite: bool = True,
    stop_on_error: bool = False,
    dry_run: bool = False,
) -> dict[str, Any]:
    """Batch-create or update multiple DAX measures with a single SaveChanges call.

    measures: list of {name, expression, format_string?, description?, display_folder?, is_hidden?}

    With ``dry_run=True`` the model is not modified; the response carries a
    ``plan`` describing the per-measure ``would_create``/``would_update``/
    ``would_fail`` outcome — useful as a preflight before committing.
    """
    return _run(
        "pbi_create_measures",
        pbi_create_measures_tool,
        CONNECTION_MANAGER,
        table=table,
        measures=measures,
        overwrite=overwrite,
        stop_on_error=stop_on_error,
        dry_run=dry_run,
    )


@mcp.tool()
def pbi_create_time_intelligence_pack(
    table: str,
    base_measure: str,
    date_table: str,
    date_column: str,
    patterns: list[str] | None = None,
    display_folder: str = "Time intelligence",
    format_inherit: bool = True,
    format_string: str = "",
    overwrite: bool = False,
    dry_run: bool = False,
) -> dict[str, Any]:
    """Create a family of time-intelligence measures from one base measure.

    Default patterns: YTD, MTD, QTD, SPY, YOY, YOY %, MA3. Generated names
    are ``"{base} {suffix}"`` (e.g. ``"Sales YTD"``). Dependency-aware:
    requesting ``YOY%`` auto-adds ``YOY`` and ``SPY``. Use ``dry_run=True``
    to preview the plan without writing.
    """
    return _run(
        "pbi_create_time_intelligence_pack",
        pbi_create_time_intelligence_pack_tool,
        CONNECTION_MANAGER,
        table=table,
        base_measure=base_measure,
        date_table=date_table,
        date_column=date_column,
        patterns=patterns,
        display_folder=display_folder,
        format_inherit=format_inherit,
        format_string=format_string,
        overwrite=overwrite,
        dry_run=dry_run,
    )


@mcp.tool()
def pbi_create_ytd_measure(
    table: str,
    base_measure: str,
    date_table: str,
    date_column: str,
    format_string: str = "",
    overwrite: bool = False,
) -> dict[str, Any]:
    """Create a single YTD companion of ``base_measure``."""
    return _run(
        "pbi_create_ytd_measure",
        pbi_create_ytd_measure_tool,
        CONNECTION_MANAGER,
        table=table,
        base_measure=base_measure,
        date_table=date_table,
        date_column=date_column,
        format_string=format_string,
        overwrite=overwrite,
    )


@mcp.tool()
def pbi_create_mtd_measure(
    table: str,
    base_measure: str,
    date_table: str,
    date_column: str,
    format_string: str = "",
    overwrite: bool = False,
) -> dict[str, Any]:
    """Create a single MTD companion of ``base_measure``."""
    return _run(
        "pbi_create_mtd_measure",
        pbi_create_mtd_measure_tool,
        CONNECTION_MANAGER,
        table=table,
        base_measure=base_measure,
        date_table=date_table,
        date_column=date_column,
        format_string=format_string,
        overwrite=overwrite,
    )


@mcp.tool()
def pbi_create_spy_measure(
    table: str,
    base_measure: str,
    date_table: str,
    date_column: str,
    format_string: str = "",
    overwrite: bool = False,
) -> dict[str, Any]:
    """Create the Same-Period-Last-Year companion of ``base_measure``."""
    return _run(
        "pbi_create_spy_measure",
        pbi_create_spy_measure_tool,
        CONNECTION_MANAGER,
        table=table,
        base_measure=base_measure,
        date_table=date_table,
        date_column=date_column,
        format_string=format_string,
        overwrite=overwrite,
    )


@mcp.tool()
def pbi_create_yoy_measure(
    table: str,
    base_measure: str,
    date_table: str,
    date_column: str,
    format_string: str = "",
    overwrite: bool = False,
) -> dict[str, Any]:
    """Create the Year-over-Year delta + SPY companion of ``base_measure``."""
    return _run(
        "pbi_create_yoy_measure",
        pbi_create_yoy_measure_tool,
        CONNECTION_MANAGER,
        table=table,
        base_measure=base_measure,
        date_table=date_table,
        date_column=date_column,
        format_string=format_string,
        overwrite=overwrite,
    )


@mcp.tool()
def pbi_create_variance_measure(
    table: str,
    base_measure: str,
    date_table: str,
    date_column: str,
    measure_name: str | None = None,
    compare_period_offset: int = -1,
    granularity: str = "year",
    format_string: str = "",
    display_folder: str = "Variance",
    overwrite: bool = False,
) -> dict[str, Any]:
    """Create a period-over-period variance measure.

    ``[base] - CALCULATE([base], DATEADD(date_table[date_column], offset, granularity))``.
    Default offset = -1 ⇒ "current vs previous period".
    """
    return _run(
        "pbi_create_variance_measure",
        pbi_create_variance_measure_tool,
        CONNECTION_MANAGER,
        table=table,
        base_measure=base_measure,
        date_table=date_table,
        date_column=date_column,
        measure_name=measure_name,
        compare_period_offset=compare_period_offset,
        granularity=granularity,
        format_string=format_string,
        display_folder=display_folder,
        overwrite=overwrite,
    )


@mcp.tool()
def pbi_create_contribution_measure(
    table: str,
    base_measure: str,
    scope_columns: list[str],
    measure_name: str | None = None,
    format_string: str = "0.00%",
    display_folder: str = "Contribution",
    overwrite: bool = False,
) -> dict[str, Any]:
    """Create a %-of-total contribution measure.

    ``DIVIDE([base], CALCULATE([base], ALL(scope_columns)))``. ``scope_columns``
    take the ``"Table.Column"`` form, e.g. ``["Categorie.Nom catégorie"]``.
    """
    return _run(
        "pbi_create_contribution_measure",
        pbi_create_contribution_measure_tool,
        CONNECTION_MANAGER,
        table=table,
        base_measure=base_measure,
        scope_columns=scope_columns,
        measure_name=measure_name,
        format_string=format_string,
        display_folder=display_folder,
        overwrite=overwrite,
    )


@mcp.tool()
def pbi_create_topn_measure(
    table: str,
    base_measure: str,
    n: int,
    dimension_table: str,
    dimension_column: str,
    measure_name: str | None = None,
    rank_measure: str | None = None,
    format_string: str = "",
    display_folder: str = "Top-N",
    overwrite: bool = False,
) -> dict[str, Any]:
    """Create a Top-N filter measure using RANKX over ALL(dim_table[dim_column])."""
    return _run(
        "pbi_create_topn_measure",
        pbi_create_topn_measure_tool,
        CONNECTION_MANAGER,
        table=table,
        base_measure=base_measure,
        n=n,
        dimension_table=dimension_table,
        dimension_column=dimension_column,
        measure_name=measure_name,
        rank_measure=rank_measure,
        format_string=format_string,
        display_folder=display_folder,
        overwrite=overwrite,
    )


@mcp.tool()
def pbi_create_rolling_average_measure(
    table: str,
    base_measure: str,
    window: int,
    date_table: str,
    date_column: str,
    granularity: str = "month",
    measure_name: str | None = None,
    format_string: str = "",
    display_folder: str = "Rolling",
    overwrite: bool = False,
) -> dict[str, Any]:
    """Create a trailing rolling-average measure using DATESINPERIOD."""
    return _run(
        "pbi_create_rolling_average_measure",
        pbi_create_rolling_average_measure_tool,
        CONNECTION_MANAGER,
        table=table,
        base_measure=base_measure,
        window=window,
        date_table=date_table,
        date_column=date_column,
        granularity=granularity,
        measure_name=measure_name,
        format_string=format_string,
        display_folder=display_folder,
        overwrite=overwrite,
    )


@mcp.tool()
def pbi_list_format_presets(filter_substring: str | None = None) -> dict[str, Any]:
    """Return the format-string preset catalogue (currency, percent, thousands,
    millions, dates, etc.). Pass ``filter_substring`` to narrow the result.
    """
    return _run(
        "pbi_list_format_presets",
        pbi_list_format_presets_tool,
        filter_substring=filter_substring,
    )


@mcp.tool()
def pbi_apply_format_preset(
    table: str,
    names: list[str],
    preset: str,
    object_type: str = "measure",
) -> dict[str, Any]:
    """Apply a named format preset to a list of measures (or columns).

    Examples of preset names: ``currency_eur_k``, ``percent_4dp``, ``thousands``,
    ``date_iso``, ``date_short_fr``. See ``pbi_list_format_presets`` for the
    full catalogue.
    """
    return _run(
        "pbi_apply_format_preset",
        pbi_apply_format_preset_tool,
        CONNECTION_MANAGER,
        table=table,
        names=names,
        preset=preset,
        object_type=object_type,
    )


@mcp.tool()
def pbi_delete_measure(table: str, name: str) -> dict[str, Any]:
    """Delete a DAX measure."""
    return _run(
        "pbi_delete_measure",
        pbi_delete_measure_tool,
        CONNECTION_MANAGER,
        table=table,
        name=name,
    )


@mcp.tool()
def pbi_rename_measure(table: str, name: str, new_name: str) -> dict[str, Any]:
    """Rename a DAX measure. Dependent DAX expressions must be updated separately."""
    return _run(
        "pbi_rename_measure",
        pbi_rename_measure_tool,
        CONNECTION_MANAGER,
        table=table,
        name=name,
        new_name=new_name,
    )


@mcp.tool()
def pbi_measure_dependencies(
    measure: str | None = None,
    table: str | None = None,
) -> dict[str, Any]:
    """Return DISCOVER_CALC_DEPENDENCY rows, optionally filtered by measure/table."""
    return _run(
        "pbi_measure_dependencies",
        pbi_measure_dependencies_tool,
        CONNECTION_MANAGER,
        measure=measure,
        table=table,
    )


@mcp.tool()
def pbi_import_dax_file(
    path: str,
    table: str = "Measures",
    overwrite: bool = True,
    default_format_string: str = "",
    default_display_folder: str = "",
    stop_on_error: bool = False,
) -> dict[str, Any]:
    """Bulk-create measures from a .dax file."""
    return _run(
        "pbi_import_dax_file",
        pbi_import_dax_file_tool,
        CONNECTION_MANAGER,
        path=path,
        table=table,
        overwrite=overwrite,
        default_format_string=default_format_string,
        default_display_folder=default_display_folder,
        stop_on_error=stop_on_error,
    )


@mcp.tool()
def pbi_set_format(
    table: str,
    names: list[str],
    format_string: str,
    object_type: str = "measure",
) -> dict[str, Any]:
    """Batch-apply a format string to measures or columns."""
    return _run(
        "pbi_set_format",
        pbi_set_format_tool,
        CONNECTION_MANAGER,
        table=table,
        names=names,
        format_string=format_string,
        object_type=object_type,
    )

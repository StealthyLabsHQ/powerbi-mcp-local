"""FastMCP server exposing Power BI Desktop model operations over stdio."""

from __future__ import annotations

import os
from pathlib import Path
from typing import Any

from mcp.server.fastmcp import FastMCP

from pbi_connection import PowerBIConnectionManager, error_payload, logger
from security import SECURITY
from tools import (
    pbi_create_measures_tool,
    pbi_validate_model_tool,
    excel_auto_width_tool,
    excel_create_sheet_tool,
    excel_create_workbook_tool,
    excel_delete_sheet_tool,
    excel_format_range_tool,
    excel_list_sheets_tool,
    excel_read_cell_tool,
    excel_read_sheet_tool,
    excel_search_tool,
    excel_to_pbi_check_tool,
    excel_workbook_info_tool,
    excel_write_cell_tool,
    excel_write_range_tool,
    pbi_connect_tool,
    pbi_create_column_tool,
    pbi_create_measure_tool,
    pbi_create_relationship_tool,
    pbi_create_table_tool,
    pbi_delete_column_tool,
    pbi_delete_measure_tool,
    pbi_delete_relationship_tool,
    pbi_delete_table_tool,
    pbi_execute_dax_as_role_tool,
    pbi_execute_dax_tool,
    pbi_export_model_tool,
    pbi_import_dax_file_tool,
    pbi_list_instances_tool,
    pbi_list_measures_tool,
    pbi_list_relationships_tool,
    pbi_list_tables_tool,
    pbi_measure_dependencies_tool,
    pbi_model_info_tool,
    pbi_refresh_metadata_tool,
    pbi_refresh_tool,
    pbi_rename_column_tool,
    pbi_rename_measure_tool,
    pbi_rename_table_tool,
    pbi_set_format_tool,
    pbi_trace_query_tool,
    pbi_update_relationship_tool,
    pbi_validate_dax_tool,
    pbi_bulk_import_excel_tool,
    pbi_create_csv_import_query_tool,
    pbi_create_folder_import_query_tool,
    pbi_create_import_query_tool,
    pbi_import_excel_workbook_tool,
    pbi_add_bar_chart_tool,
    pbi_add_card_tool,
    pbi_add_donut_chart_tool,
    pbi_add_gauge_tool,
    pbi_add_line_chart_tool,
    pbi_add_role_member_tool,
    pbi_add_slicer_tool,
    pbi_add_table_visual_tool,
    pbi_add_text_box_tool,
    pbi_add_visual_tool,
    pbi_add_waterfall_tool,
    pbi_create_calc_group_tool,
    pbi_create_role_tool,
    pbi_delete_calc_group_tool,
    pbi_delete_role_tool,
    pbi_list_calc_groups_tool,
    pbi_list_roles_tool,
    pbi_remove_role_member_tool,
    pbi_set_role_filter_tool,
    pbi_apply_design_tool,
    pbi_apply_theme_tool,
    pbi_build_dashboard_tool,
    pbi_compile_report_tool,
    pbi_create_page_tool,
    pbi_delete_page_tool,
    pbi_extract_report_tool,
    pbi_get_power_query_tool,
    pbi_get_page_tool,
    pbi_list_pages_tool,
    pbi_list_power_queries_tool,
    pbi_move_visual_tool,
    pbi_patch_layout_tool,
    pbi_remove_visual_tool,
    pbi_set_power_query_tool,
    pbi_set_page_size_tool,
)


mcp = FastMCP(
    "powerbi-desktop",
    instructions=(
        "Connects to the local Power BI Desktop Analysis Services instance, "
        "lets clients inspect the semantic model, manage measures and "
        "relationships, run DAX queries, trigger model refreshes, manage "
        "Power Query partitions, and read or write Excel workbooks used in "
        "the Power BI pipeline. It can also extract, modify, and compile "
        "report layouts for page and visual automation."
    ),
    json_response=True,
    log_level="INFO",
)

CONNECTION_MANAGER = PowerBIConnectionManager(logger)


def _run(tool_name: str, callback: Any, *args: Any, **kwargs: Any) -> dict[str, Any]:
    """Execute a tool callback with audit logging and error normalization."""
    # Audit log: every tool call, before execution
    safe_kwargs = {
        key: SECURITY.sanitize_for_logging(value)
        for key, value in kwargs.items()
        if key != "manager" and not key.startswith("_")
    }
    logger.info("TOOL_CALL tool=%s params=%s", tool_name, safe_kwargs)
    try:
        SECURITY.validate_tool_call(tool_name, kwargs)
        result = callback(*args, **kwargs)
        status = result.get("status", "unknown") if isinstance(result, dict) else "ok"
        logger.info("TOOL_OK tool=%s status=%s", tool_name, status)
        return result
    except Exception as exc:
        logger.warning("TOOL_FAIL tool=%s error=%s", tool_name, str(exc)[:300])
        logger.exception("Tool '%s' failed", tool_name)
        return error_payload(exc)


def find_pbi_port(preferred_port: int | None = None) -> int:
    """Compatibility helper for standalone scripts and README examples."""
    instances = CONNECTION_MANAGER.list_instances()
    if not instances:
        raise ValueError("No running Power BI Desktop instances were found.")
    if preferred_port is None:
        return int(instances[0]["port"])
    for instance in instances:
        if instance["port"] == preferred_port:
            return int(instance["port"])
    raise ValueError(f"No Power BI instance found on port {preferred_port}.")


@mcp.tool()
def pbi_connect(preferred_port: int | None = None, force_reconnect: bool = False) -> dict[str, Any]:
    """Find and connect to a running Power BI Desktop instance."""
    return _run(
        "pbi_connect",
        pbi_connect_tool,
        CONNECTION_MANAGER,
        preferred_port=preferred_port,
        force_reconnect=force_reconnect,
    )


@mcp.tool()
def pbi_list_tables(
    include_hidden: bool = False,
    include_row_counts: bool = False,
) -> dict[str, Any]:
    """List tables and columns in the active Power BI model."""
    return _run(
        "pbi_list_tables",
        pbi_list_tables_tool,
        CONNECTION_MANAGER,
        include_hidden=include_hidden,
        include_row_counts=include_row_counts,
    )


@mcp.tool()
def pbi_list_instances() -> dict[str, Any]:
    """List discovered Power BI Desktop instances without connecting."""
    return _run("pbi_list_instances", pbi_list_instances_tool, CONNECTION_MANAGER)


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
def pbi_list_relationships() -> dict[str, Any]:
    """List relationships in the active Power BI model."""
    return _run("pbi_list_relationships", pbi_list_relationships_tool, CONNECTION_MANAGER)


@mcp.tool()
def pbi_execute_dax(
    query: str,
    max_rows: int = 1000,
    timeout_seconds: int | None = None,
) -> dict[str, Any]:
    """Execute a DAX or DMV query and return rows. timeout_seconds=0 disables the timeout."""
    return _run(
        "pbi_execute_dax",
        pbi_execute_dax_tool,
        CONNECTION_MANAGER,
        query=query,
        max_rows=max_rows,
        timeout_seconds=timeout_seconds,
    )


@mcp.tool()
def pbi_execute_dax_as_role(query: str, role: str, username: str | None = None) -> dict[str, Any]:
    """Execute a DAX query under a specific RLS role and optional effective user."""
    return _run(
        "pbi_execute_dax_as_role",
        pbi_execute_dax_as_role_tool,
        CONNECTION_MANAGER,
        query=query,
        role=role,
        username=username,
    )


@mcp.tool()
def pbi_trace_query(
    query: str,
    timeout_seconds: int | None = None,
) -> dict[str, Any]:
    """Execute a DAX query and return rows plus performance diagnostics."""
    return _run(
        "pbi_trace_query",
        pbi_trace_query_tool,
        CONNECTION_MANAGER,
        query=query,
        timeout_seconds=timeout_seconds,
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
) -> dict[str, Any]:
    """Batch-create or update multiple DAX measures with a single SaveChanges call.

    measures: list of {name, expression, format_string?, description?, display_folder?, is_hidden?}
    """
    return _run(
        "pbi_create_measures",
        pbi_create_measures_tool,
        CONNECTION_MANAGER,
        table=table,
        measures=measures,
        overwrite=overwrite,
        stop_on_error=stop_on_error,
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
def pbi_create_relationship(
    from_table: str,
    from_column: str,
    to_table: str,
    to_column: str,
    cardinality: str = "oneToMany",
    direction: str = "oneDirection",
    is_active: bool = True,
    relationship_name: str | None = None,
) -> dict[str, Any]:
    """Create a relationship between two columns."""
    return _run(
        "pbi_create_relationship",
        pbi_create_relationship_tool,
        CONNECTION_MANAGER,
        from_table=from_table,
        from_column=from_column,
        to_table=to_table,
        to_column=to_column,
        cardinality=cardinality,
        direction=direction,
        is_active=is_active,
        relationship_name=relationship_name,
    )


@mcp.tool()
def pbi_delete_relationship(
    name: str | None = None,
    from_table: str | None = None,
    from_column: str | None = None,
    to_table: str | None = None,
    to_column: str | None = None,
) -> dict[str, Any]:
    """Delete a relationship by name or by endpoint columns."""
    return _run(
        "pbi_delete_relationship",
        pbi_delete_relationship_tool,
        CONNECTION_MANAGER,
        name=name,
        from_table=from_table,
        from_column=from_column,
        to_table=to_table,
        to_column=to_column,
    )


@mcp.tool()
def pbi_update_relationship(
    name: str | None = None,
    from_table: str | None = None,
    from_column: str | None = None,
    to_table: str | None = None,
    to_column: str | None = None,
    cardinality: str | None = None,
    direction: str | None = None,
    is_active: bool | None = None,
    new_name: str | None = None,
) -> dict[str, Any]:
    """Update properties of an existing relationship."""
    return _run(
        "pbi_update_relationship",
        pbi_update_relationship_tool,
        CONNECTION_MANAGER,
        name=name,
        from_table=from_table,
        from_column=from_column,
        to_table=to_table,
        to_column=to_column,
        cardinality=cardinality,
        direction=direction,
        is_active=is_active,
        new_name=new_name,
    )


@mcp.tool()
def pbi_delete_table(name: str) -> dict[str, Any]:
    """Delete a table from the model."""
    return _run("pbi_delete_table", pbi_delete_table_tool, CONNECTION_MANAGER, name=name)


@mcp.tool()
def pbi_delete_column(table: str, name: str) -> dict[str, Any]:
    """Delete a column from a table."""
    return _run(
        "pbi_delete_column",
        pbi_delete_column_tool,
        CONNECTION_MANAGER,
        table=table,
        name=name,
    )


@mcp.tool()
def pbi_rename_table(name: str, new_name: str) -> dict[str, Any]:
    """Rename a table. Dependent DAX expressions must be updated separately."""
    return _run(
        "pbi_rename_table",
        pbi_rename_table_tool,
        CONNECTION_MANAGER,
        name=name,
        new_name=new_name,
    )


@mcp.tool()
def pbi_rename_column(table: str, name: str, new_name: str) -> dict[str, Any]:
    """Rename a column. Dependent DAX expressions must be updated separately."""
    return _run(
        "pbi_rename_column",
        pbi_rename_column_tool,
        CONNECTION_MANAGER,
        table=table,
        name=name,
        new_name=new_name,
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
def pbi_refresh_metadata() -> dict[str, Any]:
    """Reload the cached TOM schema (cheaper than pbi_connect force_reconnect)."""
    return _run("pbi_refresh_metadata", pbi_refresh_metadata_tool, CONNECTION_MANAGER)


@mcp.tool()
def pbi_validate_dax(expression: str, kind: str = "scalar") -> dict[str, Any]:
    """Parse-check a DAX expression. kind='scalar' or 'table'."""
    return _run(
        "pbi_validate_dax",
        pbi_validate_dax_tool,
        CONNECTION_MANAGER,
        expression=expression,
        kind=kind,
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
def pbi_model_info(
    include_hidden: bool = False,
    include_row_counts: bool = False,
) -> dict[str, Any]:
    """Return a full model snapshot."""
    return _run(
        "pbi_model_info",
        pbi_model_info_tool,
        CONNECTION_MANAGER,
        include_hidden=include_hidden,
        include_row_counts=include_row_counts,
    )


@mcp.tool()
def pbi_refresh(target: str = "model", refresh_type: str = "full") -> dict[str, Any]:
    """Trigger a model or table refresh."""
    return _run(
        "pbi_refresh",
        pbi_refresh_tool,
        CONNECTION_MANAGER,
        target=target,
        refresh_type=refresh_type,
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
def pbi_create_table(
    name: str,
    expression: str,
    is_hidden: bool = False,
    overwrite: bool = False,
    refresh_after_create: bool = True,
) -> dict[str, Any]:
    """Create or update a calculated table."""
    return _run(
        "pbi_create_table",
        pbi_create_table_tool,
        CONNECTION_MANAGER,
        name=name,
        expression=expression,
        is_hidden=is_hidden,
        overwrite=overwrite,
        refresh_after_create=refresh_after_create,
    )


@mcp.tool()
def pbi_create_column(
    table: str,
    name: str,
    expression: str,
    data_type: str | None = None,
    format_string: str = "",
    display_folder: str = "",
    is_hidden: bool = False,
    overwrite: bool = False,
) -> dict[str, Any]:
    """Create or update a calculated column."""
    return _run(
        "pbi_create_column",
        pbi_create_column_tool,
        CONNECTION_MANAGER,
        table=table,
        name=name,
        expression=expression,
        data_type=data_type,
        format_string=format_string,
        display_folder=display_folder,
        is_hidden=is_hidden,
        overwrite=overwrite,
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


@mcp.tool()
def pbi_export_model(
    path: str | None = None,
    include_hidden: bool = False,
    include_row_counts: bool = False,
) -> dict[str, Any]:
    """Export the full model as JSON and optionally write it to disk."""
    return _run(
        "pbi_export_model",
        pbi_export_model_tool,
        CONNECTION_MANAGER,
        path=path,
        include_hidden=include_hidden,
        include_row_counts=include_row_counts,
    )


@mcp.tool()
def pbi_validate_model(include_warnings: bool = True) -> dict[str, Any]:
    """Audit the model for issues: empty expressions, missing format strings, orphan tables, duplicate measure names."""
    return _run(
        "pbi_validate_model",
        pbi_validate_model_tool,
        CONNECTION_MANAGER,
        include_warnings=include_warnings,
    )


@mcp.tool()
def excel_list_sheets(file_path: str) -> dict[str, Any]:
    """List workbook sheets with row and column counts."""
    return _run("excel_list_sheets", excel_list_sheets_tool, file_path=file_path)


@mcp.tool()
def excel_read_sheet(
    file_path: str,
    sheet: str,
    range: str | None = None,
    limit: int = 500,
) -> dict[str, Any]:
    """Read rows from a worksheet or range."""
    return _run(
        "excel_read_sheet",
        excel_read_sheet_tool,
        file_path=file_path,
        sheet=sheet,
        range=range,
        limit=limit,
    )


@mcp.tool()
def excel_read_cell(file_path: str, sheet: str, cell: str) -> dict[str, Any]:
    """Read a single worksheet cell."""
    return _run(
        "excel_read_cell",
        excel_read_cell_tool,
        file_path=file_path,
        sheet=sheet,
        cell=cell,
    )


@mcp.tool()
def excel_search(file_path: str, query: str, sheet: str | None = None) -> dict[str, Any]:
    """Search workbook values across one or all sheets."""
    return _run(
        "excel_search",
        excel_search_tool,
        file_path=file_path,
        query=query,
        sheet=sheet,
    )


@mcp.tool()
def excel_write_cell(
    file_path: str,
    sheet: str,
    cell: str,
    value: Any,
    format: str = "",
) -> dict[str, Any]:
    """Write a single cell value."""
    return _run(
        "excel_write_cell",
        excel_write_cell_tool,
        file_path=file_path,
        sheet=sheet,
        cell=cell,
        value=value,
        format=format,
    )


@mcp.tool()
def excel_write_range(
    file_path: str,
    sheet: str,
    start_cell: str,
    data: list[list[Any]],
) -> dict[str, Any]:
    """Write a 2D array starting at a worksheet cell."""
    return _run(
        "excel_write_range",
        excel_write_range_tool,
        file_path=file_path,
        sheet=sheet,
        start_cell=start_cell,
        data=data,
    )


@mcp.tool()
def excel_create_sheet(file_path: str, name: str, position: int | None = None) -> dict[str, Any]:
    """Create a worksheet in an existing workbook."""
    return _run(
        "excel_create_sheet",
        excel_create_sheet_tool,
        file_path=file_path,
        name=name,
        position=position,
    )


@mcp.tool()
def excel_delete_sheet(file_path: str, name: str) -> dict[str, Any]:
    """Delete a worksheet from an existing workbook."""
    return _run(
        "excel_delete_sheet",
        excel_delete_sheet_tool,
        file_path=file_path,
        name=name,
    )


@mcp.tool()
def excel_format_range(
    file_path: str,
    sheet: str,
    range: str,
    format: dict[str, Any],
) -> dict[str, Any]:
    """Apply formatting to a worksheet range."""
    return _run(
        "excel_format_range",
        excel_format_range_tool,
        file_path=file_path,
        sheet=sheet,
        range=range,
        format=format,
    )


@mcp.tool()
def excel_auto_width(file_path: str, sheet: str) -> dict[str, Any]:
    """Auto-fit worksheet column widths."""
    return _run("excel_auto_width", excel_auto_width_tool, file_path=file_path, sheet=sheet)


@mcp.tool()
def excel_create_workbook(file_path: str, sheets: list[str] | None = None) -> dict[str, Any]:
    """Create a new workbook."""
    return _run(
        "excel_create_workbook",
        excel_create_workbook_tool,
        file_path=file_path,
        sheets=sheets,
    )


@mcp.tool()
def excel_workbook_info(file_path: str) -> dict[str, Any]:
    """Return workbook metadata and sheet summaries."""
    return _run("excel_workbook_info", excel_workbook_info_tool, file_path=file_path)


@mcp.tool()
def excel_to_pbi_check(file_path: str) -> dict[str, Any]:
    """Compare an Excel workbook with the current Power BI model."""
    return _run(
        "excel_to_pbi_check",
        excel_to_pbi_check_tool,
        file_path=file_path,
        manager=CONNECTION_MANAGER,
    )


# ── Power Query tools ────────────────────────────────────────────────


@mcp.tool()
def pbi_get_power_query(table: str, partition_name: str | None = None) -> dict[str, Any]:
    """Read the Power Query (M) expression for a table."""
    return _run(
        "pbi_get_power_query",
        pbi_get_power_query_tool,
        CONNECTION_MANAGER,
        table=table,
        partition_name=partition_name,
    )


@mcp.tool()
def pbi_list_power_queries(include_hidden: bool = False) -> dict[str, Any]:
    """List table partitions with their current source expressions."""
    return _run(
        "pbi_list_power_queries",
        pbi_list_power_queries_tool,
        CONNECTION_MANAGER,
        include_hidden=include_hidden,
    )


@mcp.tool()
def pbi_set_power_query(
    table: str,
    m_expression: str,
    partition_name: str | None = None,
    refresh_after: bool = False,
) -> dict[str, Any]:
    """Write or update the Power Query (M) expression for a table."""
    return _run(
        "pbi_set_power_query",
        pbi_set_power_query_tool,
        CONNECTION_MANAGER,
        table=table,
        m_expression=m_expression,
        partition_name=partition_name,
        refresh_after=refresh_after,
    )


@mcp.tool()
def pbi_create_import_query(
    table: str,
    excel_path: str,
    sheet_name: str,
    partition_name: str | None = None,
    promote_headers: bool = True,
    refresh_after: bool = True,
) -> dict[str, Any]:
    """Generate and inject an Excel import Power Query for a table."""
    return _run(
        "pbi_create_import_query",
        pbi_create_import_query_tool,
        CONNECTION_MANAGER,
        table=table,
        excel_path=excel_path,
        sheet_name=sheet_name,
        partition_name=partition_name,
        promote_headers=promote_headers,
        refresh_after=refresh_after,
    )


@mcp.tool()
def pbi_create_csv_import_query(
    table: str,
    csv_path: str,
    partition_name: str | None = None,
    delimiter: str = ",",
    encoding: int = 65001,
    quote_style: str = "csv",
    promote_headers: bool = True,
    refresh_after: bool = True,
) -> dict[str, Any]:
    """Generate and inject a CSV import Power Query for a table."""
    return _run(
        "pbi_create_csv_import_query",
        pbi_create_csv_import_query_tool,
        CONNECTION_MANAGER,
        table=table,
        csv_path=csv_path,
        partition_name=partition_name,
        delimiter=delimiter,
        encoding=encoding,
        quote_style=quote_style,
        promote_headers=promote_headers,
        refresh_after=refresh_after,
    )


@mcp.tool()
def pbi_create_folder_import_query(
    table: str,
    folder_path: str,
    partition_name: str | None = None,
    extension_filter: str | None = None,
    include_hidden_files: bool = False,
    refresh_after: bool = True,
) -> dict[str, Any]:
    """Generate and inject a folder import Power Query for a table."""
    return _run(
        "pbi_create_folder_import_query",
        pbi_create_folder_import_query_tool,
        CONNECTION_MANAGER,
        table=table,
        folder_path=folder_path,
        partition_name=partition_name,
        extension_filter=extension_filter,
        include_hidden_files=include_hidden_files,
        refresh_after=refresh_after,
    )


@mcp.tool()
def pbi_bulk_import_excel(
    excel_path: str,
    sheet_table_map: dict[str, str] | None = None,
    promote_headers: bool = True,
    refresh_after: bool = True,
) -> dict[str, Any]:
    """Bulk-create Excel import queries for multiple tables at once."""
    return _run(
        "pbi_bulk_import_excel",
        pbi_bulk_import_excel_tool,
        CONNECTION_MANAGER,
        excel_path=excel_path,
        sheet_table_map=sheet_table_map,
        promote_headers=promote_headers,
        refresh_after=refresh_after,
    )


@mcp.tool()
def pbi_import_excel_workbook(
    excel_path: str,
    sheet_table_map: dict[str, str] | None = None,
    promote_headers: bool = True,
    refresh_after: bool = True,
) -> dict[str, Any]:
    """Import an Excel workbook into Power BI tables in one call."""
    return _run(
        "pbi_import_excel_workbook",
        pbi_import_excel_workbook_tool,
        CONNECTION_MANAGER,
        excel_path=excel_path,
        sheet_table_map=sheet_table_map,
        promote_headers=promote_headers,
        refresh_after=refresh_after,
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
def pbi_patch_layout(extract_folder: str, pbix_path: str, force: bool = False) -> dict[str, Any]:
    """Patch Report/Layout directly into an existing .pbix archive."""
    return _run(
        "pbi_patch_layout",
        pbi_patch_layout_tool,
        extract_folder=extract_folder,
        pbix_path=pbix_path,
        force=force,
    )


@mcp.tool()
def pbi_list_pages(extract_folder: str) -> dict[str, Any]:
    """List pages in an extracted report."""
    return _run("pbi_list_pages", pbi_list_pages_tool, extract_folder=extract_folder)


@mcp.tool()
def pbi_get_page(extract_folder: str, page: str) -> dict[str, Any]:
    """Get page details and visual metadata from an extracted report."""
    return _run("pbi_get_page", pbi_get_page_tool, extract_folder=extract_folder, page=page)


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
    """Add a card visual to a report page."""
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
    """Add a clustered bar chart visual."""
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
    """Add a line chart visual."""
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
    """Add a donut chart visual."""
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
    """Add a table visual."""
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
    """Add a waterfall chart visual."""
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
    """Add a slicer visual."""
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
) -> dict[str, Any]:
    """Add a gauge visual."""
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
) -> dict[str, Any]:
    """Add any visual via a generic dispatcher. visual_type: card|bar_chart|line_chart|donut|table|waterfall|slicer|gauge|text_box."""
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
    )


@mcp.tool()
def pbi_list_roles() -> dict[str, Any]:
    """List RLS roles, members, and table filters."""
    return _run("pbi_list_roles", pbi_list_roles_tool, CONNECTION_MANAGER)


@mcp.tool()
def pbi_create_role(
    name: str,
    permission: str = "Read",
    description: str = "",
    overwrite: bool = False,
) -> dict[str, Any]:
    """Create or update an RLS role. permission: None|Read|ReadRefresh|Refresh|Administrator."""
    return _run(
        "pbi_create_role",
        pbi_create_role_tool,
        CONNECTION_MANAGER,
        name=name,
        permission=permission,
        description=description,
        overwrite=overwrite,
    )


@mcp.tool()
def pbi_delete_role(name: str) -> dict[str, Any]:
    """Delete an RLS role."""
    return _run("pbi_delete_role", pbi_delete_role_tool, CONNECTION_MANAGER, name=name)


@mcp.tool()
def pbi_set_role_filter(
    role: str,
    table: str,
    filter_expression: str | None,
) -> dict[str, Any]:
    """Apply or clear a DAX RLS filter on a table for a role (None/empty clears)."""
    return _run(
        "pbi_set_role_filter",
        pbi_set_role_filter_tool,
        CONNECTION_MANAGER,
        role=role,
        table=table,
        filter_expression=filter_expression,
    )


@mcp.tool()
def pbi_add_role_member(
    role: str,
    member_name: str,
    member_type: str = "external",
    identity_provider: str = "AzureAD",
) -> dict[str, Any]:
    """Add a member to an RLS role. member_type: external|windows."""
    return _run(
        "pbi_add_role_member",
        pbi_add_role_member_tool,
        CONNECTION_MANAGER,
        role=role,
        member_name=member_name,
        member_type=member_type,
        identity_provider=identity_provider,
    )


@mcp.tool()
def pbi_remove_role_member(role: str, member_name: str) -> dict[str, Any]:
    """Remove a member from an RLS role (matched on MemberName)."""
    return _run(
        "pbi_remove_role_member",
        pbi_remove_role_member_tool,
        CONNECTION_MANAGER,
        role=role,
        member_name=member_name,
    )


@mcp.tool()
def pbi_list_calc_groups() -> dict[str, Any]:
    """List calculation groups and their calculation items."""
    return _run("pbi_list_calc_groups", pbi_list_calc_groups_tool, CONNECTION_MANAGER)


@mcp.tool()
def pbi_create_calc_group(
    table_name: str,
    column_name: str = "Name",
    precedence: int = 0,
    items: list[dict[str, Any]] | None = None,
    overwrite: bool = False,
) -> dict[str, Any]:
    """Create or replace a calculation group. items: [{name, expression, format_string_expression?, ordinal?}]."""
    return _run(
        "pbi_create_calc_group",
        pbi_create_calc_group_tool,
        CONNECTION_MANAGER,
        table_name=table_name,
        column_name=column_name,
        precedence=precedence,
        items=items,
        overwrite=overwrite,
    )


@mcp.tool()
def pbi_delete_calc_group(table_name: str) -> dict[str, Any]:
    """Delete a calculation group table."""
    return _run(
        "pbi_delete_calc_group",
        pbi_delete_calc_group_tool,
        CONNECTION_MANAGER,
        table_name=table_name,
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


# ── MCP Resources ────────────────────────────────────────────────────
# Expose live model data as MCP Resources so clients can subscribe/fetch
# without burning a tool call. Cache in the manager invalidates on writes.


@mcp.resource("powerbi://model/schema")
def resource_model_schema() -> str:
    """Full model snapshot: tables, columns, measures, relationships."""
    import json
    result = _run("pbi_model_info", pbi_model_info_tool, CONNECTION_MANAGER)
    return json.dumps(result, ensure_ascii=False, indent=2)


@mcp.resource("powerbi://model/measures")
def resource_model_measures() -> str:
    """All DAX measures in the active model."""
    import json
    result = _run("pbi_list_measures", pbi_list_measures_tool, CONNECTION_MANAGER)
    return json.dumps(result, ensure_ascii=False, indent=2)


@mcp.resource("powerbi://model/relationships")
def resource_model_relationships() -> str:
    """All relationships in the active model."""
    import json
    result = _run("pbi_list_relationships", pbi_list_relationships_tool, CONNECTION_MANAGER)
    return json.dumps(result, ensure_ascii=False, indent=2)


# ── MCP Prompts ───────────────────────────────────────────────────────
# Ready-to-use workflow prompts surfaced natively to any MCP client.


@mcp.prompt()
def model_audit() -> str:
    """Full model audit: tables, measures, relationships, and improvement suggestions."""
    return (
        "Connect to Power BI Desktop. Give me a compact audit of the active model:\n"
        "- Table count + row counts for each fact table\n"
        "- Measure count per table, flag any obviously misnamed or empty measures\n"
        "- Relationship graph (from → to + cardinality + active/inactive)\n"
        "- Any table with multiple partitions or complex measure dependencies\n"
        "Then run pbi_validate_model() and list all issues and warnings.\n"
        "End with 3 concrete improvement suggestions ranked by impact."
    )


@mcp.prompt()
def time_intelligence_kit(base_measure: str = "Revenue", date_table: str = "Date", date_column: str = "Date") -> str:
    """Generate a full time-intelligence measure kit (MTD, QTD, YTD, YoY, YoY%) for a base measure."""
    return (
        f"Assuming '{date_table}' is marked as the date table, generate these measures for [{base_measure}]:\n"
        f"- {base_measure} MTD  = TOTALMTD([{base_measure}], {date_table}[{date_column}])\n"
        f"- {base_measure} QTD  = TOTALQTD([{base_measure}], {date_table}[{date_column}])\n"
        f"- {base_measure} YTD  = TOTALYTD([{base_measure}], {date_table}[{date_column}])\n"
        f"- {base_measure} YoY  = CALCULATE([{base_measure}], SAMEPERIODLASTYEAR({date_table}[{date_column}]))\n"
        f"- {base_measure} YoY% = DIVIDE([{base_measure}] - [{base_measure} YoY], [{base_measure} YoY])\n\n"
        "Use pbi_validate_dax on each expression before creating. "
        "Use pbi_create_measures (batch) to write all 5 in one call. "
        "Apply format string '#,##0.00' to MTD/QTD/YTD/YoY and '0.00%' to YoY%."
    )


@mcp.prompt()
def star_schema_builder(fact_table: str = "FactSales") -> str:
    """Guide for wiring a star schema: relationships + key measures."""
    return (
        f"Inspect the model with pbi_list_tables and pbi_list_relationships.\n"
        f"For fact table '{fact_table}':\n"
        "1. Identify all dimension tables by looking for columns that match FK columns in the fact.\n"
        "2. Create missing Many-to-One relationships (fact → dimension, oneDirection).\n"
        "3. Flag any existing Many-to-Many relationships and suggest a bridge table fix.\n"
        "4. Create a basic measure '[Row Count]' = COUNTROWS(fact_table) as a sanity check.\n"
        "5. Run pbi_validate_model() at the end and report any remaining issues."
    )


@mcp.prompt()
def rls_setup(table: str = "Sales", filter_column: str = "Region") -> str:
    """Set up Row-Level Security for a given table and filter column."""
    return (
        f"Set up Row-Level Security on '{table}[{filter_column}]':\n"
        "1. pbi_list_roles() — check if a role already exists.\n"
        "2. pbi_create_role(role='RegionFilter') — create a new role.\n"
        f"3. pbi_set_role_filter(role='RegionFilter', table='{table}', filter_expression='[{filter_column}] = USERNAME()') — apply filter.\n"
        "4. pbi_execute_dax_as_role(query='EVALUATE ROW(\"User\", USERNAME())', role='RegionFilter') — validate.\n"
        "5. Summarize the final RLS setup."
    )


@mcp.prompt()
def dead_measure_scan() -> str:
    """Find measures not referenced by any other measure and suggest cleanup."""
    return (
        "Find measures that are not referenced by any other measure:\n"
        "1. pbi_measure_dependencies() — get the full dependency graph.\n"
        "2. pbi_list_measures() — list all measures.\n"
        "3. Cross-reference: which measures appear only as roots (nothing depends on them)?\n"
        "4. For each orphan, show its expression and suggest: keep / rename / delete.\n"
        "Do NOT delete anything — only report recommendations."
    )


@mcp.prompt()
def bulk_measure_format_fix(table: str = "Measures", format_string: str = "#,##0") -> str:
    """Apply a format string to all measures in a table that are missing one."""
    return (
        f"Find all measures in table '{table}' that have no format string set.\n"
        "Use pbi_list_measures(include_hidden=False) to get the full list.\n"
        "Filter to those where format_string is empty or null.\n"
        f"Apply format string '{format_string}' to all of them using pbi_set_format(table='{table}', names=[...], format_string='{format_string}').\n"
        "Report how many were updated."
    )


@mcp.prompt()
def excel_to_pbi_pipeline(excel_path: str = "") -> str:
    """Full pipeline: inspect Excel, create import queries, refresh, validate."""
    path_hint = f"'{excel_path}'" if excel_path else "<path/to/file.xlsx>"
    return (
        f"Run the full Excel → Power BI import pipeline for {path_hint}:\n"
        "1. excel_workbook_info() — list sheets and row counts.\n"
        "2. pbi_bulk_import_excel(excel_path=..., refresh_after=False) — inject import queries for all sheets.\n"
        "3. pbi_refresh(target='model', refresh_type='full') — refresh the model.\n"
        "4. pbi_list_tables(include_row_counts=True) — verify row counts match the Excel source.\n"
        "5. Report any discrepancies."
    )


@mcp.prompt()
def model_snapshot_export(output_path: str = "./docs/model.json") -> str:
    """Export the full model as JSON for documentation or version control."""
    return (
        f"Export the full model definition to '{output_path}':\n"
        f"pbi_export_model(path='{output_path}', include_hidden=True, include_row_counts=False)\n\n"
        "Then summarize:\n"
        "- Total tables, measures, relationships\n"
        "- Top 5 most complex measures by expression length\n"
        "- Any tables with more than 2 partitions"
    )


def _bearer_auth_middleware(app: Any, token: str) -> Any:
    """ASGI middleware that requires Authorization: Bearer <token> on HTTP requests."""
    expected = f"Bearer {token}".encode("utf-8")

    async def wrapped(scope: dict, receive: Any, send: Any) -> None:
        if scope.get("type") != "http":
            await app(scope, receive, send)
            return
        headers = dict(scope.get("headers") or [])
        provided = headers.get(b"authorization", b"")
        if provided != expected:
            await send({
                "type": "http.response.start",
                "status": 401,
                "headers": [(b"content-type", b"text/plain"), (b"www-authenticate", b"Bearer")],
            })
            await send({"type": "http.response.body", "body": b"Unauthorized"})
            return
        await app(scope, receive, send)

    return wrapped


async def _run_sse_with_auth(host: str, port: int) -> None:
    """Mirror FastMCP.run_sse_async but allow wrapping with Bearer auth middleware."""
    import uvicorn

    mcp.settings.host = host
    mcp.settings.port = port

    asgi_app = mcp.sse_app()
    token = os.getenv("PBI_MCP_AUTH_TOKEN", "").strip()
    if token:
        asgi_app = _bearer_auth_middleware(asgi_app, token)
        logger.info("SECURITY: SSE Bearer auth enabled (PBI_MCP_AUTH_TOKEN set).")
    else:
        logger.warning(
            "SECURITY: SSE has no Bearer auth. Set PBI_MCP_AUTH_TOKEN to require "
            "'Authorization: Bearer <token>' on HTTP requests.",
        )

    config = uvicorn.Config(
        asgi_app,
        host=host,
        port=port,
        log_level=mcp.settings.log_level.lower(),
    )
    server = uvicorn.Server(config)
    await server.serve()


def main() -> None:
    """Entry point — supports stdio (default) and sse transport."""
    import argparse

    parser = argparse.ArgumentParser(description="Power BI Desktop MCP Server")
    parser.add_argument(
        "--transport",
        choices=["stdio", "sse"],
        default="stdio",
        help="MCP transport: stdio (CLI tools) or sse (web/IDE clients)",
    )
    parser.add_argument(
        "--port",
        type=int,
        default=8765,
        help="Port for SSE transport (default: 8765)",
    )
    parser.add_argument(
        "--host",
        default="127.0.0.1",
        help="Host for SSE transport (default: 127.0.0.1 — localhost only)",
    )
    parser.add_argument(
        "--readonly",
        action="store_true",
        help="Disable write and destructive tools for this server process.",
    )
    parser.add_argument(
        "--profile",
        choices=["readonly", "write", "all"],
        default="all",
        help="Filter exposed tool surface: readonly, write (read+write), or all (default).",
    )
    args = parser.parse_args()
    SECURITY.policy(reload=True, cwd=Path(__file__).parent)
    if args.readonly:
        SECURITY.set_runtime_readonly(True)
        logger.info("SECURITY: readonly mode enabled via --readonly")

    _apply_profile(args.profile)

    if args.transport == "sse":
        logger.info(
            "SSE server starting on %s:%d (localhost-only by default)",
            args.host, args.port,
        )
        if args.host != "127.0.0.1":
            logger.warning(
                "SECURITY: SSE bound to %s — exposed beyond localhost. "
                "Ensure network is trusted or use --host 127.0.0.1",
                args.host,
            )
        import anyio
        anyio.run(_run_sse_with_auth, args.host, args.port)
    else:
        mcp.run(transport="stdio")


def _apply_profile(profile: str) -> None:
    """Prune FastMCP's registered tools based on the selected profile."""
    if profile == "all":
        return
    from security import READ_TOOLS, WRITE_TOOLS, DESTRUCTIVE_TOOLS

    if profile == "readonly":
        allowed = set(READ_TOOLS)
    elif profile == "write":
        allowed = set(READ_TOOLS) | set(WRITE_TOOLS)
    else:
        return

    manager = getattr(mcp, "_tool_manager", None)
    tools_map = getattr(manager, "_tools", None)
    if tools_map is None:
        logger.warning("profile filter skipped: FastMCP tool registry not accessible")
        return
    removed = [name for name in list(tools_map.keys()) if name not in allowed]
    for name in removed:
        tools_map.pop(name, None)
    logger.info("PROFILE %s applied: removed %d tools, exposing %d", profile, len(removed), len(tools_map))


if __name__ == "__main__":
    main()

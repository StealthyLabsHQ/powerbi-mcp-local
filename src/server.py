"""FastMCP server exposing Power BI Desktop model operations over stdio."""

from __future__ import annotations

from pathlib import Path
from typing import Any

from mcp.server.fastmcp import FastMCP

from pbi_connection import PowerBIConnectionManager, error_payload, logger
from security import SECURITY
from tools import (
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
    pbi_delete_measure_tool,
    pbi_execute_dax_as_role_tool,
    pbi_execute_dax_tool,
    pbi_export_model_tool,
    pbi_import_dax_file_tool,
    pbi_list_instances_tool,
    pbi_list_measures_tool,
    pbi_list_relationships_tool,
    pbi_list_tables_tool,
    pbi_model_info_tool,
    pbi_refresh_tool,
    pbi_set_format_tool,
    pbi_trace_query_tool,
    pbi_bulk_import_excel_tool,
    pbi_create_csv_import_query_tool,
    pbi_create_folder_import_query_tool,
    pbi_create_import_query_tool,
    pbi_add_bar_chart_tool,
    pbi_add_card_tool,
    pbi_add_donut_chart_tool,
    pbi_add_gauge_tool,
    pbi_add_line_chart_tool,
    pbi_add_slicer_tool,
    pbi_add_table_visual_tool,
    pbi_add_text_box_tool,
    pbi_add_waterfall_tool,
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
def pbi_execute_dax(query: str, max_rows: int = 1000) -> dict[str, Any]:
    """Execute a DAX or DMV query and return rows."""
    return _run(
        "pbi_execute_dax",
        pbi_execute_dax_tool,
        CONNECTION_MANAGER,
        query=query,
        max_rows=max_rows,
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
def pbi_trace_query(query: str) -> dict[str, Any]:
    """Execute a DAX query and return rows plus performance diagnostics."""
    return _run(
        "pbi_trace_query",
        pbi_trace_query_tool,
        CONNECTION_MANAGER,
        query=query,
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
def pbi_extract_report(pbix_path: str, extract_folder: str | None = None) -> dict[str, Any]:
    """Extract a .pbix report into a pbi-tools folder structure."""
    return _run(
        "pbi_extract_report",
        pbi_extract_report_tool,
        pbix_path=pbix_path,
        extract_folder=extract_folder,
    )


@mcp.tool()
def pbi_compile_report(extract_folder: str, output_path: str) -> dict[str, Any]:
    """Compile an extracted report folder back into a .pbix."""
    return _run(
        "pbi_compile_report",
        pbi_compile_report_tool,
        extract_folder=extract_folder,
        output_path=output_path,
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
def pbi_build_dashboard(extract_folder: str, page: str, layout: list[dict[str, Any]]) -> dict[str, Any]:
    """Build a dashboard page from a bulk layout specification."""
    return _run(
        "pbi_build_dashboard",
        pbi_build_dashboard_tool,
        extract_folder=extract_folder,
        page=page,
        layout=layout,
    )


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
    args = parser.parse_args()
    SECURITY.policy(reload=True, cwd=Path(__file__).parent)
    if args.readonly:
        SECURITY.set_runtime_readonly(True)
        logger.info("SECURITY: readonly mode enabled via --readonly")

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
        mcp.run(transport="sse", sse_params={"host": args.host, "port": args.port})
    else:
        mcp.run(transport="stdio")


if __name__ == "__main__":
    main()

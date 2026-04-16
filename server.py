"""FastMCP server exposing Power BI Desktop model operations over stdio."""

from __future__ import annotations

from typing import Any

from mcp.server.fastmcp import FastMCP

from pbi_connection import PowerBIConnectionManager, error_payload, logger
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
    pbi_get_power_query_tool,
    pbi_set_power_query_tool,
    pbi_create_import_query_tool,
    pbi_bulk_import_excel_tool,
)


mcp = FastMCP(
    "powerbi-desktop",
    instructions=(
        "Connects to the local Power BI Desktop Analysis Services instance, "
        "lets clients inspect the semantic model, manage measures and "
        "relationships, run DAX queries, trigger model refreshes, and "
        "read or write Excel workbooks used in the Power BI pipeline."
    ),
    json_response=True,
    log_level="INFO",
)

CONNECTION_MANAGER = PowerBIConnectionManager(logger)


def _run(tool_name: str, callback: Any, *args: Any, **kwargs: Any) -> dict[str, Any]:
    """Execute a tool callback and normalize failures to JSON."""
    try:
        return callback(*args, **kwargs)
    except Exception as exc:  # pragma: no cover - exercised on Windows
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
def pbi_get_power_query(table: str) -> dict[str, Any]:
    """Read the Power Query (M) expression for a table."""
    return _run(
        "pbi_get_power_query",
        pbi_get_power_query_tool,
        CONNECTION_MANAGER,
        table=table,
    )


@mcp.tool()
def pbi_set_power_query(
    table: str,
    m_expression: str,
    refresh_after: bool = False,
) -> dict[str, Any]:
    """Write or update the Power Query (M) expression for a table."""
    return _run(
        "pbi_set_power_query",
        pbi_set_power_query_tool,
        CONNECTION_MANAGER,
        table=table,
        m_expression=m_expression,
        refresh_after=refresh_after,
    )


@mcp.tool()
def pbi_create_import_query(
    table: str,
    excel_path: str,
    sheet_name: str,
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
        promote_headers=promote_headers,
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


if __name__ == "__main__":
    mcp.run(transport="stdio")

"""MCP wrappers — domain: excel."""

from __future__ import annotations

from typing import Any

from mcp_core import CONNECTION_MANAGER, _run, mcp
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

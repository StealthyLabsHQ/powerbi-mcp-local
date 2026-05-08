"""MCP wrappers — domain: excel."""

from __future__ import annotations

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

from ._helpers import register_tool

register_tool(excel_list_sheets_tool)
register_tool(excel_read_sheet_tool)
register_tool(excel_read_cell_tool)
register_tool(excel_search_tool)
register_tool(excel_write_cell_tool)
register_tool(excel_write_range_tool)
register_tool(excel_create_sheet_tool)
register_tool(excel_delete_sheet_tool)
register_tool(excel_format_range_tool)
register_tool(excel_auto_width_tool)
register_tool(excel_create_workbook_tool)
register_tool(excel_workbook_info_tool)
register_tool(excel_to_pbi_check_tool)

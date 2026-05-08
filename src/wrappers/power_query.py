"""MCP wrappers — domain: power_query."""

from __future__ import annotations

from tools import (
    pbi_bulk_import_excel_tool,
    pbi_create_csv_import_query_tool,
    pbi_create_folder_import_query_tool,
    pbi_create_import_query_tool,
    pbi_excel_import_workflow_tool,
    pbi_get_power_query_tool,
    pbi_import_excel_workbook_tool,
    pbi_list_power_queries_tool,
    pbi_parameterize_data_source_tool,
    pbi_relocate_data_source_tool,
    pbi_set_power_query_tool,
)

from ._helpers import register_tool

register_tool(pbi_get_power_query_tool)
register_tool(pbi_list_power_queries_tool)
register_tool(pbi_set_power_query_tool)
register_tool(pbi_parameterize_data_source_tool)
register_tool(pbi_relocate_data_source_tool)
register_tool(pbi_create_import_query_tool)
register_tool(pbi_create_csv_import_query_tool)
register_tool(pbi_create_folder_import_query_tool)
register_tool(pbi_bulk_import_excel_tool)
register_tool(pbi_import_excel_workbook_tool)
register_tool(pbi_excel_import_workflow_tool)

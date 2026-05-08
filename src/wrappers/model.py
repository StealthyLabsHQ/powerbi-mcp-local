"""MCP wrappers — domain: model."""

from __future__ import annotations

from tools import (
    pbi_create_column_tool,
    pbi_create_table_tool,
    pbi_delete_column_tool,
    pbi_delete_table_tool,
    pbi_list_tables_tool,
    pbi_model_info_tool,
    pbi_rename_column_tool,
    pbi_rename_table_tool,
    pbi_set_column_data_type_tool,
    pbi_validate_model_tool,
)

from ._helpers import register_tool

register_tool(pbi_list_tables_tool)
register_tool(pbi_delete_table_tool)
register_tool(pbi_delete_column_tool)
register_tool(pbi_rename_table_tool)
register_tool(pbi_rename_column_tool)
register_tool(pbi_model_info_tool)
register_tool(pbi_create_table_tool)
register_tool(pbi_create_column_tool)
register_tool(pbi_set_column_data_type_tool)
register_tool(pbi_validate_model_tool)

"""MCP wrappers — domain: connection."""

from __future__ import annotations

from tools import (
    pbi_connect_tool,
    pbi_export_model_tool,
    pbi_list_instances_tool,
    pbi_operation_history_tool,
    pbi_refresh_metadata_tool,
    pbi_refresh_tool,
    pbi_system_health_tool,
)

from ._helpers import register_tool

register_tool(pbi_connect_tool)
register_tool(pbi_list_instances_tool)
register_tool(pbi_refresh_metadata_tool)
register_tool(pbi_system_health_tool)
register_tool(pbi_operation_history_tool)
register_tool(pbi_refresh_tool)
register_tool(pbi_export_model_tool)

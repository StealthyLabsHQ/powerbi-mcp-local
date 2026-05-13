"""MCP wrappers — domain: project."""

from __future__ import annotations

from tools import (
    pbi_create_persistent_report_tool,
    pbi_list_scaffold_templates_tool,
    pbi_scaffold_pbix_tool,
)

from ._helpers import register_tool

register_tool(pbi_create_persistent_report_tool)
register_tool(pbi_scaffold_pbix_tool)
register_tool(pbi_list_scaffold_templates_tool)

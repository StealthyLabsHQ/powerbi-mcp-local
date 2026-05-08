"""MCP wrappers — domain: calc_groups."""

from __future__ import annotations

from tools import (
    pbi_create_calc_group_tool,
    pbi_delete_calc_group_tool,
    pbi_list_calc_groups_tool,
)

from ._helpers import register_tool

register_tool(pbi_list_calc_groups_tool)
register_tool(pbi_create_calc_group_tool)
register_tool(pbi_delete_calc_group_tool)

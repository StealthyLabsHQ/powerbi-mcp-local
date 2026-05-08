"""MCP wrappers — domain: rls."""

from __future__ import annotations

from tools import (
    pbi_add_role_member_tool,
    pbi_create_role_tool,
    pbi_delete_role_tool,
    pbi_list_roles_tool,
    pbi_remove_role_member_tool,
    pbi_set_role_filter_tool,
)

from ._helpers import register_tool

register_tool(pbi_list_roles_tool)
register_tool(pbi_create_role_tool)
register_tool(pbi_delete_role_tool)
register_tool(pbi_set_role_filter_tool)
register_tool(pbi_add_role_member_tool)
register_tool(pbi_remove_role_member_tool)

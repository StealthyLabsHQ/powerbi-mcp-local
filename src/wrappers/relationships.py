"""MCP wrappers — domain: relationships."""

from __future__ import annotations

from tools import (
    pbi_create_relationship_tool,
    pbi_delete_relationship_tool,
    pbi_list_relationships_tool,
    pbi_update_relationship_tool,
)

from ._helpers import register_tool

register_tool(pbi_list_relationships_tool)
register_tool(pbi_create_relationship_tool)
register_tool(pbi_update_relationship_tool)
register_tool(pbi_delete_relationship_tool)

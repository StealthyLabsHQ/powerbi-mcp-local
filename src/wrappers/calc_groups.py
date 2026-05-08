"""MCP wrappers — domain: calc_groups."""

from __future__ import annotations

from typing import Any

from mcp_core import CONNECTION_MANAGER, _run, mcp
from tools import (
    pbi_create_calc_group_tool,
    pbi_delete_calc_group_tool,
    pbi_list_calc_groups_tool,
)


@mcp.tool()
def pbi_list_calc_groups() -> dict[str, Any]:
    """List calculation groups and their calculation items."""
    return _run("pbi_list_calc_groups", pbi_list_calc_groups_tool, CONNECTION_MANAGER)


@mcp.tool()
def pbi_create_calc_group(
    table_name: str,
    column_name: str = "Name",
    precedence: int = 0,
    items: list[dict[str, Any]] | None = None,
    overwrite: bool = False,
) -> dict[str, Any]:
    """Create or replace a calculation group. items: [{name, expression, format_string_expression?, ordinal?}]."""
    return _run(
        "pbi_create_calc_group",
        pbi_create_calc_group_tool,
        CONNECTION_MANAGER,
        table_name=table_name,
        column_name=column_name,
        precedence=precedence,
        items=items,
        overwrite=overwrite,
    )


@mcp.tool()
def pbi_delete_calc_group(table_name: str) -> dict[str, Any]:
    """Delete a calculation group table."""
    return _run(
        "pbi_delete_calc_group",
        pbi_delete_calc_group_tool,
        CONNECTION_MANAGER,
        table_name=table_name,
    )

"""MCP wrappers — domain: rls."""

from __future__ import annotations

from typing import Any

from mcp_core import CONNECTION_MANAGER, _run, mcp
from tools import (
    pbi_add_role_member_tool,
    pbi_create_role_tool,
    pbi_delete_role_tool,
    pbi_list_roles_tool,
    pbi_remove_role_member_tool,
    pbi_set_role_filter_tool,
)


@mcp.tool()
def pbi_list_roles() -> dict[str, Any]:
    """List RLS roles, members, and table filters."""
    return _run("pbi_list_roles", pbi_list_roles_tool, CONNECTION_MANAGER)


@mcp.tool()
def pbi_create_role(
    name: str,
    permission: str = "Read",
    description: str = "",
    overwrite: bool = False,
) -> dict[str, Any]:
    """Create or update an RLS role. permission: None|Read|ReadRefresh|Refresh|Administrator."""
    return _run(
        "pbi_create_role",
        pbi_create_role_tool,
        CONNECTION_MANAGER,
        name=name,
        permission=permission,
        description=description,
        overwrite=overwrite,
    )


@mcp.tool()
def pbi_delete_role(name: str) -> dict[str, Any]:
    """Delete an RLS role."""
    return _run("pbi_delete_role", pbi_delete_role_tool, CONNECTION_MANAGER, name=name)


@mcp.tool()
def pbi_set_role_filter(
    role: str,
    table: str,
    filter_expression: str | None,
) -> dict[str, Any]:
    """Apply or clear a DAX RLS filter on a table for a role (None/empty clears)."""
    return _run(
        "pbi_set_role_filter",
        pbi_set_role_filter_tool,
        CONNECTION_MANAGER,
        role=role,
        table=table,
        filter_expression=filter_expression,
    )


@mcp.tool()
def pbi_add_role_member(
    role: str,
    member_name: str,
    member_type: str = "external",
    identity_provider: str = "AzureAD",
    overwrite: bool = False,
) -> dict[str, Any]:
    """Add (or update) a member on an RLS role. member_type: external|windows.

    With ``overwrite=True``, an existing member with the same name is updated
    in place (identity_provider refreshed when external) instead of raising
    ``PowerBIDuplicateError``.
    """
    return _run(
        "pbi_add_role_member",
        pbi_add_role_member_tool,
        CONNECTION_MANAGER,
        role=role,
        member_name=member_name,
        member_type=member_type,
        identity_provider=identity_provider,
        overwrite=overwrite,
    )


@mcp.tool()
def pbi_remove_role_member(role: str, member_name: str) -> dict[str, Any]:
    """Remove a member from an RLS role (matched on MemberName)."""
    return _run(
        "pbi_remove_role_member",
        pbi_remove_role_member_tool,
        CONNECTION_MANAGER,
        role=role,
        member_name=member_name,
    )

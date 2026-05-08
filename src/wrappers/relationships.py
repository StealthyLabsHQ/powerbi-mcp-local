"""MCP wrappers — domain: relationships."""

from __future__ import annotations

from typing import Any

from mcp_core import CONNECTION_MANAGER, _run, mcp
from tools import (
    pbi_create_relationship_tool,
    pbi_delete_relationship_tool,
    pbi_list_relationships_tool,
    pbi_update_relationship_tool,
)


@mcp.tool()
def pbi_list_relationships() -> dict[str, Any]:
    """List relationships in the active Power BI model."""
    return _run("pbi_list_relationships", pbi_list_relationships_tool, CONNECTION_MANAGER)


@mcp.tool()
def pbi_create_relationship(
    from_table: str,
    from_column: str,
    to_table: str,
    to_column: str,
    cardinality: str = "oneToMany",
    direction: str = "oneDirection",
    is_active: bool = True,
    relationship_name: str | None = None,
    overwrite: bool = False,
) -> dict[str, Any]:
    """Create or update a relationship between two columns.

    With ``overwrite=True``, an existing relationship between the same endpoint
    columns is updated in place (cardinality/direction/is_active refreshed)
    instead of raising ``PowerBIDuplicateError``.
    """
    return _run(
        "pbi_create_relationship",
        pbi_create_relationship_tool,
        CONNECTION_MANAGER,
        from_table=from_table,
        from_column=from_column,
        to_table=to_table,
        to_column=to_column,
        cardinality=cardinality,
        direction=direction,
        is_active=is_active,
        relationship_name=relationship_name,
        overwrite=overwrite,
    )


@mcp.tool()
def pbi_delete_relationship(
    name: str | None = None,
    from_table: str | None = None,
    from_column: str | None = None,
    to_table: str | None = None,
    to_column: str | None = None,
) -> dict[str, Any]:
    """Delete a relationship by name or by endpoint columns."""
    return _run(
        "pbi_delete_relationship",
        pbi_delete_relationship_tool,
        CONNECTION_MANAGER,
        name=name,
        from_table=from_table,
        from_column=from_column,
        to_table=to_table,
        to_column=to_column,
    )


@mcp.tool()
def pbi_update_relationship(
    name: str | None = None,
    from_table: str | None = None,
    from_column: str | None = None,
    to_table: str | None = None,
    to_column: str | None = None,
    cardinality: str | None = None,
    direction: str | None = None,
    is_active: bool | None = None,
    new_name: str | None = None,
) -> dict[str, Any]:
    """Update properties of an existing relationship."""
    return _run(
        "pbi_update_relationship",
        pbi_update_relationship_tool,
        CONNECTION_MANAGER,
        name=name,
        from_table=from_table,
        from_column=from_column,
        to_table=to_table,
        to_column=to_column,
        cardinality=cardinality,
        direction=direction,
        is_active=is_active,
        new_name=new_name,
    )

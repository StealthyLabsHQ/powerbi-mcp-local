"""MCP wrappers — domain: connection."""

from __future__ import annotations

from typing import Any

from mcp_core import CONNECTION_MANAGER, _run, mcp
from tools import (
    pbi_connect_tool,
    pbi_export_model_tool,
    pbi_list_instances_tool,
    pbi_operation_history_tool,
    pbi_refresh_metadata_tool,
    pbi_refresh_tool,
    pbi_system_health_tool,
)


@mcp.tool()
def pbi_connect(preferred_port: int | None = None, force_reconnect: bool = False) -> dict[str, Any]:
    """Find and connect to a running Power BI Desktop instance."""
    return _run(
        "pbi_connect",
        pbi_connect_tool,
        CONNECTION_MANAGER,
        preferred_port=preferred_port,
        force_reconnect=force_reconnect,
    )


@mcp.tool()
def pbi_list_instances() -> dict[str, Any]:
    """List discovered Power BI Desktop instances without connecting."""
    return _run("pbi_list_instances", pbi_list_instances_tool, CONNECTION_MANAGER)


@mcp.tool()
def pbi_refresh_metadata() -> dict[str, Any]:
    """Reload the cached TOM schema (cheaper than pbi_connect force_reconnect)."""
    return _run("pbi_refresh_metadata", pbi_refresh_metadata_tool, CONNECTION_MANAGER)


@mcp.tool()
def pbi_system_health() -> dict[str, Any]:
    """Single-call self-diagnostic: connection state, port/PID match, dependency
    availability, model loaded?, table/measure counts, cache state, last op
    timestamp. Read-only and safe to call without an active connection.
    """
    return _run("pbi_system_health", pbi_system_health_tool, CONNECTION_MANAGER)


@mcp.tool()
def pbi_operation_history(last_n: int = 20) -> dict[str, Any]:
    """Return the last N tool operations recorded by the connection manager
    (newest first). Useful for self-debugging after a failure: an LLM can pull
    the most recent calls to see which writes already landed.
    """
    return _run(
        "pbi_operation_history",
        pbi_operation_history_tool,
        CONNECTION_MANAGER,
        last_n=last_n,
    )


@mcp.tool()
def pbi_refresh(target: str = "model", refresh_type: str = "full") -> dict[str, Any]:
    """Trigger a model or table refresh."""
    return _run(
        "pbi_refresh",
        pbi_refresh_tool,
        CONNECTION_MANAGER,
        target=target,
        refresh_type=refresh_type,
    )


@mcp.tool()
def pbi_export_model(
    path: str | None = None,
    include_hidden: bool = False,
    include_row_counts: bool = False,
) -> dict[str, Any]:
    """Export the full model as JSON and optionally write it to disk."""
    return _run(
        "pbi_export_model",
        pbi_export_model_tool,
        CONNECTION_MANAGER,
        path=path,
        include_hidden=include_hidden,
        include_row_counts=include_row_counts,
    )

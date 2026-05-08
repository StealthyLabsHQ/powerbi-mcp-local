"""MCP wrappers — domain: workflows."""

from __future__ import annotations

from typing import Any

from mcp_core import CONNECTION_MANAGER, _run, mcp
from tools import (
    pbi_measure_workflow_tool,
    pbi_model_audit_workflow_tool,
)


@mcp.tool()
def pbi_model_audit_workflow(include_hidden: bool = False, include_row_counts: bool = True) -> dict[str, Any]:
    """Run a compact model audit workflow for agent planning."""
    return _run(
        "pbi_model_audit_workflow",
        pbi_model_audit_workflow_tool,
        CONNECTION_MANAGER,
        include_hidden=include_hidden,
        include_row_counts=include_row_counts,
    )


@mcp.tool()
def pbi_measure_workflow(
    table: str,
    measures: list[dict[str, Any]],
    overwrite: bool = True,
    apply: bool = False,
) -> dict[str, Any]:
    """Plan or run validated batch measure creation."""
    return _run(
        "pbi_measure_workflow",
        pbi_measure_workflow_tool,
        CONNECTION_MANAGER,
        table=table,
        measures=measures,
        overwrite=overwrite,
        apply=apply,
    )


# ── MCP Resources ────────────────────────────────────────────────────
# Expose live model data as MCP Resources so clients can subscribe/fetch
# without burning a tool call. Cache in the manager invalidates on writes.

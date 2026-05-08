"""MCP wrappers — domain: project."""

from __future__ import annotations

from typing import Any

from mcp_core import CONNECTION_MANAGER, _run, mcp
from tools import (
    pbi_create_persistent_report_tool,
)


@mcp.tool()
def pbi_create_persistent_report(
    output_path: str,
    tables: list[dict[str, Any]],
    measures: list[dict[str, Any]] | None = None,
    relationships: list[dict[str, Any]] | None = None,
    pages: list[dict[str, Any]] | None = None,
    open_after_create: bool = False,
) -> dict[str, Any]:
    """Create a persistent .pbix with model data, measures, pages, and native visuals."""
    return _run(
        "pbi_create_persistent_report",
        pbi_create_persistent_report_tool,
        output_path=output_path,
        tables=tables,
        measures=measures,
        relationships=relationships,
        pages=pages,
        open_after_create=open_after_create,
    )

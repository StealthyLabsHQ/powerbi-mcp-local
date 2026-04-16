"""Query execution and refresh tools for the Power BI MCP server."""

from __future__ import annotations

import os
import re
from typing import Any

from pbi_connection import PowerBINotFoundError, find_named, map_enum, ok


# ── DAX safety ───────────────────────────────────────────────────────

# DMV queries that expose server internals — blocked by default
_DMV_BLOCKED_PATTERNS = [
    re.compile(r"\$SYSTEM\.", re.IGNORECASE),
    re.compile(r"DISCOVER_", re.IGNORECASE),
    re.compile(r"DBSCHEMA_", re.IGNORECASE),
    re.compile(r"MDSCHEMA_", re.IGNORECASE),
]

# Allow override via env var (set to "1" to allow DMV for advanced users)
_ALLOW_DMV = os.environ.get("PBI_MCP_ALLOW_DMV", "0") == "1"


def _validate_dax_query(query: str) -> None:
    """Block dangerous DMV/system queries unless explicitly allowed."""
    if _ALLOW_DMV:
        return
    stripped = query.strip()
    for pattern in _DMV_BLOCKED_PATTERNS:
        if pattern.search(stripped):
            raise ValueError(
                f"DMV/system query blocked for security. "
                f"Set PBI_MCP_ALLOW_DMV=1 to allow. "
                f"Matched: {pattern.pattern}"
            )


def pbi_execute_dax_tool(
    manager: Any,
    *,
    query: str,
    max_rows: int = 1000,
) -> dict[str, Any]:
    """Execute a DAX or DMV query."""
    _validate_dax_query(query)
    result = manager.run_adomd_query(query, max_rows=max_rows)
    return ok(
        "Query executed successfully.",
        query=query,
        max_rows=max_rows,
        columns=result["columns"],
        rows=result["rows"],
        row_count=result["row_count"],
        truncated=result["truncated"],
    )


def pbi_refresh_tool(
    manager: Any,
    *,
    target: str = "model",
    refresh_type: str = "full",
) -> dict[str, Any]:
    """Trigger a model or table refresh."""

    def _mutator(state: Any, database: Any, model: Any) -> dict[str, Any]:
        tom = manager.tom
        refresh_enum = map_enum(tom.RefreshType, refresh_type)
        if target.strip().casefold() in {"model", "database"}:
            model.RequestRefresh(refresh_enum)
            scope = {"target_type": "model", "target": str(database.Name)}
        else:
            table = find_named(model.Tables, target)
            if table is None:
                raise PowerBINotFoundError(
                    f"Table '{target}' was not found.",
                    details={"table": target},
                )
            table.RequestRefresh(refresh_enum)
            scope = {"target_type": "table", "target": str(table.Name)}
        return {
            "refresh": {
                **scope,
                "refresh_type": refresh_type,
            }
        }

    payload = manager.execute_write("refresh", _mutator)
    return ok(
        f"Refresh requested successfully for {payload['refresh']['target_type']} '{payload['refresh']['target']}'.",
        refresh=payload["refresh"],
        save_result=payload["save_result"],
        connection=payload["connection"],
    )


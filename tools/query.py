"""Query execution and refresh tools for the Power BI MCP server."""

from __future__ import annotations

from typing import Any

from pbi_connection import PowerBINotFoundError, find_named, map_enum, ok


def pbi_execute_dax_tool(
    manager: Any,
    *,
    query: str,
    max_rows: int = 1000,
) -> dict[str, Any]:
    """Execute a DAX or DMV query."""
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


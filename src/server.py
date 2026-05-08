"""FastMCP server exposing Power BI Desktop model operations over stdio.

Tool registration lives here (every ``@mcp.tool()`` wrapper). The runtime
plumbing — FastMCP instance, connection manager, ``_run`` helper, PID lock
and parent-watcher lifecycle — is in :mod:`mcp_core` so this file stays
focused on the API surface.
"""

from __future__ import annotations

import os
from pathlib import Path
from typing import Any

from mcp_core import (
    CONNECTION_MANAGER,
    _acquire_single_instance_lock,
    _apply_profile,
    _audit_tool_registry,
    _run,
    _start_parent_watcher,
    logger,
    mcp,
)
from security import SECURITY

# Imports needed only by the @mcp.resource() handlers below — every other
# tools.* import has moved into the corresponding wrappers/<domain>.py module.
from tools import (
    pbi_list_measures_tool,
    pbi_list_relationships_tool,
    pbi_model_info_tool,
)


def find_pbi_port(preferred_port: int | None = None) -> int:
    """Compatibility helper for standalone scripts and README examples."""
    instances = CONNECTION_MANAGER.list_instances()
    if not instances:
        raise ValueError("No running Power BI Desktop instances were found.")
    if preferred_port is None:
        return int(instances[0]["port"])
    for instance in instances:
        if instance["port"] == preferred_port:
            return int(instance["port"])
    raise ValueError(f"No Power BI instance found on port {preferred_port}.")




# Side-effect imports: each wrappers module registers its @mcp.tool()
# wrappers as it loads, populating the FastMCP tool registry.
from wrappers import connection as _wrappers_connection  # noqa: F401
from wrappers import model as _wrappers_model  # noqa: F401
from wrappers import measures as _wrappers_measures  # noqa: F401
from wrappers import relationships as _wrappers_relationships  # noqa: F401
from wrappers import query as _wrappers_query  # noqa: F401
from wrappers import quality as _wrappers_quality  # noqa: F401
from wrappers import excel as _wrappers_excel  # noqa: F401
from wrappers import power_query as _wrappers_power_query  # noqa: F401
from wrappers import tmdl as _wrappers_tmdl  # noqa: F401
from wrappers import visuals as _wrappers_visuals  # noqa: F401
from wrappers import project as _wrappers_project  # noqa: F401
from wrappers import rls as _wrappers_rls  # noqa: F401
from wrappers import calc_groups as _wrappers_calc_groups  # noqa: F401
from wrappers import workflows as _wrappers_workflows  # noqa: F401

@mcp.resource("powerbi://model/schema")
def resource_model_schema() -> str:
    """Full model snapshot: tables, columns, measures, relationships."""
    import json
    result = _run("pbi_model_info", pbi_model_info_tool, CONNECTION_MANAGER)
    return json.dumps(result, ensure_ascii=False, indent=2)


@mcp.resource("powerbi://model/measures")
def resource_model_measures() -> str:
    """All DAX measures in the active model."""
    import json
    result = _run("pbi_list_measures", pbi_list_measures_tool, CONNECTION_MANAGER)
    return json.dumps(result, ensure_ascii=False, indent=2)


@mcp.resource("powerbi://model/relationships")
def resource_model_relationships() -> str:
    """All relationships in the active model."""
    import json
    result = _run("pbi_list_relationships", pbi_list_relationships_tool, CONNECTION_MANAGER)
    return json.dumps(result, ensure_ascii=False, indent=2)


# ── MCP Prompts ───────────────────────────────────────────────────────
# Ready-to-use workflow prompts surfaced natively to any MCP client.


@mcp.prompt()
def model_audit() -> str:
    """Full model audit: tables, measures, relationships, and improvement suggestions."""
    return (
        "Connect to Power BI Desktop. Give me a compact audit of the active model:\n"
        "- Table count + row counts for each fact table\n"
        "- Measure count per table, flag any obviously misnamed or empty measures\n"
        "- Relationship graph (from → to + cardinality + active/inactive)\n"
        "- Any table with multiple partitions or complex measure dependencies\n"
        "Then run pbi_validate_model() and list all issues and warnings.\n"
        "End with 3 concrete improvement suggestions ranked by impact."
    )


@mcp.prompt()
def time_intelligence_kit(base_measure: str = "Revenue", date_table: str = "Date", date_column: str = "Date") -> str:
    """Generate a full time-intelligence measure kit (MTD, QTD, YTD, YoY, YoY%) for a base measure."""
    return (
        f"Assuming '{date_table}' is marked as the date table, generate these measures for [{base_measure}]:\n"
        f"- {base_measure} MTD  = TOTALMTD([{base_measure}], {date_table}[{date_column}])\n"
        f"- {base_measure} QTD  = TOTALQTD([{base_measure}], {date_table}[{date_column}])\n"
        f"- {base_measure} YTD  = TOTALYTD([{base_measure}], {date_table}[{date_column}])\n"
        f"- {base_measure} YoY  = CALCULATE([{base_measure}], SAMEPERIODLASTYEAR({date_table}[{date_column}]))\n"
        f"- {base_measure} YoY% = DIVIDE([{base_measure}] - [{base_measure} YoY], [{base_measure} YoY])\n\n"
        "Use pbi_validate_dax on each expression before creating. "
        "Use pbi_create_measures (batch) to write all 5 in one call. "
        "Apply format string '#,##0.00' to MTD/QTD/YTD/YoY and '0.00%' to YoY%."
    )


@mcp.prompt()
def star_schema_builder(fact_table: str = "FactSales") -> str:
    """Guide for wiring a star schema: relationships + key measures."""
    return (
        f"Inspect the model with pbi_list_tables and pbi_list_relationships.\n"
        f"For fact table '{fact_table}':\n"
        "1. Identify all dimension tables by looking for columns that match FK columns in the fact.\n"
        "2. Create missing Many-to-One relationships (fact → dimension, oneDirection).\n"
        "3. Flag any existing Many-to-Many relationships and suggest a bridge table fix.\n"
        "4. Create a basic measure '[Row Count]' = COUNTROWS(fact_table) as a sanity check.\n"
        "5. Run pbi_validate_model() at the end and report any remaining issues."
    )


@mcp.prompt()
def rls_setup(table: str = "Sales", filter_column: str = "Region") -> str:
    """Set up Row-Level Security for a given table and filter column."""
    return (
        f"Set up Row-Level Security on '{table}[{filter_column}]':\n"
        "1. pbi_list_roles() — check if a role already exists.\n"
        "2. pbi_create_role(role='RegionFilter') — create a new role.\n"
        f"3. pbi_set_role_filter(role='RegionFilter', table='{table}', filter_expression='[{filter_column}] = USERNAME()') — apply filter.\n"
        "4. pbi_execute_dax_as_role(query='EVALUATE ROW(\"User\", USERNAME())', role='RegionFilter') — validate.\n"
        "5. Summarize the final RLS setup."
    )


@mcp.prompt()
def dead_measure_scan() -> str:
    """Find measures not referenced by any other measure and suggest cleanup."""
    return (
        "Find measures that are not referenced by any other measure:\n"
        "1. pbi_measure_dependencies() — get the full dependency graph.\n"
        "2. pbi_list_measures() — list all measures.\n"
        "3. Cross-reference: which measures appear only as roots (nothing depends on them)?\n"
        "4. For each orphan, show its expression and suggest: keep / rename / delete.\n"
        "Do NOT delete anything — only report recommendations."
    )


@mcp.prompt()
def bulk_measure_format_fix(table: str = "Measures", format_string: str = "#,##0") -> str:
    """Apply a format string to all measures in a table that are missing one."""
    return (
        f"Find all measures in table '{table}' that have no format string set.\n"
        "Use pbi_list_measures(include_hidden=False) to get the full list.\n"
        "Filter to those where format_string is empty or null.\n"
        f"Apply format string '{format_string}' to all of them using pbi_set_format(table='{table}', names=[...], format_string='{format_string}').\n"
        "Report how many were updated."
    )


@mcp.prompt()
def excel_to_pbi_pipeline(excel_path: str = "") -> str:
    """Full pipeline: inspect Excel, create import queries, refresh, validate."""
    path_hint = f"'{excel_path}'" if excel_path else "<path/to/file.xlsx>"
    return (
        f"Run the full Excel → Power BI import pipeline for {path_hint}:\n"
        "1. excel_workbook_info() — list sheets and row counts.\n"
        "2. pbi_bulk_import_excel(excel_path=..., refresh_after=False) — inject import queries for all sheets.\n"
        "3. pbi_refresh(target='model', refresh_type='full') — refresh the model.\n"
        "4. pbi_list_tables(include_row_counts=True) — verify row counts match the Excel source.\n"
        "5. Report any discrepancies."
    )


@mcp.prompt()
def model_snapshot_export(output_path: str = "./docs/model.json") -> str:
    """Export the full model as JSON for documentation or version control."""
    return (
        f"Export the full model definition to '{output_path}':\n"
        f"pbi_export_model(path='{output_path}', include_hidden=True, include_row_counts=False)\n\n"
        "Then summarize:\n"
        "- Total tables, measures, relationships\n"
        "- Top 5 most complex measures by expression length\n"
        "- Any tables with more than 2 partitions"
    )


def _bearer_auth_middleware(app: Any, token: str) -> Any:
    """ASGI middleware that requires Authorization: Bearer <token> on HTTP requests."""
    expected = f"Bearer {token}".encode("utf-8")

    async def wrapped(scope: dict, receive: Any, send: Any) -> None:
        if scope.get("type") != "http":
            await app(scope, receive, send)
            return
        headers = dict(scope.get("headers") or [])
        provided = headers.get(b"authorization", b"")
        if provided != expected:
            await send({
                "type": "http.response.start",
                "status": 401,
                "headers": [(b"content-type", b"text/plain"), (b"www-authenticate", b"Bearer")],
            })
            await send({"type": "http.response.body", "body": b"Unauthorized"})
            return
        await app(scope, receive, send)

    return wrapped


async def _run_sse_with_auth(host: str, port: int) -> None:
    """Mirror FastMCP.run_sse_async but allow wrapping with Bearer auth middleware."""
    import uvicorn

    mcp.settings.host = host
    mcp.settings.port = port

    asgi_app = mcp.sse_app()
    token = os.getenv("PBI_MCP_AUTH_TOKEN", "").strip()
    if token:
        asgi_app = _bearer_auth_middleware(asgi_app, token)
        logger.info("SECURITY: SSE Bearer auth enabled (PBI_MCP_AUTH_TOKEN set).")
    else:
        logger.warning(
            "SECURITY: SSE has no Bearer auth. Set PBI_MCP_AUTH_TOKEN to require "
            "'Authorization: Bearer <token>' on HTTP requests.",
        )

    config = uvicorn.Config(
        asgi_app,
        host=host,
        port=port,
        log_level=mcp.settings.log_level.lower(),
    )
    server = uvicorn.Server(config)
    await server.serve()


def main() -> None:
    """Entry point — supports stdio (default) and sse transport."""
    import argparse

    parser = argparse.ArgumentParser(description="Power BI Desktop MCP Server")
    parser.add_argument(
        "--transport",
        choices=["stdio", "sse"],
        default="stdio",
        help="MCP transport: stdio (CLI tools) or sse (web/IDE clients)",
    )
    parser.add_argument(
        "--port",
        type=int,
        default=8765,
        help="Port for SSE transport (default: 8765)",
    )
    parser.add_argument(
        "--host",
        default="127.0.0.1",
        help="Host for SSE transport (default: 127.0.0.1 — localhost only)",
    )
    parser.add_argument(
        "--readonly",
        action="store_true",
        help="Disable write and destructive tools for this server process.",
    )
    parser.add_argument(
        "--profile",
        choices=["readonly", "write", "all", "grading"],
        default="all",
        help="Filter exposed tool surface: readonly, write (read+write), grading (analysis + scoring only), or all (default).",
    )
    args = parser.parse_args()
    SECURITY.policy(reload=True, cwd=Path(__file__).parent)
    if args.readonly:
        SECURITY.set_runtime_readonly(True)
        logger.info("SECURITY: readonly mode enabled via --readonly")

    # Pre-flight: detect registration drift between tools/__all__ and @mcp.tool()
    # wrappers. Opt-in via env var so production servers skip the introspection
    # cost at every startup. Strict mode (CI) fails on missing wrappers.
    if os.environ.get("PBI_MCP_AUDIT", "0") == "1" or os.environ.get("PBI_MCP_STRICT_REGISTRY", "0") == "1":
        _audit_tool_registry(strict=os.environ.get("PBI_MCP_STRICT_REGISTRY", "0") == "1")

    _apply_profile(args.profile)

    if args.transport == "sse":
        logger.info(
            "SSE server starting on %s:%d (localhost-only by default)",
            args.host, args.port,
        )
        if args.host != "127.0.0.1":
            logger.warning(
                "SECURITY: SSE bound to %s — exposed beyond localhost. "
                "Ensure network is trusted or use --host 127.0.0.1",
                args.host,
            )
        import anyio
        anyio.run(_run_sse_with_auth, args.host, args.port)
    else:
        _acquire_single_instance_lock()
        _start_parent_watcher()
        mcp.run(transport="stdio")


if __name__ == "__main__":
    main()

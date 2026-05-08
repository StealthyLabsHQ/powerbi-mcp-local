"""MCP wrappers — domain: query."""

from __future__ import annotations

from typing import Any

from mcp_core import CONNECTION_MANAGER, _run, mcp
from tools import (
    pbi_execute_dax_as_role_tool,
    pbi_execute_dax_tool,
    pbi_generate_dax_context_prompt_tool,
    pbi_trace_query_tool,
    pbi_validate_dax_semantic_tool,
    pbi_validate_dax_tool,
    pbi_validate_filter_expression_tool,
)


@mcp.tool()
def pbi_execute_dax(
    query: str,
    max_rows: int = 1000,
    timeout_seconds: int | None = None,
) -> dict[str, Any]:
    """Execute a DAX or DMV query and return rows. timeout_seconds=0 disables the timeout."""
    return _run(
        "pbi_execute_dax",
        pbi_execute_dax_tool,
        CONNECTION_MANAGER,
        query=query,
        max_rows=max_rows,
        timeout_seconds=timeout_seconds,
    )


@mcp.tool()
def pbi_execute_dax_as_role(query: str, role: str, username: str | None = None) -> dict[str, Any]:
    """Execute a DAX query under a specific RLS role and optional effective user."""
    return _run(
        "pbi_execute_dax_as_role",
        pbi_execute_dax_as_role_tool,
        CONNECTION_MANAGER,
        query=query,
        role=role,
        username=username,
    )


@mcp.tool()
def pbi_trace_query(
    query: str,
    timeout_seconds: int | None = None,
) -> dict[str, Any]:
    """Execute a DAX query and return rows plus performance diagnostics."""
    return _run(
        "pbi_trace_query",
        pbi_trace_query_tool,
        CONNECTION_MANAGER,
        query=query,
        timeout_seconds=timeout_seconds,
    )


@mcp.tool()
def pbi_validate_dax(expression: str, kind: str = "scalar") -> dict[str, Any]:
    """Parse-check a DAX expression. kind='scalar' or 'table'."""
    return _run(
        "pbi_validate_dax",
        pbi_validate_dax_tool,
        CONNECTION_MANAGER,
        expression=expression,
        kind=kind,
    )


@mcp.tool()
def pbi_validate_dax_semantic(
    expression: str,
    kind: str = "scalar",
    format_string: str = "",
    include_hidden: bool = False,
) -> dict[str, Any]:
    """Validate a DAX expression with three layers: reference existence
    (column / measure typo detection against the live model), format-string
    sanity heuristic (percent format on a money expression, currency on a
    ratio), and the runtime ASEngine probe.

    Returns ``{valid, syntax: ok|error, semantic: {unknown_references,
    suspicious_format, columns_referenced, measures_referenced},
    runtime_error?}``. Use as the canonical preflight before committing a
    new DAX measure.
    """
    return _run(
        "pbi_validate_dax_semantic",
        pbi_validate_dax_semantic_tool,
        CONNECTION_MANAGER,
        expression=expression,
        kind=kind,
        format_string=format_string,
        include_hidden=include_hidden,
    )


@mcp.tool()
def pbi_validate_filter_expression(filter_expression: str) -> dict[str, Any]:
    """Validate a DAX boolean filter expression before visual probes."""
    return _run(
        "pbi_validate_filter_expression",
        pbi_validate_filter_expression_tool,
        CONNECTION_MANAGER,
        filter_expression=filter_expression,
    )


@mcp.tool()
def pbi_generate_dax_context_prompt(
    include_hidden: bool = False,
    include_dax: bool = True,
    include_relationships: bool = True,
    max_chars: int = 12000,
) -> dict[str, Any]:
    """Render a compact markdown snapshot of the model (tables, columns,
    measures, relationships) for direct paste into an LLM system prompt.

    Use before asking another LLM to write DAX so it has the schema in one
    round-trip. Output is truncated to ``max_chars`` (default 12 000) with a
    trailing notice when truncation kicks in. Set ``include_dax=False`` for
    a terser version that omits measure expressions.
    """
    return _run(
        "pbi_generate_dax_context_prompt",
        pbi_generate_dax_context_prompt_tool,
        CONNECTION_MANAGER,
        include_hidden=include_hidden,
        include_dax=include_dax,
        include_relationships=include_relationships,
        max_chars=max_chars,
    )

"""MCP wrappers — domain: query."""

from __future__ import annotations

from tools import (
    pbi_execute_dax_as_role_tool,
    pbi_execute_dax_tool,
    pbi_generate_dax_context_prompt_tool,
    pbi_trace_query_tool,
    pbi_validate_dax_semantic_tool,
    pbi_validate_dax_tool,
    pbi_validate_filter_expression_tool,
)

from ._helpers import register_tool

register_tool(pbi_execute_dax_tool)
register_tool(pbi_execute_dax_as_role_tool)
register_tool(pbi_trace_query_tool)
register_tool(pbi_validate_dax_tool)
register_tool(pbi_validate_dax_semantic_tool)
register_tool(pbi_validate_filter_expression_tool)
register_tool(pbi_generate_dax_context_prompt_tool)

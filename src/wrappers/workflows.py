"""MCP wrappers — domain: workflows."""

from __future__ import annotations

from tools import (
    pbi_measure_workflow_tool,
    pbi_model_audit_workflow_tool,
)

from ._helpers import register_tool

register_tool(pbi_model_audit_workflow_tool)
register_tool(pbi_measure_workflow_tool)

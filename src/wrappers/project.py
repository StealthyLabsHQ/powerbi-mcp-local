"""MCP wrappers — domain: project."""

from __future__ import annotations

from tools import (
    pbi_apply_style_preset_tool,
    pbi_check_scaffold_spec_dbcc_risks_tool,
    pbi_create_persistent_report_tool,
    pbi_diagnose_pbix_dbcc_tool,
    pbi_list_scaffold_templates_tool,
    pbi_list_style_presets_tool,
    pbi_scaffold_pbix_tool,
)

from ._helpers import register_tool

register_tool(pbi_create_persistent_report_tool)
register_tool(pbi_scaffold_pbix_tool)
register_tool(pbi_list_scaffold_templates_tool)
register_tool(pbi_diagnose_pbix_dbcc_tool)
register_tool(pbi_check_scaffold_spec_dbcc_risks_tool)
register_tool(pbi_apply_style_preset_tool)
register_tool(pbi_list_style_presets_tool)

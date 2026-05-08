"""MCP wrappers — domain: measures."""

from __future__ import annotations

from tools import (
    pbi_apply_format_preset_tool,
    pbi_create_contribution_measure_tool,
    pbi_create_measure_tool,
    pbi_create_measures_tool,
    pbi_create_mtd_measure_tool,
    pbi_create_rolling_average_measure_tool,
    pbi_create_spy_measure_tool,
    pbi_create_time_intelligence_pack_tool,
    pbi_create_topn_measure_tool,
    pbi_create_variance_measure_tool,
    pbi_create_yoy_measure_tool,
    pbi_create_ytd_measure_tool,
    pbi_delete_measure_tool,
    pbi_import_dax_file_tool,
    pbi_list_format_presets_tool,
    pbi_list_measures_tool,
    pbi_measure_dependencies_tool,
    pbi_rename_measure_tool,
    pbi_set_format_tool,
)

from ._helpers import register_tool

register_tool(pbi_list_measures_tool)
register_tool(pbi_create_measure_tool)
register_tool(pbi_create_measures_tool)
register_tool(pbi_create_time_intelligence_pack_tool)
register_tool(pbi_create_ytd_measure_tool)
register_tool(pbi_create_mtd_measure_tool)
register_tool(pbi_create_spy_measure_tool)
register_tool(pbi_create_yoy_measure_tool)
register_tool(pbi_create_variance_measure_tool)
register_tool(pbi_create_contribution_measure_tool)
register_tool(pbi_create_topn_measure_tool)
register_tool(pbi_create_rolling_average_measure_tool)
register_tool(pbi_list_format_presets_tool)
register_tool(pbi_apply_format_preset_tool)
register_tool(pbi_delete_measure_tool)
register_tool(pbi_rename_measure_tool)
register_tool(pbi_measure_dependencies_tool)
register_tool(pbi_import_dax_file_tool)
register_tool(pbi_set_format_tool)

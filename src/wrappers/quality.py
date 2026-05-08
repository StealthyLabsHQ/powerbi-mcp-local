"""MCP wrappers — domain: quality."""

from __future__ import annotations

from tools import (
    pbi_audit_model_tool,
    pbi_compare_report_versions_tool,
    pbi_detect_circular_dependencies_tool,
    pbi_detect_dirty_dates_tool,
    pbi_detect_empty_visuals_tool,
    pbi_detect_missing_visuals_tool,
    pbi_detect_name_collisions_tool,
    pbi_export_correction_report_tool,
    pbi_export_validation_report_tool,
    pbi_generate_measure_tests_tool,
    pbi_lint_dax_tool,
    pbi_lint_report_layout_tool,
    pbi_run_scenario_tool,
    pbi_score_dashboard_tool,
    pbi_score_rubric_tool,
    pbi_validate_pbix_persistence_tool,
    pbi_validate_pbix_reopen_tool,
    pbi_validate_power_query_steps_tool,
    pbi_validate_relationship_plan_tool,
    pbi_validate_star_schema_tool,
    pbi_validate_visual_bindings_tool,
)

from ._helpers import register_tool

register_tool(pbi_audit_model_tool)
register_tool(pbi_validate_star_schema_tool)
register_tool(pbi_detect_circular_dependencies_tool)
register_tool(pbi_validate_power_query_steps_tool)
register_tool(pbi_detect_missing_visuals_tool)
register_tool(pbi_score_rubric_tool)
register_tool(pbi_export_correction_report_tool)
register_tool(pbi_lint_dax_tool)
register_tool(pbi_detect_name_collisions_tool)
register_tool(pbi_detect_dirty_dates_tool)
register_tool(pbi_validate_relationship_plan_tool)
register_tool(pbi_detect_empty_visuals_tool)
register_tool(pbi_generate_measure_tests_tool)
register_tool(pbi_validate_pbix_persistence_tool)
register_tool(pbi_validate_pbix_reopen_tool)
register_tool(pbi_export_validation_report_tool)
register_tool(pbi_lint_report_layout_tool)
register_tool(pbi_validate_visual_bindings_tool)
register_tool(pbi_score_dashboard_tool)
register_tool(pbi_run_scenario_tool)
register_tool(pbi_compare_report_versions_tool)

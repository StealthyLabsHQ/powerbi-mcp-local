"""Quality gates for Power BI models, DAX, report layout, and scenario runs.

Re-export façade for the quality package. The implementation lives in:

- ``_shared``: layout loading, visual config/bounds/overlap helpers, constants.
- ``_model_audit``: model audit, name collisions, dirty dates, relationship
  plan, star schema, circular dependencies.
- ``_dax_lint``: DAX lint, filter-expression and Power Query validation,
  measure smoke tests.
- ``_layout_lint``: report layout lint, visual bindings, empty/missing
  visuals, report version diffs.
- ``_persistence``: PBIX persistence validation and the Power BI Desktop
  reopen probe (UIAutomation + screenshot + Windows OCR).
- ``_scoring``: dashboard scoring, rubric scoring, scenario runs, and
  validation/correction report exports.

The reopen probe in ``_persistence`` matches Power BI Desktop's modal dialog
text against known repair/DBCC signals, including: "Fix this", "Something's
wrong with one or more fields", "See details", "Something went wrong",
"Database consistency checks", "DBCC", "Vertipaq", "string store",
"An error occurred while loading", "Report this issue",
"Copy details to clipboard", and "multiple tables".

Cross-tool calls inside the package resolve helpers through this module at
call time (``from . import _model_snapshot``), so patching
``tools.quality.<name>`` keeps working exactly as it did when quality was a
single module.
"""

from __future__ import annotations

import json
import os
import subprocess
import zipfile
from datetime import UTC, datetime
from pathlib import Path
from typing import Any

from pbi_connection import PowerBIValidationError, dax_quote_table_name, ok
from security import resolve_local_path

from ._dax_lint import (
    _format_matches,
    _measure_expected_format,
    _measure_ref,
    _selected_measures,
    pbi_generate_measure_tests_tool,
    pbi_lint_dax_tool,
    pbi_validate_filter_expression_tool,
    pbi_validate_power_query_steps_tool,
)
from ._layout_lint import (
    _filtered_table_query,
    _measure_aliases,
    _visual_query_parts,
    pbi_compare_report_versions_tool,
    pbi_detect_empty_visuals_tool,
    pbi_detect_missing_visuals_tool,
    pbi_lint_report_layout_tool,
    pbi_validate_visual_bindings_tool,
)
from ._model_audit import (
    _column_profile,
    _duplicate_relationship_key_issues,
    _find_column,
    _find_table,
    _graph_paths,
    _model_audit_from_snapshot,
    _table_map,
    pbi_audit_model_tool,
    pbi_detect_circular_dependencies_tool,
    pbi_detect_dirty_dates_tool,
    pbi_detect_name_collisions_tool,
    pbi_validate_relationship_plan_tool,
    pbi_validate_star_schema_tool,
)
from ._persistence import (
    _analyze_reopen_screenshot,
    _layout_summary,
    _ocr_reopen_screenshot,
    _run_reopen_probe,
    pbi_validate_pbix_persistence_tool,
    pbi_validate_pbix_reopen_tool,
)
from ._scoring import (
    _score_parts,
    pbi_export_correction_report_tool,
    pbi_export_validation_report_tool,
    pbi_run_scenario_tool,
    pbi_score_dashboard_tool,
    pbi_score_rubric_tool,
)
from ._shared import (
    _PS_UTF8_PRELUDE,
    DATE_PARSE_FORMATS,
    LAYOUT_RELATIVE_PATH,
    MAX_VISUALS_PER_PAGE,
    MIN_VISUAL_HEIGHT,
    MIN_VISUAL_WIDTH,
    _bounds,
    _dax_column,
    _load_layout,
    _model_snapshot,
    _overlap_area,
    _row_value,
    _visual_config,
    _visual_has_title,
    _visual_name,
    _visual_type,
)

__all__ = [
    "pbi_audit_model_tool",
    "pbi_compare_report_versions_tool",
    "pbi_detect_circular_dependencies_tool",
    "pbi_detect_dirty_dates_tool",
    "pbi_detect_empty_visuals_tool",
    "pbi_detect_missing_visuals_tool",
    "pbi_detect_name_collisions_tool",
    "pbi_export_correction_report_tool",
    "pbi_export_validation_report_tool",
    "pbi_generate_measure_tests_tool",
    "pbi_lint_dax_tool",
    "pbi_lint_report_layout_tool",
    "pbi_run_scenario_tool",
    "pbi_score_dashboard_tool",
    "pbi_score_rubric_tool",
    "pbi_validate_filter_expression_tool",
    "pbi_validate_pbix_persistence_tool",
    "pbi_validate_pbix_reopen_tool",
    "pbi_validate_power_query_steps_tool",
    "pbi_validate_relationship_plan_tool",
    "pbi_validate_star_schema_tool",
    "pbi_validate_visual_bindings_tool",
]

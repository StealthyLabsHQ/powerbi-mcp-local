"""Report page and visual automation tools.

This package re-exports the entire visuals tool surface from focused
submodules. Existing ``from tools.visuals import …`` imports keep working
unchanged. Internally each concern lives in its own file:

- ``_base``      constants + error classes
- ``_paths``     path resolution
- ``_refs``      field reference normalisation
- ``_layout``    layout I/O + atomic write + dry_run + page helpers
- ``_formatting`` literal / color / title encoders
- ``_home_tables`` measure → home table resolution
- ``_bindings``  prototype/select builders + validators
- ``_containers`` visual container construction + append flow
- ``_io``        pbi-tools CLI + zip extraction + PowerShell helpers
- ``_pages``     page-level tools
- ``_charts`` / ``_cards`` / ``_structure`` 14 ``pbi_add_*_tool`` builders
- ``_ops``       remove/move/format/convert/auto-grid/patch
- ``_design``    DESIGN_PRESETS + apply_theme + apply_design + build_dashboard
- ``_repair``    validate / repair report fields
- ``_dispatcher`` ``pbi_add_visual_tool`` + ``_VISUAL_TYPE_DISPATCH``

The ``pbi_model_info_tool`` re-export at module level is intentional: it
keeps existing test patches against ``tools.visuals.pbi_model_info_tool``
working, since several submodules look the function up through this
namespace at call time.
"""

from __future__ import annotations

import logging
import os
import shutil
import subprocess
import tempfile
import time
import zipfile
from collections.abc import Callable
from pathlib import Path
from typing import Any

from pbi_connection import error_payload

from ..model import pbi_model_info_tool
from ._base import (
    DEFAULT_PAGE_HEIGHT,
    DEFAULT_PAGE_WIDTH,
    DEFAULT_VISUAL_SIZES,
    DESIGN_THEME_RELATIVE_PATH,
    HEX_COLOR_RE,
    LAYOUT_RELATIVE_PATH,
    MODEL_TABLES_RELATIVE_DIR,
    THEMES_RELATIVE_DIR,
    VISUAL_FIELD_ROLES,
    VISUAL_ROLE_KINDS,
    PageNotFoundError,
    PBIToolsNotInstalledError,
    ReportLayoutError,
    VisualNotFoundError,
    VisualToolError,
)
from ._layout import (
    _LAYOUT_WRITE_TL,
    _dump_embedded_json,
    _find_page,
    _is_dry_run,
    _load_layout,
    _next_page_name,
    _normalize_page_name,
    _page_summary,
    _parse_embedded_json,
    _record_dry_run_write,
    _save_layout,
    dry_run_layout_writes,
)
from ._paths import (
    _layout_path,
    _resolve_extract_folder,
    _resolve_pbix_path,
    _resolve_theme_path,
)
from ._refs import (
    _BRACKET_REF_RE,
    _normalize_reference,
    _query_ref,
    _split_column_ref,
)

logger = logging.getLogger(__name__)

# Re-export the offline ``_run`` from ``_base`` so historic
# ``from tools.visuals import _run`` imports keep working.
from ._base import _run  # noqa: E402, F401
from ._bindings import (
    _assert_container_bindings,
    _build_prototype_query,
    _build_select_entry,
    _from_entity_by_alias,
    _live_model_field_index,
    _next_alias,
    _scan_visual_bindings,
    _select_name_map,
    _sync_container_query,
    _validate_field_references_live,
    _validate_projection_roles,
    _visual_binding_issues,
)
from ._cards import (
    pbi_add_card_tool,
    pbi_add_gauge_tool,
    pbi_add_kpi_tool,
    pbi_add_labelled_card_tool,
    pbi_add_text_box_tool,
)
from ._charts import (
    pbi_add_area_chart_tool,
    pbi_add_bar_chart_tool,
    pbi_add_clustered_column_chart_tool,
    pbi_add_combo_chart_tool,
    pbi_add_donut_chart_tool,
    pbi_add_funnel_tool,
    pbi_add_hundred_percent_stacked_area_chart_tool,
    pbi_add_hundred_percent_stacked_bar_chart_tool,
    pbi_add_hundred_percent_stacked_column_chart_tool,
    pbi_add_line_chart_tool,
    pbi_add_multi_row_card_tool,
    pbi_add_pie_chart_tool,
    pbi_add_ribbon_chart_tool,
    pbi_add_scatter_chart_tool,
    pbi_add_stacked_area_chart_tool,
    pbi_add_stacked_bar_chart_tool,
    pbi_add_stacked_column_chart_tool,
    pbi_add_treemap_tool,
    pbi_add_waterfall_tool,
)
from ._containers import (
    _append_visual,
    _base_visual_config,
    _create_chart_container,
    _find_visual,
    _make_visual_container,
    _page_next_z,
    _unique_visual_id,
    _validate_dimensions,
    _visual_payload,
)
from ._design import (
    DESIGN_PRESETS,
    pbi_apply_design_tool,
    pbi_apply_theme_tool,
    pbi_build_dashboard_tool,
)
from ._dispatcher import (
    _VISUAL_TYPE_DISPATCH,
    pbi_add_visual_tool,
)
from ._formatting import (
    _VISUAL_FORMAT_TYPES,
    _datapoint_fill_objects,
    _decimal_literal,
    _encode_visual_format_value,
    _gauge_axis_objects,
    _int_literal,
    _literal_value,
    _solid_color,
    _text_literal,
    _title_objects,
)
from ._home_tables import (
    _augment_measure_home_map_with_live,
    _inspect_value_measures,
    _persistence_risks,
    _resolve_measure_home_map,
    _scan_measure_home_tables,
)
from ._io import (
    _extract_pbix_zip_natively,
    _find_pbi_tools,
    _force_kill_powerbi,
    _maybe_force_close_powerbi,
    _page_names_from_layout_bytes,
    _run_pbi_tools,
    _run_powershell,
    _save_and_close_powerbi_gracefully,
    pbi_compile_report_tool,
    pbi_extract_report_tool,
)
from ._ops import (
    pbi_add_conditional_formatting_tool,
    pbi_auto_grid_layout_tool,
    pbi_convert_visual_type_tool,
    pbi_disable_card_autoscale_tool,
    pbi_move_visual_tool,
    pbi_patch_layout_tool,
    pbi_remove_visual_tool,
    pbi_set_series_color_tool,
    pbi_set_visual_format_property_tool,
    pbi_update_visual_bindings_tool,
)
from ._pages import (
    pbi_create_page_tool,
    pbi_delete_page_tool,
    pbi_describe_page_tool,
    pbi_get_page_tool,
    pbi_list_pages_tool,
    pbi_set_page_size_tool,
)
from ._repair import (
    pbi_diagnose_render_risks_tool,
    pbi_repair_report_fields_tool,
    pbi_validate_report_fields_tool,
)
from ._structure import (
    pbi_add_map_tool,
    pbi_add_matrix_tool,
    pbi_add_slicer_tool,
    pbi_add_table_visual_tool,
)

__all__ = [
    "pbi_add_visual_tool",
    "pbi_add_area_chart_tool",
    "pbi_add_bar_chart_tool",
    "pbi_add_clustered_column_chart_tool",
    "pbi_add_funnel_tool",
    "pbi_add_hundred_percent_stacked_area_chart_tool",
    "pbi_add_hundred_percent_stacked_bar_chart_tool",
    "pbi_add_hundred_percent_stacked_column_chart_tool",
    "pbi_add_multi_row_card_tool",
    "pbi_add_pie_chart_tool",
    "pbi_add_ribbon_chart_tool",
    "pbi_add_stacked_area_chart_tool",
    "pbi_add_stacked_bar_chart_tool",
    "pbi_add_stacked_column_chart_tool",
    "pbi_add_treemap_tool",
    "pbi_add_card_tool",
    "pbi_add_combo_chart_tool",
    "pbi_add_donut_chart_tool",
    "pbi_add_gauge_tool",
    "pbi_add_kpi_tool",
    "pbi_add_labelled_card_tool",
    "pbi_add_line_chart_tool",
    "pbi_add_map_tool",
    "pbi_add_matrix_tool",
    "pbi_add_scatter_chart_tool",
    "pbi_add_slicer_tool",
    "pbi_add_table_visual_tool",
    "pbi_add_text_box_tool",
    "pbi_add_waterfall_tool",
    "pbi_apply_design_tool",
    "pbi_apply_theme_tool",
    "pbi_build_dashboard_tool",
    "pbi_compile_report_tool",
    "pbi_create_page_tool",
    "pbi_delete_page_tool",
    "pbi_extract_report_tool",
    "pbi_get_page_tool",
    "pbi_list_pages_tool",
    "pbi_move_visual_tool",
    "pbi_patch_layout_tool",
    "pbi_repair_report_fields_tool",
    "pbi_auto_grid_layout_tool",
    "pbi_convert_visual_type_tool",
    "pbi_describe_page_tool",
    "pbi_diagnose_render_risks_tool",
    "pbi_disable_card_autoscale_tool",
    "pbi_remove_visual_tool",
    "pbi_set_page_size_tool",
    "pbi_set_series_color_tool",
    "pbi_set_visual_format_property_tool",
    "pbi_add_conditional_formatting_tool",
    "pbi_update_visual_bindings_tool",
    "pbi_validate_report_fields_tool",
]

"""Measure operations for the Power BI MCP server.

Re-export façade for the measures package. The implementation lives in:

- ``_crud``: create/rename/delete/list measures, batch create, format strings.
- ``_time_intelligence``: DAX ref quoting, the time-intelligence template
  catalogue and pattern resolution, the TI pack + single-pattern wrappers
  (YTD/MTD/SPY/YOY), and the variance/contribution/Top-N/rolling generators.
- ``_dax_import``: .dax file parsing, comment stripping, and bulk import.

Cross-tool calls inside the package resolve helpers through this module at
call time (``from . import pbi_create_measure_tool``), so patching
``tools.measures.<name>`` keeps working exactly as it did when measures was a
single module.
"""

from __future__ import annotations

from ._crud import (
    pbi_create_measure_tool,
    pbi_create_measures_tool,
    pbi_delete_measure_tool,
    pbi_list_measures_tool,
    pbi_rename_measure_tool,
    pbi_set_format_tool,
)
from ._dax_import import (
    _parse_dax_file,
    _strip_dax_comments,
    pbi_import_dax_file_tool,
)
from ._time_intelligence import (
    _DEFAULT_TIME_INTELLIGENCE_PATTERNS,
    _TIME_INTELLIGENCE_TEMPLATES,
    _create_ti_single,
    _dax_column_ref,
    _dax_table_ref,
    _resolve_ti_patterns,
    pbi_create_contribution_measure_tool,
    pbi_create_mtd_measure_tool,
    pbi_create_rolling_average_measure_tool,
    pbi_create_spy_measure_tool,
    pbi_create_time_intelligence_pack_tool,
    pbi_create_topn_measure_tool,
    pbi_create_variance_measure_tool,
    pbi_create_yoy_measure_tool,
    pbi_create_ytd_measure_tool,
)

__all__ = [
    "pbi_create_contribution_measure_tool",
    "pbi_create_measure_tool",
    "pbi_create_measures_tool",
    "pbi_create_mtd_measure_tool",
    "pbi_create_rolling_average_measure_tool",
    "pbi_create_spy_measure_tool",
    "pbi_create_time_intelligence_pack_tool",
    "pbi_create_topn_measure_tool",
    "pbi_create_variance_measure_tool",
    "pbi_create_yoy_measure_tool",
    "pbi_create_ytd_measure_tool",
    "pbi_delete_measure_tool",
    "pbi_import_dax_file_tool",
    "pbi_list_measures_tool",
    "pbi_rename_measure_tool",
    "pbi_set_format_tool",
]

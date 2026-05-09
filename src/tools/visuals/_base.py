"""Visuals package: shared constants, error types, and the offline ``_run``
helper.

Lives at the bottom of the package import graph so other submodules can
freely import from it without risk of cycles.
"""

from __future__ import annotations

import re
from collections.abc import Callable
from pathlib import Path
from typing import Any

from pbi_connection import PowerBIError, error_payload


def _run(callback: Callable[[], dict[str, Any]]) -> dict[str, Any]:
    """Execute ``callback`` and translate any raised exception into the
    standard error payload. Each visuals tool wraps its body with this so
    callers always get a well-shaped ``{ok: False, error: …}`` response
    instead of a stack trace.
    """
    try:
        return callback()
    except Exception as exc:
        return error_payload(exc)


DEFAULT_PAGE_WIDTH = 1280
DEFAULT_PAGE_HEIGHT = 720
LAYOUT_RELATIVE_PATH = Path("Report") / "Layout"
THEMES_RELATIVE_DIR = Path("Report") / "StaticResources" / "Themes"
DESIGN_THEME_RELATIVE_PATH = Path("Report") / "StaticResources" / "SharedResources" / "BaseThemes" / "CY26SU02.json"
MODEL_TABLES_RELATIVE_DIR = Path("Model") / "tables"
HEX_COLOR_RE = re.compile(r"^#[0-9A-Fa-f]{6}$")

DEFAULT_VISUAL_SIZES = {
    "card": (200, 120),
    "bar_chart": (400, 300),
    "line_chart": (420, 300),
    "donut": (320, 280),
    "table": (520, 320),
    "waterfall": (420, 300),
    "slicer": (220, 120),
    "text": (280, 80),
    "gauge": (280, 220),
    "kpi": (260, 140),
    "map": (420, 320),
}

VISUAL_FIELD_ROLES = {
    "card": {"Values"},
    "multiRowCard": {"Category", "Values"},
    "clusteredBarChart": {"Category", "Y", "Series"},
    "clusteredColumnChart": {"Category", "Y", "Series"},
    "stackedBarChart": {"Category", "Y", "Series"},
    "stackedColumnChart": {"Category", "Y", "Series"},
    "hundredPercentStackedBarChart": {"Category", "Y", "Series"},
    "hundredPercentStackedColumnChart": {"Category", "Y", "Series"},
    "ribbonChart": {"Category", "Y", "Series"},
    "lineChart": {"Category", "Y", "Series"},
    "areaChart": {"Category", "Y", "Series"},
    "stackedAreaChart": {"Category", "Y", "Series"},
    "hundredPercentStackedAreaChart": {"Category", "Y", "Series"},
    "donutChart": {"Category", "Y"},
    "pieChart": {"Category", "Y"},
    "treemap": {"Category", "Details", "Values"},
    "funnel": {"Group", "Values"},
    "tableEx": {"Values"},
    "waterfallChart": {"Category", "Y"},
    "slicer": {"Values"},
    "gauge": {"Y", "Goal"},
    "kpi": {"Indicator", "TrendLine", "Goal"},
    "map": {"Category", "Y"},
    "scatterChart": {"Category", "X", "Y", "Size", "Series"},
    "lineClusteredColumnComboChart": {"Category", "Y", "Y2", "Series"},
    "pivotTable": {"Rows", "Columns", "Values"},
}

_CATEGORICAL_ROLE_KINDS = {"Category": "column", "Y": "measure", "Series": "column"}

VISUAL_ROLE_KINDS: dict[str, dict[str, str]] = {
    "card": {"Values": "measure"},
    "multiRowCard": {"Category": "column", "Values": "measure"},
    "clusteredBarChart": dict(_CATEGORICAL_ROLE_KINDS),
    "clusteredColumnChart": dict(_CATEGORICAL_ROLE_KINDS),
    "stackedBarChart": dict(_CATEGORICAL_ROLE_KINDS),
    "stackedColumnChart": dict(_CATEGORICAL_ROLE_KINDS),
    "hundredPercentStackedBarChart": dict(_CATEGORICAL_ROLE_KINDS),
    "hundredPercentStackedColumnChart": dict(_CATEGORICAL_ROLE_KINDS),
    "ribbonChart": dict(_CATEGORICAL_ROLE_KINDS),
    "lineChart": {"Category": "column", "Y": "measure", "Series": "column"},
    "areaChart": {"Category": "column", "Y": "measure", "Series": "column"},
    "stackedAreaChart": {"Category": "column", "Y": "measure", "Series": "column"},
    "hundredPercentStackedAreaChart": {"Category": "column", "Y": "measure", "Series": "column"},
    "donutChart": {"Category": "column", "Y": "measure"},
    "pieChart": {"Category": "column", "Y": "measure"},
    "treemap": {"Category": "column", "Details": "column", "Values": "measure"},
    "funnel": {"Group": "column", "Values": "measure"},
    "tableEx": {"Values": "any"},
    "waterfallChart": {"Category": "column", "Y": "measure"},
    "slicer": {"Values": "column"},
    "gauge": {"Y": "measure", "Goal": "measure"},
    "kpi": {"Indicator": "measure", "TrendLine": "column", "Goal": "measure"},
    "map": {"Category": "column", "Y": "measure"},
    "scatterChart": {"Category": "column", "X": "measure", "Y": "measure", "Size": "measure", "Series": "column"},
    "lineClusteredColumnComboChart": {"Category": "column", "Y": "measure", "Y2": "measure", "Series": "column"},
    "pivotTable": {"Rows": "column", "Columns": "column", "Values": "measure"},
}


class VisualToolError(PowerBIError):
    code = "visual_error"


class PBIToolsNotInstalledError(VisualToolError):
    code = "pbi_tools_not_found"


class ReportLayoutError(VisualToolError):
    code = "report_layout_error"


class PageNotFoundError(VisualToolError):
    code = "report_page_not_found"


class VisualNotFoundError(VisualToolError):
    code = "report_visual_not_found"

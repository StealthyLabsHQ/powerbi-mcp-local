"""Visuals package: shared constants, error types, and design presets.

Lives at the bottom of the package import graph so other submodules can
freely import from it without risk of cycles.
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Any

from pbi_connection import PowerBIError


DEFAULT_PAGE_WIDTH = 1280
DEFAULT_PAGE_HEIGHT = 720
LAYOUT_RELATIVE_PATH = Path("Report") / "Layout"
THEMES_RELATIVE_DIR = Path("Report") / "StaticResources" / "Themes"
DESIGN_THEME_RELATIVE_PATH = (
    Path("Report") / "StaticResources" / "SharedResources" / "BaseThemes" / "CY26SU02.json"
)
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
    "clusteredBarChart": {"Category", "Y", "Series"},
    "clusteredColumnChart": {"Category", "Y", "Series"},
    "lineChart": {"Category", "Y"},
    "donutChart": {"Category", "Y"},
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

VISUAL_ROLE_KINDS: dict[str, dict[str, str]] = {
    "card": {"Values": "measure"},
    "clusteredBarChart": {"Category": "column", "Y": "measure", "Series": "column"},
    "clusteredColumnChart": {"Category": "column", "Y": "measure", "Series": "column"},
    "lineChart": {"Category": "column", "Y": "measure"},
    "donutChart": {"Category": "column", "Y": "measure"},
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

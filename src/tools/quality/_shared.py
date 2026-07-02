"""Shared constants and layout/visual/model helpers for the quality package."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from pbi_connection import PowerBIValidationError, dax_quote_table_name

LAYOUT_RELATIVE_PATH = Path("Report") / "Layout"

# Force UTF-8 I/O on Windows PowerShell 5.1. Default Out-File encoding is
# UTF-16 LE and the host codepage drives stdout — both can mangle paths /
# JSON when piped back through subprocess + json.loads on Python's side.
_PS_UTF8_PRELUDE = (
    "$OutputEncoding = [System.Text.UTF8Encoding]::new($false);"
    "[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false);"
    "[Console]::InputEncoding = [System.Text.UTF8Encoding]::new($false);"
    "$PSDefaultParameterValues['Out-File:Encoding']='utf8';"
    "$PSDefaultParameterValues['*:Encoding']='utf8';\n"
)

MIN_VISUAL_WIDTH = 120
MIN_VISUAL_HEIGHT = 80
MAX_VISUALS_PER_PAGE = 12
DATE_PARSE_FORMATS = (
    "%Y-%m-%d",
    "%Y/%m/%d",
    "%d/%m/%Y",
    "%m/%d/%Y",
    "%d-%m-%Y",
    "%Y%m%d",
)


def _load_layout(extract_folder: str) -> tuple[Path, dict[str, Any]]:
    from . import resolve_local_path

    folder = resolve_local_path(extract_folder, must_exist=True)
    layout_path = folder / LAYOUT_RELATIVE_PATH
    if not layout_path.exists():
        raise PowerBIValidationError("Report/Layout was not found.", details={"extract_folder": str(folder)})
    return folder, json.loads(layout_path.read_bytes().decode("utf-16-le"))


def _visual_config(container: dict[str, Any]) -> dict[str, Any]:
    raw = container.get("config", "{}")
    if isinstance(raw, str):
        try:
            return json.loads(raw)
        except json.JSONDecodeError:
            return {}
    return raw if isinstance(raw, dict) else {}


def _visual_name(container: dict[str, Any]) -> str:
    return str(_visual_config(container).get("name", ""))


def _visual_type(container: dict[str, Any]) -> str:
    cfg = _visual_config(container)
    return str((cfg.get("singleVisual") or {}).get("visualType", ""))


def _visual_has_title(container: dict[str, Any]) -> bool:
    single = _visual_config(container).get("singleVisual") or {}
    objects = single.get("objects") or {}
    title = objects.get("title") or []
    return bool(title)


def _bounds(container: dict[str, Any]) -> tuple[float, float, float, float]:
    return (
        float(container.get("x", 0) or 0),
        float(container.get("y", 0) or 0),
        float(container.get("width", 0) or 0),
        float(container.get("height", 0) or 0),
    )


def _overlap_area(a: dict[str, Any], b: dict[str, Any]) -> float:
    ax, ay, aw, ah = _bounds(a)
    bx, by, bw, bh = _bounds(b)
    x_overlap = max(0.0, min(ax + aw, bx + bw) - max(ax, bx))
    y_overlap = max(0.0, min(ay + ah, by + bh) - max(ay, by))
    return x_overlap * y_overlap


def _model_snapshot(manager: Any, *, include_hidden: bool = False) -> dict[str, Any]:
    from ..model import pbi_model_info_tool

    return pbi_model_info_tool(manager, include_hidden=include_hidden, include_row_counts=False)


def _dax_column(table: str, column: str) -> str:
    return f"{dax_quote_table_name(table)}[{column.replace(']', ']]')}]"


def _row_value(row: dict[str, Any], alias: str) -> Any:
    return row.get(alias, row.get(f"[{alias}]"))

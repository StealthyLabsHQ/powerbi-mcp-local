"""Generic visual dispatcher: ``pbi_add_visual_tool`` + per-type registry."""

from __future__ import annotations

from typing import Any, Callable

from pbi_connection import PowerBIValidationError

from ._base import DEFAULT_VISUAL_SIZES
from ._cards import (
    pbi_add_card_tool,
    pbi_add_gauge_tool,
    pbi_add_labelled_card_tool,
    pbi_add_text_box_tool,
)
from ._charts import (
    pbi_add_bar_chart_tool,
    pbi_add_donut_chart_tool,
    pbi_add_line_chart_tool,
    pbi_add_waterfall_tool,
)
from ._layout import dry_run_layout_writes
from ._structure import (
    pbi_add_map_tool,
    pbi_add_slicer_tool,
    pbi_add_table_visual_tool,
)


_VISUAL_TYPE_DISPATCH: dict[str, Callable[..., dict[str, Any]]] = {}


def pbi_add_visual_tool(
    extract_folder: str,
    page: str,
    visual_type: str,
    x: int,
    y: int,
    width: int | None = None,
    height: int | None = None,
    title: str = "",
    config: dict[str, Any] | None = None,
    *,
    manager: Any | None = None,
    dry_run: bool = False,
) -> dict[str, Any]:
    """Generic visual dispatcher. Keeps the per-type tools as stable API surface.

    visual_type: one of card, bar_chart, line_chart, donut, table, waterfall,
                 slicer, gauge, kpi, scatter_chart, combo_chart, matrix, map,
                 text_box, labelled_card.
    config: per-type keyword arguments.
    dry_run: when True, run all validation and binding logic but skip the
             layout disk write — response carries ``dry_run=True`` and a
             ``write_log``.
    """
    cfg = dict(config or {})
    visual_key = visual_type.strip().casefold()
    size = DEFAULT_VISUAL_SIZES.get(visual_key)
    effective_width = width if width is not None else (size[0] if size else 320)
    effective_height = height if height is not None else (size[1] if size else 240)

    handler = _VISUAL_TYPE_DISPATCH.get(visual_key)
    if handler is None:
        raise PowerBIValidationError(
            f"Unknown visual_type '{visual_type}'. Allowed: {sorted(_VISUAL_TYPE_DISPATCH)}",
            details={"visual_type": visual_type},
        )
    if manager is not None and "__manager__" not in cfg:
        cfg["__manager__"] = manager

    if dry_run:
        with dry_run_layout_writes() as write_log:
            try:
                result = handler(extract_folder, page, x, y, effective_width, effective_height, title, cfg)
            except Exception:
                raise
        result = dict(result or {})
        result["dry_run"] = True
        result["write_log"] = list(write_log)
        result["message"] = "[dry-run] " + str(result.get("message", ""))
        return result
    return handler(extract_folder, page, x, y, effective_width, effective_height, title, cfg)


def _dispatch_card(extract, page, x, y, w, h, title, cfg):
    measure = cfg.get("measure")
    if not measure:
        raise PowerBIValidationError("card visual requires config.measure", details={"visual_type": "card"})
    return pbi_add_card_tool(extract, page, measure, x, y, w, h, title, manager=cfg.get("__manager__"))


def _dispatch_bar(extract, page, x, y, w, h, title, cfg):
    cat = cfg.get("category_column")
    value = cfg.get("value_measure")
    if not cat or not value:
        raise PowerBIValidationError(
            "bar_chart requires config.category_column and config.value_measure",
            details={"visual_type": "bar_chart"},
        )
    return pbi_add_bar_chart_tool(extract, page, cat, value, x, y, w, h, title, cfg.get("legend_column"), manager=cfg.get("__manager__"))


def _dispatch_line(extract, page, x, y, w, h, title, cfg):
    axis = cfg.get("axis_column")
    measures = cfg.get("value_measures") or []
    if not axis or not measures:
        raise PowerBIValidationError(
            "line_chart requires config.axis_column and config.value_measures (list)",
            details={"visual_type": "line_chart"},
        )
    return pbi_add_line_chart_tool(extract, page, axis, measures, x, y, w, h, title, manager=cfg.get("__manager__"))


def _dispatch_donut(extract, page, x, y, w, h, title, cfg):
    cat = cfg.get("category_column")
    value = cfg.get("value_measure")
    if not cat or not value:
        raise PowerBIValidationError(
            "donut requires config.category_column and config.value_measure",
            details={"visual_type": "donut"},
        )
    return pbi_add_donut_chart_tool(extract, page, cat, value, x, y, w, h, title, manager=cfg.get("__manager__"))


def _dispatch_table(extract, page, x, y, w, h, title, cfg):
    columns = cfg.get("columns") or []
    if not columns:
        raise PowerBIValidationError("table requires config.columns (list)", details={"visual_type": "table"})
    return pbi_add_table_visual_tool(extract, page, columns, x, y, w, h, title, manager=cfg.get("__manager__"))


def _dispatch_waterfall(extract, page, x, y, w, h, title, cfg):
    cat = cfg.get("category_column")
    value = cfg.get("value_measure")
    if not cat or not value:
        raise PowerBIValidationError(
            "waterfall requires config.category_column and config.value_measure",
            details={"visual_type": "waterfall"},
        )
    return pbi_add_waterfall_tool(extract, page, cat, value, x, y, w, h, title, manager=cfg.get("__manager__"))


def _dispatch_slicer(extract, page, x, y, w, h, title, cfg):
    column = cfg.get("column")
    if not column:
        raise PowerBIValidationError("slicer requires config.column", details={"visual_type": "slicer"})
    return pbi_add_slicer_tool(extract, page, column, x, y, w, h, cfg.get("slicer_type", "dropdown"), manager=cfg.get("__manager__"))


def _dispatch_gauge(extract, page, x, y, w, h, title, cfg):
    measure = cfg.get("measure")
    if not measure:
        raise PowerBIValidationError("gauge requires config.measure", details={"visual_type": "gauge"})
    return pbi_add_gauge_tool(
        extract,
        page,
        measure,
        x,
        y,
        w,
        h,
        title,
        cfg.get("target_measure"),
        min_value=cfg.get("min_value"),
        max_value=cfg.get("max_value"),
        target_value=cfg.get("target_value"),
        fill_color=cfg.get("fill_color"),
        target_color=cfg.get("target_color"),
        fill_color_measure=cfg.get("fill_color_measure"),
        manager=cfg.get("__manager__"),
    )


def _dispatch_labelled_card(extract, page, x, y, w, h, title, cfg):
    measure = cfg.get("measure")
    label = cfg.get("label") or title
    if not measure or not label:
        raise PowerBIValidationError(
            "labelled_card requires config.measure and config.label (or title)",
            details={"visual_type": "labelled_card"},
        )
    return pbi_add_labelled_card_tool(
        extract,
        page,
        measure,
        str(label),
        x,
        y,
        w,
        h,
        label_height=int(cfg.get("label_height", 28)),
        label_font_size=int(cfg.get("label_font_size", 11)),
        label_bold=bool(cfg.get("label_bold", True)),
        label_color=str(cfg.get("label_color", "#1F2937")),
        manager=cfg.get("__manager__"),
    )


def _dispatch_map(extract, page, x, y, w, h, title, cfg):
    location = cfg.get("location") or cfg.get("category_column") or cfg.get("category")
    measure = cfg.get("measure") or cfg.get("value_measure")
    if not location:
        raise PowerBIValidationError(
            "map requires config.location (Table.Column with the geographic field)",
            details={"visual_type": "map"},
        )
    return pbi_add_map_tool(extract, page, location, measure, x, y, w, h, title, manager=cfg.get("__manager__"))


def _dispatch_text_box(extract, page, x, y, w, h, title, cfg):
    text = cfg.get("text")
    if text is None:
        raise PowerBIValidationError("text_box requires config.text", details={"visual_type": "text_box"})
    return pbi_add_text_box_tool(
        extract,
        page,
        str(text),
        x,
        y,
        w,
        h,
        int(cfg.get("font_size", 16)),
        bool(cfg.get("bold", False)),
        str(cfg.get("color", "#222222")),
    )


_VISUAL_TYPE_DISPATCH.update({
    "card": _dispatch_card,
    "labelled_card": _dispatch_labelled_card,
    "labeled_card": _dispatch_labelled_card,
    "bar_chart": _dispatch_bar,
    "line_chart": _dispatch_line,
    "donut": _dispatch_donut,
    "table": _dispatch_table,
    "waterfall": _dispatch_waterfall,
    "slicer": _dispatch_slicer,
    "gauge": _dispatch_gauge,
    "map": _dispatch_map,
    "text_box": _dispatch_text_box,
    "textbox": _dispatch_text_box,
})

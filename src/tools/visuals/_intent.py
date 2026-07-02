"""Visual Intent Layer: business-intent spec → deterministic visual choice.

Instead of asking an LLM to pick ``pbi_add_bar_chart`` directly, the client
supplies a constrained intent spec (metric, dimension, comparison, trend,
target, breakdown, …). The MCP picks the visual type and the field roles
deterministically, so the same intent always yields the same — renderable —
visual.

Tools
-----
- ``pbi_plan_visual_tool``  — pure planner, no disk access; returns the
  chosen visual type, dispatcher config, and rationale.
- ``pbi_add_visual_from_intent_tool`` — plans then delegates the build to
  ``pbi_add_visual_tool`` (supports ``dry_run``).
"""

from __future__ import annotations

from typing import Any

from pbi_connection import PowerBIValidationError, ok

from ._base import _run

INTENT_ALLOWED_KEYS = {
    "metric",
    "metrics",
    "dimension",
    "time",
    "breakdown",
    "target",
    "comparison",
    "trend",
    "parts_of_whole",
    "correlation",
    "geographic",
    "detail_table",
    "many_categories",
    "filter_control",
}


def _normalize_intent(intent: dict[str, Any]) -> dict[str, Any]:
    if not isinstance(intent, dict):
        raise PowerBIValidationError("intent must be an object.", details={"intent": repr(intent)})
    unknown = sorted(set(intent) - INTENT_ALLOWED_KEYS)
    if unknown:
        raise PowerBIValidationError(
            f"Unknown intent key(s): {unknown}. Allowed: {sorted(INTENT_ALLOWED_KEYS)}",
            details={"unknown_keys": unknown},
        )
    metrics: list[str] = []
    if intent.get("metric"):
        metrics.append(str(intent["metric"]))
    for item in intent.get("metrics") or []:
        ref = str(item)
        if ref not in metrics:
            metrics.append(ref)
    normalized = {
        "metrics": metrics,
        "dimension": str(intent["dimension"]) if intent.get("dimension") else None,
        "time": str(intent["time"]) if intent.get("time") else None,
        "breakdown": str(intent["breakdown"]) if intent.get("breakdown") else None,
        "target": intent.get("target"),
        "comparison": bool(intent.get("comparison")),
        "trend": bool(intent.get("trend")) or bool(intent.get("time")),
        "parts_of_whole": bool(intent.get("parts_of_whole")),
        "correlation": bool(intent.get("correlation")),
        "geographic": bool(intent.get("geographic")),
        "detail_table": bool(intent.get("detail_table")),
        "many_categories": bool(intent.get("many_categories")),
        "filter_control": bool(intent.get("filter_control")),
    }
    for key in ("metrics",):
        for ref in normalized[key]:
            _require_field_ref(ref, "metric")
    for key in ("dimension", "time", "breakdown"):
        if normalized[key]:
            _require_field_ref(normalized[key], key)
    if isinstance(normalized["target"], str):
        _require_field_ref(normalized["target"], "target")
    return normalized


def _require_field_ref(ref: str, label: str) -> None:
    if "." not in ref.strip():
        raise PowerBIValidationError(
            f"intent.{label} must be a 'Table.Field' reference, got '{ref}'.",
            details={label: ref},
        )


def _plan(intent: dict[str, Any]) -> dict[str, Any]:
    """Deterministic intent → (visual_type, config, rationale). Rules are
    evaluated in priority order; the first match wins.
    """
    spec = _normalize_intent(intent)
    metrics = spec["metrics"]
    dimension = spec["dimension"]
    time_col = spec["time"]
    breakdown = spec["breakdown"]
    target = spec["target"]

    def plan(visual_type: str, config: dict[str, Any], rationale: str) -> dict[str, Any]:
        return {
            "visual_type": visual_type,
            "config": {key: value for key, value in config.items() if value not in (None, [], "")},
            "rationale": rationale,
            "intent": spec,
        }

    if spec["filter_control"]:
        column = dimension or time_col
        if not column:
            raise PowerBIValidationError("filter_control intent requires a dimension (or time) column.")
        return plan("slicer", {"column": column}, "filter_control → slicer on the dimension")

    if spec["detail_table"]:
        columns = [c for c in (dimension, breakdown, time_col) if c] + metrics
        if not columns:
            raise PowerBIValidationError("detail_table intent requires at least one metric or dimension.")
        return plan("table", {"columns": columns}, "detail_table → table with dimensions then metrics")

    if spec["geographic"]:
        if not dimension:
            raise PowerBIValidationError("geographic intent requires a dimension (the location column).")
        return plan(
            "map",
            {"location": dimension, "measure": metrics[0] if metrics else None},
            "geographic dimension → map (location + optional size measure)",
        )

    if spec["correlation"]:
        if len(metrics) < 2 or not dimension:
            raise PowerBIValidationError(
                "correlation intent requires two metrics and a dimension.",
                details={"metrics": metrics, "dimension": dimension},
            )
        return plan(
            "scatter_chart",
            {
                "category_column": dimension,
                "x_measure": metrics[0],
                "y_measure": metrics[1],
                "size_measure": metrics[2] if len(metrics) > 2 else None,
                "legend_column": breakdown,
            },
            "correlation of two metrics → scatter chart (X=first metric, Y=second)",
        )

    if spec["trend"]:
        axis = time_col or dimension
        if not axis or not metrics:
            raise PowerBIValidationError(
                "trend intent requires a time (or dimension) column and at least one metric.",
                details={"time": time_col, "dimension": dimension, "metrics": metrics},
            )
        if target and len(metrics) == 1 and isinstance(target, str):
            return plan(
                "kpi",
                {"indicator_measure": metrics[0], "trend_column": axis, "goal_measure": target},
                "single metric + target + trend → KPI (indicator, trend axis, goal)",
            )
        if spec["parts_of_whole"] and breakdown:
            return plan(
                "stacked_area_chart",
                {"axis_column": axis, "value_measures": metrics, "legend_column": breakdown},
                "trend + composition by breakdown → stacked area chart",
            )
        return plan(
            "line_chart",
            {"axis_column": axis, "value_measures": metrics, "legend_column": breakdown},
            "trend over time → line chart (one line per metric / breakdown)",
        )

    if target is not None:
        if not metrics:
            raise PowerBIValidationError("target intent requires a metric.")
        config: dict[str, Any] = {"measure": metrics[0]}
        if isinstance(target, str):
            config["target_measure"] = target
        else:
            config["target_value"] = float(target)
        return plan("gauge", config, "metric vs target without trend → gauge")

    if spec["parts_of_whole"]:
        if not dimension or not metrics:
            raise PowerBIValidationError("parts_of_whole intent requires a dimension and a metric.")
        if spec["many_categories"]:
            return plan(
                "treemap",
                {"category_column": dimension, "value_measure": metrics[0]},
                "composition with many categories → treemap (donut would be unreadable)",
            )
        return plan(
            "donut",
            {"category_column": dimension, "value_measure": metrics[0]},
            "composition with few categories → donut",
        )

    if dimension and metrics:
        if len(metrics) > 1:
            return plan(
                "combo_chart",
                {
                    "category_column": dimension,
                    "bar_measures": [metrics[0]],
                    "line_measures": metrics[1:],
                    "legend_column": breakdown,
                },
                "multiple metrics across one dimension → combo chart (first metric as bars)",
            )
        if breakdown:
            return plan(
                "stacked_column_chart",
                {"category_column": dimension, "value_measure": metrics[0], "legend_column": breakdown},
                "comparison with breakdown → stacked column chart",
            )
        return plan(
            "bar_chart",
            {"category_column": dimension, "value_measure": metrics[0]},
            "comparison across a dimension → bar chart",
        )

    if metrics:
        if len(metrics) > 1:
            return plan("multi_row_card", {"measures": metrics}, "several headline metrics → multi-row card")
        return plan("card", {"measure": metrics[0]}, "single headline metric → card")

    if dimension:
        return plan("slicer", {"column": dimension}, "dimension without metric → slicer")

    raise PowerBIValidationError(
        "intent must contain at least a metric or a dimension.",
        details={"intent": spec},
    )


def pbi_plan_visual_tool(intent: dict[str, Any]) -> dict[str, Any]:
    """Plan a visual from a business-intent spec without touching the report.

    intent keys (all optional unless a rule needs them):
      metric / metrics  — 'Table.Measure' reference(s)
      dimension         — 'Table.Column' categorical axis
      time              — 'Table.Column' time axis (implies trend)
      breakdown         — 'Table.Column' secondary split (legend/series)
      target            — 'Table.Measure' reference or numeric goal
      comparison, trend, parts_of_whole, correlation, geographic,
      detail_table, many_categories, filter_control — boolean hints

    Returns the chosen ``visual_type``, the matching ``pbi_add_visual``
    ``config``, and a ``rationale`` explaining the rule that fired.
    """

    def _impl() -> dict[str, Any]:
        planned = _plan(intent)
        return ok(
            f"Planned visual_type '{planned['visual_type']}': {planned['rationale']}",
            **planned,
        )

    return _run(_impl)


def pbi_add_visual_from_intent_tool(
    extract_folder: str,
    page: str,
    intent: dict[str, Any],
    x: int,
    y: int,
    width: int | None = None,
    height: int | None = None,
    title: str = "",
    *,
    manager: Any | None = None,
    dry_run: bool = False,
) -> dict[str, Any]:
    """Plan a visual from a business-intent spec, then build it.

    Same intent schema as ``pbi_plan_visual_tool``. The chosen visual is
    created through the generic dispatcher, so all field validation and
    binding logic applies. ``dry_run=True`` plans and validates without
    writing the layout.
    """

    def _impl() -> dict[str, Any]:
        from ._dispatcher import pbi_add_visual_tool

        planned = _plan(intent)
        result = pbi_add_visual_tool(
            extract_folder,
            page,
            planned["visual_type"],
            x,
            y,
            width,
            height,
            title,
            planned["config"],
            manager=manager,
            dry_run=dry_run,
        )
        result = dict(result or {})
        result["decision"] = {
            "visual_type": planned["visual_type"],
            "rationale": planned["rationale"],
            "config": planned["config"],
            "intent": planned["intent"],
        }
        return result

    return _run(_impl)

"""Phase 1: Visual Intent Layer + repairable-error registry.

The intent layer must map every business-intent combination to a single
deterministic visual choice, and the repair loop must classify every known
failure mode into a stable, LLM-actionable error code.
"""

from __future__ import annotations

import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

from tools.visuals._dispatcher import _VISUAL_TYPE_DISPATCH
from tools.visuals._errors import (
    REPAIRABLE_ERRORS,
    _classify,
    _detect_double_display_units,
    _detect_gauge_target_issues,
    _parse_pbi_literal_number,
    pbi_list_repairable_errors_tool,
)
from tools.visuals._intent import _plan, pbi_plan_visual_tool

METRIC = "Sales.Total Sales"
METRIC2 = "Sales.Total Cost"
DIM = "Product.Category"
TIME = "Calendar.Month"
BREAK = "Region.Region"


class IntentPlanningTests(unittest.TestCase):
    def _type(self, intent):
        return _plan(intent)["visual_type"]

    def test_single_metric_is_card(self):
        self.assertEqual(self._type({"metric": METRIC}), "card")

    def test_multiple_metrics_is_multi_row_card(self):
        self.assertEqual(self._type({"metrics": [METRIC, METRIC2]}), "multi_row_card")

    def test_metric_with_numeric_target_is_gauge(self):
        plan = _plan({"metric": METRIC, "target": 1000})
        self.assertEqual(plan["visual_type"], "gauge")
        self.assertEqual(plan["config"]["target_value"], 1000.0)

    def test_metric_with_measure_target_is_gauge(self):
        plan = _plan({"metric": METRIC, "target": "Sales.Target"})
        self.assertEqual(plan["visual_type"], "gauge")
        self.assertEqual(plan["config"]["target_measure"], "Sales.Target")

    def test_metric_target_trend_is_kpi(self):
        plan = _plan({"metric": METRIC, "target": "Sales.Target", "time": TIME})
        self.assertEqual(plan["visual_type"], "kpi")
        self.assertEqual(plan["config"]["trend_column"], TIME)

    def test_time_implies_line_chart(self):
        plan = _plan({"metric": METRIC, "time": TIME})
        self.assertEqual(plan["visual_type"], "line_chart")
        self.assertEqual(plan["config"]["axis_column"], TIME)

    def test_trend_composition_is_stacked_area(self):
        plan = _plan({"metric": METRIC, "time": TIME, "breakdown": BREAK, "parts_of_whole": True})
        self.assertEqual(plan["visual_type"], "stacked_area_chart")

    def test_parts_of_whole_few_categories_is_donut(self):
        self.assertEqual(self._type({"metric": METRIC, "dimension": DIM, "parts_of_whole": True}), "donut")

    def test_parts_of_whole_many_categories_is_treemap(self):
        self.assertEqual(
            self._type({"metric": METRIC, "dimension": DIM, "parts_of_whole": True, "many_categories": True}),
            "treemap",
        )

    def test_comparison_is_bar_chart(self):
        self.assertEqual(self._type({"metric": METRIC, "dimension": DIM, "comparison": True}), "bar_chart")

    def test_comparison_with_breakdown_is_stacked_column(self):
        self.assertEqual(self._type({"metric": METRIC, "dimension": DIM, "breakdown": BREAK}), "stacked_column_chart")

    def test_two_metrics_one_dimension_is_combo(self):
        plan = _plan({"metrics": [METRIC, METRIC2], "dimension": DIM})
        self.assertEqual(plan["visual_type"], "combo_chart")
        self.assertEqual(plan["config"]["bar_measures"], [METRIC])
        self.assertEqual(plan["config"]["line_measures"], [METRIC2])

    def test_correlation_is_scatter(self):
        plan = _plan({"metrics": [METRIC, METRIC2], "dimension": DIM, "correlation": True})
        self.assertEqual(plan["visual_type"], "scatter_chart")
        self.assertEqual(plan["config"]["x_measure"], METRIC)
        self.assertEqual(plan["config"]["y_measure"], METRIC2)

    def test_geographic_is_map(self):
        plan = _plan({"metric": METRIC, "dimension": "Geo.City", "geographic": True})
        self.assertEqual(plan["visual_type"], "map")
        self.assertEqual(plan["config"]["location"], "Geo.City")

    def test_detail_table_orders_dimensions_before_metrics(self):
        plan = _plan({"metrics": [METRIC], "dimension": DIM, "detail_table": True})
        self.assertEqual(plan["visual_type"], "table")
        self.assertEqual(plan["config"]["columns"], [DIM, METRIC])

    def test_dimension_only_is_slicer(self):
        self.assertEqual(self._type({"dimension": DIM}), "slicer")

    def test_filter_control_wins_over_everything(self):
        self.assertEqual(self._type({"metric": METRIC, "dimension": DIM, "filter_control": True}), "slicer")

    def test_unknown_key_rejected(self):
        result = pbi_plan_visual_tool({"metric": METRIC, "chart": "bar"})
        self.assertFalse(result.get("ok", True))
        self.assertIn("chart", str(result))

    def test_empty_intent_rejected(self):
        result = pbi_plan_visual_tool({})
        self.assertFalse(result.get("ok", True))

    def test_bare_field_ref_rejected(self):
        result = pbi_plan_visual_tool({"metric": "Total Sales"})
        self.assertFalse(result.get("ok", True))

    def test_every_planned_type_is_dispatchable(self):
        intents = [
            {"metric": METRIC},
            {"metrics": [METRIC, METRIC2]},
            {"metric": METRIC, "target": 100},
            {"metric": METRIC, "target": "Sales.Target", "time": TIME},
            {"metric": METRIC, "time": TIME},
            {"metric": METRIC, "time": TIME, "breakdown": BREAK, "parts_of_whole": True},
            {"metric": METRIC, "dimension": DIM, "parts_of_whole": True},
            {"metric": METRIC, "dimension": DIM, "parts_of_whole": True, "many_categories": True},
            {"metric": METRIC, "dimension": DIM},
            {"metric": METRIC, "dimension": DIM, "breakdown": BREAK},
            {"metrics": [METRIC, METRIC2], "dimension": DIM},
            {"metrics": [METRIC, METRIC2], "dimension": DIM, "correlation": True},
            {"metric": METRIC, "dimension": DIM, "geographic": True},
            {"metric": METRIC, "dimension": DIM, "detail_table": True},
            {"dimension": DIM},
        ]
        for intent in intents:
            visual_type = _plan(intent)["visual_type"]
            self.assertIn(visual_type, _VISUAL_TYPE_DISPATCH, f"intent {intent} planned undispatchable {visual_type}")


class DispatcherCompletenessTests(unittest.TestCase):
    def test_kpi_and_matrix_are_registered(self):
        self.assertIn("kpi", _VISUAL_TYPE_DISPATCH)
        self.assertIn("matrix", _VISUAL_TYPE_DISPATCH)


def _layout_with_visual(visual_type: str, objects: dict, projections: dict | None = None) -> dict:
    return {
        "sections": [
            {
                "name": "ReportSection1",
                "displayName": "Page 1",
                "visualContainers": [
                    {
                        "config": {
                            "name": "v1",
                            "singleVisual": {
                                "visualType": visual_type,
                                "projections": projections or {},
                                "objects": objects,
                            },
                        }
                    }
                ],
            }
        ]
    }


def _decimal(value: float) -> dict:
    return {"expr": {"Literal": {"Value": f"{value}D"}}}


class RepairableErrorTests(unittest.TestCase):
    def test_registry_covers_all_plan_errors(self):
        for code in (
            "measure_not_found",
            "column_not_found",
            "unexpected_projection_role",
            "double_display_units",
            "gauge_target_invalid",
            "empty_visual",
        ):
            self.assertIn(code, REPAIRABLE_ERRORS)
            self.assertTrue(REPAIRABLE_ERRORS[code]["llm_action"])

    def test_list_tool_returns_registry(self):
        result = pbi_list_repairable_errors_tool()
        self.assertTrue(result["ok"])
        self.assertIn("measure_not_found", result["errors"])

    def test_classify_known_and_unknown(self):
        known = _classify({"issue": "measure_not_found", "page": "P1", "visual_id": "v1"})
        self.assertEqual(known["code"], "measure_not_found")
        self.assertFalse(known["auto_repair"])
        unknown = _classify({"issue": "weird_new_thing"})
        self.assertEqual(unknown["severity"], "error")

    def test_parse_pbi_literal_number(self):
        self.assertEqual(_parse_pbi_literal_number(_decimal(100.0)), 100.0)
        self.assertEqual(_parse_pbi_literal_number({"expr": {"Literal": {"Value": "5L"}}}), 5.0)
        self.assertIsNone(_parse_pbi_literal_number({"expr": {"Literal": {"Value": "'text'"}}}))
        self.assertIsNone(_parse_pbi_literal_number(None))

    def test_gauge_target_outside_range_flagged(self):
        layout = _layout_with_visual(
            "gauge",
            {"axis": [{"properties": {"min": _decimal(0), "max": _decimal(100), "target": _decimal(150)}}]},
        )
        issues = _detect_gauge_target_issues(layout, None)
        self.assertEqual(len(issues), 1)
        self.assertEqual(issues[0]["issue"], "gauge_target_invalid")

    def test_gauge_min_ge_max_flagged(self):
        layout = _layout_with_visual(
            "gauge",
            {"axis": [{"properties": {"min": _decimal(100), "max": _decimal(100)}}]},
        )
        self.assertEqual(len(_detect_gauge_target_issues(layout, None)), 1)

    def test_valid_gauge_passes(self):
        layout = _layout_with_visual(
            "gauge",
            {"axis": [{"properties": {"min": _decimal(0), "max": _decimal(200), "target": _decimal(150)}}]},
        )
        self.assertEqual(_detect_gauge_target_issues(layout, None), [])

    def test_double_display_units_flagged(self):
        layout = _layout_with_visual(
            "card",
            {"labels": [{"properties": {"labelDisplayUnits": _decimal(1000)}}]},
            projections={"Values": [{"queryRef": METRIC}]},
        )
        format_index = {METRIC.casefold(): '#,0,.0"K"'}
        issues = _detect_double_display_units(layout, None, format_index)
        self.assertEqual(len(issues), 1)
        self.assertEqual(issues[0]["issue"], "double_display_units")

    def test_plain_format_with_display_units_passes(self):
        layout = _layout_with_visual(
            "card",
            {"labels": [{"properties": {"labelDisplayUnits": _decimal(1000)}}]},
            projections={"Values": [{"queryRef": METRIC}]},
        )
        self.assertEqual(_detect_double_display_units(layout, None, {METRIC.casefold(): "#,0.00"}), [])

    def test_auto_display_units_passes(self):
        layout = _layout_with_visual(
            "card",
            {"labels": [{"properties": {"labelDisplayUnits": _decimal(0)}}]},
            projections={"Values": [{"queryRef": METRIC}]},
        )
        self.assertEqual(_detect_double_display_units(layout, None, {METRIC.casefold(): '#,0,.0"K"'}), [])


if __name__ == "__main__":
    unittest.main()

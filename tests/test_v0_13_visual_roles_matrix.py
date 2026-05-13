"""v0.13: Visual roles + role-kinds completeness matrix.

The renderer-compat work in v0.12.7 / v0.12.8 made it clear that the
canonical projection roles per visual type are load-bearing. Pin the
whole matrix so a future chart addition can't silently regress roles
or kinds (which the data-shape pass would then drop).
"""

from __future__ import annotations

import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

from tools.visuals._base import VISUAL_FIELD_ROLES, VISUAL_ROLE_KINDS


_EXPECTED_ROLES = {
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


class VisualRolesMatrixTests(unittest.TestCase):
    def test_every_known_visual_pins_its_role_set(self) -> None:
        for visual_type, expected in _EXPECTED_ROLES.items():
            with self.subTest(visual=visual_type):
                self.assertIn(visual_type, VISUAL_FIELD_ROLES)
                self.assertEqual(VISUAL_FIELD_ROLES[visual_type], expected)

    def test_role_kinds_cover_each_role(self) -> None:
        # Every role in VISUAL_FIELD_ROLES must have a matching kind in
        # VISUAL_ROLE_KINDS (column / measure / any). Without it, the
        # binding validator can't decide whether a measure or column is
        # expected for a given slot.
        for visual_type, roles in VISUAL_FIELD_ROLES.items():
            with self.subTest(visual=visual_type):
                self.assertIn(visual_type, VISUAL_ROLE_KINDS)
                kinds = VISUAL_ROLE_KINDS[visual_type]
                missing = roles - set(kinds.keys())
                self.assertEqual(missing, set(), f"missing kinds for {visual_type}: {missing}")
                for role, kind in kinds.items():
                    self.assertIn(kind, {"column", "measure", "any"})

    def test_no_cartesian_y_in_categorical_only_visuals(self) -> None:
        # Visuals that the data-shape pass treats as categorical (treemap,
        # donut, pie, funnel) must not declare the cartesian Y role on its
        # own — that combination is exactly what dropped the v0.12.7 /
        # v0.12.8 treemap into an empty render.
        for visual_type in ("treemap", "funnel"):
            with self.subTest(visual=visual_type):
                self.assertNotIn(
                    "Y",
                    VISUAL_FIELD_ROLES[visual_type],
                    f"{visual_type} should not use cartesian Y",
                )


class DispatcherCoverageTests(unittest.TestCase):
    def test_dispatcher_lists_every_chart_family(self) -> None:
        from tools.visuals._dispatcher import _VISUAL_TYPE_DISPATCH

        expected_keys = {
            "card",
            "bar_chart",
            "clustered_column_chart",
            "stacked_bar_chart",
            "stacked_column_chart",
            "hundred_percent_stacked_bar_chart",
            "hundred_percent_stacked_column_chart",
            "ribbon_chart",
            "line_chart",
            "area_chart",
            "stacked_area_chart",
            "hundred_percent_stacked_area_chart",
            "donut",
            "pie_chart",
            "treemap",
            "funnel",
            "multi_row_card",
            "scatter_chart",
            "combo_chart",
        }
        missing = expected_keys - set(_VISUAL_TYPE_DISPATCH.keys())
        self.assertEqual(missing, set(), f"dispatcher missing entries: {missing}")


if __name__ == "__main__":
    unittest.main(verbosity=2)

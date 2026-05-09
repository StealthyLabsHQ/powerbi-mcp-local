"""Smoke tests for the v0.12.6 chart pack.

The new tools (`pie`, `stacked` / `100%` bar+column variants, area family,
ribbon, treemap, funnel, multiRowCard) all reuse the same shared
`_append_visual` plumbing as the existing chart builders. We don't try
to validate Power BI Desktop's eventual rendering — we only assert that
the layout disk-write produces a well-formed visual container with the
right ``visualType`` and projection roles.
"""

from __future__ import annotations

import json
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

from tools.visuals._charts import (
    pbi_add_area_chart_tool,
    pbi_add_clustered_column_chart_tool,
    pbi_add_funnel_tool,
    pbi_add_hundred_percent_stacked_area_chart_tool,
    pbi_add_hundred_percent_stacked_bar_chart_tool,
    pbi_add_hundred_percent_stacked_column_chart_tool,
    pbi_add_multi_row_card_tool,
    pbi_add_pie_chart_tool,
    pbi_add_ribbon_chart_tool,
    pbi_add_stacked_area_chart_tool,
    pbi_add_stacked_bar_chart_tool,
    pbi_add_stacked_column_chart_tool,
    pbi_add_treemap_tool,
)


def _build_empty_layout(folder: Path) -> None:
    layout = {
        "id": 0,
        "resourcePackages": [],
        "sections": [
            {
                "id": 1,
                "name": "ReportSection1",
                "displayName": "Page 1",
                "visualContainers": [],
            }
        ],
    }
    layout_path = folder / "Report" / "Layout"
    layout_path.parent.mkdir(parents=True, exist_ok=True)
    layout_path.write_bytes(json.dumps(layout).encode("utf-16-le"))


def _last_visual_type(folder: Path) -> str:
    layout = json.loads((folder / "Report" / "Layout").read_bytes().decode("utf-16-le"))
    container = layout["sections"][0]["visualContainers"][-1]
    config = json.loads(container["config"])
    return config["singleVisual"]["visualType"]


class ChartPackSmokeTests(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp = tempfile.TemporaryDirectory()
        self.root = Path(self.tmp.name)
        from security import SECURITY, configure_allowed_dirs

        self.previous_allowed = [str(item) for item in SECURITY.allowed_base_dirs()]
        configure_allowed_dirs([str(self.root)])
        SECURITY.policy(reload=True, cwd=self.root)
        _build_empty_layout(self.root)

    def tearDown(self) -> None:
        from security import SECURITY, configure_allowed_dirs

        configure_allowed_dirs(self.previous_allowed)
        SECURITY.policy(reload=True, cwd=Path.cwd())
        self.tmp.cleanup()

    def _common_kw(self) -> dict[str, object]:
        return {
            "extract_folder": str(self.root),
            "page": "Page 1",
            "x": 0,
            "y": 0,
            "manager": None,
        }

    def test_categorical_variants(self) -> None:
        cases = [
            (pbi_add_pie_chart_tool, "pieChart"),
            (pbi_add_stacked_bar_chart_tool, "stackedBarChart"),
            (pbi_add_stacked_column_chart_tool, "stackedColumnChart"),
            (pbi_add_clustered_column_chart_tool, "clusteredColumnChart"),
            (pbi_add_hundred_percent_stacked_bar_chart_tool, "hundredPercentStackedBarChart"),
            (pbi_add_hundred_percent_stacked_column_chart_tool, "hundredPercentStackedColumnChart"),
            (pbi_add_ribbon_chart_tool, "ribbonChart"),
            (pbi_add_treemap_tool, "treemap"),
        ]
        for builder, expected_type in cases:
            with self.subTest(visual_type=expected_type):
                result = builder(
                    category_column="Sales.Region",
                    value_measure="Total Sales",
                    **self._common_kw(),
                )
                self.assertTrue(result["ok"], result)
                self.assertEqual(_last_visual_type(self.root), expected_type)

    def test_axis_variants_accept_multiple_measures(self) -> None:
        for builder, expected_type in (
            (pbi_add_area_chart_tool, "areaChart"),
            (pbi_add_stacked_area_chart_tool, "stackedAreaChart"),
            (pbi_add_hundred_percent_stacked_area_chart_tool, "hundredPercentStackedAreaChart"),
        ):
            with self.subTest(visual_type=expected_type):
                result = builder(
                    axis_column="Date.Date",
                    value_measures=["Revenue", "Cost"],
                    **self._common_kw(),
                )
                self.assertTrue(result["ok"], result)
                self.assertEqual(_last_visual_type(self.root), expected_type)

    def test_funnel_uses_group_values_roles(self) -> None:
        result = pbi_add_funnel_tool(
            group_column="Stage.Name",
            value_measure="Pipeline Value",
            **self._common_kw(),
        )
        self.assertTrue(result["ok"], result)
        layout = json.loads((self.root / "Report" / "Layout").read_bytes().decode("utf-16-le"))
        config = json.loads(layout["sections"][0]["visualContainers"][-1]["config"])
        sv = config["singleVisual"]
        self.assertEqual(sv["visualType"], "funnel")
        self.assertIn("Group", sv["projections"])
        self.assertIn("Values", sv["projections"])

    def test_multi_row_card_with_and_without_category(self) -> None:
        result = pbi_add_multi_row_card_tool(
            measures=["Revenue", "Margin"],
            **self._common_kw(),
        )
        self.assertTrue(result["ok"], result)
        self.assertEqual(_last_visual_type(self.root), "multiRowCard")
        # With category — second visual on the same page.
        result2 = pbi_add_multi_row_card_tool(
            measures=["Revenue"],
            category_column="Region.Name",
            **self._common_kw(),
        )
        self.assertTrue(result2["ok"], result2)
        self.assertEqual(_last_visual_type(self.root), "multiRowCard")

    def test_multi_row_card_rejects_empty_measures(self) -> None:
        # Direct calls to ``*_tool`` functions raise validation errors;
        # the MCP wrapper layer (`_run`) translates those into the
        # error envelope. We're testing the underlying tool here.
        from pbi_connection import PowerBIValidationError

        with self.assertRaises(PowerBIValidationError):
            pbi_add_multi_row_card_tool(
                measures=[],
                **self._common_kw(),
            )


if __name__ == "__main__":
    unittest.main(verbosity=2)

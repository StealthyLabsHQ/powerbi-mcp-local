"""Regressions for the v0.12.7 renderer-compat fixes.

Power BI Desktop's data-shape pass silently drops:
- ``From`` entries that lack ``Type: 0`` (the standard "Table" entity-source kind)
- projection items that lack ``"active": True``

The visuals open with an empty data area despite the projections + select
entries being correct. These regressions assert that every emitted
visual carries both fields so the next "empty visual" report points at
a different root cause.
"""

from __future__ import annotations

import json
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

from tools.visuals._bindings import _build_prototype_query
from tools.visuals._charts import (
    pbi_add_bar_chart_tool,
    pbi_add_line_chart_tool,
    pbi_add_pie_chart_tool,
    pbi_add_treemap_tool,
)
from tools.visuals._refs import _projection


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


def _last_config(folder: Path) -> dict:
    layout = json.loads((folder / "Report" / "Layout").read_bytes().decode("utf-16-le"))
    return json.loads(layout["sections"][0]["visualContainers"][-1]["config"])


class FromTypeFieldTests(unittest.TestCase):
    def test_build_prototype_query_emits_type_zero(self) -> None:
        proto = _build_prototype_query(["Sales.Region", "Total Sales"], {"Total Sales": "Sales"})
        self.assertEqual(proto["Version"], 2)
        self.assertGreater(len(proto["From"]), 0)
        for entry in proto["From"]:
            self.assertEqual(entry["Type"], 0, f"From entry must carry Type: 0 — {entry!r}")
            self.assertIn("Name", entry)
            self.assertIn("Entity", entry)


class ProjectionHelperTests(unittest.TestCase):
    def test_projection_helper_emits_active_true(self) -> None:
        item = _projection("Total Sales")
        self.assertEqual(item, {"queryRef": "Total Sales", "active": True})

    def test_projection_helper_can_disable_active(self) -> None:
        item = _projection("Total Sales", active=False)
        self.assertEqual(item, {"queryRef": "Total Sales"})


class EndToEndRendererCompatTests(unittest.TestCase):
    """Build a real visual through the public tool and assert the on-disk
    layout has every field the renderer needs."""

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

    def _common(self) -> dict:
        return {
            "extract_folder": str(self.root),
            "page": "Page 1",
            "x": 0,
            "y": 0,
            "manager": None,
        }

    def _assert_renderer_compat(self) -> None:
        config = _last_config(self.root)
        sv = config["singleVisual"]
        self.assertIn("prototypeQuery", sv)
        proto = sv["prototypeQuery"]
        self.assertEqual(proto["Version"], 2)
        for entry in proto["From"]:
            self.assertEqual(entry["Type"], 0, f"missing Type: 0 → {entry}")
        for role, items in (sv.get("projections") or {}).items():
            for item in items:
                self.assertTrue(
                    item.get("active"),
                    f"role {role!r} has projection item without active: True → {item}",
                )

    def test_treemap_emits_renderer_compat_layout(self) -> None:
        result = pbi_add_treemap_tool(
            category_column="Sales.Region",
            value_measure="Total Sales",
            **self._common(),
        )
        self.assertTrue(result["ok"], result)
        self._assert_renderer_compat()

    def test_pie_chart_emits_renderer_compat_layout(self) -> None:
        result = pbi_add_pie_chart_tool(
            category_column="Sales.Region",
            value_measure="Total Sales",
            **self._common(),
        )
        self.assertTrue(result["ok"], result)
        self._assert_renderer_compat()

    def test_bar_chart_with_legend_emits_renderer_compat_layout(self) -> None:
        result = pbi_add_bar_chart_tool(
            category_column="Sales.Region",
            value_measure="Total Sales",
            legend_column="Sales.Channel",
            **self._common(),
        )
        self.assertTrue(result["ok"], result)
        self._assert_renderer_compat()

    def test_line_chart_with_multiple_measures(self) -> None:
        result = pbi_add_line_chart_tool(
            axis_column="Date.Date",
            value_measures=["Revenue", "Cost"],
            **self._common(),
        )
        self.assertTrue(result["ok"], result)
        self._assert_renderer_compat()


class DispatcherCoverageTests(unittest.TestCase):
    def test_dispatcher_covers_v0126_chart_pack(self) -> None:
        from tools.visuals._dispatcher import _VISUAL_TYPE_DISPATCH

        for visual_type in (
            "pie_chart",
            "treemap",
            "stacked_bar_chart",
            "stacked_column_chart",
            "clustered_column_chart",
            "hundred_percent_stacked_bar_chart",
            "hundred_percent_stacked_column_chart",
            "ribbon_chart",
            "area_chart",
            "stacked_area_chart",
            "hundred_percent_stacked_area_chart",
            "funnel",
            "multi_row_card",
            "scatter_chart",
            "combo_chart",
        ):
            with self.subTest(visual_type=visual_type):
                self.assertIn(visual_type, _VISUAL_TYPE_DISPATCH)


if __name__ == "__main__":
    unittest.main(verbosity=2)

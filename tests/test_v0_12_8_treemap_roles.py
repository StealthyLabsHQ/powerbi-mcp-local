"""Regressions for the v0.12.8 treemap role-name fix.

Power BI Desktop's treemap visual exposes three projection roles:
``Category``, ``Details``, and ``Values``. Earlier builds of
``pbi_add_treemap`` emitted ``Y`` (a cartesian role), which the
data-shape pass dropped silently — the visual rendered as an empty
white rectangle. Pin the canonical roles so a future refactor can't
regress.
"""

from __future__ import annotations

import json
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

from tools.visuals._base import VISUAL_FIELD_ROLES, VISUAL_ROLE_KINDS
from tools.visuals._charts import pbi_add_treemap_tool


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


class TreemapRoleTests(unittest.TestCase):
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

    def test_visual_field_roles_treemap_uses_canonical_names(self) -> None:
        self.assertEqual(VISUAL_FIELD_ROLES["treemap"], {"Category", "Details", "Values"})
        self.assertNotIn("Y", VISUAL_FIELD_ROLES["treemap"])

    def test_visual_role_kinds_treemap_categorises_each_role(self) -> None:
        self.assertEqual(
            VISUAL_ROLE_KINDS["treemap"],
            {"Category": "column", "Details": "column", "Values": "measure"},
        )

    def test_treemap_emits_values_role_not_y(self) -> None:
        result = pbi_add_treemap_tool(
            extract_folder=str(self.root),
            page="Page 1",
            category_column="Sales.Region",
            value_measure="Total Sales",
            x=0,
            y=0,
        )
        self.assertTrue(result["ok"], result)
        config = _last_config(self.root)
        projections = config["singleVisual"]["projections"]
        self.assertIn("Category", projections)
        self.assertIn("Values", projections)
        self.assertNotIn("Y", projections, "treemap must not emit the cartesian Y role")
        # Sanity: each projection item still carries the renderer-compat flag.
        for items in projections.values():
            for item in items:
                self.assertTrue(item.get("active"))

    def test_treemap_with_details_column(self) -> None:
        result = pbi_add_treemap_tool(
            extract_folder=str(self.root),
            page="Page 1",
            category_column="Sales.Region",
            value_measure="Total Sales",
            details_column="Sales.SubRegion",
            x=0,
            y=0,
        )
        self.assertTrue(result["ok"], result)
        projections = _last_config(self.root)["singleVisual"]["projections"]
        self.assertIn("Details", projections)
        self.assertEqual(projections["Details"][0]["queryRef"], "SubRegion")


if __name__ == "__main__":
    unittest.main(verbosity=2)

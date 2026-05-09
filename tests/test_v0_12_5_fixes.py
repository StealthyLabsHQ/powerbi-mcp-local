"""Regressions for the v0.12.5 fixes.

Coverage:
- Bug 1: ``register_tool`` injects ``manager`` as a keyword regardless of
  where it appears in the underlying tool's signature.
- Bug 3: ``pbi_set_series_color_tool`` writes a per-series ``dataPoint``
  override with the correct selector, leaving sibling series alone.
- Smoke: every wrapper module imports cleanly after the rewiring.
"""

from __future__ import annotations

import json
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

from tools.visuals._ops import (
    _build_series_selector,
    pbi_set_series_color_tool,
)
from wrappers._helpers import register_tool


def _build_layout(extract_folder: Path, projections: dict[str, list[dict[str, str]]]) -> None:
    """Write a minimal Report/Layout that ``_load_layout`` accepts."""
    layout = {
        "id": 0,
        "resourcePackages": [],
        "sections": [
            {
                "id": 1,
                "name": "ReportSection1",
                "displayName": "Page 1",
                "visualContainers": [
                    {
                        "x": 0,
                        "y": 0,
                        "width": 400,
                        "height": 300,
                        "config": json.dumps(
                            {
                                "name": "visual-1",
                                "singleVisual": {
                                    "visualType": "lineChart",
                                    "projections": projections,
                                    "prototypeQuery": {
                                        "From": [{"Name": "s", "Entity": "Sales"}],
                                        "Select": [
                                            {
                                                "Measure": {
                                                    "Expression": {"SourceRef": {"Source": "s"}},
                                                    "Property": "Revenue",
                                                },
                                                "Name": "Sales.Revenue",
                                            },
                                            {
                                                "Measure": {
                                                    "Expression": {"SourceRef": {"Source": "s"}},
                                                    "Property": "Cost",
                                                },
                                                "Name": "Sales.Cost",
                                            },
                                        ],
                                    },
                                },
                            }
                        ),
                    }
                ],
            }
        ],
    }
    layout_path = extract_folder / "Report" / "Layout"
    layout_path.parent.mkdir(parents=True, exist_ok=True)
    layout_path.write_bytes(json.dumps(layout).encode("utf-16-le"))


class ManagerInjectionTests(unittest.TestCase):
    """Bug 1 regression: ``register_tool`` must not bind CONNECTION_MANAGER
    positionally when ``manager`` appears late in the underlying tool's
    signature. The previous wrapper sent it as ``*args[0]``, which collided
    with the first positional parameter (``extract_folder`` etc.).
    """

    def test_manager_late_signature_does_not_collide(self) -> None:
        captured: dict[str, object] = {}

        # Tool with NO path parameters so the wrapper-level security
        # validator does not get in the way of the test's assertion. The
        # bug we're regressing on is purely about argument routing.
        def underlying_tool(
            label: str,
            count: int = 0,
            force: bool = False,
            manager: object | None = None,
        ) -> dict[str, object]:
            captured["label"] = label
            captured["count"] = count
            captured["force"] = force
            captured["manager"] = manager
            return {"ok": True}

        wrapper = register_tool(underlying_tool, name=f"pbi_test_late_manager_{id(underlying_tool)}")
        wrapper(label="abc", count=3, force=False)
        self.assertEqual(captured["label"], "abc")
        self.assertEqual(captured["count"], 3)
        self.assertFalse(captured["force"])
        self.assertIsNotNone(captured["manager"], "manager must be auto-injected")

    def test_caller_provided_manager_wins(self) -> None:
        captured: dict[str, object] = {}

        def underlying_tool(label: str, manager: object | None = None) -> dict[str, object]:
            captured["manager"] = manager
            return {"ok": True}

        wrapper = register_tool(underlying_tool, name=f"pbi_test_explicit_manager_{id(underlying_tool)}")
        sentinel = object()
        wrapper(label="abc", manager=sentinel)
        self.assertIs(captured["manager"], sentinel)


class SeriesTargetTests(unittest.TestCase):
    """Bug 3 regression: ``pbi_set_series_color`` must emit a selector that
    pins the colour to a single series instead of overriding every series
    via ``defaultColor``.
    """

    def setUp(self) -> None:
        self.tmp = tempfile.TemporaryDirectory()
        self.root = Path(self.tmp.name)
        from security import SECURITY, configure_allowed_dirs

        self.previous_allowed = [str(item) for item in SECURITY.allowed_base_dirs()]
        configure_allowed_dirs([str(self.root)])
        SECURITY.policy(reload=True, cwd=self.root)
        _build_layout(
            self.root,
            projections={
                "Y": [{"queryRef": "Sales.Revenue"}, {"queryRef": "Sales.Cost"}],
            },
        )

    def tearDown(self) -> None:
        from security import SECURITY, configure_allowed_dirs

        configure_allowed_dirs(self.previous_allowed)
        SECURITY.policy(reload=True, cwd=Path.cwd())
        self.tmp.cleanup()

    def _read_data_point(self) -> list[dict[str, object]]:
        layout = json.loads((self.root / "Report" / "Layout").read_bytes().decode("utf-16-le"))
        config = json.loads(layout["sections"][0]["visualContainers"][0]["config"])
        return config["singleVisual"].get("objects", {}).get("dataPoint", [])

    def test_series_index_targets_single_series(self) -> None:
        result = pbi_set_series_color_tool(
            extract_folder=str(self.root),
            page="Page 1",
            visual_id="visual-1",
            color="#FF8800",
            series_index=1,
        )
        self.assertTrue(result["ok"], result)
        data_point = self._read_data_point()
        self.assertEqual(len(data_point), 1)
        entry = data_point[0]
        # The selector must reference Cost, not Revenue, so the override
        # only paints the second series.
        selector = entry.get("selector") or {}
        measure_meta = selector.get("measure") or {}
        self.assertEqual(measure_meta.get("Property"), "Cost")
        self.assertEqual(result["target"]["property"], "Cost")
        self.assertEqual(result["target"]["role"], "Y")

    def test_series_name_match(self) -> None:
        result = pbi_set_series_color_tool(
            extract_folder=str(self.root),
            page="Page 1",
            visual_id="visual-1",
            color="#0044CC",
            series_name="Revenue",
        )
        self.assertTrue(result["ok"], result)
        self.assertEqual(result["target"]["property"], "Revenue")

    def test_series_index_out_of_range(self) -> None:
        # _run wraps the validation error into the MCP error envelope.
        result = pbi_set_series_color_tool(
            extract_folder=str(self.root),
            page="Page 1",
            visual_id="visual-1",
            color="#000000",
            series_index=99,
        )
        self.assertFalse(result.get("ok", True))
        self.assertIn("series_index", json.dumps(result))

    def test_selector_shape(self) -> None:
        target = {"kind": "measure", "entity": "Sales", "property": "Cost", "role": "Y", "query_ref": "Sales.Cost"}
        sel = _build_series_selector(target)
        self.assertIn("measure", sel)
        self.assertEqual(sel["measure"]["Property"], "Cost")

    def test_resolve_target_default_role_order(self) -> None:
        # When the visual has both Series and Y projections,
        # _SERIES_COLOR_DEFAULT_ROLE_ORDER puts Y before Series.
        from tools.visuals._ops import _resolve_series_target

        single_visual = {
            "projections": {
                "Series": [{"queryRef": "Sales.Cost"}],
                "Y": [{"queryRef": "Sales.Revenue"}],
            },
            "prototypeQuery": {
                "From": [{"Name": "s", "Entity": "Sales"}],
                "Select": [
                    {
                        "Measure": {"Expression": {"SourceRef": {"Source": "s"}}, "Property": "Revenue"},
                        "Name": "Sales.Revenue",
                    },
                    {
                        "Measure": {"Expression": {"SourceRef": {"Source": "s"}}, "Property": "Cost"},
                        "Name": "Sales.Cost",
                    },
                ],
            },
        }
        target = _resolve_series_target(single_visual, series_index=0, series_name=None)
        self.assertEqual(target["role"], "Y")
        self.assertEqual(target["property"], "Revenue")


class WrapperImportSmokeTests(unittest.TestCase):
    """Sanity check: every wrappers/<domain>.py imports without errors."""

    def test_all_wrappers_import(self) -> None:
        # Importing server wires every wrapper module via side-effect imports.
        import importlib

        import server  # noqa: F401

        for name in ("connection", "measures", "model", "power_query", "visuals", "quality"):
            module = importlib.import_module(f"wrappers.{name}")
            self.assertIsNotNone(module)


if __name__ == "__main__":
    unittest.main(verbosity=2)

"""Standalone tests for report visual field validation and repair."""

from __future__ import annotations

import json
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

from security import SECURITY, tool_category
from tools.visuals import (
    LAYOUT_RELATIVE_PATH,
    pbi_patch_layout_tool,
    pbi_repair_report_fields_tool,
    pbi_validate_report_fields_tool,
)


def _dump(value: object) -> str:
    return json.dumps(value, ensure_ascii=False, separators=(",", ":"))


def _layout_with_visual(single_visual: dict) -> dict:
    config = {"name": "visual1", "singleVisual": single_visual}
    return {
        "sections": [
            {
                "name": "ReportSection1",
                "displayName": "Overview",
                "visualContainers": [
                    {
                        "x": 0,
                        "y": 0,
                        "width": 200,
                        "height": 120,
                        "config": _dump(config),
                        "query": _dump(
                            {
                                "Commands": [
                                    {"SemanticQueryDataShapeCommand": {"Query": single_visual["prototypeQuery"]}}
                                ]
                            }
                        ),
                        "filters": "[]",
                        "dataTransforms": "{}",
                    }
                ],
            }
        ]
    }


def _write_layout(folder: Path, layout: dict) -> None:
    path = folder / LAYOUT_RELATIVE_PATH
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(layout), encoding="utf-16-le")


def _read_visual(folder: Path) -> dict:
    layout = json.loads((folder / LAYOUT_RELATIVE_PATH).read_text(encoding="utf-16-le"))
    container = layout["sections"][0]["visualContainers"][0]
    return json.loads(container["config"])["singleVisual"]


class VisualFieldValidationTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp_dir = tempfile.TemporaryDirectory()
        self.root = Path(self.temp_dir.name)
        self.extract_folder = self.root / "extract"
        self.extract_folder.mkdir(parents=True, exist_ok=True)
        self.previous_allowed = [str(item) for item in SECURITY.allowed_base_dirs()]
        SECURITY.configure_allowed_dirs([str(self.root)])

    def tearDown(self) -> None:
        SECURITY.configure_allowed_dirs(self.previous_allowed)
        self.temp_dir.cleanup()

    def test_repair_query_ref_mismatch_dry_run_does_not_modify_layout(self) -> None:
        single_visual = {
            "visualType": "tableEx",
            "projections": {"Values": [{"queryRef": "Sales.Year"}]},
            "prototypeQuery": {
                "Version": 2,
                "From": [{"Name": "s0", "Entity": "Sales"}],
                "Select": [
                    {"Column": {"Expression": {"SourceRef": {"Source": "s0"}}, "Property": "Year"}, "Name": "Year"}
                ],
            },
        }
        _write_layout(self.extract_folder, _layout_with_visual(single_visual))
        before = (self.extract_folder / LAYOUT_RELATIVE_PATH).read_bytes()

        result = pbi_repair_report_fields_tool(str(self.extract_folder), apply=False)

        self.assertTrue(result["ok"], result)
        self.assertTrue(result["needs_apply"])
        self.assertEqual((self.extract_folder / LAYOUT_RELATIVE_PATH).read_bytes(), before)

    def test_repair_query_ref_mismatch_apply_updates_projection(self) -> None:
        single_visual = {
            "visualType": "tableEx",
            "projections": {"Values": [{"queryRef": "Sales.Year"}]},
            "prototypeQuery": {
                "Version": 2,
                "From": [{"Name": "s0", "Entity": "Sales"}],
                "Select": [
                    {"Column": {"Expression": {"SourceRef": {"Source": "s0"}}, "Property": "Year"}, "Name": "Year"}
                ],
            },
        }
        _write_layout(self.extract_folder, _layout_with_visual(single_visual))

        result = pbi_repair_report_fields_tool(str(self.extract_folder), apply=True)

        self.assertTrue(result["ok"], result)
        repaired = _read_visual(self.extract_folder)
        self.assertEqual(repaired["projections"]["Values"][0]["queryRef"], "Year")

    def test_repair_gauge_value_role_to_y(self) -> None:
        single_visual = {
            "visualType": "gauge",
            "projections": {"Value": [{"queryRef": "SalesTarget"}]},
            "prototypeQuery": {
                "Version": 2,
                "From": [{"Name": "s0", "Entity": "Measures"}],
                "Select": [
                    {
                        "Measure": {"Expression": {"SourceRef": {"Source": "s0"}}, "Property": "SalesTarget"},
                        "Name": "SalesTarget",
                    }
                ],
            },
        }
        _write_layout(self.extract_folder, _layout_with_visual(single_visual))

        result = pbi_repair_report_fields_tool(str(self.extract_folder), apply=True)

        self.assertTrue(result["ok"], result)
        repaired = _read_visual(self.extract_folder)
        self.assertIn("Y", repaired["projections"])
        self.assertNotIn("Value", repaired["projections"])

    def test_repair_measure_home_table_from_metadata(self) -> None:
        measures_dir = self.extract_folder / "Model" / "tables" / "Sales" / "measures"
        measures_dir.mkdir(parents=True)
        (measures_dir / "Total Sales.dax").write_text("Total Sales = SUM(Sales[Amount])", encoding="utf-8")
        single_visual = {
            "visualType": "card",
            "projections": {"Values": [{"queryRef": "Total Sales"}]},
            "prototypeQuery": {
                "Version": 2,
                "From": [{"Name": "s0", "Entity": "$Measures"}],
                "Select": [
                    {
                        "Measure": {"Expression": {"SourceRef": {"Source": "s0"}}, "Property": "Total Sales"},
                        "Name": "Total Sales",
                    }
                ],
            },
        }
        _write_layout(self.extract_folder, _layout_with_visual(single_visual))

        result = pbi_repair_report_fields_tool(str(self.extract_folder), apply=True)

        self.assertTrue(result["ok"], result)
        repaired = _read_visual(self.extract_folder)
        self.assertEqual(repaired["prototypeQuery"]["From"][0]["Entity"], "Sales")

    def test_unresolved_field_is_reported_not_removed(self) -> None:
        single_visual = {
            "visualType": "tableEx",
            "projections": {"Values": [{"queryRef": "Missing"}]},
            "prototypeQuery": {"Version": 2, "From": [], "Select": []},
        }
        _write_layout(self.extract_folder, _layout_with_visual(single_visual))

        result = pbi_validate_report_fields_tool(str(self.extract_folder))

        self.assertTrue(result["ok"], result)
        self.assertFalse(result["valid"])
        self.assertEqual(result["issues"][0]["issue"], "query_ref_not_found")
        self.assertEqual(
            len(
                json.loads((self.extract_folder / LAYOUT_RELATIVE_PATH).read_text(encoding="utf-16-le"))["sections"][0][
                    "visualContainers"
                ]
            ),
            1,
        )

    def test_live_model_validation_reports_missing_column(self) -> None:
        single_visual = {
            "visualType": "tableEx",
            "projections": {"Values": [{"queryRef": "MissingColumn"}]},
            "prototypeQuery": {
                "Version": 2,
                "From": [{"Name": "s0", "Entity": "Sales"}],
                "Select": [
                    {
                        "Column": {"Expression": {"SourceRef": {"Source": "s0"}}, "Property": "MissingColumn"},
                        "Name": "MissingColumn",
                    }
                ],
            },
        }
        _write_layout(self.extract_folder, _layout_with_visual(single_visual))
        with patch(
            "tools.visuals.pbi_model_info_tool",
            return_value={
                "ok": True,
                "tables": [{"name": "Sales", "columns": [{"name": "Amount"}]}],
                "measures": [],
                "relationships": [],
            },
        ):
            result = pbi_validate_report_fields_tool(str(self.extract_folder), manager=object())

        self.assertTrue(result["ok"], result)
        self.assertFalse(result["valid"])
        self.assertEqual(result["model_validation"]["status"], "available")
        self.assertIn("column_not_found", {item["issue"] for item in result["issues"]})

    def test_live_model_validation_reports_missing_measure(self) -> None:
        single_visual = {
            "visualType": "card",
            "projections": {"Values": [{"queryRef": "Missing Measure"}]},
            "prototypeQuery": {
                "Version": 2,
                "From": [{"Name": "s0", "Entity": "Sales"}],
                "Select": [
                    {
                        "Measure": {"Expression": {"SourceRef": {"Source": "s0"}}, "Property": "Missing Measure"},
                        "Name": "Missing Measure",
                    }
                ],
            },
        }
        _write_layout(self.extract_folder, _layout_with_visual(single_visual))
        with patch(
            "tools.visuals.pbi_model_info_tool",
            return_value={
                "ok": True,
                "tables": [{"name": "Sales", "columns": []}],
                "measures": [{"table": "Sales", "name": "Existing Measure"}],
                "relationships": [],
            },
        ):
            result = pbi_validate_report_fields_tool(str(self.extract_folder), manager=object())

        self.assertTrue(result["ok"], result)
        self.assertFalse(result["valid"])
        self.assertIn("measure_not_found", {item["issue"] for item in result["issues"]})

    def test_live_model_home_table_resolves_measure_missing_from_extract_metadata(self) -> None:
        single_visual = {
            "visualType": "card",
            "projections": {"Values": [{"queryRef": "Total Value"}]},
            "prototypeQuery": {
                "Version": 2,
                "From": [{"Name": "s0", "Entity": "$Measures"}],
                "Select": [
                    {
                        "Measure": {"Expression": {"SourceRef": {"Source": "s0"}}, "Property": "Total Value"},
                        "Name": "Total Value",
                    }
                ],
            },
        }
        _write_layout(self.extract_folder, _layout_with_visual(single_visual))
        with patch(
            "tools.visuals.pbi_model_info_tool",
            return_value={
                "ok": True,
                "tables": [{"name": "TestData", "columns": [{"name": "Value"}]}],
                "measures": [{"table": "TestData", "name": "Total Value"}],
                "relationships": [],
            },
        ):
            result = pbi_repair_report_fields_tool(str(self.extract_folder), manager=object(), apply=True)

        self.assertTrue(result["ok"], result)
        self.assertEqual(result["unresolved"], [])
        self.assertEqual(result["persistence_risk_count"], 1)
        self.assertEqual(result["persistence_risks"][0]["source"], "live_model")
        repaired = _read_visual(self.extract_folder)
        self.assertEqual(repaired["prototypeQuery"]["From"][0]["Entity"], "TestData")

    def test_patch_layout_can_block_live_only_persistence_risk(self) -> None:
        single_visual = {
            "visualType": "card",
            "projections": {"Values": [{"queryRef": "Total Value"}]},
            "prototypeQuery": {
                "Version": 2,
                "From": [{"Name": "s0", "Entity": "$Measures"}],
                "Select": [
                    {
                        "Measure": {"Expression": {"SourceRef": {"Source": "s0"}}, "Property": "Total Value"},
                        "Name": "Total Value",
                    }
                ],
            },
        }
        _write_layout(self.extract_folder, _layout_with_visual(single_visual))
        pbix_path = self.root / "report.pbix"
        pbix_path.write_bytes(b"fake")
        with patch(
            "tools.visuals.pbi_model_info_tool",
            return_value={
                "ok": True,
                "tables": [{"name": "TestData", "columns": [{"name": "Value"}]}],
                "measures": [{"table": "TestData", "name": "Total Value"}],
                "relationships": [],
            },
        ):
            result = pbi_patch_layout_tool(
                str(self.extract_folder),
                str(pbix_path),
                fail_on_persistence_risk=True,
                manager=object(),
            )

        self.assertFalse(result["ok"], result)
        self.assertEqual(result["error"]["code"], "validation_error")
        self.assertEqual(result["error"]["details"]["persistence_risk_count"], 1)

    def test_repair_tool_security_category_uses_apply_flag(self) -> None:
        self.assertEqual(tool_category("pbi_repair_report_fields", {"apply": False}), "read")
        self.assertEqual(tool_category("pbi_repair_report_fields", {"apply": True}), "write")


if __name__ == "__main__":
    unittest.main(verbosity=2)

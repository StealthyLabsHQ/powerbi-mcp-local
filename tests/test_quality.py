"""Standalone tests for quality gate helpers."""

from __future__ import annotations

import json
import os
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

from security import SECURITY, tool_category
from tools.quality import (
    LAYOUT_RELATIVE_PATH,
    _duplicate_relationship_key_issues,
    _model_audit_from_snapshot,
    pbi_compare_report_versions_tool,
    pbi_detect_dirty_dates_tool,
    pbi_detect_empty_visuals_tool,
    pbi_detect_name_collisions_tool,
    pbi_export_validation_report_tool,
    pbi_generate_measure_tests_tool,
    pbi_lint_report_layout_tool,
    pbi_validate_filter_expression_tool,
    pbi_validate_pbix_persistence_tool,
    pbi_validate_pbix_reopen_tool,
    pbi_validate_relationship_plan_tool,
)


def _write_layout(folder: Path, visuals: list[dict], *, width: int = 1280, height: int = 720) -> None:
    path = folder / LAYOUT_RELATIVE_PATH
    path.parent.mkdir(parents=True, exist_ok=True)
    layout = {
        "sections": [
            {
                "name": "ReportSection1",
                "displayName": "Overview",
                "width": width,
                "height": height,
                "visualContainers": visuals,
            }
        ]
    }
    path.write_bytes(json.dumps(layout).encode("utf-16-le"))


def _container(
    name: str, x: int, y: int, width: int, height: int, visual_type: str = "card", title: bool = True
) -> dict:
    objects = {"title": [{"properties": {"show": {"expr": {"Literal": {"Value": "true"}}}}}]} if title else {}
    return {
        "x": x,
        "y": y,
        "width": width,
        "height": height,
        "config": json.dumps({"name": name, "singleVisual": {"visualType": visual_type, "objects": objects}}),
    }


class QualityToolTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp_dir = tempfile.TemporaryDirectory()
        self.root = Path(self.temp_dir.name)
        self.previous_allowed = [str(item) for item in SECURITY.allowed_base_dirs()]
        SECURITY.configure_allowed_dirs([str(self.root)])

    def tearDown(self) -> None:
        SECURITY.configure_allowed_dirs(self.previous_allowed)
        self.temp_dir.cleanup()

    def test_model_audit_detects_unrelated_table_and_ambiguous_triangle(self) -> None:
        snapshot = {
            "tables": [
                {"name": "Fact", "is_hidden": False, "columns": []},
                {"name": "Customer", "is_hidden": False, "columns": []},
                {"name": "Region", "is_hidden": False, "columns": []},
                {"name": "Unused", "is_hidden": False, "columns": []},
            ],
            "measures": [{"table": "Fact", "name": "Revenue"}],
            "relationships": [
                {"from_table": "Fact", "to_table": "Customer", "direction": "OneDirection"},
                {"from_table": "Fact", "to_table": "Region", "direction": "OneDirection"},
                {"from_table": "Customer", "to_table": "Region", "direction": "OneDirection"},
            ],
        }

        result = _model_audit_from_snapshot(snapshot)

        warning_types = {item["type"] for item in result["warnings"]}
        self.assertIn("unrelated_table", warning_types)
        self.assertIn("ambiguous_relationship_triangle", warning_types)

    def test_layout_lint_detects_overlap_small_visual_and_missing_title(self) -> None:
        folder = self.root / "extract"
        _write_layout(
            folder,
            [
                _container("A", 10, 10, 100, 70, title=False),
                _container("B", 50, 40, 200, 120),
            ],
        )

        result = pbi_lint_report_layout_tool(str(folder))

        self.assertFalse(result["valid"], result)
        issue_types = {item["type"] for item in result["issues"]}
        warning_types = {item["type"] for item in result["warnings"]}
        self.assertIn("visual_overlap", issue_types)
        self.assertIn("visual_too_small", warning_types)
        self.assertIn("missing_title", warning_types)

    def test_duplicate_relationship_key_check_detects_duplicate_one_side(self) -> None:
        class Manager:
            def run_adomd_query(self, query, max_rows=1):
                return {"rows": [{"[__Rows]": 3, "[__Distinct]": 2}]}

        issues = _duplicate_relationship_key_issues(
            Manager(),
            [
                {
                    "from_table": "FactSales",
                    "from_column": "CustomerKey",
                    "to_table": "Customer",
                    "to_column": "CustomerKey",
                    "cardinality": "ManyToOne",
                }
            ],
        )

        self.assertEqual(issues[0]["type"], "duplicate_relationship_key")
        self.assertEqual(issues[0]["table"], "Customer")

    def test_duplicate_relationship_key_check_detects_non_fact_many_side(self) -> None:
        class Manager:
            def run_adomd_query(self, query, max_rows=1):
                if "CustomerDuplicate" in query:
                    return {"rows": [{"__Rows": 3, "__Distinct": 2}]}
                return {"rows": [{"__Rows": 2, "__Distinct": 2}]}

        issues = _duplicate_relationship_key_issues(
            Manager(),
            [
                {
                    "from_table": "CustomerDuplicate",
                    "from_column": "CustomerKey",
                    "to_table": "FactSales",
                    "to_column": "CustomerKey",
                    "cardinality": "ManyToOne",
                }
            ],
        )

        self.assertEqual(issues[0]["relationship_role"], "non_fact_many_side")

    def test_duplicate_relationship_key_check_ignores_fact_many_side(self) -> None:
        class Manager:
            def run_adomd_query(self, query, max_rows=1):
                if "Customer" in query:
                    return {"rows": [{"__Rows": 2, "__Distinct": 2}]}
                return {"rows": [{"__Rows": 20, "__Distinct": 2}]}

        issues = _duplicate_relationship_key_issues(
            Manager(),
            [
                {
                    "from_table": "Sales",
                    "from_column": "CustomerKey",
                    "to_table": "Customer",
                    "to_column": "CustomerKey",
                    "cardinality": "ManyToOne",
                }
            ],
        )

        self.assertEqual(issues, [])

    def test_compare_report_versions_returns_visual_delta(self) -> None:
        a = self.root / "a"
        b = self.root / "b"
        _write_layout(a, [_container("A", 10, 10, 200, 120)])
        _write_layout(b, [_container("A", 10, 10, 200, 120), _container("B", 240, 10, 200, 120)])

        result = pbi_compare_report_versions_tool(extract_folder_a=str(a), extract_folder_b=str(b))

        self.assertTrue(result["ok"], result)
        self.assertEqual(result["delta"]["visual_count"], 1)

    def test_detect_name_collisions_flags_measure_column_collision(self) -> None:
        snapshot = {
            "tables": [{"name": "Sales", "columns": [{"name": "Revenue"}]}],
            "measures": [{"table": "Sales", "name": "Revenue"}],
            "relationships": [],
        }

        with patch("tools.quality._model_snapshot", return_value=snapshot):
            result = pbi_detect_name_collisions_tool(object())

        self.assertFalse(result["valid"], result)
        self.assertEqual(result["issues"][0]["type"], "measure_column_name_collision")

    def test_detect_dirty_dates_flags_invalid_text_dates(self) -> None:
        snapshot = {
            "tables": [{"name": "TextDates", "columns": [{"name": "Date texte", "data_type": "String"}]}],
            "measures": [],
            "relationships": [],
        }

        class Manager:
            def run_adomd_query(self, query, max_rows=200):
                return {"rows": [{"[__Value]": "2025-01-01"}, {"[__Value]": "bad-date"}, {"[__Value]": ""}]}

        with patch("tools.quality._model_snapshot", return_value=snapshot):
            result = pbi_detect_dirty_dates_tool(Manager(), table="TextDates", min_parse_success_rate=0.8)

        self.assertFalse(result["valid"], result)
        self.assertEqual(result["issues"][0]["type"], "dirty_text_date")

    def test_validate_relationship_plan_blocks_many_to_many(self) -> None:
        snapshot = {
            "tables": [
                {"name": "Sales", "columns": [{"name": "CustomerID"}]},
                {"name": "Customers", "columns": [{"name": "CustomerID"}]},
            ],
            "measures": [],
            "relationships": [],
        }

        class Manager:
            def run_adomd_query(self, query, max_rows=1):
                return {"rows": [{"[__Rows]": 3, "[__Distinct]": 2, "[__Blank]": 0}]}

        with patch("tools.quality._model_snapshot", return_value=snapshot):
            result = pbi_validate_relationship_plan_tool(
                Manager(),
                from_table="Sales",
                from_column="CustomerID",
                to_table="Customers",
                to_column="CustomerID",
                cardinality="manyToMany",
            )

        self.assertFalse(result["safe_to_create"], result)
        self.assertIn("many_to_many_relationship", {item["type"] for item in result["issues"]})

    def test_detect_empty_visuals_flags_zero_row_probe(self) -> None:
        folder = self.root / "extract"
        visual = _container("A", 10, 10, 200, 120)
        config = json.loads(visual["config"])
        config["singleVisual"]["prototypeQuery"] = {
            "From": [{"Name": "s", "Entity": "Sales"}],
            "Select": [
                {
                    "Column": {"Expression": {"SourceRef": {"Source": "s"}}, "Property": "Region"},
                    "Name": "Sales.Region",
                },
                {
                    "Measure": {"Expression": {"SourceRef": {"Source": "s"}}, "Property": "Revenue"},
                    "Name": "Sales.Revenue",
                },
            ],
        }
        visual["config"] = json.dumps(config)
        _write_layout(folder, [visual])

        class Manager:
            def run_adomd_query(self, query, max_rows=1):
                return {"rows": []}

        result = pbi_detect_empty_visuals_tool(Manager(), extract_folder=str(folder))

        self.assertFalse(result["valid"], result)
        self.assertEqual(result["issues"][0]["type"], "empty_visual")

    def test_detect_empty_visuals_warns_all_blank_measures(self) -> None:
        folder = self.root / "extract"
        visual = _container("A", 10, 10, 200, 120)
        config = json.loads(visual["config"])
        config["singleVisual"]["prototypeQuery"] = {
            "From": [{"Name": "s", "Entity": "Sales"}],
            "Select": [
                {
                    "Measure": {"Expression": {"SourceRef": {"Source": "s"}}, "Property": "Revenue"},
                    "Name": "Sales.Revenue",
                },
            ],
        }
        visual["config"] = json.dumps(config)
        _write_layout(folder, [visual])

        class Manager:
            def run_adomd_query(self, query, max_rows=1):
                return {"rows": [{"[__M0]": None}]}

        result = pbi_detect_empty_visuals_tool(
            Manager(), extract_folder=str(folder), filter_expression="'Sales'[Region] = \"Nowhere\""
        )

        self.assertTrue(result["valid"], result)
        self.assertEqual(result["warnings"][0]["type"], "visual_measures_all_blank")

    def test_detect_empty_visuals_wraps_table_probe_filter_in_calculatetable(self) -> None:
        folder = self.root / "extract"
        visual = _container("A", 10, 10, 200, 120)
        config = json.loads(visual["config"])
        config["singleVisual"]["prototypeQuery"] = {
            "From": [{"Name": "s", "Entity": "Sales"}],
            "Select": [
                {
                    "Column": {"Expression": {"SourceRef": {"Source": "s"}}, "Property": "Region"},
                    "Name": "Sales.Region",
                },
                {
                    "Measure": {"Expression": {"SourceRef": {"Source": "s"}}, "Property": "Revenue"},
                    "Name": "Sales.Revenue",
                },
            ],
        }
        visual["config"] = json.dumps(config)
        _write_layout(folder, [visual])

        class Manager:
            query = ""

            def run_adomd_query(self, query, max_rows=1):
                self.query = query
                return {"rows": [{"[__M0]": 1}]}

        manager = Manager()
        result = pbi_detect_empty_visuals_tool(
            manager, extract_folder=str(folder), filter_expression="'Calendar'[Year] = 2025"
        )

        self.assertTrue(result["valid"], result)
        self.assertIn("CALCULATETABLE(SUMMARIZECOLUMNS", manager.query)
        self.assertNotIn("ALLSELECTED()", manager.query)

    def test_validate_filter_expression_reports_invalid_filter(self) -> None:
        class Manager:
            def run_adomd_query(self, query, max_rows=1):
                raise RuntimeError("bad filter")

        result = pbi_validate_filter_expression_tool(Manager(), filter_expression="'Calendar'[Year] ==")

        self.assertFalse(result["valid"], result)
        self.assertIn("bad filter", result["error"])

    def test_detect_empty_visuals_skips_when_filter_invalid(self) -> None:
        folder = self.root / "extract"
        _write_layout(folder, [_container("A", 10, 10, 200, 120)])

        class Manager:
            def run_adomd_query(self, query, max_rows=1):
                raise RuntimeError("bad filter")

        result = pbi_detect_empty_visuals_tool(
            Manager(), extract_folder=str(folder), filter_expression="'Calendar'[Year] =="
        )

        self.assertFalse(result["valid"], result)
        self.assertEqual(result["issues"][0]["type"], "invalid_filter_expression")

    def test_export_validation_report_writes_json(self) -> None:
        output = self.root / "validation.json"
        with (
            patch(
                "tools.quality.pbi_audit_model_tool", return_value={"valid": True, "issue_count": 0, "warning_count": 0}
            ),
            patch(
                "tools.quality.pbi_lint_dax_tool", return_value={"valid": True, "issue_count": 0, "warning_count": 0}
            ),
            patch("tools.quality.pbi_detect_name_collisions_tool", return_value={"valid": True, "issue_count": 0}),
            patch("tools.quality.pbi_detect_dirty_dates_tool", return_value={"valid": True, "issue_count": 0}),
            patch("tools.quality.pbi_score_dashboard_tool", return_value={"score_total": 100}),
        ):
            result = pbi_export_validation_report_tool(object(), output_path=str(output))

        self.assertTrue(result["ok"], result)
        self.assertTrue(result["overall_valid"], result)
        self.assertTrue(output.exists())
        payload = json.loads(output.read_text(encoding="utf-8"))
        self.assertEqual(payload["score"]["score_total"], 100)
        self.assertTrue(payload["summary"]["overall_valid"])

    def test_generate_measure_tests_executes_selected_measure(self) -> None:
        snapshot = {
            "tables": [],
            "relationships": [],
            "measures": [
                {"table": "Measures", "name": "Revenue", "expression": "SUM(Sales[Revenue])", "format_string": "$#,0"},
                {"table": "Measures", "name": "Bad Ratio", "expression": "[Revenue] / [Cost]", "format_string": "0.0%"},
            ],
        }

        class Manager:
            def run_adomd_query(self, query, max_rows=1):
                return {"rows": [{"[__Value]": 10}]}

        with patch("tools.quality._model_snapshot", return_value=snapshot):
            result = pbi_generate_measure_tests_tool(Manager(), measures=["Revenue"])

        self.assertTrue(result["valid"], result)
        self.assertEqual(result["tested_count"], 1)
        self.assertEqual(result["tests"][0]["measure"], "Revenue")

    def test_generate_measure_tests_warns_unsafe_division(self) -> None:
        snapshot = {
            "tables": [],
            "relationships": [],
            "measures": [
                {"table": "Measures", "name": "Bad Ratio", "expression": "[Revenue] / [Cost]", "format_string": "0.0%"}
            ],
        }

        class Manager:
            def run_adomd_query(self, query, max_rows=1):
                return {"rows": [{"[__Value]": 0}]}

        with patch("tools.quality._model_snapshot", return_value=snapshot):
            result = pbi_generate_measure_tests_tool(Manager())

        self.assertTrue(result["valid"], result)
        self.assertIn("unsafe_division_operator", {item["type"] for item in result["warnings"]})

    def test_generate_measure_tests_treats_ratio_names_as_numbers(self) -> None:
        snapshot = {
            "tables": [],
            "relationships": [],
            "measures": [
                {"table": "Measures", "name": "LTV/CAC", "expression": "DIVIDE([LTV], [CAC])", "format_string": "0.0"},
                {
                    "table": "Measures",
                    "name": "Pipeline Coverage",
                    "expression": "DIVIDE([Pipeline], [Target])",
                    "format_string": "0.0",
                },
            ],
        }

        class Manager:
            def run_adomd_query(self, query, max_rows=1):
                return {"rows": [{"[__Value]": 1.5}]}

        with patch("tools.quality._model_snapshot", return_value=snapshot):
            result = pbi_generate_measure_tests_tool(Manager())

        self.assertEqual([item for item in result["warnings"] if item["type"] == "unexpected_measure_format"], [])

    def test_validate_pbix_persistence_detects_security_bindings(self) -> None:
        folder = self.root / "extract"
        _write_layout(folder, [_container("A", 10, 10, 200, 120)])
        pbix = self.root / "report.pbix"
        layout_bytes = (folder / LAYOUT_RELATIVE_PATH).read_bytes()
        with zipfile.ZipFile(pbix, "w") as archive:
            archive.writestr("Report/Layout", layout_bytes)
            archive.writestr("SecurityBindings", b"stale")

        result = pbi_validate_pbix_persistence_tool(pbix_path=str(pbix), extract_folder=str(folder))

        self.assertFalse(result["valid"], result)
        self.assertIn("security_bindings_present", {item["type"] for item in result["issues"]})

    def test_validate_pbix_persistence_compares_visual_counts(self) -> None:
        folder = self.root / "extract"
        _write_layout(folder, [_container("A", 10, 10, 200, 120), _container("B", 230, 10, 200, 120)])
        other = self.root / "other"
        _write_layout(other, [_container("A", 10, 10, 200, 120)])
        pbix = self.root / "report.pbix"
        with zipfile.ZipFile(pbix, "w") as archive:
            archive.writestr("Report/Layout", (other / LAYOUT_RELATIVE_PATH).read_bytes())

        result = pbi_validate_pbix_persistence_tool(pbix_path=str(pbix), extract_folder=str(folder))

        self.assertFalse(result["valid"], result)
        self.assertIn("visual_count_mismatch", {item["type"] for item in result["issues"]})

    @unittest.skipUnless(os.name == "nt", "patching os.name to 'nt' makes pathlib instantiate WindowsPath")
    def test_validate_pbix_reopen_flags_fix_this_signal(self) -> None:
        pbix = self.root / "report.pbix"
        with zipfile.ZipFile(pbix, "w") as archive:
            archive.writestr("Report/Layout", json.dumps({"sections": []}).encode("utf-16-le"))

        with (
            patch("tools.quality.os.name", "nt"),
            patch(
                "tools.quality._run_reopen_probe",
                return_value={
                    "opened": True,
                    "process_id": 1,
                    "process_name": "PBIDesktop",
                    "window_title": "report - Power BI Desktop",
                    "ui_text_count": 3,
                    "ui_text_matches": ["Fix this"],
                    "screenshot_path": None,
                },
            ),
        ):
            result = pbi_validate_pbix_reopen_tool(pbix_path=str(pbix), timeout_seconds=10)

        self.assertFalse(result["valid"], result)
        self.assertIn("powerbi_fix_this_signal", {item["type"] for item in result["issues"]})

    @unittest.skipUnless(os.name == "nt", "patching os.name to 'nt' makes pathlib instantiate WindowsPath")
    def test_validate_pbix_reopen_accepts_clean_probe(self) -> None:
        pbix = self.root / "report.pbix"
        with zipfile.ZipFile(pbix, "w") as archive:
            archive.writestr(
                "Report/Layout", json.dumps({"sections": [{"visualContainers": [{}]}]}).encode("utf-16-le")
            )

        with (
            patch("tools.quality.os.name", "nt"),
            patch(
                "tools.quality._run_reopen_probe",
                return_value={
                    "opened": True,
                    "process_id": 1,
                    "process_name": "PBIDesktop",
                    "window_title": "report - Power BI Desktop",
                    "ui_text_count": 0,
                    "ui_text_matches": [],
                    "screenshot_path": None,
                },
            ),
        ):
            result = pbi_validate_pbix_reopen_tool(pbix_path=str(pbix), timeout_seconds=10)

        self.assertTrue(result["valid"], result)

    @unittest.skipUnless(os.name == "nt", "patching os.name to 'nt' makes pathlib instantiate WindowsPath")
    def test_validate_pbix_reopen_flags_screenshot_signal(self) -> None:
        pbix = self.root / "report.pbix"
        screenshot = self.root / "probe.png"
        with zipfile.ZipFile(pbix, "w") as archive:
            archive.writestr(
                "Report/Layout", json.dumps({"sections": [{"visualContainers": [{}]}]}).encode("utf-16-le")
            )

        with (
            patch("tools.quality.os.name", "nt"),
            patch(
                "tools.quality._run_reopen_probe",
                return_value={
                    "opened": True,
                    "process_id": 1,
                    "process_name": "PBIDesktop",
                    "window_title": "report - Power BI Desktop",
                    "ui_text_count": 0,
                    "ui_text_matches": [],
                    "screenshot_path": str(screenshot),
                },
            ),
            patch(
                "tools.quality._analyze_reopen_screenshot",
                return_value={
                    "available": True,
                    "dark_pixel_ratio": 0.5,
                    "teal_pixel_ratio": 0.01,
                    "fix_this_like": True,
                },
            ),
            patch("tools.quality._ocr_reopen_screenshot", return_value={"available": True, "matches": []}),
        ):
            result = pbi_validate_pbix_reopen_tool(
                pbix_path=str(pbix), timeout_seconds=10, screenshot_path=str(screenshot)
            )

        self.assertFalse(result["valid"], result)
        self.assertIn("screenshot_fix_this_like_regions", {item["type"] for item in result["issues"]})

    @unittest.skipUnless(os.name == "nt", "patching os.name to 'nt' makes pathlib instantiate WindowsPath")
    def test_validate_pbix_reopen_flags_windows_ocr_signal(self) -> None:
        pbix = self.root / "report.pbix"
        screenshot = self.root / "probe.png"
        with zipfile.ZipFile(pbix, "w") as archive:
            archive.writestr(
                "Report/Layout", json.dumps({"sections": [{"visualContainers": [{}]}]}).encode("utf-16-le")
            )

        with (
            patch("tools.quality.os.name", "nt"),
            patch(
                "tools.quality._run_reopen_probe",
                return_value={
                    "opened": True,
                    "process_id": 1,
                    "process_name": "PBIDesktop",
                    "window_title": "report - Power BI Desktop",
                    "ui_text_count": 0,
                    "ui_text_matches": [],
                    "screenshot_path": str(screenshot),
                },
            ),
            patch("tools.quality._analyze_reopen_screenshot", return_value={"available": True, "fix_this_like": False}),
            patch(
                "tools.quality._ocr_reopen_screenshot",
                return_value={"available": True, "matches": ["Fix this"], "text": "Fix this"},
            ),
        ):
            result = pbi_validate_pbix_reopen_tool(
                pbix_path=str(pbix),
                timeout_seconds=10,
                screenshot_path=str(screenshot),
                use_windows_ocr=True,
            )

        self.assertFalse(result["valid"], result)
        self.assertIn("windows_ocr_fix_this_signal", {item["type"] for item in result["issues"]})

    def test_tool_categories(self) -> None:
        for name in [
            "pbi_audit_model",
            "pbi_lint_dax",
            "pbi_detect_name_collisions",
            "pbi_detect_dirty_dates",
            "pbi_validate_relationship_plan",
            "pbi_validate_filter_expression",
            "pbi_detect_empty_visuals",
            "pbi_generate_measure_tests",
            "pbi_validate_pbix_persistence",
            "pbi_lint_report_layout",
            "pbi_validate_visual_bindings",
            "pbi_score_dashboard",
            "pbi_run_scenario",
            "pbi_compare_report_versions",
        ]:
            self.assertEqual(tool_category(name), "read")
        self.assertEqual(tool_category("pbi_export_validation_report"), "write")
        # pbi_validate_pbix_reopen captures screen + may close PBI Desktop —
        # treat as write so --readonly blocks it.
        self.assertEqual(tool_category("pbi_validate_pbix_reopen"), "write")


if __name__ == "__main__":
    unittest.main(verbosity=2)

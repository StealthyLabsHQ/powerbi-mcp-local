"""Offline tests for v0.10.4 analysis tools."""

from __future__ import annotations

import json
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

from tools.quality import (
    LAYOUT_RELATIVE_PATH,
    pbi_detect_circular_dependencies_tool,
    pbi_detect_missing_visuals_tool,
    pbi_export_correction_report_tool,
    pbi_score_rubric_tool,
    pbi_validate_power_query_steps_tool,
    pbi_validate_star_schema_tool,
)


class _StubManager:
    pass


def _patch_snapshot(snapshot: dict) -> object:
    return patch("tools.quality._model_snapshot", return_value=snapshot)


class StarSchemaTests(unittest.TestCase):
    def test_clean_star_schema(self) -> None:
        snapshot = {
            "tables": [
                {"name": "FactSales", "is_hidden": False, "columns": []},
                {"name": "DimDate", "is_hidden": False, "columns": []},
                {"name": "DimProduct", "is_hidden": False, "columns": []},
            ],
            "relationships": [
                {
                    "from_table": "FactSales",
                    "from_column": "DateKey",
                    "to_table": "DimDate",
                    "to_column": "DateKey",
                    "direction": "OneDirection",
                },
                {
                    "from_table": "FactSales",
                    "from_column": "ProductKey",
                    "to_table": "DimProduct",
                    "to_column": "ProductKey",
                    "direction": "OneDirection",
                },
            ],
            "measures": [],
        }
        with _patch_snapshot(snapshot):
            result = pbi_validate_star_schema_tool(_StubManager())
        self.assertTrue(result["is_star_schema"])
        self.assertEqual(set(result["fact_tables"]), {"FactSales"})
        self.assertEqual(set(result["dim_tables"]), {"DimDate", "DimProduct"})
        self.assertEqual(result["issue_count"], 0)

    def test_snowflake_violation(self) -> None:
        snapshot = {
            "tables": [
                {"name": "FactSales", "is_hidden": False, "columns": []},
                {"name": "DimProduct", "is_hidden": False, "columns": []},
                {"name": "DimCategory", "is_hidden": False, "columns": []},
            ],
            "relationships": [
                {
                    "from_table": "FactSales",
                    "from_column": "PK",
                    "to_table": "DimProduct",
                    "to_column": "PK",
                    "direction": "OneDirection",
                },
                {
                    "from_table": "DimProduct",
                    "from_column": "CK",
                    "to_table": "DimCategory",
                    "to_column": "CK",
                    "direction": "OneDirection",
                },
            ],
            "measures": [],
        }
        with _patch_snapshot(snapshot):
            result = pbi_validate_star_schema_tool(_StubManager())
        self.assertFalse(result["is_star_schema"])
        types = {issue["type"] for issue in result["issues"]}
        self.assertIn("snowflake_dim_to_dim", types)


class CircularDependencyTests(unittest.TestCase):
    def test_no_cycles(self) -> None:
        snapshot = {
            "tables": [],
            "relationships": [],
            "measures": [
                {"name": "A", "table": "F", "expression": "SUM(F[X])"},
                {"name": "B", "table": "F", "expression": "[A] * 2"},
            ],
        }
        with _patch_snapshot(snapshot):
            result = pbi_detect_circular_dependencies_tool(_StubManager())
        self.assertTrue(result["valid"])
        self.assertEqual(result["cycle_count"], 0)
        self.assertEqual(result["self_reference_count"], 0)

    def test_simple_cycle(self) -> None:
        snapshot = {
            "tables": [],
            "relationships": [],
            "measures": [
                {"name": "A", "table": "F", "expression": "[B] + 1"},
                {"name": "B", "table": "F", "expression": "[A] * 2"},
            ],
        }
        with _patch_snapshot(snapshot):
            result = pbi_detect_circular_dependencies_tool(_StubManager())
        self.assertFalse(result["valid"])
        self.assertGreaterEqual(result["cycle_count"], 1)

    def test_self_reference(self) -> None:
        snapshot = {
            "tables": [],
            "relationships": [],
            "measures": [{"name": "Recursive", "table": "F", "expression": "[Recursive] + 1"}],
        }
        with _patch_snapshot(snapshot):
            result = pbi_detect_circular_dependencies_tool(_StubManager())
        self.assertFalse(result["valid"])
        self.assertEqual(result["self_reference_count"], 1)


class PowerQueryStepsTests(unittest.TestCase):
    def test_substring_match(self) -> None:
        m = """
let
    Source = Csv.Document(File.Contents("data.csv")),
    PaddedZip = Table.TransformColumns(Source, {"Zip", each Text.PadStart(_, 5, "0")}),
    Filtered = Table.SelectRows(PaddedZip, each [CustomerId] <> null)
in
    Filtered
"""
        with patch("tools.power_query.pbi_get_power_query_tool", return_value={"expression": m}):
            result = pbi_validate_power_query_steps_tool(
                _StubManager(),
                table="Sales",
                expected_steps=["Text.PadStart", "[CustomerId] <> null", "re:Table\\.\\w+"],
            )
        self.assertTrue(result["valid"])
        self.assertEqual(result["found_count"], 3)

    def test_missing_step(self) -> None:
        with patch(
            "tools.power_query.pbi_get_power_query_tool", return_value={"expression": "let Source = X in Source"}
        ):
            result = pbi_validate_power_query_steps_tool(
                _StubManager(),
                table="Sales",
                expected_steps=["Text.PadStart"],
            )
        self.assertFalse(result["valid"])
        self.assertEqual(result["missing_count"], 1)


class MissingVisualsTests(unittest.TestCase):
    def _layout(self, folder: Path, visuals: list[dict]) -> None:
        path = folder / LAYOUT_RELATIVE_PATH
        path.parent.mkdir(parents=True, exist_ok=True)
        layout = {
            "sections": [{"name": "S1", "displayName": "P1", "width": 1280, "height": 720, "visualContainers": visuals}]
        }
        path.write_bytes(json.dumps(layout).encode("utf-16-le"))

    def _container(self, visual_type: str, fields: list[str]) -> dict:
        return {
            "x": 0,
            "y": 0,
            "width": 200,
            "height": 200,
            "config": json.dumps(
                {
                    "singleVisual": {
                        "visualType": visual_type,
                        "prototypeQuery": {"Select": [{"Name": f} for f in fields]},
                    }
                }
            ),
        }

    def test_satisfied(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            self._layout(Path(tmp), [self._container("map", ["City"]), self._container("card", ["Total"])])
            with patch("tools.quality.resolve_local_path", return_value=Path(tmp)):
                result = pbi_detect_missing_visuals_tool(
                    tmp,
                    page="P1",
                    requirements=[
                        {"visual_type": "map", "contains_field": "City"},
                        {"visual_type": "card"},
                    ],
                )
        self.assertTrue(result["valid"])
        self.assertEqual(result["found_count"], 2)

    def test_missing(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            self._layout(Path(tmp), [self._container("card", ["Total"])])
            with patch("tools.quality.resolve_local_path", return_value=Path(tmp)):
                result = pbi_detect_missing_visuals_tool(
                    tmp,
                    page="P1",
                    requirements=[{"visual_type": "map", "label": "Carte géographique"}],
                )
        self.assertFalse(result["valid"])
        self.assertEqual(result["missing_count"], 1)


class RubricTests(unittest.TestCase):
    def test_aggregate(self) -> None:
        snapshot = {
            "tables": [
                {"name": "FactSales", "is_hidden": False, "columns": []},
                {"name": "DimDate", "is_hidden": False, "columns": []},
            ],
            "relationships": [
                {
                    "from_table": "FactSales",
                    "to_table": "DimDate",
                    "from_column": "K",
                    "to_column": "K",
                    "direction": "OneDirection",
                },
            ],
            "measures": [{"name": "Total", "table": "FactSales", "expression": "SUM(FactSales[X])"}],
        }
        with _patch_snapshot(snapshot):
            result = pbi_score_rubric_tool(
                _StubManager(),
                criteria=[
                    {"id": "star", "label": "Star schema", "check": "star_schema", "weight": 2.0},
                    {"id": "no_loop", "label": "No cycles", "check": "no_circular_deps", "weight": 1.0},
                    {
                        "id": "has_total",
                        "label": "Total measure",
                        "check": "measure_exists",
                        "weight": 1.0,
                        "params": {"name": "Total"},
                    },
                    {
                        "id": "missing",
                        "label": "Missing measure",
                        "check": "measure_exists",
                        "weight": 1.0,
                        "params": {"name": "Nope"},
                    },
                ],
            )
        self.assertEqual(result["passed_count"], 3)
        self.assertEqual(result["criterion_count"], 4)
        self.assertAlmostEqual(result["score"], 4.0 / 5.0, places=4)


class CorrectionReportTests(unittest.TestCase):
    def test_writes_markdown(self) -> None:
        snapshot = {
            "tables": [
                {"name": "FactSales", "is_hidden": False, "columns": []},
                {"name": "DimDate", "is_hidden": False, "columns": []},
            ],
            "relationships": [
                {
                    "from_table": "FactSales",
                    "to_table": "DimDate",
                    "from_column": "K",
                    "to_column": "K",
                    "direction": "OneDirection",
                },
            ],
            "measures": [{"name": "Total", "table": "FactSales", "expression": "SUM(FactSales[X])"}],
        }
        with tempfile.TemporaryDirectory() as tmp:
            out = Path(tmp) / "report.md"
            with (
                _patch_snapshot(snapshot),
                patch("tools.quality._duplicate_relationship_key_issues", return_value=[]),
                patch("tools.quality.resolve_local_path", return_value=out),
            ):
                result = pbi_export_correction_report_tool(_StubManager(), output_path=str(out))
            self.assertTrue(out.exists())
            content = out.read_text(encoding="utf-8")
            self.assertIn("# Power BI correction report", content)
            self.assertIn("Star schema", content)
            self.assertIn("Circular dependencies", content)
        self.assertEqual(result["audit_issue_count"], 0)


if __name__ == "__main__":
    unittest.main()

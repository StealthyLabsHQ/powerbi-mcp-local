"""Standalone tests for persistent PBIX report builder helpers."""

from __future__ import annotations

import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

from pbi_connection import PowerBIConfigurationError
from security import SECURITY, tool_category
from tools.persistent_report import pbi_create_persistent_report_tool


class FakeBuilder:
    instances: list[FakeBuilder] = []

    def __init__(self) -> None:
        self.tables = []
        self.measures = []
        self.relationships = []
        self.pages = []
        FakeBuilder.instances.append(self)

    def add_table(self, name, columns, rows, source_csv=None, source_db=None, mode="import"):
        self.tables.append(
            {
                "name": name,
                "columns": columns,
                "rows": rows,
                "source_csv": source_csv,
                "source_db": source_db,
                "mode": mode,
            }
        )

    def add_measure(self, table, name, expression):
        self.measures.append({"table": table, "name": name, "expression": expression})
        self._measures = self.measures

    def add_relationship(self, from_table, from_column, to_table, to_column):
        self.relationships.append(
            {
                "from_table": from_table,
                "from_column": from_column,
                "to_table": to_table,
                "to_column": to_column,
            }
        )

    def add_page(self, name, visuals):
        self.pages.append({"name": name, "visuals": visuals})

    def _pre_build_checks(self):
        return []

    def save(self, path):
        Path(path).write_bytes(b"pbix")

    def validate(self):
        return []


class PersistentReportTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp_dir = tempfile.TemporaryDirectory()
        self.root = Path(self.temp_dir.name)
        self.previous_allowed = [str(item) for item in SECURITY.allowed_base_dirs()]
        SECURITY.configure_allowed_dirs([str(self.root)])
        FakeBuilder.instances = []

    def tearDown(self) -> None:
        SECURITY.configure_allowed_dirs(self.previous_allowed)
        self.temp_dir.cleanup()

    def test_create_persistent_report_writes_pbix(self) -> None:
        output = self.root / "stock.pbix"

        with patch("tools.persistent_report._load_pbix_builder", return_value=FakeBuilder):
            result = pbi_create_persistent_report_tool(
                str(output),
                tables=[
                    {
                        "name": "Stock",
                        "columns": [{"name": "SKU", "data_type": "string"}, {"name": "Qty", "data_type": "int64"}],
                        "rows": [{"SKU": "A-001", "Qty": 12}],
                    }
                ],
                measures=[
                    {"table": "Stock", "name": "Total Qty", "expression": "SUM(Stock[Qty])", "format_string": "#,##0"}
                ],
                pages=[
                    {
                        "name": "Overview",
                        "visuals": [{"type": "card", "config": {"measure": "Total Qty"}}],
                    }
                ],
            )

        self.assertTrue(result["ok"], result)
        # Compare resolved paths so Windows short/long-form temp dirs (e.g.
        # ``RUNNER~1`` vs ``runneradmin`` on GitHub Actions) don't fail the
        # equality check.
        self.assertEqual(Path(result["output_path"]).resolve(), output.resolve())
        self.assertEqual(result["table_count"], 1)
        self.assertEqual(result["measure_count"], 1)
        self.assertEqual(result["page_count"], 1)
        self.assertTrue(output.exists())
        self.assertEqual(FakeBuilder.instances[0].measures[0]["name"], "Total Qty")
        self.assertEqual(FakeBuilder.instances[0].measures[0]["format_string"], "#,##0")

    def test_rejects_non_pbix_output(self) -> None:
        with self.assertRaises(Exception):
            pbi_create_persistent_report_tool(
                str(self.root / "stock.txt"),
                tables=[{"name": "Stock", "columns": [{"name": "SKU", "data_type": "string"}]}],
            )

    def test_missing_optional_builder_raises_configuration_error(self) -> None:
        with patch("tools.persistent_report._load_pbix_builder", side_effect=PowerBIConfigurationError("missing")):
            with self.assertRaises(PowerBIConfigurationError):
                pbi_create_persistent_report_tool(
                    str(self.root / "stock.pbix"),
                    tables=[{"name": "Stock", "columns": [{"name": "SKU", "data_type": "string"}]}],
                )

    def test_tool_category(self) -> None:
        self.assertEqual(tool_category("pbi_create_persistent_report"), "write")


if __name__ == "__main__":
    unittest.main(verbosity=2)

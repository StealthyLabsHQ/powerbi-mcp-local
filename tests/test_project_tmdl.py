"""Standalone tests for PBIP/TMDL project helpers."""

from __future__ import annotations

import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

from pbi_connection import PowerBIValidationError
from security import SECURITY, tool_category
from tools.project import (
    pbi_list_tmdl_files_tool,
    pbi_patch_tmdl_measure_tool,
    pbi_read_tmdl_file_tool,
    pbi_write_tmdl_file_tool,
)


class ProjectTmdlToolTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp_dir = tempfile.TemporaryDirectory()
        self.root = Path(self.temp_dir.name)
        self.project = self.root / "Demo"
        self.definition = self.project / "Demo.SemanticModel" / "definition"
        self.definition.mkdir(parents=True)
        (self.project / "Demo.pbip").write_text("{}", encoding="utf-8")
        (self.definition / "model.tmdl").write_text("model Model\n", encoding="utf-8")
        tables = self.definition / "tables"
        tables.mkdir()
        (tables / "Sales.tmdl").write_text("table Sales\n\tcolumn Amount\n", encoding="utf-8")
        self.previous_allowed = [str(item) for item in SECURITY.allowed_base_dirs()]
        SECURITY.configure_allowed_dirs([str(self.root)])

    def tearDown(self) -> None:
        SECURITY.configure_allowed_dirs(self.previous_allowed)
        self.temp_dir.cleanup()

    def test_list_tmdl_files_from_pbip_path(self) -> None:
        result = pbi_list_tmdl_files_tool(str(self.project / "Demo.pbip"))

        self.assertTrue(result["ok"], result)
        self.assertEqual(result["file_count"], 2)
        self.assertIn("tables/Sales.tmdl", {item["relative_file"] for item in result["files"]})
        self.assertIn("tab_indentation", {item["issue"] for item in result["issues"]})

    def test_read_tmdl_file(self) -> None:
        result = pbi_read_tmdl_file_tool(str(self.project), "tables/Sales.tmdl")

        self.assertTrue(result["ok"], result)
        self.assertEqual(result["relative_file"], "tables/Sales.tmdl")
        self.assertIn("table Sales", result["content"])

    def test_write_tmdl_file_updates_existing(self) -> None:
        result = pbi_write_tmdl_file_tool(str(self.project), "tables/Sales.tmdl", "table Sales\n    column Amount\n")

        self.assertTrue(result["ok"], result)
        self.assertEqual(result["action"], "updated")
        self.assertIn("column Amount", (self.definition / "tables" / "Sales.tmdl").read_text(encoding="utf-8"))

    def test_write_tmdl_file_creates_when_allowed(self) -> None:
        result = pbi_write_tmdl_file_tool(str(self.project), "tables/Customers.tmdl", "table Customers\n", create=True)

        self.assertTrue(result["ok"], result)
        self.assertEqual(result["action"], "created")
        self.assertTrue((self.definition / "tables" / "Customers.tmdl").exists())

    def test_write_tmdl_file_blocks_traversal(self) -> None:
        with self.assertRaises(PowerBIValidationError):
            pbi_write_tmdl_file_tool(str(self.project), "../outside.tmdl", "table Bad\n", create=True)

    def test_patch_tmdl_measure_creates_measure_block(self) -> None:
        result = pbi_patch_tmdl_measure_tool(
            str(self.project),
            "tables/Sales.tmdl",
            "Total Sales",
            "SUM(Sales[Amount])",
            format_string="$#,0",
            display_folder="Core",
        )

        self.assertTrue(result["ok"], result)
        self.assertEqual(result["action"], "created")
        content = (self.definition / "tables" / "Sales.tmdl").read_text(encoding="utf-8")
        self.assertIn("measure 'Total Sales' = SUM(Sales[Amount])", content)
        self.assertIn("formatString: '$#,0'", content)
        self.assertIn("displayFolder: 'Core'", content)

    def test_patch_tmdl_measure_updates_existing_measure_block(self) -> None:
        pbi_patch_tmdl_measure_tool(str(self.project), "tables/Sales.tmdl", "Total Sales", "SUM(Sales[Amount])")

        result = pbi_patch_tmdl_measure_tool(str(self.project), "tables/Sales.tmdl", "Total Sales", "SUM(Sales[Net])")

        self.assertTrue(result["ok"], result)
        self.assertEqual(result["action"], "updated")
        content = (self.definition / "tables" / "Sales.tmdl").read_text(encoding="utf-8")
        self.assertIn("SUM(Sales[Net])", content)
        self.assertNotIn("SUM(Sales[Amount])", content)

    def test_tool_categories(self) -> None:
        self.assertEqual(tool_category("pbi_list_tmdl_files"), "read")
        self.assertEqual(tool_category("pbi_read_tmdl_file"), "read")
        self.assertEqual(tool_category("pbi_write_tmdl_file"), "write")
        self.assertEqual(tool_category("pbi_patch_tmdl_measure"), "write")


if __name__ == "__main__":
    unittest.main(verbosity=2)

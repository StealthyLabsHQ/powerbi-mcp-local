"""Standalone tests for workflow tools."""

from __future__ import annotations

import sys
import unittest
from pathlib import Path
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

from pbi_connection import PowerBIValidationError
from security import tool_category
from tools.workflows import (
    pbi_excel_import_workflow_tool,
    pbi_measure_workflow_tool,
    pbi_model_audit_workflow_tool,
)


class WorkflowToolTests(unittest.TestCase):
    def test_excel_import_workflow_dry_run_does_not_apply(self) -> None:
        with (
            patch(
                "tools.workflows.excel_workbook_info_tool",
                return_value={
                    "ok": True,
                    "file_path": "sales.xlsx",
                    "sheets": [{"name": "Sales", "has_data": True, "rows": 3}],
                },
            ),
            patch(
                "tools.workflows.pbi_list_tables_tool",
                return_value={"ok": True, "tables": [{"name": "Sales"}]},
            ),
            patch("tools.workflows.pbi_import_excel_workbook_tool") as importer,
        ):
            result = pbi_excel_import_workflow_tool(object(), excel_path="sales.xlsx")

        self.assertTrue(result["ok"], result)
        self.assertTrue(result["needs_apply"])
        self.assertEqual(result["sheet_table_map"], {"Sales": "Sales"})
        importer.assert_not_called()

    def test_excel_import_workflow_apply_blocks_missing_table(self) -> None:
        with (
            patch(
                "tools.workflows.excel_workbook_info_tool",
                return_value={
                    "ok": True,
                    "file_path": "sales.xlsx",
                    "sheets": [{"name": "Sales", "has_data": True}],
                },
            ),
            patch(
                "tools.workflows.pbi_list_tables_tool",
                return_value={"ok": True, "tables": []},
            ),
        ):
            with self.assertRaises(PowerBIValidationError):
                pbi_excel_import_workflow_tool(
                    object(),
                    excel_path="sales.xlsx",
                    sheet_table_map={"Sales": "Sales"},
                    apply=True,
                )

    def test_measure_workflow_dry_run_blocks_existing_without_overwrite(self) -> None:
        with (
            patch(
                "tools.workflows.pbi_list_tables_tool",
                return_value={"ok": True, "tables": [{"name": "Sales"}]},
            ),
            patch(
                "tools.workflows.pbi_model_info_tool",
                return_value={"ok": True, "measures": [{"table": "Sales", "name": "Total Sales"}]},
            ),
            patch(
                "tools.workflows.pbi_validate_dax_tool",
                return_value={"ok": True, "valid": True},
            ),
            patch("tools.workflows.pbi_create_measures_tool") as creator,
        ):
            result = pbi_measure_workflow_tool(
                object(),
                table="Sales",
                measures=[{"name": "Total Sales", "expression": "SUM(Sales[Amount])"}],
                overwrite=False,
            )

        self.assertTrue(result["ok"], result)
        self.assertFalse(result["validation"]["ready"])
        creator.assert_not_called()

    def test_measure_workflow_apply_creates_after_validation(self) -> None:
        with (
            patch(
                "tools.workflows.pbi_list_tables_tool",
                return_value={"ok": True, "tables": [{"name": "Sales"}]},
            ),
            patch(
                "tools.workflows.pbi_model_info_tool",
                return_value={"ok": True, "measures": []},
            ),
            patch(
                "tools.workflows.pbi_validate_dax_tool",
                return_value={"ok": True, "valid": True},
            ),
            patch(
                "tools.workflows.pbi_create_measures_tool",
                return_value={"ok": True, "created": 1, "updated": 0, "failed": 0},
            ) as creator,
        ):
            result = pbi_measure_workflow_tool(
                object(),
                table="Sales",
                measures=[{"name": "Total Sales", "expression": "SUM(Sales[Amount])"}],
                apply=True,
            )

        self.assertTrue(result["ok"], result)
        self.assertFalse(result["needs_apply"])
        creator.assert_called_once()

    def test_model_audit_workflow_returns_recommendations(self) -> None:
        with (
            patch(
                "tools.workflows.pbi_model_info_tool",
                return_value={"ok": True, "tables": [{"name": "Sales"}], "measures": [], "relationships": []},
            ),
            patch(
                "tools.workflows.pbi_validate_model_tool",
                return_value={"ok": True, "issues": [], "warnings": [], "issue_count": 0, "warning_count": 0},
            ),
            patch(
                "tools.workflows.pbi_measure_dependencies_tool",
                return_value={"ok": True, "truncated": False},
            ),
            patch(
                "tools.workflows.pbi_list_power_queries_tool",
                return_value={"ok": True, "queries": []},
            ),
        ):
            result = pbi_model_audit_workflow_tool(object())

        self.assertTrue(result["ok"], result)
        self.assertEqual(result["summary"]["table_count"], 1)
        self.assertTrue(result["recommendations"])

    def test_workflow_security_category_uses_apply_flag(self) -> None:
        self.assertEqual(tool_category("pbi_excel_import_workflow", {"apply": False}), "read")
        self.assertEqual(tool_category("pbi_excel_import_workflow", {"apply": True}), "write")
        self.assertEqual(tool_category("pbi_measure_workflow", {"apply": False}), "read")
        self.assertEqual(tool_category("pbi_measure_workflow", {"apply": True}), "write")


if __name__ == "__main__":
    unittest.main(verbosity=2)

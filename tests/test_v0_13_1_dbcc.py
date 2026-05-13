"""v0.13.1: DBCC string-store diagnostics + prime_string_store coverage.

Pins:
- The pre-build spec checker flags empty Import tables that declare a
  String column (the root cause of the user-reported "Database
  consistency checks (DBCC) failed while checking the string store"
  dialog after build).
- The static PBIX diagnoser surfaces missing/undersized DataModel parts.
- ``pbi_create_persistent_report`` + ``pbi_scaffold_pbix`` inject one
  sentinel row per affected table when ``prime_string_store=True``
  (default).
- The reopen probe signal list now matches the DBCC dialog.
"""

from __future__ import annotations

import json
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

from tools import (
    pbi_check_scaffold_spec_dbcc_risks_tool,
    pbi_diagnose_pbix_dbcc_tool,
)
from tools.dbcc import DBCC_KNOWN_SIGNALS, MIN_DATA_MODEL_BYTES
from tools.persistent_report import _SENTINEL_VALUES_BY_TYPE, _prime_string_store


class SpecRiskCheckerTests(unittest.TestCase):
    def test_empty_string_table_flagged(self) -> None:
        spec = [{"name": "Dim", "columns": [{"name": "Name", "data_type": "String"}], "rows": []}]
        result = pbi_check_scaffold_spec_dbcc_risks_tool(spec)
        self.assertFalse(result["valid"])
        self.assertEqual(result["issues"][0]["type"], "empty_string_table")
        self.assertIn("Name", result["issues"][0]["string_columns"])

    def test_empty_numeric_table_is_warning_not_error(self) -> None:
        spec = [{"name": "Facts", "columns": [{"name": "Amount", "data_type": "Decimal"}], "rows": []}]
        result = pbi_check_scaffold_spec_dbcc_risks_tool(spec)
        self.assertTrue(result["valid"])
        self.assertGreaterEqual(result["warning_count"], 1)

    def test_table_with_rows_not_flagged(self) -> None:
        spec = [
            {
                "name": "Dim",
                "columns": [{"name": "Name", "data_type": "String"}],
                "rows": [["x"]],
            }
        ]
        result = pbi_check_scaffold_spec_dbcc_risks_tool(spec)
        self.assertTrue(result["valid"])
        self.assertEqual(result["issue_count"], 0)

    def test_table_with_source_csv_not_flagged(self) -> None:
        spec = [
            {
                "name": "Dim",
                "columns": [{"name": "Name", "data_type": "String"}],
                "rows": [],
                "source_csv": "some.csv",
            }
        ]
        result = pbi_check_scaffold_spec_dbcc_risks_tool(spec)
        self.assertTrue(result["valid"])

    def test_directquery_table_not_flagged(self) -> None:
        spec = [
            {
                "name": "Dim",
                "columns": [{"name": "Name", "data_type": "String"}],
                "rows": [],
                "mode": "directquery",
            }
        ]
        result = pbi_check_scaffold_spec_dbcc_risks_tool(spec)
        self.assertTrue(result["valid"])

    def test_non_list_input_rejected(self) -> None:
        from pbi_connection import PowerBIValidationError

        with self.assertRaises(PowerBIValidationError):
            pbi_check_scaffold_spec_dbcc_risks_tool({"not": "a list"})

    def test_invalid_table_entry_flagged(self) -> None:
        result = pbi_check_scaffold_spec_dbcc_risks_tool(["string-not-dict"])
        self.assertFalse(result["valid"])
        self.assertEqual(result["issues"][0]["type"], "invalid_table_entry")


class PrimeStringStoreTests(unittest.TestCase):
    def test_sentinel_values_cover_every_type(self) -> None:
        for type_name in ("String", "Int64", "Double", "Decimal", "DateTime", "Boolean"):
            self.assertIn(type_name, _SENTINEL_VALUES_BY_TYPE)

    def test_priming_injects_one_row_for_string_table(self) -> None:
        table = {
            "name": "Dim",
            "columns": [
                {"name": "Id", "data_type": "Int64"},
                {"name": "Name", "data_type": "String"},
            ],
            "rows": [],
        }
        primed = _prime_string_store(table)
        self.assertTrue(primed.get("__primed_string_store__"))
        self.assertEqual(len(primed["rows"]), 1)
        self.assertEqual(primed["rows"][0][0], 0)  # Int64 sentinel
        self.assertEqual(primed["rows"][0][1], "_seed_")  # String sentinel

    def test_priming_no_op_when_rows_present(self) -> None:
        table = {
            "name": "Dim",
            "columns": [{"name": "Name", "data_type": "String"}],
            "rows": [["real-value"]],
        }
        self.assertIs(_prime_string_store(table), table)

    def test_priming_no_op_for_numeric_only_table(self) -> None:
        table = {
            "name": "Facts",
            "columns": [{"name": "Amount", "data_type": "Decimal"}],
            "rows": [],
        }
        primed = _prime_string_store(table)
        self.assertNotIn("__primed_string_store__", primed)
        self.assertEqual(primed.get("rows", []), [])

    def test_priming_no_op_when_source_csv(self) -> None:
        table = {
            "name": "Dim",
            "columns": [{"name": "Name", "data_type": "String"}],
            "rows": [],
            "source_csv": "x.csv",
        }
        primed = _prime_string_store(table)
        self.assertNotIn("__primed_string_store__", primed)


class StaticPbixDiagnoserTests(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp = tempfile.TemporaryDirectory()
        self.root = Path(self.tmp.name)
        from security import SECURITY, configure_allowed_dirs

        self.previous_allowed = [str(item) for item in SECURITY.allowed_base_dirs()]
        configure_allowed_dirs([str(self.root)])
        SECURITY.policy(reload=True, cwd=self.root)

    def tearDown(self) -> None:
        from security import SECURITY, configure_allowed_dirs

        configure_allowed_dirs(self.previous_allowed)
        SECURITY.policy(reload=True, cwd=Path.cwd())
        self.tmp.cleanup()

    def _make_pbix(self, name: str, parts: dict[str, bytes]) -> Path:
        path = self.root / name
        with zipfile.ZipFile(path, "w") as zf:
            for part_name, data in parts.items():
                zf.writestr(part_name, data)
        return path

    def test_non_zip_path_is_invalid(self) -> None:
        path = self.root / "broken.pbix"
        path.write_bytes(b"not-a-zip")
        result = pbi_diagnose_pbix_dbcc_tool(str(path))
        self.assertFalse(result["valid"])
        self.assertEqual(result["issues"][0]["type"], "pbix_not_zip")

    def test_pbix_without_datamodel_flagged(self) -> None:
        path = self._make_pbix(
            "no_model.pbix",
            {"Report/Layout": json.dumps({"sections": []}).encode("utf-16-le")},
        )
        result = pbi_diagnose_pbix_dbcc_tool(str(path))
        types = {issue["type"] for issue in result["issues"]}
        self.assertIn("no_data_model", types)

    def test_pbix_with_undersized_datamodel_flagged(self) -> None:
        path = self._make_pbix(
            "tiny.pbix",
            {
                "DataModel": b"x" * 16,
                "Report/Layout": json.dumps({"sections": []}).encode("utf-16-le"),
            },
        )
        result = pbi_diagnose_pbix_dbcc_tool(str(path))
        types = {issue["type"] for issue in result["issues"]}
        self.assertIn("undersized_data_model", types)

    def test_pbix_with_large_datamodel_passes(self) -> None:
        path = self._make_pbix(
            "ok.pbix",
            {
                "DataModel": b"x" * (MIN_DATA_MODEL_BYTES + 1024),
                "Report/Layout": json.dumps({"sections": []}).encode("utf-16-le"),
                "Connections": b"{}",
                "Metadata": b"{}",
            },
        )
        result = pbi_diagnose_pbix_dbcc_tool(str(path))
        self.assertTrue(result["valid"], result)

    def test_known_signals_are_surfaced(self) -> None:
        # The probe pattern list is documented in the tool response so a
        # client can cross-reference reopen-probe matches with DBCC.
        path = self._make_pbix(
            "ok.pbix",
            {"DataModel": b"x" * (MIN_DATA_MODEL_BYTES + 1024)},
        )
        result = pbi_diagnose_pbix_dbcc_tool(str(path))
        self.assertEqual(tuple(result["known_signals"]), DBCC_KNOWN_SIGNALS)


class ReopenProbeSignalTests(unittest.TestCase):
    def test_quality_module_lists_dbcc_signals(self) -> None:
        # The reopen probe matches Power BI Desktop's modal dialog text
        # via UIAutomation. The signal list is embedded in the PS script.
        from tools import quality

        source = Path(quality.__file__).read_text(encoding="utf-8")
        for needle in (
            "Database consistency checks",
            "DBCC",
            "Vertipaq",
            "string store",
            "An error occurred while loading",
        ):
            self.assertIn(needle, source, f"reopen probe missing signal: {needle}")


if __name__ == "__main__":
    unittest.main(verbosity=2)

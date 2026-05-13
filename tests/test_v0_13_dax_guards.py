"""v0.13: DAX guard regressions.

Pins the DMV/debug blocklist for ``pbi_execute_dax``: ``INFO.*`` and
``EVALUATEANDLOG`` are added in v0.13 because both expose server-side
metadata or write side-channel debug output. Block them by default so a
crafted LLM-issued DAX cannot surface metadata or hot-path tracing.
"""

from __future__ import annotations

import os
import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

from pbi_connection import PowerBIValidationError
from tools.query import _validate_dax_query


class DAXGuardTests(unittest.TestCase):
    def setUp(self) -> None:
        # Make sure the opt-in env-var isn't leaking from another test.
        self._previous = os.environ.pop("PBI_MCP_ALLOW_DMV", None)

    def tearDown(self) -> None:
        if self._previous is not None:
            os.environ["PBI_MCP_ALLOW_DMV"] = self._previous

    def test_legitimate_dax_passes(self) -> None:
        _validate_dax_query("EVALUATE Sales")  # query-only kind, but execute_dax accepts EVALUATE
        _validate_dax_query("CALCULATE([Total Sales])")

    def test_system_query_blocked(self) -> None:
        with self.assertRaises(PowerBIValidationError):
            _validate_dax_query("SELECT * FROM $SYSTEM.DISCOVER_KEYWORDS")

    def test_discover_blocked(self) -> None:
        with self.assertRaises(PowerBIValidationError):
            _validate_dax_query("SELECT * FROM $SYSTEM.DISCOVER_OBJECTS")

    def test_dbschema_blocked(self) -> None:
        with self.assertRaises(PowerBIValidationError):
            _validate_dax_query("SELECT * FROM DBSCHEMA_TABLES")

    def test_mdschema_blocked(self) -> None:
        with self.assertRaises(PowerBIValidationError):
            _validate_dax_query("SELECT * FROM MDSCHEMA_PROPERTIES")

    def test_info_function_blocked(self) -> None:
        # v0.13: INFO.* surfaces server metadata, treat like a DMV.
        with self.assertRaises(PowerBIValidationError):
            _validate_dax_query("EVALUATE INFO.TABLES()")

    def test_info_function_blocked_case_insensitive(self) -> None:
        with self.assertRaises(PowerBIValidationError):
            _validate_dax_query("evaluate info.measures()")

    def test_evaluateandlog_blocked(self) -> None:
        # v0.13: EVALUATEANDLOG writes side-channel debug output.
        with self.assertRaises(PowerBIValidationError):
            _validate_dax_query("EVALUATE EVALUATEANDLOG(SUM(Sales[Amount]))")

    def test_evaluateandlog_inline_blocked(self) -> None:
        with self.assertRaises(PowerBIValidationError):
            _validate_dax_query("EVALUATE ADDCOLUMNS('Date', \"x\", EVALUATEANDLOG([Total Sales]))")

    def test_normal_function_named_like_info_is_not_blocked(self) -> None:
        # "INFO" alone (no dot-call) is not the blocked pattern. The guard
        # specifically targets the dotted callable form ``INFO.NAME(``.
        _validate_dax_query('EVALUATE FILTER(Customer, Customer[INFO] = "x")')

    def test_opt_in_env_var_bypasses_guard(self) -> None:
        os.environ["PBI_MCP_ALLOW_DMV"] = "1"
        try:
            _validate_dax_query("EVALUATE INFO.TABLES()")
        finally:
            os.environ.pop("PBI_MCP_ALLOW_DMV", None)


if __name__ == "__main__":
    unittest.main(verbosity=2)

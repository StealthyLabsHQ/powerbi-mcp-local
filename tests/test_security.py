"""Security regression tests for the Power BI MCP server."""

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

from pbi_connection import PowerBIValidationError
from security import SECURITY, SecurityManager, SecurityPolicyError, inspect_excel_archive, resolve_local_path, validate_measure_name
from tools.model import pbi_export_model_tool
from tools.power_query import _validate_m_expression
from tools.query import _validate_dax_query


class SecurityTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp_dir = tempfile.TemporaryDirectory()
        self.root = Path(self.temp_dir.name)
        self.outside_dir = tempfile.TemporaryDirectory()
        self.outside_root = Path(self.outside_dir.name)
        self.previous_allowed = [str(item) for item in SECURITY.allowed_base_dirs()]
        self.previous_policy = os.environ.get("PBI_MCP_SECURITY_POLICY")
        self.previous_readonly = os.environ.get("PBI_MCP_READONLY")
        SECURITY.configure_allowed_dirs([str(self.root)])
        SECURITY.set_runtime_readonly(False)
        SECURITY.policy(reload=True, cwd=self.root)

    def tearDown(self) -> None:
        self.temp_dir.cleanup()
        self.outside_dir.cleanup()
        SECURITY.configure_allowed_dirs(self.previous_allowed)
        SECURITY.set_runtime_readonly(False)
        if self.previous_policy is None:
            os.environ.pop("PBI_MCP_SECURITY_POLICY", None)
        else:
            os.environ["PBI_MCP_SECURITY_POLICY"] = self.previous_policy
        if self.previous_readonly is None:
            os.environ.pop("PBI_MCP_READONLY", None)
        else:
            os.environ["PBI_MCP_READONLY"] = self.previous_readonly
        SECURITY.policy(reload=True, cwd=Path.cwd())

    def test_path_traversal_and_symlink_blocked(self) -> None:
        inside = self.root / "safe.xlsx"
        inside.write_bytes(b"not-a-zip")
        self.assertEqual(
            resolve_local_path(str(inside), must_exist=True, allowed_extensions={".xlsx"}),
            inside.resolve(),
        )

        outside = self.outside_root / "outside.xlsx"
        outside.write_bytes(b"not-a-zip")
        with self.assertRaises(SecurityPolicyError):
            resolve_local_path(str(outside), must_exist=True, allowed_extensions={".xlsx"})
        with self.assertRaises(SecurityPolicyError):
            resolve_local_path("../escape.xlsx", must_exist=False, allowed_extensions={".xlsx"})

        link_path = self.root / "linked.xlsx"
        try:
            link_path.symlink_to(outside)
        except (NotImplementedError, OSError):
            self.skipTest("Symlinks are not available on this platform.")
        with self.assertRaises(SecurityPolicyError):
            resolve_local_path(str(link_path), must_exist=True, allowed_extensions={".xlsx"})

    def test_dmv_queries_blocked(self) -> None:
        with patch.dict(os.environ, {"PBI_MCP_ALLOW_DMV": "0"}, clear=False):
            with self.assertRaises(PowerBIValidationError):
                _validate_dax_query("EVALUATE $SYSTEM.TMSCHEMA_TABLES")

    def test_blocked_m_functions_rejected(self) -> None:
        with patch.dict(os.environ, {"PBI_MCP_ALLOW_EXTERNAL_M": "0"}, clear=False):
            with self.assertRaises(PowerBIValidationError):
                _validate_m_expression('let Source = Web.Contents("https://example.com") in Source')

    def test_measure_name_injection_rejected(self) -> None:
        for name in ('Bad[Measure]', 'Bad"Measure', "Bad'Measure", "Bad\nMeasure"):
            with self.subTest(name=name):
                with self.assertRaises(SecurityPolicyError):
                    validate_measure_name(name)

    def test_zip_bomb_protection(self) -> None:
        workbook = self.root / "bomb.xlsx"
        with zipfile.ZipFile(workbook, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=9) as archive:
            archive.writestr("[Content_Types].xml", "A" * 5000)
            archive.writestr("xl/workbook.xml", "B" * 5000)
        with self.assertRaises(SecurityPolicyError):
            inspect_excel_archive(str(workbook), max_ratio=5.0)

    def test_export_model_redaction(self) -> None:
        export_path = self.root / "model.json"
        snapshot = {
            "connection": {"database": "UnitTest"},
            "tables": [
                {
                    "name": "Secrets",
                    "columns": [
                        {
                            "name": "Conn",
                            "expression": 'Provider=SQLNCLI11;Server=demo;Password=hunter2;User Id=sa',
                        }
                    ],
                }
            ],
            "measures": [
                {
                    "name": "Exposure",
                    "table": "Secrets",
                    "expression": "token=abcd1234; pwd=unsafe",
                }
            ],
            "relationships": [],
        }
        with patch("tools.model.pbi_model_info_tool", return_value=snapshot):
            response = pbi_export_model_tool(object(), path=str(export_path))
        self.assertTrue(response["ok"], response)
        model = response["model"]
        self.assertIn("[REDACTED]", model["tables"][0]["columns"][0]["expression"])
        self.assertIn("[REDACTED]", model["measures"][0]["expression"])
        self.assertTrue(export_path.exists())

    def test_readonly_mode_blocks_writes(self) -> None:
        SECURITY.set_runtime_readonly(True)
        with self.assertRaises(SecurityPolicyError):
            SECURITY.validate_tool_call(
                "excel_write_cell",
                {"file_path": str(self.root / "book.xlsx"), "sheet": "Sheet1", "cell": "A1", "value": "blocked"},
            )

    def test_security_policy_enforcement(self) -> None:
        policy_path = self.root / "security_policy.json"
        policy_path.write_text(
            json.dumps(
                {
                    "deny_categories": ["write"],
                    "disabled_tools": ["excel_search"],
                    "max_dax_rows": 10,
                    "allowed_base_dirs": [str(self.root)],
                }
            ),
            encoding="utf-8",
        )
        manager = SecurityManager()
        manager.policy(reload=True, cwd=self.root)

        with self.assertRaises(SecurityPolicyError):
            manager.validate_tool_call(
                "excel_write_cell",
                {"file_path": str(self.root / "book.xlsx"), "sheet": "Sheet1", "cell": "A1", "value": "x"},
            )
        with self.assertRaises(SecurityPolicyError):
            manager.validate_tool_call(
                "excel_search",
                {"file_path": str(self.root / "book.xlsx"), "query": "Revenue"},
            )
        with self.assertRaises(SecurityPolicyError):
            manager.validate_tool_call(
                "pbi_execute_dax",
                {"query": 'EVALUATE ROW("Value", 1)', "max_rows": 11},
            )


class ErrorPayloadTests(unittest.TestCase):
    """error_payload preserves the full exception chain."""

    def test_python_cause_chain_is_preserved(self) -> None:
        from pbi_connection import error_payload

        try:
            try:
                raise ValueError("low-level boom")
            except ValueError as exc:
                raise RuntimeError("outer wrapper") from exc
        except RuntimeError as outer:
            payload = error_payload(outer)

        self.assertFalse(payload["ok"])
        self.assertEqual(payload["error"]["code"], "internal_error")
        # Message flattens the whole chain, so the underlying cause is visible.
        self.assertIn("outer wrapper", payload["error"]["message"])
        self.assertIn("low-level boom", payload["error"]["message"])
        # Structured chain in details has at least 2 links (outer + cause).
        chain = payload["error"]["details"]["cause_chain"]
        self.assertGreaterEqual(len(chain), 2)
        types = [link["type"] for link in chain]
        self.assertIn("RuntimeError", types)
        self.assertIn("ValueError", types)

    def test_dotnet_inner_exception_traversed(self) -> None:
        """Mimics the .NET exception shape (.InnerException) used by pythonnet."""
        from pbi_connection import error_payload

        class Inner(Exception):
            pass

        class Outer(Exception):
            def __init__(self, message: str, inner: Exception) -> None:
                super().__init__(message)
                self.InnerException = inner

        outer = Outer("save failed", Inner("constraint x"))
        payload = error_payload(outer)
        self.assertIn("save failed", payload["error"]["message"])
        self.assertIn("constraint x", payload["error"]["message"])
        chain_types = [link["type"] for link in payload["error"]["details"]["cause_chain"]]
        self.assertEqual(chain_types[:2], ["Outer", "Inner"])


class SymlinkParentRejectionTests(unittest.TestCase):
    """``resolve_local_path`` rejects paths whose ancestors are symlinks."""

    def setUp(self) -> None:
        self.temp_dir = tempfile.TemporaryDirectory()
        self.root = Path(self.temp_dir.name).resolve()
        self.previous_allowed = [str(item) for item in SECURITY.allowed_base_dirs()]
        SECURITY.configure_allowed_dirs([str(self.root)])

    def tearDown(self) -> None:
        SECURITY.configure_allowed_dirs(self.previous_allowed)
        self.temp_dir.cleanup()

    def test_rejects_symlinked_parent_directory(self) -> None:
        target = self.root / "real"
        target.mkdir()
        link = self.root / "via_symlink"
        try:
            link.symlink_to(target, target_is_directory=True)
        except (OSError, NotImplementedError):
            self.skipTest("symlink creation not permitted on this platform")

        suspect = link / "data.txt"
        with self.assertRaises(SecurityPolicyError) as ctx:
            resolve_local_path(str(suspect), must_exist=False)
        self.assertIn("Symlink", str(ctx.exception))


if __name__ == "__main__":
    unittest.main(verbosity=2)

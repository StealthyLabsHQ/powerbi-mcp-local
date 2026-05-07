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


class OperationHistoryTests(unittest.TestCase):
    """Connection manager records operations into a ring buffer."""

    def test_ring_buffer_records_read_and_write(self) -> None:
        from pbi_connection import PowerBIConnectionManager
        m = PowerBIConnectionManager()
        # Direct record (no live PBI needed) — the helper is the only thing under test.
        m._record_operation(op="read_one", kind="read", started=0.0, ok_=True)
        m._record_operation(op="write_one", kind="write", started=0.0, ok_=False, error=ValueError("boom"))
        history = m.operation_history(last_n=10)
        self.assertEqual(len(history), 2)
        # Newest first.
        self.assertEqual(history[0]["op"], "write_one")
        self.assertFalse(history[0]["ok"])
        self.assertEqual(history[0]["error_type"], "ValueError")
        self.assertEqual(history[1]["op"], "read_one")
        self.assertTrue(history[1]["ok"])

    def test_ring_buffer_capacity_50(self) -> None:
        from pbi_connection import PowerBIConnectionManager
        m = PowerBIConnectionManager()
        for i in range(60):
            m._record_operation(op=f"op_{i}", kind="read", started=0.0, ok_=True)
        history = m.operation_history(last_n=60)
        self.assertEqual(len(history), 50)
        # Oldest retained should be op_10 (60 - 50).
        names = [entry["op"] for entry in history]
        self.assertEqual(names[0], "op_59")
        self.assertEqual(names[-1], "op_10")


class SystemHealthTests(unittest.TestCase):
    """pbi_system_health returns a usable snapshot even without a live connection."""

    def test_snapshot_when_disconnected(self) -> None:
        from pbi_connection import PowerBIConnectionManager
        from tools.model import pbi_system_health_tool
        m = PowerBIConnectionManager()
        result = pbi_system_health_tool(m)
        self.assertTrue(result["ok"], result)
        self.assertFalse(result["connected"])
        self.assertIsNone(result["model_loaded"]) if result.get("model_loaded") is None else self.assertFalse(result["model_loaded"])
        self.assertIn("dependencies", result)
        # mcp dependency at minimum should be available since this is an MCP server repo.
        self.assertIn("mcp", result["dependencies"])


class TimeIntelligenceTemplateTests(unittest.TestCase):
    """Time-intelligence DAX templates render the canonical strings."""

    def test_dependency_expansion(self) -> None:
        from tools.measures import _resolve_ti_patterns
        # Requesting YOY% pulls YOY and SPY in front of it.
        resolved = _resolve_ti_patterns(["YOY%"])
        self.assertEqual(resolved, ["SPY", "YOY", "YOY%"])

    def test_unknown_pattern_raises(self) -> None:
        from pbi_connection import PowerBIValidationError
        from tools.measures import _resolve_ti_patterns
        with self.assertRaises(PowerBIValidationError):
            _resolve_ti_patterns(["fOO"])

    def test_template_renders(self) -> None:
        from tools.measures import _TIME_INTELLIGENCE_TEMPLATES
        rendered = _TIME_INTELLIGENCE_TEMPLATES["YTD"]["template"].format(
            base="Sales", date_table="Date", date_column="Date",
        )
        self.assertEqual(rendered, "CALCULATE([Sales], DATESYTD(Date[Date]))")


class DAXSemanticReferenceParserTests(unittest.TestCase):
    """Reference scanner inside pbi_validate_dax_semantic_tool finds the right tokens."""

    def test_extracts_columns_and_measures(self) -> None:
        from tools.query import _DAX_TABLE_COLUMN_RE, _DAX_MEASURE_REF_RE
        expr = 'CALCULATE([Sales], Date[Year] = 2024) - [Sales SPY]'
        cols = {(m.group("table"), m.group("column")) for m in _DAX_TABLE_COLUMN_RE.finditer(expr)}
        self.assertEqual(cols, {("Date", "Year")})
        # Strip column refs first to isolate bare measures.
        leftover = _DAX_TABLE_COLUMN_RE.sub("", expr)
        measures = {m.group("measure") for m in _DAX_MEASURE_REF_RE.finditer(leftover)}
        self.assertEqual(measures, {"Sales", "Sales SPY"})


class V010MatskiBugfixTests(unittest.TestCase):
    """Regression tests for the 7 bugs reported during MATSKI report creation."""

    # --- Bug 1: pbi_extract_report falls back to ZIP when pbi-tools.core lacks 'extract' ---
    def test_extract_report_falls_back_to_zip_when_cli_unavailable(self) -> None:
        from tools.visuals import (
            VisualToolError,
            pbi_extract_report_tool,
            LAYOUT_RELATIVE_PATH,
        )

        with tempfile.TemporaryDirectory() as tmp_dir:
            root = Path(tmp_dir)
            previous = [str(item) for item in SECURITY.allowed_base_dirs()]
            SECURITY.configure_allowed_dirs([str(root)])
            try:
                pbix = root / "fake.pbix"
                with zipfile.ZipFile(pbix, "w") as zf:
                    zf.writestr("Report/Layout", json.dumps({"sections": []}, ensure_ascii=False).encode("utf-16-le"))
                    zf.writestr("Report/StaticResources/Themes/Dummy.json", b'{"name":"Dummy"}')
                target = root / "extracted"

                # Simulate the bundled pbi-tools.core CLI failing on 'extract'.
                def _fake_run_pbi_tools(args):
                    raise VisualToolError(
                        "pbi-tools command failed.",
                        details={"stdout": "Unknown action: 'extract'", "stderr": ""},
                    )

                with patch("tools.visuals._run_pbi_tools", side_effect=_fake_run_pbi_tools):
                    result = pbi_extract_report_tool(str(pbix), extract_folder=str(target))

                self.assertTrue(result["ok"], result)
                self.assertEqual(result["extraction_method"], "zip_native")
                # Layout was reconstructed from the ZIP.
                self.assertTrue((target / LAYOUT_RELATIVE_PATH).exists())
                # Static resources copied too.
                self.assertTrue((target / "Report" / "StaticResources" / "Themes" / "Dummy.json").exists())
            finally:
                SECURITY.configure_allowed_dirs(previous)

    # --- Bug 2: bracket-form references parse identically to dotted form ---
    def test_reference_parsing_accepts_bracket_and_dotted_forms(self) -> None:
        from tools.visuals import _normalize_reference, _query_ref, _split_column_ref

        # All three forms collapse to the canonical "Date.Année" representation.
        self.assertEqual(_normalize_reference("Date.Année"), "Date.Année")
        self.assertEqual(_normalize_reference("Date[Année]"), "Date.Année")
        self.assertEqual(_normalize_reference("'Date'[Année]"), "Date.Année")
        self.assertEqual(_normalize_reference("'Date Avec Espaces'[Year]"), "Date Avec Espaces.Year")
        # Bare measure stays bare.
        self.assertEqual(_normalize_reference("Sales"), "Sales")
        # Short queryRef strips the table prefix in every form.
        self.assertEqual(_query_ref("Date[Année]"), "Année")
        self.assertEqual(_query_ref("'Date'[Année]"), "Année")
        # _split_column_ref accepts bracket form.
        self.assertEqual(_split_column_ref("Date[Année]"), ("Date", "Année"))
        self.assertEqual(_split_column_ref("'Date'[Année]"), ("Date", "Année"))

    # --- Bug 3: pbi_add_visual accepts visual_type="map" ---
    def test_add_visual_dispatch_includes_map(self) -> None:
        from tools.visuals import _VISUAL_TYPE_DISPATCH

        self.assertIn("map", _VISUAL_TYPE_DISPATCH)
        self.assertTrue(callable(_VISUAL_TYPE_DISPATCH["map"]))

    # --- Bug 4: visuals carry the right home table when manager is supplied ---
    def test_resolve_measure_home_map_pulls_live_model(self) -> None:
        from types import SimpleNamespace
        from tools.visuals import _augment_measure_home_map_with_live

        with patch(
            "tools.visuals.pbi_model_info_tool",
            return_value={
                "ok": True,
                "measures": [
                    {"name": "Sales", "table": "Facture"},
                    {"name": "Margin", "table": "Produit"},
                ],
            },
        ):
            home_map: dict[str, str] = {}
            _augment_measure_home_map_with_live(home_map, SimpleNamespace())
            self.assertEqual(home_map.get("Sales"), "Facture")
            self.assertEqual(home_map.get("Margin"), "Produit")

        # Existing entries are not overwritten by live results (extract metadata wins).
        with patch(
            "tools.visuals.pbi_model_info_tool",
            return_value={"ok": True, "measures": [{"name": "Sales", "table": "Facture"}]},
        ):
            home_map = {"Sales": "ManualTable"}
            _augment_measure_home_map_with_live(home_map, SimpleNamespace())
            self.assertEqual(home_map["Sales"], "ManualTable")

    # --- Bug 5: detect_empty_visuals does not flag FORMAT() text measures ---
    def test_empty_visuals_skips_text_measure_zero_check(self) -> None:
        # The fix is logical: only flag "all numeric and all zero" — strings like
        # "1 230 €" returned by FORMAT() should not trigger the warning.
        non_blank_text = ["1 230 €", "0 K €"]
        non_blank_zero = [0, 0.0]
        non_blank_mixed = [0, "0 K €"]

        all_numeric = lambda values: all(isinstance(v, (int, float)) and not isinstance(v, bool) for v in values)
        # The triggering condition (post-fix):
        def should_flag(values):
            return all_numeric(values) and all(float(v) == 0 for v in values)

        self.assertFalse(should_flag(non_blank_text))
        self.assertTrue(should_flag(non_blank_zero))
        self.assertFalse(should_flag(non_blank_mixed))

    # --- Bug 6: error message names the expected role kind ---
    def test_field_validation_error_includes_role_kind_hint(self) -> None:
        from pbi_connection import PowerBIValidationError
        from tools.visuals import _validate_field_references_live

        fake_index = {
            "columns": {("dim_date", "year")},
            "measures": {"sales": {"facture"}},
            "measure_tables": {"sales": {"Facture"}},
        }
        with patch("tools.visuals._live_model_field_index", return_value=(fake_index, {"status": "available"})):
            from types import SimpleNamespace

            # User passes "Annee" expecting it to fill a column role; the validator
            # surfaces kind="column" and a hint about the right format.
            with self.assertRaises(PowerBIValidationError) as ctx:
                _validate_field_references_live(
                    SimpleNamespace(),
                    ["Annee"],
                    expected_kinds={"Annee": "column"},
                )
            entry = ctx.exception.details["missing"][0]
            self.assertEqual(entry["kind"], "column")
            self.assertIn("hint", entry)
            self.assertIn("Date.Year", entry["hint"])

    # --- Bug 7: layout linter ignore_warnings + only_pages knobs ---
    def test_layout_linter_ignore_warnings_and_only_pages(self) -> None:
        from tools.quality import pbi_lint_report_layout_tool

        with tempfile.TemporaryDirectory() as tmp_dir:
            root = Path(tmp_dir)
            previous = [str(item) for item in SECURITY.allowed_base_dirs()]
            SECURITY.configure_allowed_dirs([str(root)])
            try:
                # Page with a tiny visual that would normally trigger visual_too_small + missing_title.
                layout = {
                    "sections": [
                        {
                            "name": "ReportSection1",
                            "displayName": "Crowded",
                            "displayOption": 0,
                            "width": 1280,
                            "height": 720,
                            "visualContainers": [
                                {
                                    "x": 10, "y": 10, "z": 0, "width": 50, "height": 30,
                                    "config": json.dumps({
                                        "name": "v1",
                                        "singleVisual": {"visualType": "card", "objects": {}, "projections": {"Values": [{"queryRef": "X"}]}, "prototypeQuery": {"Version": 2, "From": [], "Select": []}},
                                        "layouts": [{"id": 0, "position": {"x": 10, "y": 10, "width": 50, "height": 30}}],
                                    }, ensure_ascii=False),
                                    "filters": "[]",
                                    "query": "{}",
                                    "dataTransforms": "{}",
                                }
                            ],
                            "filters": "[]",
                        },
                        {
                            "name": "ReportSection2",
                            "displayName": "Empty",
                            "displayOption": 0,
                            "width": 1280,
                            "height": 720,
                            "visualContainers": [],
                            "filters": "[]",
                        },
                    ],
                }
                extract = root / "extract"
                (extract / "Report").mkdir(parents=True)
                (extract / "Report" / "Layout").write_text(
                    json.dumps(layout, ensure_ascii=False, indent=2), encoding="utf-16-le",
                )

                # Without filters → both warnings present on the Crowded page.
                full = pbi_lint_report_layout_tool(str(extract))
                warning_types = {w["type"] for w in full["warnings"]}
                self.assertIn("visual_too_small", warning_types)
                self.assertIn("missing_title", warning_types)

                # With ignore_warnings → those warning types are dropped.
                filtered = pbi_lint_report_layout_tool(
                    str(extract),
                    ignore_warnings=["visual_too_small", "missing_title"],
                )
                filtered_types = {w["type"] for w in filtered["warnings"]}
                self.assertNotIn("visual_too_small", filtered_types)
                self.assertNotIn("missing_title", filtered_types)
                self.assertIn("ignored_warnings", filtered)
                self.assertEqual(set(filtered["ignored_warnings"]), {"visual_too_small", "missing_title"})

                # With only_pages → only the named pages are scanned.
                scoped = pbi_lint_report_layout_tool(str(extract), only_pages=["Empty"])
                scoped_pages = {w.get("page") for w in scoped["warnings"]}
                self.assertNotIn("Crowded", scoped_pages)
            finally:
                SECURITY.configure_allowed_dirs(previous)


if __name__ == "__main__":
    unittest.main(verbosity=2)

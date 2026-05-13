"""v0.13.2: pbix-mcp 0.9.2 upstream patch regressions.

The 0.9.2 release ships a partly-fixed builder. This suite pins every
fix delivered in v0.13.2 (bugs #1, #3, #4, #5) plus the DBCC DAX-risk
scanner (#2) so that an accidental ``pip install --upgrade pbix-mcp``
into an un-patched build is caught here.

For #2 the upstream Vertipaq dictionary encoder cannot be repaired
without a full encoder rewrite; the patch surfaces the known-bad DAX
patterns as a structured warning instead of silently producing a
corrupted ``.pbix``.
"""

from __future__ import annotations

import json
import sys
import tempfile
import unittest
import warnings
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))


def _layout_dict(pbix_bytes_or_path) -> dict:
    if isinstance(pbix_bytes_or_path, (str, Path)):
        with open(pbix_bytes_or_path, "rb") as handle:
            data = handle.read()
    else:
        data = pbix_bytes_or_path
    import io

    with zipfile.ZipFile(io.BytesIO(data), "r") as zf:
        return json.loads(zf.read("Report/Layout").decode("utf-16-le"))


def _make_minimal_table_spec() -> list[dict]:
    return [
        {
            "name": "Sales",
            "columns": [
                {"name": "Region", "data_type": "String"},
                {"name": "Amount", "data_type": "Decimal"},
            ],
            "rows": [{"Region": "North", "Amount": 100.0}],
        }
    ]


class Bug1_FormatStringPersisted(unittest.TestCase):
    """Bug #1: ``Measure.FormatString`` is persisted, not literal NULL."""

    def test_insert_measure_uses_format_string_placeholder(self) -> None:
        from pbix_mcp import builder

        source = Path(builder.__file__).read_text(encoding="utf-8")
        # The hardened INSERT carries FormatString as a ? placeholder
        # and the parameter tuple includes mdef.get("format_string"). A
        # literal NULL in the VALUES clause would mean the legacy
        # always-NULL bug is back.
        insert_block_start = source.index("INSERT INTO [Measure]")
        # The block extends until the close paren / parameter list. Grab
        # 1200 chars which always contains the full statement.
        insert_block = source[insert_block_start : insert_block_start + 1200]
        self.assertIn('mdef.get("format_string")', insert_block)
        # The legacy bug had `?, NULL, 0, 1,` — a literal NULL where the
        # FormatString placeholder should sit. Hardened code never has
        # ``, NULL,`` immediately after the Expression placeholder.
        self.assertNotIn("?, NULL, 0, 1,", insert_block)


class Bug2_DBCCRiskScanner(unittest.TestCase):
    """Bug #2: DAX-risk scanner emits a structured warning for the
    HASONEVALUE+VALUES pattern over an embedded-rows table."""

    def test_scanner_flags_hasonevalue_values_on_embedded_table(self) -> None:
        from pbix_mcp.dbcc_guard import scan_measures_for_dbcc_risks

        measures = [
            {
                "table": "F",
                "name": "Apport CP",
                "expression": (
                    "VAR sel = IF(HASONEVALUE(T[Scenario]), VALUES(T[Scenario]), \"default\") "
                    "RETURN CALCULATE(SUM(T[Amount]), T[Source] = sel)"
                ),
            }
        ]
        tables = [
            {"name": "T", "rows": [{"Scenario": "A"}], "columns": [{"name": "Scenario", "data_type": "String"}]},
            {"name": "F", "rows": [{"Amount": 1}], "columns": [{"name": "Amount", "data_type": "Decimal"}]},
        ]
        findings = scan_measures_for_dbcc_risks(measures, tables)
        self.assertEqual(len(findings), 1)
        self.assertEqual(findings[0]["pattern"], "hasonevalue_values_string")
        self.assertEqual(findings[0]["measure"], "Apport CP")
        self.assertEqual(findings[0]["referenced_table"], "T")

    def test_scanner_ignores_csv_sourced_table(self) -> None:
        from pbix_mcp.dbcc_guard import scan_measures_for_dbcc_risks

        measures = [
            {
                "table": "F",
                "name": "Safe",
                "expression": "IF(HASONEVALUE(T[Scenario]), VALUES(T[Scenario]), \"x\")",
            }
        ]
        tables = [
            {"name": "T", "rows": [], "columns": [], "source_csv": "C:/data/t.csv"},
            {"name": "F", "rows": [], "columns": []},
        ]
        findings = scan_measures_for_dbcc_risks(measures, tables)
        self.assertEqual(findings, [])

    def test_scanner_flags_treatas_values_over_embedded_table(self) -> None:
        from pbix_mcp.dbcc_guard import scan_measures_for_dbcc_risks

        measures = [
            {
                "table": "F",
                "name": "TreatAsRisk",
                "expression": "CALCULATE([Total], TREATAS(VALUES(T[Scenario]), F[Scenario]))",
            }
        ]
        tables = [{"name": "T", "rows": [{"Scenario": "A"}], "columns": []}, {"name": "F", "rows": [], "columns": []}]
        findings = scan_measures_for_dbcc_risks(measures, tables)
        patterns = {f["pattern"] for f in findings}
        self.assertIn("treatas_string", patterns)

    def test_emit_runtime_warnings_uses_dbcc_category(self) -> None:
        from pbix_mcp.dbcc_guard import (
            DBCCRiskWarning,
            emit_runtime_warnings,
            scan_measures_for_dbcc_risks,
        )

        measures = [
            {
                "table": "F",
                "name": "Risky",
                "expression": "IF(HASONEVALUE(T[Scenario]), VALUES(T[Scenario]), \"x\")",
            }
        ]
        tables = [{"name": "T", "rows": [{"Scenario": "A"}], "columns": []}, {"name": "F", "rows": [], "columns": []}]
        findings = scan_measures_for_dbcc_risks(measures, tables)
        with warnings.catch_warnings(record=True) as captured:
            warnings.simplefilter("always")
            emit_runtime_warnings(findings)
        categories = {w.category for w in captured}
        self.assertIn(DBCCRiskWarning, categories)


class Bug3_VisualConfigPassthroughs(unittest.TestCase):
    """Bug #3: ``series``, ``objects``, ``vcObjects`` and page ``config``
    survive a Layout build round-trip."""

    def _build_layout(self, page: dict) -> dict:
        from pbix_mcp.builder import PBIXBuilder

        builder = PBIXBuilder()
        builder._tables = [
            {
                "name": "Sales",
                "columns": [
                    {"name": "Region", "data_type": "String"},
                    {"name": "Product", "data_type": "String"},
                ],
                "rows": [],
                "hidden": False,
                "source_csv": None,
                "source_db": None,
                "mode": "import",
            }
        ]
        builder._measures = [{"table": "Sales", "name": "Total", "expression": "1", "description": ""}]
        builder._pages = [page]
        raw = builder._build_layout()
        return json.loads(raw.decode("utf-16-le"))

    def test_series_projection_added_when_present(self) -> None:
        layout = self._build_layout(
            {
                "name": "P1",
                "visuals": [
                    {
                        "type": "stackedBarChart",
                        "config": {
                            "category": {"table": "Sales", "column": "Region"},
                            "measure": "Total",
                            "series": {"table": "Sales", "column": "Product"},
                        },
                    }
                ],
            }
        )
        config = json.loads(layout["sections"][0]["visualContainers"][0]["config"])
        projections = config["singleVisual"]["projections"]
        self.assertIn("Series", projections)
        self.assertTrue(projections["Series"][0]["active"])

    def test_objects_and_vcObjects_passed_through(self) -> None:
        layout = self._build_layout(
            {
                "name": "P1",
                "visuals": [
                    {
                        "type": "card",
                        "config": {"measure": "Total"},
                        "objects": {"general": [{"show": True}]},
                        "vcObjects": {"background": [{"show": True}]},
                    }
                ],
            }
        )
        config = json.loads(layout["sections"][0]["visualContainers"][0]["config"])
        single_visual = config["singleVisual"]
        self.assertEqual(single_visual["objects"], {"general": [{"show": True}]})
        self.assertEqual(single_visual["vcObjects"], {"background": [{"show": True}]})

    def test_page_config_passed_through(self) -> None:
        layout = self._build_layout(
            {
                "name": "P1",
                "visuals": [],
                "config": {"background": {"color": {"solid": {"color": "#FFFFFF"}}}},
            }
        )
        self.assertEqual(
            layout["sections"][0]["config"],
            {"background": {"color": {"solid": {"color": "#FFFFFF"}}}},
        )


class Bug4_AddMeasureFormatStringSignature(unittest.TestCase):
    """Bug #4: ``add_measure`` accepts ``format_string`` directly; the
    stored measure dict carries it without post-hoc mutation."""

    def test_add_measure_accepts_format_string_kwarg(self) -> None:
        from pbix_mcp.builder import PBIXBuilder

        builder = PBIXBuilder()
        builder.add_measure("Sales", "Margin %", "DIVIDE(1, 2)", format_string="0.0%")
        self.assertEqual(builder._measures[-1]["format_string"], "0.0%")

    def test_add_measure_defaults_format_string_to_none(self) -> None:
        from pbix_mcp.builder import PBIXBuilder

        builder = PBIXBuilder()
        builder.add_measure("Sales", "M", "1")
        self.assertIsNone(builder._measures[-1]["format_string"])

    def test_persistent_report_forwards_format_string(self) -> None:
        # Patch the loader to inject our fake builder, then ensure the
        # persistent_report path forwards the format string via the new
        # keyword API rather than the post-hoc mutation fallback.
        from tools import persistent_report

        class _FakeBuilder:
            def __init__(self) -> None:
                self._measures: list[dict] = []
                self._tables: list[dict] = []
                self.last_format_string: str | None = None

            def add_table(self, *a, **kw) -> None:
                self._tables.append({"name": a[0], "columns": a[1], "rows": a[2]})

            def add_measure(self, table, name, expression, format_string=None):
                self.last_format_string = format_string
                self._measures.append(
                    {"table": table, "name": name, "expression": expression, "format_string": format_string}
                )

            def add_relationship(self, *a, **kw) -> None:
                pass

            def add_page(self, *a, **kw) -> None:
                pass

            def _pre_build_checks(self):
                return []

            def validate(self):
                return []

            def save(self, output_path):
                Path(output_path).write_bytes(b"PK\x03\x04stub")

        instance = _FakeBuilder()
        original_loader = persistent_report._load_pbix_builder
        persistent_report._load_pbix_builder = lambda: (lambda: instance)
        try:
            from security import SECURITY, configure_allowed_dirs

            with tempfile.TemporaryDirectory() as tmp:
                previous = [str(p) for p in SECURITY.allowed_base_dirs()]
                configure_allowed_dirs([tmp])
                SECURITY.policy(reload=True, cwd=Path(tmp))
                try:
                    out = Path(tmp) / "x.pbix"
                    persistent_report.pbi_create_persistent_report_tool(
                        output_path=str(out),
                        tables=_make_minimal_table_spec(),
                        measures=[
                            {
                                "table": "Sales",
                                "name": "Margin %",
                                "expression": "DIVIDE(1, 2)",
                                "format_string": "0.0%",
                            }
                        ],
                    )
                finally:
                    configure_allowed_dirs(previous)
                    SECURITY.policy(reload=True, cwd=Path.cwd())
        finally:
            persistent_report._load_pbix_builder = original_loader

        self.assertEqual(instance.last_format_string, "0.0%")
        self.assertEqual(instance._measures[-1]["format_string"], "0.0%")


class Bug5_RowFormatValidation(unittest.TestCase):
    """Bug #5: ``add_table`` raises a usable TypeError when rows are not
    dicts."""

    def test_list_row_raises_clear_typeerror(self) -> None:
        from pbix_mcp.builder import PBIXBuilder

        builder = PBIXBuilder()
        with self.assertRaises(TypeError) as cm:
            builder.add_table(
                "T",
                columns=[{"name": "a", "data_type": "Int64"}],
                rows=[[1]],
            )
        msg = str(cm.exception)
        self.assertIn("T", msg)
        self.assertIn("row 0", msg)
        self.assertIn("dict", msg)
        # The message should suggest the expected shape so callers can fix.
        self.assertIn("col1", msg)

    def test_tuple_row_raises_typeerror(self) -> None:
        from pbix_mcp.builder import PBIXBuilder

        builder = PBIXBuilder()
        with self.assertRaises(TypeError):
            builder.add_table(
                "T",
                columns=[{"name": "a", "data_type": "Int64"}],
                rows=[(1,)],
            )

    def test_dict_rows_accepted(self) -> None:
        from pbix_mcp.builder import PBIXBuilder

        builder = PBIXBuilder()
        builder.add_table(
            "T",
            columns=[{"name": "a", "data_type": "Int64"}],
            rows=[{"a": 1}, {"a": 2}],
        )
        self.assertEqual(len(builder._tables[-1]["rows"]), 2)


class Bug2_DBCCWarningOnSave(unittest.TestCase):
    """Bug #2 end-to-end: a build that includes the risky pattern emits a
    ``DBCCRiskWarning`` via ``save()`` so callers can never miss it."""

    def test_save_emits_dbcc_warning_on_risky_pattern(self) -> None:
        from pbix_mcp.builder import PBIXBuilder
        from pbix_mcp.dbcc_guard import DBCCRiskWarning

        builder = PBIXBuilder()
        builder.add_table(
            "T",
            columns=[
                {"name": "Scenario", "data_type": "String"},
                {"name": "Source", "data_type": "String"},
                {"name": "Amount", "data_type": "Decimal"},
            ],
            rows=[
                {"Scenario": "A", "Source": "x", "Amount": 1.0},
                {"Scenario": "B", "Source": "y", "Amount": 2.0},
            ],
        )
        builder.add_measure(
            "T",
            "Apport CP",
            (
                "VAR sel = IF(HASONEVALUE(T[Scenario]), VALUES(T[Scenario]), \"default\") "
                "RETURN CALCULATE(SUM(T[Amount]), T[Source] = sel)"
            ),
        )

        with warnings.catch_warnings(record=True) as captured:
            warnings.simplefilter("always")
            # Build to bytes without writing — we only care about the
            # warning emission path. ``build()`` triggers
            # ``_pre_build_checks`` which surfaces the same finding.
            try:
                builder.build()
            except Exception:
                pass
        # The pre-build path surfaces the same finding as a string.
        msg_blob = "\n".join(str(w.message) for w in captured)
        self.assertIn("DBCC string-store risk", msg_blob)


if __name__ == "__main__":
    unittest.main(verbosity=2)

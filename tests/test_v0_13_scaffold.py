"""v0.13: PBIX scaffold tool coverage.

The scaffold tool wraps ``pbi_create_persistent_report_tool`` with named
templates. Validate template names, table/measure shape, theme-rejection
on invalid JSON, and the public list-templates helper. The underlying
PBIX builder is monkeypatched so these tests do not need ``pbix-mcp``
installed.
"""

from __future__ import annotations

import json
import sys
import tempfile
import types
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

from tools import (
    pbi_list_scaffold_templates_tool,
    pbi_scaffold_pbix_tool,
)
from tools import scaffold as scaffold_module
from tools.scaffold import SCAFFOLD_TEMPLATES, list_scaffold_templates


class FakePBIXBuilder:
    def __init__(self) -> None:
        self._tables: list[dict] = []
        self._measures: list[dict] = []
        self._relationships: list[dict] = []
        self._pages: list[dict] = []

    def add_table(self, name, columns, rows, source_csv=None, source_db=None, mode="import"):
        self._tables.append({"name": name, "columns": columns, "rows": rows, "mode": mode})

    def add_measure(self, table, name, expression):
        self._measures.append({"table": table, "name": name, "expression": expression, "format_string": None})

    def add_relationship(self, from_table, from_column, to_table, to_column):
        self._relationships.append(
            {"from_table": from_table, "from_column": from_column, "to_table": to_table, "to_column": to_column}
        )

    def add_page(self, name, visuals):
        self._pages.append({"name": name, "visuals": visuals})

    def _pre_build_checks(self):
        return []

    def validate(self):
        return []

    def save(self, output_path):
        Path(output_path).write_bytes(b"PK\x03\x04fake-pbix")


class ScaffoldTemplateListingTests(unittest.TestCase):
    def test_templates_dict_has_baseline_keys(self) -> None:
        self.assertIn("blank", SCAFFOLD_TEMPLATES)
        self.assertIn("finance", SCAFFOLD_TEMPLATES)
        self.assertIn("sales", SCAFFOLD_TEMPLATES)
        self.assertIn("analytics", SCAFFOLD_TEMPLATES)

    def test_list_helper_summarises_each_template(self) -> None:
        templates = list_scaffold_templates()
        keys = {entry["key"] for entry in templates}
        self.assertEqual(keys, {"blank", "finance", "sales", "analytics"})
        for entry in templates:
            self.assertIn("description", entry)
            self.assertIn("table_count", entry)
            self.assertIn("measure_count", entry)
            self.assertGreaterEqual(entry["table_count"], 1)

    def test_list_tool_emits_ok(self) -> None:
        result = pbi_list_scaffold_templates_tool()
        self.assertTrue(result["ok"], result)
        self.assertEqual(len(result["templates"]), len(SCAFFOLD_TEMPLATES))

    def test_blank_template_has_only_date_table(self) -> None:
        blank = SCAFFOLD_TEMPLATES["blank"]
        self.assertEqual(len(blank["tables"]), 1)
        self.assertEqual(blank["tables"][0]["name"], "DateTable")
        self.assertEqual(blank["measures"], [])

    def test_finance_template_has_baseline_measures(self) -> None:
        finance = SCAFFOLD_TEMPLATES["finance"]
        measure_names = {m["name"] for m in finance["measures"]}
        self.assertIn("Total", measure_names)
        self.assertIn("Total YTD", measure_names)
        self.assertIn("Total MTD", measure_names)
        self.assertIn("Total YoY %", measure_names)

    def test_sales_template_has_two_relationships(self) -> None:
        sales = SCAFFOLD_TEMPLATES["sales"]
        self.assertEqual(len(sales["relationships"]), 2)

    def test_analytics_template_has_events_and_user(self) -> None:
        analytics = SCAFFOLD_TEMPLATES["analytics"]
        table_names = {t["name"] for t in analytics["tables"]}
        self.assertIn("Events", table_names)
        self.assertIn("User", table_names)


class ScaffoldExecutionTests(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp = tempfile.TemporaryDirectory()
        self.root = Path(self.tmp.name)
        from security import SECURITY, configure_allowed_dirs

        self.previous_allowed = [str(item) for item in SECURITY.allowed_base_dirs()]
        configure_allowed_dirs([str(self.root)])
        SECURITY.policy(reload=True, cwd=self.root)
        # monkeypatch the pbix builder loader inside persistent_report.
        from tools import persistent_report

        self._original_loader = persistent_report._load_pbix_builder
        persistent_report._load_pbix_builder = lambda: FakePBIXBuilder

    def tearDown(self) -> None:
        from security import SECURITY, configure_allowed_dirs
        from tools import persistent_report

        persistent_report._load_pbix_builder = self._original_loader
        configure_allowed_dirs(self.previous_allowed)
        SECURITY.policy(reload=True, cwd=Path.cwd())
        self.tmp.cleanup()

    def test_unknown_template_rejected(self) -> None:
        from pbi_connection import PowerBIValidationError

        out = self.root / "bad.pbix"
        with self.assertRaises(PowerBIValidationError) as cm:
            pbi_scaffold_pbix_tool(str(out), template="bogus")
        self.assertIn("template", cm.exception.details)

    def test_blank_scaffold_writes_pbix(self) -> None:
        out = self.root / "blank.pbix"
        result = pbi_scaffold_pbix_tool(str(out), template="blank")
        self.assertTrue(result["ok"], result)
        self.assertTrue(out.exists())
        self.assertEqual(result["template"], "blank")
        self.assertEqual(result["table_count"], 1)

    def test_finance_scaffold_reports_measure_count(self) -> None:
        out = self.root / "fin.pbix"
        result = pbi_scaffold_pbix_tool(str(out), template="finance")
        self.assertTrue(result["ok"], result)
        # baseline 4 measures
        self.assertEqual(result["measure_count"], 4)

    def test_extra_measures_appended(self) -> None:
        out = self.root / "fin.pbix"
        result = pbi_scaffold_pbix_tool(
            str(out),
            template="finance",
            extra_measures=[{"table": "GL", "name": "Custom", "expression": "SUM(GL[Amount])"}],
        )
        self.assertTrue(result["ok"], result)
        # baseline 4 + 1 custom = 5
        self.assertEqual(result["measure_count"], 5)

    def test_invalid_extra_measures_payload(self) -> None:
        from pbi_connection import PowerBIValidationError

        out = self.root / "fin.pbix"
        with self.assertRaises(PowerBIValidationError):
            pbi_scaffold_pbix_tool(str(out), template="finance", extra_measures="not-a-list")

    def test_theme_json_path_must_validate(self) -> None:
        from tools.visuals._themes import ThemeValidationError

        out = self.root / "fin.pbix"
        bad_theme = self.root / "bad-theme.json"
        bad_theme.write_text(json.dumps({"rogue": True}), encoding="utf-8")
        with self.assertRaises(ThemeValidationError):
            pbi_scaffold_pbix_tool(
                str(out),
                template="finance",
                theme_json_path=str(bad_theme),
            )

    def test_valid_theme_is_accepted(self) -> None:
        out = self.root / "fin.pbix"
        good_theme = self.root / "good-theme.json"
        good_theme.write_text(json.dumps({"name": "Clean", "dataColors": ["#001122"]}), encoding="utf-8")
        result = pbi_scaffold_pbix_tool(
            str(out),
            template="finance",
            theme_json_path=str(good_theme),
        )
        self.assertTrue(result["ok"], result)
        self.assertTrue(result["theme_applied"])


if __name__ == "__main__":
    unittest.main(verbosity=2)

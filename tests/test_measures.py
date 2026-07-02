"""Offline tests for measure tools (src/tools/measures.py) with a faked TOM layer."""

from __future__ import annotations

import sys
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

from pbi_connection import (
    PowerBIDuplicateError,
    PowerBINotFoundError,
    PowerBIValidationError,
)
from security import SECURITY
from tools.measures import (
    _parse_dax_file,
    _resolve_ti_patterns,
    _strip_dax_comments,
    pbi_create_contribution_measure_tool,
    pbi_create_measure_tool,
    pbi_create_measures_tool,
    pbi_create_rolling_average_measure_tool,
    pbi_create_time_intelligence_pack_tool,
    pbi_create_topn_measure_tool,
    pbi_create_variance_measure_tool,
    pbi_delete_measure_tool,
    pbi_import_dax_file_tool,
    pbi_list_measures_tool,
    pbi_rename_measure_tool,
)


class FakeCollection(list):
    @property
    def Count(self) -> int:
        return len(self)

    def Find(self, name: str):
        for item in self:
            if str(getattr(item, "Name", "")).casefold() == name.casefold():
                return item
        return None

    def Add(self, item) -> None:
        self.append(item)

    def Remove(self, item) -> None:
        list.remove(self, item)


class FakeMeasure:
    def __init__(self) -> None:
        self.Name = ""
        self.Expression = ""
        self.FormatString = ""
        self.Description = ""
        self.DisplayFolder = ""
        self.IsHidden = False


class FakeTable:
    def __init__(self, name: str, measures=()) -> None:
        self.Name = name
        self.Measures = FakeCollection(measures)


class FakeModel:
    def __init__(self, tables=()) -> None:
        self.Tables = FakeCollection(tables)


class FakeTom:
    Measure = FakeMeasure


class FakeManager:
    def __init__(self, model: FakeModel) -> None:
        self.tom = FakeTom()
        self.database = SimpleNamespace(Model=model)
        self.state = SimpleNamespace(
            database=self.database,
            snapshot=lambda: {"connected": True, "database": "UnitTest"},
        )
        self.write_calls = 0

    def run_read(self, _operation_name, reader):
        return reader(self.state)

    def cached_run_read(self, _cache_key, _operation_name, reader):
        return reader(self.state)

    def execute_write(self, _operation_name, mutator):
        self.write_calls += 1
        payload = mutator(self.state, self.database, self.database.Model)
        payload["save_result"] = {"status": "saved"}
        payload["connection"] = self.state.snapshot()
        return payload


def make_measure(name: str, expression: str = "1", format_string: str = "") -> FakeMeasure:
    measure = FakeMeasure()
    measure.Name = name
    measure.Expression = expression
    measure.FormatString = format_string
    return measure


def make_manager(measures=()) -> FakeManager:
    return FakeManager(FakeModel(tables=[FakeTable("Sales", measures=measures)]))


class CreateMeasureTests(unittest.TestCase):
    def test_create_measure_happy_path(self) -> None:
        manager = make_manager()
        result = pbi_create_measure_tool(
            manager,
            table="Sales",
            name="Total Sales",
            expression="SUM(Sales[Amount])",
            format_string="#,0",
            description="Total amount",
            display_folder="Base",
            is_hidden=False,
        )
        self.assertTrue(result["ok"], result)
        self.assertEqual(result["action"], "created")
        measure = manager.database.Model.Tables[0].Measures.Find("Total Sales")
        self.assertIsNotNone(measure)
        self.assertEqual(measure.Expression, "SUM(Sales[Amount])")
        self.assertEqual(measure.FormatString, "#,0")
        self.assertEqual(measure.DisplayFolder, "Base")

    def test_create_measure_overwrite_default_updates(self) -> None:
        manager = make_manager(measures=[make_measure("Total Sales")])
        result = pbi_create_measure_tool(manager, table="Sales", name="Total Sales", expression="SUM(Sales[Amount])")
        self.assertTrue(result["ok"], result)
        self.assertEqual(result["action"], "updated")
        self.assertEqual(manager.database.Model.Tables[0].Measures.Count, 1)

    def test_create_measure_duplicate_without_overwrite_raises(self) -> None:
        manager = make_manager(measures=[make_measure("Total Sales")])
        with self.assertRaises(PowerBIDuplicateError):
            pbi_create_measure_tool(manager, table="Sales", name="total sales", expression="1", overwrite=False)

    def test_create_measure_table_not_found(self) -> None:
        manager = make_manager()
        with self.assertRaises(PowerBINotFoundError):
            pbi_create_measure_tool(manager, table="Ghost", name="M", expression="1")

    def test_create_measure_rejects_brackets_in_name(self) -> None:
        manager = make_manager()
        with self.assertRaises(PowerBIValidationError):
            pbi_create_measure_tool(manager, table="Sales", name="Bad[Name]", expression="1")

    def test_create_measure_rejects_query_only_dax(self) -> None:
        manager = make_manager()
        with self.assertRaises(PowerBIValidationError):
            pbi_create_measure_tool(manager, table="Sales", name="M", expression="EVALUATE VALUES(Sales)")


class CreateMeasuresBatchTests(unittest.TestCase):
    def test_batch_creates_and_updates(self) -> None:
        manager = make_manager(measures=[make_measure("Existing")])
        result = pbi_create_measures_tool(
            manager,
            table="Sales",
            measures=[
                {"name": "Existing", "expression": "2", "format_string": "#,0"},
                {"name": "New Measure", "expression": "3", "is_hidden": True},
            ],
        )
        self.assertTrue(result["ok"], result)
        self.assertEqual(result["created"], 1)
        self.assertEqual(result["updated"], 1)
        self.assertEqual(result["failed"], 0)
        table = manager.database.Model.Tables[0]
        self.assertEqual(table.Measures.Count, 2)
        self.assertEqual(table.Measures.Find("Existing").Expression, "2")
        self.assertTrue(table.Measures.Find("New Measure").IsHidden)
        # A single SaveChanges (execute_write) for the whole batch.
        self.assertEqual(manager.write_calls, 1)

    def test_batch_duplicate_without_overwrite_records_error_envelope(self) -> None:
        manager = make_manager(measures=[make_measure("Existing")])
        result = pbi_create_measures_tool(
            manager,
            table="Sales",
            measures=[
                {"name": "Existing", "expression": "2"},
                {"name": "New Measure", "expression": "3"},
            ],
            overwrite=False,
        )
        self.assertTrue(result["ok"], result)
        self.assertEqual(result["created"], 1)
        self.assertEqual(result["failed"], 1)
        failed_item = result["results"][0]
        self.assertFalse(failed_item["ok"])
        self.assertIn("error", failed_item)
        self.assertIn("already exists", failed_item["error"]["message"])

    def test_batch_stop_on_error_halts_processing(self) -> None:
        manager = make_manager(measures=[make_measure("Existing")])
        result = pbi_create_measures_tool(
            manager,
            table="Sales",
            measures=[
                {"name": "Existing", "expression": "2"},
                {"name": "Never Created", "expression": "3"},
            ],
            overwrite=False,
            stop_on_error=True,
        )
        self.assertTrue(result["ok"], result)
        self.assertEqual(result["failed"], 1)
        self.assertEqual(len(result["results"]), 1)
        self.assertIsNone(manager.database.Model.Tables[0].Measures.Find("Never Created"))

    def test_batch_empty_list_raises(self) -> None:
        manager = make_manager()
        with self.assertRaises(PowerBIValidationError):
            pbi_create_measures_tool(manager, table="Sales", measures=[])

    def test_batch_invalid_expression_raises_before_any_write(self) -> None:
        manager = make_manager()
        with self.assertRaises(PowerBIValidationError):
            pbi_create_measures_tool(
                manager,
                table="Sales",
                measures=[{"name": "M", "expression": "EVALUATE VALUES(Sales)"}],
            )
        self.assertEqual(manager.write_calls, 0)

    def test_batch_dry_run_reports_plan_without_mutation(self) -> None:
        manager = make_manager(measures=[make_measure("Existing")])
        result = pbi_create_measures_tool(
            manager,
            table="Sales",
            measures=[
                {"name": "Existing", "expression": "2"},
                {"name": "New Measure", "expression": "3"},
                {"name": "new measure", "expression": "4"},
            ],
            overwrite=False,
            dry_run=True,
        )
        self.assertTrue(result["ok"], result)
        self.assertTrue(result["dry_run"])
        actions = {item["name"]: item["planned_action"] for item in result["plan"]}
        self.assertEqual(actions["Existing"], "would_fail")
        self.assertEqual(actions["New Measure"], "would_create")
        self.assertEqual(actions["new measure"], "would_fail")
        self.assertEqual(result["would_create"], 1)
        self.assertEqual(result["would_update"], 0)
        self.assertEqual(result["would_fail"], 2)
        self.assertEqual(manager.write_calls, 0)
        self.assertEqual(manager.database.Model.Tables[0].Measures.Count, 1)

    def test_batch_dry_run_with_overwrite_reports_would_update(self) -> None:
        manager = make_manager(measures=[make_measure("Existing")])
        result = pbi_create_measures_tool(
            manager,
            table="Sales",
            measures=[{"name": "Existing", "expression": "2"}],
            overwrite=True,
            dry_run=True,
        )
        self.assertTrue(result["ok"], result)
        self.assertEqual(result["would_update"], 1)


class RenameMeasureTests(unittest.TestCase):
    def test_rename_measure_happy_path(self) -> None:
        manager = make_manager(measures=[make_measure("Old Name")])
        result = pbi_rename_measure_tool(manager, table="Sales", name="Old Name", new_name="New Name")
        self.assertTrue(result["ok"], result)
        self.assertEqual(
            result["rename"],
            {"table": "Sales", "measure_old_name": "Old Name", "measure_new_name": "New Name"},
        )
        self.assertIsNotNone(manager.database.Model.Tables[0].Measures.Find("New Name"))

    def test_rename_measure_case_only_change_allowed(self) -> None:
        manager = make_manager(measures=[make_measure("total")])
        result = pbi_rename_measure_tool(manager, table="Sales", name="total", new_name="Total")
        self.assertTrue(result["ok"], result)
        self.assertEqual(manager.database.Model.Tables[0].Measures[0].Name, "Total")

    def test_rename_measure_conflict_in_other_table_raises(self) -> None:
        model = FakeModel(
            tables=[
                FakeTable("Sales", measures=[make_measure("Old Name")]),
                FakeTable("Finance", measures=[make_measure("New Name")]),
            ]
        )
        manager = FakeManager(model)
        with self.assertRaises(PowerBIDuplicateError):
            pbi_rename_measure_tool(manager, table="Sales", name="Old Name", new_name="New Name")

    def test_rename_measure_not_found(self) -> None:
        manager = make_manager()
        with self.assertRaises(PowerBINotFoundError):
            pbi_rename_measure_tool(manager, table="Sales", name="Ghost", new_name="New Name")

    def test_rename_measure_table_not_found(self) -> None:
        manager = make_manager()
        with self.assertRaises(PowerBINotFoundError):
            pbi_rename_measure_tool(manager, table="Ghost", name="A", new_name="B")


class DeleteMeasureTests(unittest.TestCase):
    def test_delete_measure_happy_path(self) -> None:
        manager = make_manager(measures=[make_measure("Total Sales")])
        result = pbi_delete_measure_tool(manager, table="Sales", name="Total Sales")
        self.assertTrue(result["ok"], result)
        self.assertEqual(result["deleted_measure"], {"table": "Sales", "name": "Total Sales"})
        self.assertEqual(manager.database.Model.Tables[0].Measures.Count, 0)

    def test_delete_measure_not_found(self) -> None:
        manager = make_manager()
        with self.assertRaises(PowerBINotFoundError):
            pbi_delete_measure_tool(manager, table="Sales", name="Ghost")

    def test_delete_measure_table_not_found(self) -> None:
        manager = make_manager()
        with self.assertRaises(PowerBINotFoundError):
            pbi_delete_measure_tool(manager, table="Ghost", name="M")


class ListMeasuresTests(unittest.TestCase):
    def test_list_measures_skips_hidden_by_default(self) -> None:
        hidden = make_measure("Hidden Measure")
        hidden.IsHidden = True
        manager = make_manager(measures=[make_measure("Visible"), hidden])
        result = pbi_list_measures_tool(manager)
        self.assertTrue(result["ok"], result)
        self.assertEqual([item["name"] for item in result["measures"]], ["Visible"])
        result_all = pbi_list_measures_tool(manager, include_hidden=True)
        self.assertEqual(len(result_all["measures"]), 2)


class ResolveTiPatternsTests(unittest.TestCase):
    def test_none_returns_default_patterns(self) -> None:
        self.assertEqual(_resolve_ti_patterns(None), ["YTD", "MTD", "QTD", "SPY", "YOY", "YOY%", "MA3"])

    def test_empty_list_raises(self) -> None:
        with self.assertRaises(PowerBIValidationError):
            _resolve_ti_patterns([])

    def test_unknown_pattern_raises(self) -> None:
        with self.assertRaises(PowerBIValidationError):
            _resolve_ti_patterns(["WTD"])

    def test_yoy_percent_expands_dependencies_in_order(self) -> None:
        self.assertEqual(_resolve_ti_patterns(["YOY%"]), ["SPY", "YOY", "YOY%"])

    def test_duplicates_are_deduped_case_insensitively(self) -> None:
        self.assertEqual(_resolve_ti_patterns(["ytd", "YTD", " Ytd "]), ["YTD"])


class TimeIntelligencePackTests(unittest.TestCase):
    def test_dry_run_reports_plan_without_mutation(self) -> None:
        manager = make_manager(measures=[make_measure("Total Sales")])
        result = pbi_create_time_intelligence_pack_tool(
            manager,
            table="Sales",
            base_measure="Total Sales",
            date_table="Date",
            date_column="Date",
            dry_run=True,
        )
        self.assertTrue(result["ok"], result)
        self.assertTrue(result["dry_run"])
        self.assertEqual(len(result["plan"]), 7)
        specs = {spec["name"]: spec for spec in result["measures"]}
        self.assertEqual(
            specs["Total Sales YTD"]["expression"],
            "CALCULATE([Total Sales], DATESYTD('Date'[Date]))",
        )
        self.assertEqual(
            specs["Total Sales YOY"]["expression"],
            "[Total Sales] - [Total Sales SPY]",
        )
        self.assertEqual(specs["Total Sales YOY %"]["format_string"], "0.00%")
        self.assertEqual(manager.write_calls, 0)
        self.assertEqual(manager.database.Model.Tables[0].Measures.Count, 1)

    def test_pack_creates_measures_and_inherits_format(self) -> None:
        manager = make_manager(measures=[make_measure("Total Sales", format_string="#,0")])
        result = pbi_create_time_intelligence_pack_tool(
            manager,
            table="Sales",
            base_measure="Total Sales",
            date_table="Date",
            date_column="Date",
            patterns=["YTD", "SPY"],
        )
        self.assertTrue(result["ok"], result)
        self.assertEqual(result["created"], 2)
        self.assertEqual(result["failed"], 0)
        table = manager.database.Model.Tables[0]
        ytd = table.Measures.Find("Total Sales YTD")
        self.assertIsNotNone(ytd)
        self.assertEqual(ytd.FormatString, "#,0")
        self.assertEqual(ytd.DisplayFolder, "Time intelligence")

    def test_pack_unknown_pattern_raises(self) -> None:
        manager = make_manager(measures=[make_measure("Total Sales")])
        with self.assertRaises(PowerBIValidationError):
            pbi_create_time_intelligence_pack_tool(
                manager,
                table="Sales",
                base_measure="Total Sales",
                date_table="Date",
                date_column="Date",
                patterns=["BOGUS"],
            )


class TemplateGeneratorTests(unittest.TestCase):
    def test_variance_measure_dax(self) -> None:
        manager = make_manager()
        result = pbi_create_variance_measure_tool(
            manager,
            table="Sales",
            base_measure="Total Sales",
            date_table="Date",
            date_column="Date",
        )
        self.assertTrue(result["ok"], result)
        measure = manager.database.Model.Tables[0].Measures.Find("Total Sales Variance")
        self.assertEqual(
            measure.Expression,
            "[Total Sales] - CALCULATE([Total Sales], DATEADD('Date'[Date], -1, YEAR))",
        )

    def test_variance_measure_bad_granularity(self) -> None:
        manager = make_manager()
        with self.assertRaises(PowerBIValidationError):
            pbi_create_variance_measure_tool(
                manager,
                table="Sales",
                base_measure="Total Sales",
                date_table="Date",
                date_column="Date",
                granularity="week",
            )

    def test_contribution_measure_dax(self) -> None:
        manager = make_manager()
        result = pbi_create_contribution_measure_tool(
            manager,
            table="Sales",
            base_measure="Total Sales",
            scope_columns=["Product.Category"],
        )
        self.assertTrue(result["ok"], result)
        measure = manager.database.Model.Tables[0].Measures.Find("Total Sales % of total")
        self.assertEqual(
            measure.Expression,
            "DIVIDE([Total Sales], CALCULATE([Total Sales], ALL('Product'[Category])))",
        )
        self.assertEqual(measure.FormatString, "0.00%")

    def test_contribution_measure_requires_dotted_scope_columns(self) -> None:
        manager = make_manager()
        with self.assertRaises(PowerBIValidationError):
            pbi_create_contribution_measure_tool(
                manager, table="Sales", base_measure="Total Sales", scope_columns=["Category"]
            )

    def test_contribution_measure_requires_scope_columns(self) -> None:
        manager = make_manager()
        with self.assertRaises(PowerBIValidationError):
            pbi_create_contribution_measure_tool(manager, table="Sales", base_measure="Total Sales", scope_columns=[])

    def test_topn_measure_dax(self) -> None:
        manager = make_manager()
        result = pbi_create_topn_measure_tool(
            manager,
            table="Sales",
            base_measure="Total Sales",
            n=5,
            dimension_table="Product",
            dimension_column="Category",
        )
        self.assertTrue(result["ok"], result)
        measure = manager.database.Model.Tables[0].Measures.Find("Total Sales Top 5")
        self.assertEqual(
            measure.Expression,
            "IF(RANKX(ALL('Product'[Category]), [Total Sales], , DESC) <= 5, [Total Sales], BLANK())",
        )

    def test_topn_measure_rejects_non_positive_n(self) -> None:
        manager = make_manager()
        with self.assertRaises(PowerBIValidationError):
            pbi_create_topn_measure_tool(
                manager,
                table="Sales",
                base_measure="Total Sales",
                n=0,
                dimension_table="Product",
                dimension_column="Category",
            )

    def test_rolling_average_measure_dax(self) -> None:
        manager = make_manager()
        result = pbi_create_rolling_average_measure_tool(
            manager,
            table="Sales",
            base_measure="Total Sales",
            window=3,
            date_table="Date",
            date_column="Date",
        )
        self.assertTrue(result["ok"], result)
        measure = manager.database.Model.Tables[0].Measures.Find("Total Sales Rolling 3 Month")
        self.assertEqual(
            measure.Expression,
            "AVERAGEX(DATESINPERIOD('Date'[Date], LASTDATE('Date'[Date]), -3, MONTH), [Total Sales])",
        )

    def test_rolling_average_rejects_bad_window_and_granularity(self) -> None:
        manager = make_manager()
        with self.assertRaises(PowerBIValidationError):
            pbi_create_rolling_average_measure_tool(
                manager,
                table="Sales",
                base_measure="Total Sales",
                window=0,
                date_table="Date",
                date_column="Date",
            )
        with self.assertRaises(PowerBIValidationError):
            pbi_create_rolling_average_measure_tool(
                manager,
                table="Sales",
                base_measure="Total Sales",
                window=3,
                date_table="Date",
                date_column="Date",
                granularity="week",
            )


class DaxFileParsingTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp_dir = tempfile.TemporaryDirectory()
        self.root = Path(self.temp_dir.name)
        self.previous_allowed = [str(item) for item in SECURITY.allowed_base_dirs()]
        SECURITY.configure_allowed_dirs([str(self.root)])

    def tearDown(self) -> None:
        SECURITY.configure_allowed_dirs(self.previous_allowed)
        self.temp_dir.cleanup()

    def _write(self, content: str, name: str = "measures.dax") -> Path:
        path = self.root / name
        path.write_text(content, encoding="utf-8")
        return path

    def test_parse_dax_file_multiple_blocks_with_comments(self) -> None:
        path = self._write(
            "// header comment\n"
            "Total Sales = SUM(Sales[Amount])\n"
            "\n"
            "Margin % =\n"
            "DIVIDE(\n"
            "    [Profit], /* block comment */\n"
            "    [Total Sales]\n"
            ")\n"
        )
        parsed = _parse_dax_file(path)
        self.assertEqual(len(parsed), 2)
        self.assertEqual(parsed[0], {"name": "Total Sales", "expression": "SUM(Sales[Amount])"})
        self.assertEqual(parsed[1]["name"], "Margin %")
        self.assertIn("DIVIDE(", parsed[1]["expression"])
        self.assertNotIn("block comment", parsed[1]["expression"])

    def test_parse_dax_file_empty_raises(self) -> None:
        path = self._write("// only comments\n\n// nothing else\n")
        with self.assertRaises(PowerBIValidationError):
            _parse_dax_file(path)

    def test_parse_dax_file_invalid_header_raises(self) -> None:
        path = self._write("NotAMeasureHeader\nSUM(Sales[Amount])\n")
        with self.assertRaises(PowerBIValidationError):
            _parse_dax_file(path)

    def test_parse_dax_file_rejects_bracketed_measure_name(self) -> None:
        path = self._write("Bad[Name] = 1\n")
        with self.assertRaises(PowerBIValidationError):
            _parse_dax_file(path)

    def test_parse_dax_file_wrong_extension_raises(self) -> None:
        path = self._write("Total = 1\n", name="measures.txt")
        with self.assertRaises(PowerBIValidationError):
            _parse_dax_file(path)

    def test_import_dax_file_tool_creates_measures(self) -> None:
        path = self._write("Total Sales = SUM(Sales[Amount])\n\nProfit = SUM(Sales[Profit])\n")
        manager = make_manager()
        result = pbi_import_dax_file_tool(manager, path=str(path), table="Sales")
        self.assertTrue(result["ok"], result)
        self.assertEqual(result["parsed_count"], 2)
        self.assertEqual(result["created"], 2)
        self.assertEqual(result["failed"], 0)
        self.assertIsNotNone(manager.database.Model.Tables[0].Measures.Find("Profit"))

    def test_import_dax_file_tool_stop_on_error(self) -> None:
        path = self._write("Existing = 1\n\nNever Created = 2\n")
        manager = make_manager(measures=[make_measure("Existing")])
        result = pbi_import_dax_file_tool(manager, path=str(path), table="Sales", overwrite=False, stop_on_error=True)
        self.assertTrue(result["ok"], result)
        self.assertEqual(result["failed"], 1)
        self.assertEqual(len(result["results"]), 1)
        self.assertFalse(result["results"][0]["ok"])
        self.assertIsNone(manager.database.Model.Tables[0].Measures.Find("Never Created"))


class StripDaxCommentsTests(unittest.TestCase):
    def test_strips_line_and_block_comments(self) -> None:
        text = "A = 1 // trailing\n/* block\nspanning */\nB = 2\n"
        stripped = _strip_dax_comments(text)
        self.assertNotIn("trailing", stripped)
        self.assertNotIn("block", stripped)
        self.assertIn("A = 1", stripped)
        self.assertIn("B = 2", stripped)

    def test_preserves_comment_markers_inside_strings(self) -> None:
        text = 'M = "https://example.com" & "/* not a comment */"'
        self.assertEqual(_strip_dax_comments(text), text)

    def test_preserves_escaped_quotes_inside_strings(self) -> None:
        text = 'M = "say ""hi"" // still text"'
        self.assertEqual(_strip_dax_comments(text), text)


if __name__ == "__main__":
    unittest.main(verbosity=2)

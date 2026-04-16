"""Standalone tests for Power Query tools."""

from __future__ import annotations

import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

from pbi_connection import PowerBIValidationError
from tools.power_query import (
    _build_csv_m,
    _build_excel_m,
    _build_folder_m,
    _validate_m_expression,
    pbi_bulk_import_excel_tool,
    pbi_list_power_queries_tool,
    pbi_set_power_query_tool,
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


class FakeMPartitionSource:
    def __init__(self, expression: str = "") -> None:
        self.Expression = expression


class FakeQueryPartitionSource:
    def __init__(self, query: str = "") -> None:
        self.Query = query


class FakeCalculatedPartitionSource:
    def __init__(self, expression: str = "") -> None:
        self.Expression = expression


class FakePartition:
    def __init__(self, name: str, source, source_type: str | None = None) -> None:
        self.Name = name
        self.Source = source
        self.SourceType = source_type or type(source).__name__.replace("PartitionSource", "")


class FakeTable:
    def __init__(self, name: str, partitions, *, is_hidden: bool = False) -> None:
        self.Name = name
        self.IsHidden = is_hidden
        self.Partitions = FakeCollection(partitions)
        self.refresh_requests: list[str] = []

    def RequestRefresh(self, refresh_type: str) -> None:
        self.refresh_requests.append(refresh_type)


class FakeModel:
    def __init__(self, tables) -> None:
        self.Tables = FakeCollection(tables)


class FakeDatabase:
    def __init__(self, model: FakeModel, compatibility_level: int = 1500) -> None:
        self.Model = model
        self.CompatibilityLevel = compatibility_level


class FakeTom:
    class MPartitionSource(FakeMPartitionSource):
        pass

    class RefreshType:
        Full = "FULL"


class FakeManager:
    def __init__(self, model: FakeModel, compatibility_level: int = 1500) -> None:
        self.tom = FakeTom()
        self.database = FakeDatabase(model, compatibility_level=compatibility_level)
        self.state = SimpleNamespace(
            database=self.database,
            snapshot=lambda: {"connected": True, "database": "UnitTest"},
        )

    def run_read(self, _operation_name, reader):
        return reader(self.state)

    def execute_write(self, _operation_name, mutator):
        payload = mutator(self.state, self.database, self.database.Model)
        payload["save_result"] = {"status": "saved"}
        payload["connection"] = self.state.snapshot()
        return payload


class PowerQueryToolTests(unittest.TestCase):
    def test_build_excel_m_handles_quotes_spaces_and_unicode(self) -> None:
        expression = _build_excel_m(
            'C:\\Data Files\\Vélo "Premium".xlsx',
            'Q1 "VTT" spécial',
            promote_headers=True,
        )
        self.assertIn('File.Contents("C:\\Data Files\\Vélo ""Premium"".xlsx")', expression)
        self.assertIn('[Item="Q1 ""VTT"" spécial",Kind="Sheet"]', expression)
        _validate_m_expression(expression)

    def test_build_csv_and_folder_queries_validate(self) -> None:
        csv_expression = _build_csv_m("/tmp/sales data.csv", delimiter=";", quote_style="none")
        folder_expression = _build_folder_m("/tmp/Vélos source", extension_filter="csv")
        self.assertIn("Csv.Document(", csv_expression)
        self.assertIn("QuoteStyle.None", csv_expression)
        self.assertIn('Folder.Files("/tmp/Vélos source")', folder_expression)
        self.assertIn('Text.Lower([Extension]) = ".csv"', folder_expression)
        _validate_m_expression(csv_expression)
        _validate_m_expression(folder_expression)

    def test_validate_m_expression_rejects_broken_syntax(self) -> None:
        with self.assertRaises(PowerBIValidationError):
            _validate_m_expression('let Source = Csv.Document(File.Contents("broken.csv") in Source')

    def test_partition_write_replaces_query_source_with_m(self) -> None:
        partition = FakePartition("Sales", FakeQueryPartitionSource("select * from Sales"), source_type="Query")
        table = FakeTable("Sales", [partition])
        manager = FakeManager(FakeModel([table]))
        expression = "let\n    Source = 1\nin\n    Source"

        result = pbi_set_power_query_tool(manager, table="Sales", m_expression=expression, refresh_after=True)

        self.assertTrue(result["ok"], result)
        self.assertIsInstance(partition.Source, FakeTom.MPartitionSource)
        self.assertEqual(partition.Source.Expression, expression)
        self.assertEqual(table.refresh_requests, ["FULL"])
        self.assertEqual(result["query"]["source_type"], "m")

    def test_set_power_query_rejects_calculated_partition(self) -> None:
        partition = FakePartition("Calc", FakeCalculatedPartitionSource("ROW(1)"), source_type="Calculated")
        table = FakeTable("Calc", [partition])
        manager = FakeManager(FakeModel([table]))
        with self.assertRaises(PowerBIValidationError):
            pbi_set_power_query_tool(manager, table="Calc", m_expression="let\n    Source = 1\nin\n    Source")

    def test_bulk_import_excel_auto_map_skips_hidden_and_multi_partition(self) -> None:
        sales = FakeTable("Sales", [FakePartition("Sales", FakeQueryPartitionSource("select 1"), source_type="Query")])
        hidden = FakeTable("HiddenTable", [FakePartition("HiddenTable", FakeQueryPartitionSource("select 1"))], is_hidden=True)
        sharded = FakeTable(
            "Sharded",
            [
                FakePartition("Sharded_2025", FakeQueryPartitionSource("select 1")),
                FakePartition("Sharded_2026", FakeQueryPartitionSource("select 2")),
            ],
        )
        manager = FakeManager(FakeModel([sales, hidden, sharded]))

        with patch("tools.power_query.inspect_excel_archive", return_value=Path("dummy.xlsx")), patch(
            "tools.power_query._load_excel_sheet_names",
            return_value=["Sales", "HiddenTable", "Sharded", "Missing"],
        ):
            result = pbi_bulk_import_excel_tool(manager, excel_path="dummy.xlsx")

        self.assertTrue(result["ok"], result)
        self.assertEqual(result["created"], 1)
        self.assertEqual(result["skipped"], 1)
        self.assertEqual(sales.Partitions[0].Source.Expression.count('Item="Sales"'), 1)
        reasons = {item["table"]: item["reason"] for item in result["results"] if item["status"] == "skipped"}
        self.assertEqual(reasons["Sharded"], "multiple_partitions")
        self.assertNotIn("HiddenTable", {item["table"] for item in result["results"]})

    def test_list_power_queries_reports_partition_metadata(self) -> None:
        sales = FakeTable("Sales", [FakePartition("Sales", FakeMPartitionSource("let\nin\n    Source"), source_type="M")])
        manager = FakeManager(FakeModel([sales]))
        result = pbi_list_power_queries_tool(manager)
        self.assertTrue(result["ok"], result)
        self.assertEqual(result["queries"][0]["partitions"][0]["source_type"], "m")


if __name__ == "__main__":
    unittest.main(verbosity=2)

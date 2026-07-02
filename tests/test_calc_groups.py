"""Offline tests for calculation group tools (src/tools/calc_groups.py) with a faked TOM layer."""

from __future__ import annotations

import sys
import unittest
from pathlib import Path
from types import SimpleNamespace

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

from pbi_connection import (
    PowerBIDuplicateError,
    PowerBINotFoundError,
    PowerBIValidationError,
)
from tools.calc_groups import (
    pbi_create_calc_group_tool,
    pbi_delete_calc_group_tool,
    pbi_list_calc_groups_tool,
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

    def RemoveAt(self, index: int) -> None:
        self.pop(index)


class FakeCalculationGroup:
    def __init__(self) -> None:
        self.Precedence = 0
        self.Description = ""
        self.CalculationItems = FakeCollection()


class FakeCalculationItem:
    def __init__(self) -> None:
        self.Name = ""
        self.Expression = ""
        self.Ordinal = -1
        self.FormatStringDefinition = None


class FakeFormatStringDefinition:
    def __init__(self) -> None:
        self.Expression = ""


class FakeDataColumn:
    def __init__(self) -> None:
        self.Name = ""
        self.DataType = None
        self.SourceColumn = ""


class FakePartition:
    def __init__(self) -> None:
        self.Name = ""
        self.Source = None


class FakeCalculationGroupSource:
    pass


class FakeTable:
    def __init__(self, name: str = "") -> None:
        self.Name = name
        self.CalculationGroup = None
        self.Columns = FakeCollection()
        self.Partitions = FakeCollection()


class FakeModel:
    def __init__(self, tables=()) -> None:
        self.Tables = FakeCollection(tables)
        self.DiscourageImplicitMeasures = False


class FakeTom:
    Table = FakeTable
    CalculationGroup = FakeCalculationGroup
    CalculationItem = FakeCalculationItem
    FormatStringDefinition = FakeFormatStringDefinition
    DataColumn = FakeDataColumn
    Partition = FakePartition
    CalculationGroupSource = FakeCalculationGroupSource
    DataType = SimpleNamespace(String="String")


class FakeManager:
    def __init__(self, model: FakeModel) -> None:
        self.tom = FakeTom()
        self.database = SimpleNamespace(Model=model)
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


def make_calc_group_table(name: str) -> FakeTable:
    table = FakeTable(name)
    table.CalculationGroup = FakeCalculationGroup()
    column = FakeDataColumn()
    column.Name = "Name"
    column.SourceColumn = "Name"
    table.Columns.Add(column)
    return table


class CreateCalcGroupTests(unittest.TestCase):
    def test_create_calc_group_happy_path(self) -> None:
        model = FakeModel()
        manager = FakeManager(model)
        result = pbi_create_calc_group_tool(
            manager,
            table_name="Time Calcs",
            column_name="Calculation",
            precedence=10,
            items=[
                {"name": "Current", "expression": "SELECTEDMEASURE()"},
                {
                    "name": "YTD",
                    "expression": "CALCULATE(SELECTEDMEASURE(), DATESYTD('Date'[Date]))",
                    "format_string_expression": '"0.00%"',
                    "ordinal": 5,
                },
            ],
        )
        self.assertTrue(result["ok"], result)
        self.assertEqual(result["action"], "created")
        table = model.Tables.Find("Time Calcs")
        self.assertIsNotNone(table)
        self.assertTrue(model.DiscourageImplicitMeasures)
        group = table.CalculationGroup
        self.assertEqual(group.Precedence, 10)
        self.assertEqual(group.CalculationItems.Count, 2)
        self.assertEqual(group.CalculationItems[0].Ordinal, 0)
        self.assertEqual(group.CalculationItems[1].Ordinal, 5)
        self.assertEqual(group.CalculationItems[1].FormatStringDefinition.Expression, '"0.00%"')
        # Exactly one String column whose SourceColumn is the literal "Name".
        self.assertEqual(table.Columns.Count, 1)
        self.assertEqual(table.Columns[0].Name, "Calculation")
        self.assertEqual(table.Columns[0].SourceColumn, "Name")
        self.assertEqual(table.Columns[0].DataType, "String")
        # One partition backed by a CalculationGroupSource.
        self.assertEqual(table.Partitions.Count, 1)
        self.assertIsInstance(table.Partitions[0].Source, FakeCalculationGroupSource)

    def test_create_calc_group_duplicate_raises(self) -> None:
        model = FakeModel(tables=[make_calc_group_table("Time Calcs")])
        manager = FakeManager(model)
        with self.assertRaises(PowerBIDuplicateError):
            pbi_create_calc_group_tool(manager, table_name="time calcs")

    def test_create_calc_group_overwrite_replaces_items(self) -> None:
        table = make_calc_group_table("Time Calcs")
        old_item = FakeCalculationItem()
        old_item.Name = "Old"
        table.CalculationGroup.CalculationItems.Add(old_item)
        model = FakeModel(tables=[table])
        manager = FakeManager(model)

        result = pbi_create_calc_group_tool(
            manager,
            table_name="Time Calcs",
            items=[{"name": "New", "expression": "SELECTEDMEASURE()"}],
            overwrite=True,
        )

        self.assertTrue(result["ok"], result)
        self.assertEqual(result["action"], "updated")
        items = table.CalculationGroup.CalculationItems
        self.assertEqual(items.Count, 1)
        self.assertEqual(items[0].Name, "New")

    def test_create_calc_group_refuses_overwriting_regular_table(self) -> None:
        model = FakeModel(tables=[FakeTable("Sales")])
        manager = FakeManager(model)
        with self.assertRaises(PowerBIValidationError):
            pbi_create_calc_group_tool(manager, table_name="Sales", overwrite=True)

    def test_create_calc_group_item_missing_expression(self) -> None:
        manager = FakeManager(FakeModel())
        with self.assertRaises(PowerBIValidationError):
            pbi_create_calc_group_tool(manager, table_name="Time Calcs", items=[{"name": "Broken"}])

    def test_create_calc_group_item_query_only_expression(self) -> None:
        manager = FakeManager(FakeModel())
        with self.assertRaises(PowerBIValidationError):
            pbi_create_calc_group_tool(
                manager,
                table_name="Time Calcs",
                items=[{"name": "Bad", "expression": "EVALUATE VALUES(Sales)"}],
            )

    def test_create_calc_group_empty_table_name(self) -> None:
        manager = FakeManager(FakeModel())
        with self.assertRaises(PowerBIValidationError):
            pbi_create_calc_group_tool(manager, table_name="  ")


class ListCalcGroupsTests(unittest.TestCase):
    def test_list_calc_groups_sorts_items_and_skips_regular_tables(self) -> None:
        table = make_calc_group_table("Time Calcs")
        table.CalculationGroup.Precedence = 20
        second = FakeCalculationItem()
        second.Name = "YTD"
        second.Ordinal = 1
        first = FakeCalculationItem()
        first.Name = "Current"
        first.Ordinal = 0
        table.CalculationGroup.CalculationItems.Add(second)
        table.CalculationGroup.CalculationItems.Add(first)
        model = FakeModel(tables=[FakeTable("Sales"), table])
        manager = FakeManager(model)

        result = pbi_list_calc_groups_tool(manager)

        self.assertTrue(result["ok"], result)
        self.assertEqual(len(result["calc_groups"]), 1)
        group = result["calc_groups"][0]
        self.assertEqual(group["table"], "Time Calcs")
        self.assertEqual(group["precedence"], 20)
        self.assertEqual(group["column_name"], "Name")
        self.assertEqual([item["name"] for item in group["items"]], ["Current", "YTD"])

    def test_list_calc_groups_empty_model(self) -> None:
        manager = FakeManager(FakeModel())
        result = pbi_list_calc_groups_tool(manager)
        self.assertTrue(result["ok"], result)
        self.assertEqual(result["calc_groups"], [])


class DeleteCalcGroupTests(unittest.TestCase):
    def test_delete_calc_group_happy_path(self) -> None:
        model = FakeModel(tables=[make_calc_group_table("Time Calcs")])
        manager = FakeManager(model)
        result = pbi_delete_calc_group_tool(manager, table_name="Time Calcs")
        self.assertTrue(result["ok"], result)
        self.assertEqual(result["deleted_calc_group"], {"table": "Time Calcs"})
        self.assertEqual(model.Tables.Count, 0)

    def test_delete_calc_group_not_found(self) -> None:
        manager = FakeManager(FakeModel())
        with self.assertRaises(PowerBINotFoundError):
            pbi_delete_calc_group_tool(manager, table_name="Ghost")

    def test_delete_calc_group_rejects_regular_table(self) -> None:
        model = FakeModel(tables=[FakeTable("Sales")])
        manager = FakeManager(model)
        with self.assertRaises(PowerBIValidationError):
            pbi_delete_calc_group_tool(manager, table_name="Sales")
        self.assertEqual(model.Tables.Count, 1)


if __name__ == "__main__":
    unittest.main(verbosity=2)

"""Offline tests for relationship tools (src/tools/relationships.py) with a faked TOM layer."""

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
from tools.relationships import (
    pbi_create_relationship_tool,
    pbi_delete_relationship_tool,
    pbi_list_relationships_tool,
    pbi_update_relationship_tool,
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


class FakeColumn:
    def __init__(self, name: str, table) -> None:
        self.Name = name
        self.Table = table


class FakeTable:
    def __init__(self, name: str, columns=()) -> None:
        self.Name = name
        self.Columns = FakeCollection()
        for column_name in columns:
            self.Columns.Add(FakeColumn(column_name, self))


class FakeSingleColumnRelationship:
    def __init__(self) -> None:
        self.Name = ""
        self.FromColumn = None
        self.ToColumn = None
        self.FromCardinality = None
        self.ToCardinality = None
        self.CrossFilteringBehavior = None
        self.IsActive = True


class FakeModel:
    def __init__(self, tables=(), relationships=()) -> None:
        self.Tables = FakeCollection(tables)
        self.Relationships = FakeCollection(relationships)


class FakeTom:
    SingleColumnRelationship = FakeSingleColumnRelationship
    RelationshipEndCardinality = SimpleNamespace(Many="Many", One="One")
    CrossFilteringBehavior = SimpleNamespace(
        OneDirection="OneDirection",
        BothDirections="BothDirections",
        Automatic="Automatic",
    )


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

    def cached_run_read(self, _cache_key, _operation_name, reader):
        return reader(self.state)

    def execute_write(self, _operation_name, mutator):
        payload = mutator(self.state, self.database, self.database.Model)
        payload["save_result"] = {"status": "saved"}
        payload["connection"] = self.state.snapshot()
        return payload


def make_manager() -> FakeManager:
    sales = FakeTable("Sales", columns=["ProductID", "Amount"])
    product = FakeTable("Product", columns=["ProductID", "Category"])
    return FakeManager(FakeModel(tables=[sales, product]))


def create_default(manager: FakeManager, **kwargs):
    params = dict(
        from_table="Sales",
        from_column="ProductID",
        to_table="Product",
        to_column="ProductID",
    )
    params.update(kwargs)
    return pbi_create_relationship_tool(manager, **params)


class CreateRelationshipTests(unittest.TestCase):
    def test_create_relationship_happy_path(self) -> None:
        manager = make_manager()
        result = create_default(manager)
        self.assertTrue(result["ok"], result)
        self.assertEqual(result["action"], "created")
        model = manager.database.Model
        self.assertEqual(model.Relationships.Count, 1)
        relationship = model.Relationships[0]
        self.assertEqual(relationship.Name, "Sales_ProductID__Product_ProductID")
        self.assertEqual(relationship.FromCardinality, "Many")
        self.assertEqual(relationship.ToCardinality, "One")
        self.assertEqual(relationship.CrossFilteringBehavior, "OneDirection")
        self.assertTrue(relationship.IsActive)

    def test_create_relationship_custom_name_and_inactive(self) -> None:
        manager = make_manager()
        result = create_default(manager, relationship_name="Rel Sales-Product", is_active=False)
        self.assertTrue(result["ok"], result)
        relationship = manager.database.Model.Relationships[0]
        self.assertEqual(relationship.Name, "Rel Sales-Product")
        self.assertFalse(relationship.IsActive)

    def test_create_one_to_one_forces_both_directions(self) -> None:
        manager = make_manager()
        result = create_default(manager, cardinality="oneToOne", direction="oneDirection")
        self.assertTrue(result["ok"], result)
        relationship = manager.database.Model.Relationships[0]
        self.assertEqual(relationship.FromCardinality, "One")
        self.assertEqual(relationship.ToCardinality, "One")
        self.assertEqual(relationship.CrossFilteringBehavior, "BothDirections")

    def test_create_many_to_many(self) -> None:
        manager = make_manager()
        result = create_default(manager, cardinality="manyToMany", direction="both")
        self.assertTrue(result["ok"], result)
        relationship = manager.database.Model.Relationships[0]
        self.assertEqual(relationship.FromCardinality, "Many")
        self.assertEqual(relationship.ToCardinality, "Many")
        self.assertEqual(relationship.CrossFilteringBehavior, "BothDirections")

    def test_create_duplicate_same_direction_raises(self) -> None:
        manager = make_manager()
        create_default(manager)
        with self.assertRaises(PowerBIDuplicateError):
            create_default(manager)

    def test_create_duplicate_reverse_direction_raises(self) -> None:
        manager = make_manager()
        create_default(manager)
        with self.assertRaises(PowerBIDuplicateError):
            create_default(
                manager,
                from_table="Product",
                from_column="ProductID",
                to_table="Sales",
                to_column="ProductID",
            )

    def test_create_overwrite_updates_existing(self) -> None:
        manager = make_manager()
        create_default(manager)
        result = create_default(manager, direction="bothDirections", is_active=False, overwrite=True)
        self.assertTrue(result["ok"], result)
        self.assertEqual(result["action"], "updated")
        model = manager.database.Model
        self.assertEqual(model.Relationships.Count, 1)
        self.assertEqual(model.Relationships[0].CrossFilteringBehavior, "BothDirections")
        self.assertFalse(model.Relationships[0].IsActive)

    def test_create_relationship_table_not_found(self) -> None:
        manager = make_manager()
        with self.assertRaises(PowerBINotFoundError):
            create_default(manager, from_table="Ghost")
        with self.assertRaises(PowerBINotFoundError):
            create_default(manager, to_table="Ghost")

    def test_create_relationship_column_not_found(self) -> None:
        manager = make_manager()
        with self.assertRaises(PowerBINotFoundError):
            create_default(manager, from_column="Ghost")
        with self.assertRaises(PowerBINotFoundError):
            create_default(manager, to_column="Ghost")

    def test_create_relationship_bad_cardinality(self) -> None:
        manager = make_manager()
        with self.assertRaises(PowerBIValidationError):
            create_default(manager, cardinality="oneToNothing")

    def test_create_relationship_bad_direction(self) -> None:
        manager = make_manager()
        with self.assertRaises(PowerBIValidationError):
            create_default(manager, direction="sideways")

    def test_create_relationship_empty_table_name(self) -> None:
        manager = make_manager()
        with self.assertRaises(PowerBIValidationError):
            create_default(manager, from_table="   ")


class ListRelationshipsTests(unittest.TestCase):
    def test_list_relationships_reports_endpoints(self) -> None:
        manager = make_manager()
        create_default(manager)
        result = pbi_list_relationships_tool(manager)
        self.assertTrue(result["ok"], result)
        self.assertEqual(len(result["relationships"]), 1)
        item = result["relationships"][0]
        self.assertEqual(item["from_table"], "Sales")
        self.assertEqual(item["from_column"], "ProductID")
        self.assertEqual(item["to_table"], "Product")
        self.assertEqual(item["to_column"], "ProductID")
        self.assertEqual(item["cardinality"], "ManyToOne")
        self.assertTrue(item["is_active"])


class DeleteRelationshipTests(unittest.TestCase):
    def test_delete_by_name(self) -> None:
        manager = make_manager()
        create_default(manager, relationship_name="RelA")
        result = pbi_delete_relationship_tool(manager, name="rela")
        self.assertTrue(result["ok"], result)
        self.assertEqual(result["deleted_relationship"], {"name": "RelA"})
        self.assertEqual(manager.database.Model.Relationships.Count, 0)

    def test_delete_by_endpoints(self) -> None:
        manager = make_manager()
        create_default(manager)
        result = pbi_delete_relationship_tool(
            manager,
            from_table="Sales",
            from_column="ProductID",
            to_table="Product",
            to_column="ProductID",
        )
        self.assertTrue(result["ok"], result)
        self.assertEqual(manager.database.Model.Relationships.Count, 0)

    def test_delete_requires_name_or_full_endpoints(self) -> None:
        manager = make_manager()
        with self.assertRaises(PowerBIValidationError):
            pbi_delete_relationship_tool(manager, from_table="Sales", from_column="ProductID")

    def test_delete_not_found(self) -> None:
        manager = make_manager()
        with self.assertRaises(PowerBINotFoundError):
            pbi_delete_relationship_tool(manager, name="Ghost")


class UpdateRelationshipTests(unittest.TestCase):
    def test_update_all_properties(self) -> None:
        manager = make_manager()
        create_default(manager, relationship_name="RelA")
        result = pbi_update_relationship_tool(
            manager,
            name="RelA",
            cardinality="manyToMany",
            direction="bothDirections",
            is_active=False,
            new_name="RelB",
        )
        self.assertTrue(result["ok"], result)
        relationship = manager.database.Model.Relationships[0]
        self.assertEqual(relationship.Name, "RelB")
        self.assertEqual(relationship.FromCardinality, "Many")
        self.assertEqual(relationship.ToCardinality, "Many")
        self.assertEqual(relationship.CrossFilteringBehavior, "BothDirections")
        self.assertFalse(relationship.IsActive)

    def test_update_by_endpoints(self) -> None:
        manager = make_manager()
        create_default(manager)
        result = pbi_update_relationship_tool(
            manager,
            from_table="Sales",
            from_column="ProductID",
            to_table="Product",
            to_column="ProductID",
            is_active=False,
        )
        self.assertTrue(result["ok"], result)
        self.assertFalse(manager.database.Model.Relationships[0].IsActive)

    def test_update_requires_at_least_one_change(self) -> None:
        manager = make_manager()
        create_default(manager, relationship_name="RelA")
        with self.assertRaises(PowerBIValidationError):
            pbi_update_relationship_tool(manager, name="RelA")

    def test_update_requires_identifier(self) -> None:
        manager = make_manager()
        with self.assertRaises(PowerBIValidationError):
            pbi_update_relationship_tool(manager, is_active=False)

    def test_update_not_found(self) -> None:
        manager = make_manager()
        with self.assertRaises(PowerBINotFoundError):
            pbi_update_relationship_tool(manager, name="Ghost", is_active=False)


if __name__ == "__main__":
    unittest.main(verbosity=2)

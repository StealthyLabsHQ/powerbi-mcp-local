"""Offline tests for RLS tools (src/tools/rls.py) with a fully faked TOM layer."""

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
from tools.rls import (
    pbi_add_role_member_tool,
    pbi_create_role_tool,
    pbi_delete_role_tool,
    pbi_list_roles_tool,
    pbi_remove_role_member_tool,
    pbi_set_role_filter_tool,
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


class FakeModelRole:
    def __init__(self) -> None:
        self.Name = ""
        self.Description = ""
        self.ModelPermission = None
        self.Members = FakeCollection()
        self.TablePermissions = FakeCollection()


class FakeTablePermission:
    def __init__(self) -> None:
        self.Table = None
        self.FilterExpression = ""


class FakeExternalModelRoleMember:
    def __init__(self) -> None:
        self.MemberName = ""
        self.IdentityProvider = ""


class FakeWindowsModelRoleMember:
    def __init__(self) -> None:
        self.MemberName = ""


class FakeTable:
    def __init__(self, name: str) -> None:
        self.Name = name


class FakeModel:
    def __init__(self, roles=(), tables=()) -> None:
        self.Roles = FakeCollection(roles)
        self.Tables = FakeCollection(tables)


class FakeTom:
    ModelRole = FakeModelRole
    TablePermission = FakeTablePermission
    ExternalModelRoleMember = FakeExternalModelRoleMember
    WindowsModelRoleMember = FakeWindowsModelRoleMember
    ModelPermission = type(
        "ModelPermission",
        (),
        {
            "None": "None",
            "Read": "Read",
            "ReadRefresh": "ReadRefresh",
            "Refresh": "Refresh",
            "Administrator": "Administrator",
        },
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

    def execute_write(self, _operation_name, mutator):
        payload = mutator(self.state, self.database, self.database.Model)
        payload["save_result"] = {"status": "saved"}
        payload["connection"] = self.state.snapshot()
        return payload


def make_role(name: str, permission: str = "Read") -> FakeModelRole:
    role = FakeModelRole()
    role.Name = name
    role.ModelPermission = permission
    return role


class CreateRoleTests(unittest.TestCase):
    def test_create_role_happy_path(self) -> None:
        model = FakeModel()
        manager = FakeManager(model)
        result = pbi_create_role_tool(manager, name="Sales FR", permission="read", description="French sales")
        self.assertTrue(result["ok"], result)
        self.assertEqual(result["action"], "created")
        role = model.Roles.Find("Sales FR")
        self.assertIsNotNone(role)
        self.assertEqual(role.ModelPermission, "Read")
        self.assertEqual(role.Description, "French sales")

    def test_create_role_duplicate_raises(self) -> None:
        model = FakeModel(roles=[make_role("Sales FR")])
        manager = FakeManager(model)
        with self.assertRaises(PowerBIDuplicateError):
            pbi_create_role_tool(manager, name="sales fr")

    def test_create_role_overwrite_updates(self) -> None:
        model = FakeModel(roles=[make_role("Sales FR")])
        manager = FakeManager(model)
        result = pbi_create_role_tool(manager, name="Sales FR", permission="admin", overwrite=True)
        self.assertTrue(result["ok"], result)
        self.assertEqual(result["action"], "updated")
        self.assertEqual(model.Roles.Count, 1)
        self.assertEqual(model.Roles[0].ModelPermission, "Administrator")

    def test_create_role_invalid_permission(self) -> None:
        manager = FakeManager(FakeModel())
        with self.assertRaises(PowerBIValidationError):
            pbi_create_role_tool(manager, name="Sales", permission="superuser")

    def test_create_role_empty_name(self) -> None:
        manager = FakeManager(FakeModel())
        with self.assertRaises(PowerBIValidationError):
            pbi_create_role_tool(manager, name="   ")


class DeleteRoleTests(unittest.TestCase):
    def test_delete_role_happy_path(self) -> None:
        model = FakeModel(roles=[make_role("Sales FR")])
        manager = FakeManager(model)
        result = pbi_delete_role_tool(manager, name="Sales FR")
        self.assertTrue(result["ok"], result)
        self.assertEqual(result["deleted_role"], {"name": "Sales FR"})
        self.assertEqual(model.Roles.Count, 0)

    def test_delete_role_not_found(self) -> None:
        manager = FakeManager(FakeModel())
        with self.assertRaises(PowerBINotFoundError):
            pbi_delete_role_tool(manager, name="Ghost")


class ListRolesTests(unittest.TestCase):
    def test_list_roles_reports_members_and_filters(self) -> None:
        role = make_role("Sales FR")
        member = FakeExternalModelRoleMember()
        member.MemberName = "user@contoso.com"
        member.IdentityProvider = "AzureAD"
        role.Members.Add(member)
        perm = FakeTablePermission()
        perm.Table = FakeTable("Sales")
        perm.FilterExpression = '[Country] = "FR"'
        role.TablePermissions.Add(perm)
        manager = FakeManager(FakeModel(roles=[role]))

        result = pbi_list_roles_tool(manager)

        self.assertTrue(result["ok"], result)
        self.assertEqual(len(result["roles"]), 1)
        serialized = result["roles"][0]
        self.assertEqual(serialized["name"], "Sales FR")
        self.assertEqual(serialized["members"][0]["name"], "user@contoso.com")
        self.assertEqual(serialized["filters"][0]["table"], "Sales")
        self.assertIn("[Country]", serialized["filters"][0]["filter_expression"])


class SetRoleFilterTests(unittest.TestCase):
    def _manager(self) -> FakeManager:
        return FakeManager(FakeModel(roles=[make_role("Sales FR")], tables=[FakeTable("Sales")]))

    def test_set_filter_creates_permission(self) -> None:
        manager = self._manager()
        result = pbi_set_role_filter_tool(manager, role="Sales FR", table="Sales", filter_expression='[Country] = "FR"')
        self.assertTrue(result["ok"], result)
        role = manager.database.Model.Roles[0]
        self.assertEqual(role.TablePermissions.Count, 1)
        self.assertEqual(role.TablePermissions[0].FilterExpression, '[Country] = "FR"')

    def test_set_filter_updates_existing_permission(self) -> None:
        manager = self._manager()
        pbi_set_role_filter_tool(manager, role="Sales FR", table="Sales", filter_expression='[Country] = "FR"')
        result = pbi_set_role_filter_tool(manager, role="Sales FR", table="Sales", filter_expression='[Country] = "BE"')
        self.assertTrue(result["ok"], result)
        role = manager.database.Model.Roles[0]
        self.assertEqual(role.TablePermissions.Count, 1)
        self.assertEqual(role.TablePermissions[0].FilterExpression, '[Country] = "BE"')

    def test_set_filter_none_removes_permission(self) -> None:
        manager = self._manager()
        pbi_set_role_filter_tool(manager, role="Sales FR", table="Sales", filter_expression='[Country] = "FR"')
        result = pbi_set_role_filter_tool(manager, role="Sales FR", table="Sales", filter_expression=None)
        self.assertTrue(result["ok"], result)
        self.assertTrue(result["filter"]["removed"])
        self.assertEqual(manager.database.Model.Roles[0].TablePermissions.Count, 0)

    def test_set_filter_none_without_existing_reports_removed_false(self) -> None:
        manager = self._manager()
        result = pbi_set_role_filter_tool(manager, role="Sales FR", table="Sales", filter_expression=None)
        self.assertTrue(result["ok"], result)
        self.assertFalse(result["filter"]["removed"])

    def test_set_filter_role_not_found(self) -> None:
        manager = self._manager()
        with self.assertRaises(PowerBINotFoundError):
            pbi_set_role_filter_tool(manager, role="Ghost", table="Sales", filter_expression="TRUE()")

    def test_set_filter_table_not_found(self) -> None:
        manager = self._manager()
        with self.assertRaises(PowerBINotFoundError):
            pbi_set_role_filter_tool(manager, role="Sales FR", table="Ghost", filter_expression="TRUE()")

    def test_set_filter_rejects_query_only_dax(self) -> None:
        manager = self._manager()
        with self.assertRaises(PowerBIValidationError):
            pbi_set_role_filter_tool(
                manager, role="Sales FR", table="Sales", filter_expression="EVALUATE VALUES(Sales)"
            )


class RoleMemberTests(unittest.TestCase):
    def _manager(self) -> FakeManager:
        return FakeManager(FakeModel(roles=[make_role("Sales FR")]))

    def test_add_external_member(self) -> None:
        manager = self._manager()
        result = pbi_add_role_member_tool(manager, role="Sales FR", member_name="user@contoso.com")
        self.assertTrue(result["ok"], result)
        self.assertEqual(result["action"], "added")
        member = manager.database.Model.Roles[0].Members[0]
        self.assertIsInstance(member, FakeExternalModelRoleMember)
        self.assertEqual(member.MemberName, "user@contoso.com")
        self.assertEqual(member.IdentityProvider, "AzureAD")

    def test_add_windows_member(self) -> None:
        manager = self._manager()
        result = pbi_add_role_member_tool(manager, role="Sales FR", member_name="CONTOSO\\jdoe", member_type="windows")
        self.assertTrue(result["ok"], result)
        member = manager.database.Model.Roles[0].Members[0]
        self.assertIsInstance(member, FakeWindowsModelRoleMember)
        self.assertIsNone(result["member"]["identity_provider"])

    def test_add_member_duplicate_raises(self) -> None:
        manager = self._manager()
        pbi_add_role_member_tool(manager, role="Sales FR", member_name="user@contoso.com")
        with self.assertRaises(PowerBIDuplicateError):
            pbi_add_role_member_tool(manager, role="Sales FR", member_name="USER@contoso.com")

    def test_add_member_overwrite_updates_identity_provider(self) -> None:
        manager = self._manager()
        pbi_add_role_member_tool(manager, role="Sales FR", member_name="user@contoso.com")
        result = pbi_add_role_member_tool(
            manager,
            role="Sales FR",
            member_name="user@contoso.com",
            identity_provider="CustomIdP",
            overwrite=True,
        )
        self.assertTrue(result["ok"], result)
        self.assertEqual(result["action"], "updated")
        role = manager.database.Model.Roles[0]
        self.assertEqual(role.Members.Count, 1)
        self.assertEqual(role.Members[0].IdentityProvider, "CustomIdP")

    def test_add_member_invalid_type(self) -> None:
        manager = self._manager()
        with self.assertRaises(PowerBIValidationError):
            pbi_add_role_member_tool(manager, role="Sales FR", member_name="x", member_type="ldap")

    def test_add_member_empty_name(self) -> None:
        manager = self._manager()
        with self.assertRaises(PowerBIValidationError):
            pbi_add_role_member_tool(manager, role="Sales FR", member_name="   ")

    def test_add_member_role_not_found(self) -> None:
        manager = self._manager()
        with self.assertRaises(PowerBINotFoundError):
            pbi_add_role_member_tool(manager, role="Ghost", member_name="user@contoso.com")

    def test_remove_member_happy_path(self) -> None:
        manager = self._manager()
        pbi_add_role_member_tool(manager, role="Sales FR", member_name="user@contoso.com")
        result = pbi_remove_role_member_tool(manager, role="Sales FR", member_name="user@contoso.com")
        self.assertTrue(result["ok"], result)
        self.assertEqual(manager.database.Model.Roles[0].Members.Count, 0)

    def test_remove_member_not_found(self) -> None:
        manager = self._manager()
        with self.assertRaises(PowerBINotFoundError):
            pbi_remove_role_member_tool(manager, role="Sales FR", member_name="ghost@contoso.com")

    def test_remove_member_role_not_found(self) -> None:
        manager = self._manager()
        with self.assertRaises(PowerBINotFoundError):
            pbi_remove_role_member_tool(manager, role="Ghost", member_name="user@contoso.com")


if __name__ == "__main__":
    unittest.main(verbosity=2)

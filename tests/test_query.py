"""Standalone tests for query tools."""

from __future__ import annotations

import sys
import unittest
from pathlib import Path
from types import SimpleNamespace

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

from pbi_connection import PowerBIError
from tools.query import pbi_execute_dax_as_role_tool, pbi_trace_query_tool


class FakeRole:
    def __init__(self, name: str) -> None:
        self.Name = name


class FakeModel:
    def __init__(self, roles: list[str]) -> None:
        self.Roles = [FakeRole(item) for item in roles]


class FakeDatabase:
    def __init__(self, name: str, roles: list[str]) -> None:
        self.Name = name
        self.Model = FakeModel(roles)


class FakeAdomdConnection:
    def __init__(self, connection_string: str) -> None:
        self.connection_string = connection_string
        self.opened = False
        self.closed = False

    def Open(self) -> None:
        self.opened = True

    def Close(self) -> None:
        self.closed = True


class FakeAdomdClient:
    def __init__(self) -> None:
        self.last_connection: FakeAdomdConnection | None = None

    def AdomdConnection(self, connection_string: str) -> FakeAdomdConnection:
        connection = FakeAdomdConnection(connection_string)
        self.last_connection = connection
        return connection


class FakeManager:
    def __init__(self, roles: list[str] | None = None) -> None:
        self._adomd_client = FakeAdomdClient()
        self.state = SimpleNamespace(
            instance=SimpleNamespace(port=52000),
            database=FakeDatabase("UnitTestDb", roles or ["Analyst"]),
            warnings=[],
        )

    def run_read(self, _operation_name: str, reader):
        return reader(self.state)

    def _query_with_pythonnet(self, _connection, _query: str, max_rows: int) -> dict:
        return {
            "columns": ["Ping"],
            "rows": [{"Ping": 1}],
            "row_count": 1,
            "truncated": max_rows < 1,
        }

    def run_adomd_query(self, query: str, *, max_rows: int = 1000) -> dict:
        if "Storage_Table_Relationships" in query:
            return {"columns": ["X"], "rows": [{"X": 1}, {"X": 2}], "row_count": 2, "truncated": False}
        if "SE_Calls" in query:
            return {"columns": ["SE_Calls"], "rows": [{"SE_Calls": 7}], "row_count": 1, "truncated": False}
        return {
            "columns": ["Result"],
            "rows": [{"Result": "ok"}],
            "row_count": 1,
            "truncated": False,
        }


class QueryToolTests(unittest.TestCase):
    def test_execute_dax_as_role_injects_roles_in_connection_string(self) -> None:
        manager = FakeManager(roles=["Analyst", "Manager"])
        result = pbi_execute_dax_as_role_tool(
            manager,
            query='EVALUATE ROW("Ping", 1)',
            role="Analyst",
        )

        self.assertTrue(result["ok"], result)
        self.assertIsNotNone(manager._adomd_client.last_connection)
        connection_string = manager._adomd_client.last_connection.connection_string
        self.assertIn("Roles=Analyst;", connection_string)

    def test_execute_dax_as_role_injects_effective_username_when_provided(self) -> None:
        manager = FakeManager(roles=["Analyst"])
        result = pbi_execute_dax_as_role_tool(
            manager,
            query='EVALUATE ROW("Ping", 1)',
            role="Analyst",
            username="person@contoso.com",
        )

        self.assertTrue(result["ok"], result)
        self.assertIsNotNone(manager._adomd_client.last_connection)
        connection_string = manager._adomd_client.last_connection.connection_string
        self.assertIn("Roles=Analyst;", connection_string)
        self.assertIn("EffectiveUserName=person@contoso.com;", connection_string)

    def test_execute_dax_as_role_returns_role_not_found_error_code(self) -> None:
        manager = FakeManager(roles=["Manager"])
        with self.assertRaises(PowerBIError) as ctx:
            pbi_execute_dax_as_role_tool(
                manager,
                query='EVALUATE ROW("Ping", 1)',
                role="Analyst",
            )
        self.assertEqual(ctx.exception.code, "role_not_found")

    def test_trace_query_returns_duration_and_diagnostics(self) -> None:
        manager = FakeManager()
        result = pbi_trace_query_tool(manager, query='EVALUATE ROW("Ping", 1)')

        self.assertTrue(result["ok"], result)
        self.assertIn("rows", result)
        self.assertIn("diagnostics", result)
        self.assertIn("duration_ms", result["diagnostics"])
        self.assertIn("row_count", result["diagnostics"])


if __name__ == "__main__":
    unittest.main(verbosity=2)

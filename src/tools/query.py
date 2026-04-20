"""Query execution and refresh tools for the Power BI MCP server."""

from __future__ import annotations

import os
import re
import time
from typing import Any

from pbi_connection import (
    PowerBIConfigurationError,
    PowerBIError,
    PowerBINotFoundError,
    PowerBIValidationError,
    find_named,
    flatten_exception_message,
    map_enum,
    ok,
    serialize_value,
)
from security import (
    SECURITY,
    validate_connection_string_property_value,
    validate_model_object_name,
    validate_query_text,
)


# ── DAX safety ───────────────────────────────────────────────────────

# DMV queries that expose server internals — blocked by default
_DMV_BLOCKED_PATTERNS = [
    re.compile(r"\$SYSTEM\.", re.IGNORECASE),
    re.compile(r"DISCOVER_", re.IGNORECASE),
    re.compile(r"DBSCHEMA_", re.IGNORECASE),
    re.compile(r"MDSCHEMA_", re.IGNORECASE),
]

def _validate_dax_query(query: str) -> None:
    """Block dangerous DMV/system queries unless explicitly allowed."""
    validate_query_text(query, max_length=SECURITY.policy().max_query_length)
    if os.environ.get("PBI_MCP_ALLOW_DMV", "0") == "1":
        return
    stripped = query.strip()
    for pattern in _DMV_BLOCKED_PATTERNS:
        if pattern.search(stripped):
            raise PowerBIValidationError(
                f"DMV/system query blocked for security. "
                f"Set PBI_MCP_ALLOW_DMV=1 to allow. "
                f"Matched: {pattern.pattern}",
                details={"pattern": pattern.pattern},
            )


def pbi_execute_dax_tool(
    manager: Any,
    *,
    query: str,
    max_rows: int = 1000,
) -> dict[str, Any]:
    """Execute a DAX or DMV query."""
    _validate_dax_query(query)
    limit = SECURITY.policy().max_rows_for_dax
    if max_rows > limit:
        raise PowerBIValidationError(
            f"max_rows {max_rows} exceeds the configured limit of {limit}.",
            details={"max_rows": max_rows, "limit": limit},
        )
    result = manager.run_adomd_query(query, max_rows=max_rows)
    return ok(
        "Query executed successfully.",
        query=query,
        max_rows=max_rows,
        columns=result["columns"],
        rows=result["rows"],
        row_count=result["row_count"],
        truncated=result["truncated"],
    )


class RoleNotFoundError(PowerBIError):
    code = "role_not_found"


def _pythonnet_stopwatch() -> Any | None:
    try:
        from System.Diagnostics import Stopwatch  # type: ignore
    except Exception:
        return None
    return Stopwatch()


def _extract_role_names(state: Any) -> list[str]:
    model = getattr(getattr(state, "database", None), "Model", None)
    roles = getattr(model, "Roles", None)
    if roles is None:
        return []
    names: list[str] = []
    for item in roles:
        name = str(getattr(item, "Name", "")).strip()
        if name:
            names.append(name)
    return names


def _to_optional_int(value: Any) -> int | None:
    if value is None:
        return None
    if isinstance(value, bool):
        return int(value)
    if isinstance(value, int):
        return value
    if isinstance(value, float):
        return int(value)
    text = str(value).strip()
    if not text:
        return None
    try:
        return int(text)
    except ValueError:
        try:
            return int(float(text))
        except ValueError:
            return None


def _probe_se_calls(manager: Any) -> int | None:
    try:
        probe = manager.run_adomd_query('EVALUATE ROW("SE_Calls", [Storage Engine Calls])', max_rows=1)
        rows = probe.get("rows", [])
        if rows:
            value = _to_optional_int(rows[0].get("SE_Calls"))
            if value is not None:
                return value
    except Exception:
        pass

    try:
        fallback = manager.run_adomd_query("SELECT * FROM $System.Discover_Storage_Table_Relationships", max_rows=SECURITY.policy().max_rows_for_dax)
        return _to_optional_int(fallback.get("row_count"))
    except Exception:
        return None


def _probe_formula_engine_ms(manager: Any) -> int | None:
    try:
        probe = manager.run_adomd_query('EVALUATE ROW("FormulaEngineMs", [Formula Engine Duration])', max_rows=1)
    except Exception:
        return None
    rows = probe.get("rows", [])
    if not rows:
        return None
    return _to_optional_int(rows[0].get("FormulaEngineMs"))


def pbi_execute_dax_as_role_tool(
    manager: Any,
    *,
    query: str,
    role: str,
    username: str | None = None,
) -> dict[str, Any]:
    """Execute a DAX query under a specific role context."""
    _validate_dax_query(query)
    policy = SECURITY.policy()
    validate_model_object_name(role, max_length=policy.max_name_length)
    validate_connection_string_property_value(role, field="role")
    if username is not None:
        validate_model_object_name(username, max_length=policy.max_name_length)
        validate_connection_string_property_value(username, field="username")

    max_rows = min(1000, policy.max_rows_for_dax)

    def _execute(state: Any) -> dict[str, Any]:
        available_roles = _extract_role_names(state)
        if not any(item.casefold() == role.casefold() for item in available_roles):
            raise RoleNotFoundError(
                f"Role '{role}' was not found in the current model.",
                details={"role": role, "available_roles": available_roles},
            )

        adomd_client = getattr(manager, "_adomd_client", None)
        if adomd_client is None:
            raise PowerBIConfigurationError(
                "ADOMD query support is unavailable for role-scoped execution.",
                details={"warnings": getattr(state, "warnings", [])},
            )

        connection_string = (
            "Provider=MSOLAP;"
            f"Data Source=localhost:{state.instance.port};"
            f"Initial Catalog={state.database.Name};"
            "Integrated Security=SSPI;"
            f"Roles={role};"
        )
        if username:
            connection_string += f"EffectiveUserName={username};"

        connection = adomd_client.AdomdConnection(connection_string)
        try:
            connection.Open()
        except Exception as exc:
            message = flatten_exception_message(exc)
            lowered = message.casefold()
            if "role" in lowered and any(token in lowered for token in ("not found", "does not exist", "cannot find", "unknown")):
                raise RoleNotFoundError(
                    f"Role '{role}' was not found in the current model.",
                    details={"role": role, "reason": message},
                ) from exc
            translate = getattr(manager, "_translate_exception", None)
            if callable(translate):
                raise translate(exc, "execute_dax_as_role") from exc
            raise

        try:
            if hasattr(manager, "_query_with_pythonnet"):
                return manager._query_with_pythonnet(connection, query, max_rows)
            command = adomd_client.AdomdCommand(query, connection)
            reader = command.ExecuteReader()
            try:
                columns = [str(reader.GetName(index)) for index in range(reader.FieldCount)]
                rows: list[dict[str, Any]] = []
                truncated = False
                while reader.Read():
                    if len(rows) >= max_rows:
                        truncated = True
                        break
                    rows.append(
                        {columns[index]: serialize_value(reader.GetValue(index)) for index in range(reader.FieldCount)}
                    )
                return {"columns": columns, "rows": rows, "row_count": len(rows), "truncated": truncated}
            finally:
                reader.Close()
                command.Dispose()
        finally:
            try:
                connection.Close()
            except Exception:
                pass

    result = manager.run_read("execute_dax_as_role", _execute)
    return ok(
        "Query executed successfully under role context.",
        query=query,
        role=role,
        username=username,
        max_rows=max_rows,
        columns=result["columns"],
        rows=result["rows"],
        row_count=result["row_count"],
        truncated=result["truncated"],
    )


def pbi_trace_query_tool(manager: Any, *, query: str) -> dict[str, Any]:
    """Execute a DAX query and return result rows with timing diagnostics."""
    _validate_dax_query(query)
    max_rows = min(1000, SECURITY.policy().max_rows_for_dax)
    stopwatch = _pythonnet_stopwatch()
    start = time.perf_counter()
    if stopwatch is not None:
        stopwatch.Start()
    result = manager.run_adomd_query(query, max_rows=max_rows)
    duration_ms = int((time.perf_counter() - start) * 1000)
    if stopwatch is not None:
        stopwatch.Stop()
        duration_ms = int(stopwatch.ElapsedMilliseconds)

    diagnostics = {
        "duration_ms": duration_ms,
        "row_count": result["row_count"],
        "se_calls": _probe_se_calls(manager),
        "formula_engine_ms": _probe_formula_engine_ms(manager),
    }
    return ok(
        "Query traced successfully.",
        query=query,
        max_rows=max_rows,
        columns=result["columns"],
        rows=result["rows"],
        row_count=result["row_count"],
        truncated=result["truncated"],
        diagnostics=diagnostics,
    )


def pbi_refresh_tool(
    manager: Any,
    *,
    target: str = "model",
    refresh_type: str = "full",
) -> dict[str, Any]:
    """Trigger a model or table refresh."""

    def _mutator(state: Any, database: Any, model: Any) -> dict[str, Any]:
        tom = manager.tom
        refresh_enum = map_enum(tom.RefreshType, refresh_type)
        if target.strip().casefold() in {"model", "database"}:
            model.RequestRefresh(refresh_enum)
            scope = {"target_type": "model", "target": str(database.Name)}
        else:
            table = find_named(model.Tables, target)
            if table is None:
                raise PowerBINotFoundError(
                    f"Table '{target}' was not found.",
                    details={"table": target},
                )
            table.RequestRefresh(refresh_enum)
            scope = {"target_type": "table", "target": str(table.Name)}
        return {
            "refresh": {
                **scope,
                "refresh_type": refresh_type,
            }
        }

    payload = manager.execute_write("refresh", _mutator)
    return ok(
        f"Refresh requested successfully for {payload['refresh']['target_type']} '{payload['refresh']['target']}'.",
        refresh=payload["refresh"],
        save_result=payload["save_result"],
        connection=payload["connection"],
    )

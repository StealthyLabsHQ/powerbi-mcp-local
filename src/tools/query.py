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
    timeout_seconds: int | None = None,
) -> dict[str, Any]:
    """Execute a DAX or DMV query."""
    _validate_dax_query(query)
    limit = SECURITY.policy().max_rows_for_dax
    if max_rows > limit:
        raise PowerBIValidationError(
            f"max_rows {max_rows} exceeds the configured limit of {limit}.",
            details={"max_rows": max_rows, "limit": limit},
        )
    result = manager.run_adomd_query(query, max_rows=max_rows, timeout_seconds=timeout_seconds)
    return ok(
        "Query executed successfully.",
        query=query,
        max_rows=max_rows,
        timeout_seconds=timeout_seconds,
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


def pbi_trace_query_tool(
    manager: Any,
    *,
    query: str,
    timeout_seconds: int | None = None,
) -> dict[str, Any]:
    """Execute a DAX query and return result rows with timing diagnostics."""
    _validate_dax_query(query)
    max_rows = min(1000, SECURITY.policy().max_rows_for_dax)
    stopwatch = _pythonnet_stopwatch()
    start = time.perf_counter()
    if stopwatch is not None:
        stopwatch.Start()
    result = manager.run_adomd_query(query, max_rows=max_rows, timeout_seconds=timeout_seconds)
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


def pbi_validate_dax_tool(
    manager: Any,
    *,
    expression: str,
    kind: str = "scalar",
) -> dict[str, Any]:
    """Parse-check a DAX expression by running a zero/one-row probe and catching errors.

    kind='scalar' wraps the expression with EVALUATE ROW("v", <expr>).
    kind='table'  wraps the expression with EVALUATE TOPN(0, <expr>).
    """
    if not expression or not expression.strip():
        raise PowerBIValidationError("expression is required.")
    policy = SECURITY.policy()
    validate_query_text(expression, max_length=policy.max_query_length)
    normalized_kind = kind.strip().casefold()
    if normalized_kind not in {"scalar", "table"}:
        raise PowerBIValidationError(
            "kind must be 'scalar' or 'table'.", details={"kind": kind}
        )

    if normalized_kind == "scalar":
        probe = f'EVALUATE ROW("__probe", {expression})'
    else:
        probe = f"EVALUATE TOPN(0, {expression})"

    try:
        manager.run_adomd_query(probe, max_rows=1)
    except PowerBIError as exc:
        return ok(
            "DAX expression is invalid.",
            valid=False,
            kind=normalized_kind,
            error=flatten_exception_message(exc),
            error_code=getattr(exc, "code", "validation_error"),
        )
    return ok(
        "DAX expression is valid.",
        valid=True,
        kind=normalized_kind,
    )


_DAX_TABLE_COLUMN_RE = re.compile(r"(?P<table>'[^']+'|[A-Za-z_][\w]*)\s*\[(?P<column>[^\]]+)\]")
_DAX_MEASURE_REF_RE = re.compile(r"(?<![\w.\]])\[(?P<measure>[^\]]+)\]")


def pbi_validate_dax_semantic_tool(
    manager: Any,
    *,
    expression: str,
    kind: str = "scalar",
    format_string: str = "",
    include_hidden: bool = False,
) -> dict[str, Any]:
    """Validate a DAX expression with both semantic and syntax checks.

    Three layers, each cheap and skippable on failure:

    1. **Semantic — references**: scan the expression for ``Table[Column]`` and
       bare ``[Measure]`` tokens, look them up against the live model index. Each
       unknown reference is reported under ``semantic.unknown_references`` with
       its kind. Strict failure means at least one column or measure isn't in
       the model — typically a typo before commit.

    2. **Semantic — format compatibility (heuristic, never blocks)**: if a
       ``format_string`` is supplied and looks percent-shaped (``"0.00%"``,
       ``"0%"``…) but the expression looks scalar-money (``SUM`` / ``SUMX`` /
       ``DIVIDE`` over column values), surface a ``semantic.suspicious_format``
       warning. Best-effort; the runtime probe stays the source of truth.

    3. **Runtime**: delegates to ``pbi_validate_dax_tool`` (EVALUATE ROW / TOPN
       probe) and surfaces any ASEngine error verbatim.

    Returns ``{ok, valid, syntax: ok|error, semantic: {unknown_references,
    suspicious_format}, runtime_error?}``.
    """
    if not expression or not expression.strip():
        raise PowerBIValidationError("expression is required.")

    # --- Layer 1: semantic references ---
    # Strip string literals (DAX uses double quotes for strings) so we don't
    # treat literals like "[Foo]" as measure references.
    sanitized = re.sub(r'"(?:[^"\\]|\\.)*"', '""', expression)

    columns_seen: set[tuple[str, str]] = set()
    for match in _DAX_TABLE_COLUMN_RE.finditer(sanitized):
        table = match.group("table").strip("'")
        column = match.group("column")
        columns_seen.add((table, column))

    measures_seen: set[str] = set()
    # Subtract column references from the leftover tokens so we don't double-count
    # the [Column] of Table[Column] as a measure reference.
    column_only_text = _DAX_TABLE_COLUMN_RE.sub("", sanitized)
    for match in _DAX_MEASURE_REF_RE.finditer(column_only_text):
        measures_seen.add(match.group("measure"))

    unknown_references: list[dict[str, str]] = []
    semantic_status = "skipped"

    # Lazy import to avoid circulars: the field-index helper lives in tools.visuals.
    try:
        from .visuals import _live_model_field_index
        index, status = _live_model_field_index(manager, include_hidden=include_hidden)
    except Exception as exc:  # pragma: no cover
        index, status = None, {"status": "unavailable", "error": flatten_exception_message(exc)}
    if index is not None:
        semantic_status = "checked"
        for table, column in sorted(columns_seen):
            if (table.casefold(), column.casefold()) not in index["columns"]:
                unknown_references.append({"reference": f"{table}[{column}]", "kind": "column"})
        for measure_name in sorted(measures_seen):
            if measure_name.casefold() not in index["measures"]:
                # An unknown bare ``[X]`` could also be a column whose table prefix was elided.
                # Report it as ``measure_or_column`` so callers know the heuristic is fuzzy.
                # NB: use a distinct local name (``ref_kind``) so we never shadow the outer
                # ``kind`` parameter that the runtime probe needs intact.
                ref_kind = "measure_or_column"
                if any(col_lc == measure_name.casefold() for _, col_lc in index["columns"]):
                    # It IS an existing column name, just unqualified — this is technically
                    # legal DAX but flagged as a style warning.
                    continue
                unknown_references.append({"reference": f"[{measure_name}]", "kind": ref_kind})

    # --- Layer 2: format compatibility heuristic ---
    suspicious_format: list[dict[str, str]] = []
    if format_string:
        fmt_lower = format_string.lower()
        looks_percent = "%" in format_string and ("0%" in format_string or "0.0" in format_string)
        # Very rough scalar-money heuristic: SUM/SUMX over a non-percent column.
        looks_money = bool(
            re.search(r"\b(SUM|SUMX|TOTAL|REVENUE|SALES)\b", expression, re.IGNORECASE)
            and "%" not in expression
        )
        if looks_percent and looks_money:
            suspicious_format.append({
                "format_string": format_string,
                "reason": "percent format on a likely scalar-money expression",
            })
        # Currency-shape heuristic: percent expression but currency format.
        if not looks_percent and ("€" in format_string or "$" in format_string):
            if "DIVIDE" in expression.upper() and "%" not in fmt_lower:
                # DIVIDE often produces ratios — currency on a ratio is suspicious.
                suspicious_format.append({
                    "format_string": format_string,
                    "reason": "currency format on a DIVIDE expression that often returns a ratio",
                })

    # --- Layer 3: runtime probe (delegates to existing tool) ---
    runtime = pbi_validate_dax_tool(manager, expression=expression, kind=kind)
    syntax = "ok" if runtime.get("valid") else "error"

    valid = bool(runtime.get("valid")) and not unknown_references
    return ok(
        "DAX semantic validation completed."
        if valid
        else "DAX semantic validation found at least one issue.",
        valid=valid,
        kind=runtime.get("kind", kind),
        syntax=syntax,
        semantic={
            "status": semantic_status,
            "unknown_references": unknown_references,
            "suspicious_format": suspicious_format,
            "columns_referenced": sorted(f"{t}[{c}]" for t, c in columns_seen),
            "measures_referenced": sorted(f"[{m}]" for m in measures_seen),
        },
        runtime_error=runtime.get("error"),
    )


def pbi_measure_dependencies_tool(
    manager: Any,
    *,
    measure: str | None = None,
    table: str | None = None,
) -> dict[str, Any]:
    """Return calc-dependency graph rows (source → referenced object) from DISCOVER_CALC_DEPENDENCY."""
    if measure is not None:
        validate_model_object_name(measure)
    if table is not None:
        validate_model_object_name(table)

    query = "SELECT * FROM $SYSTEM.DISCOVER_CALC_DEPENDENCY"
    result = manager.run_adomd_query(query, max_rows=SECURITY.policy().max_rows_for_dax)

    rows = result.get("rows", [])
    if measure is not None or table is not None:
        def _match(row: dict[str, Any]) -> bool:
            name_match = True
            table_match = True
            if measure is not None:
                obj = str(row.get("OBJECT", "") or row.get("Object", ""))
                name_match = obj.casefold() == measure.casefold()
            if table is not None:
                tbl = str(row.get("TABLE", "") or row.get("Table", ""))
                table_match = tbl.casefold() == table.casefold()
            return name_match and table_match

        rows = [row for row in rows if _match(row)]

    return ok(
        "Measure dependencies retrieved successfully.",
        measure=measure,
        table=table,
        columns=result.get("columns", []),
        rows=rows,
        row_count=len(rows),
        truncated=result.get("truncated", False),
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
        persistence=payload.get("persistence"),
        connection=payload["connection"],
    )

"""Model inspection and table/column operations for the Power BI MCP server."""

from __future__ import annotations

import json
from typing import Any

from pbi_connection import (
    PowerBIDuplicateError,
    PowerBINotFoundError,
    PowerBIValidationError,
    dax_quote_table_name,
    find_named,
    map_enum,
    ok,
    serialize_value,
)
from security import (
    redact_sensitive_data,
    resolve_local_path,
    validate_model_expression,
    validate_model_object_name,
)


def pbi_connect_tool(
    manager: Any,
    *,
    preferred_port: int | None = None,
    force_reconnect: bool = False,
) -> dict[str, Any]:
    """Connect to Power BI Desktop and report the active instance."""
    snapshot = manager.connect(
        preferred_port=preferred_port,
        force_reconnect=force_reconnect,
    )
    return ok(
        "Connected to Power BI Desktop.",
        **snapshot,
    )


def pbi_list_instances_tool(manager: Any) -> dict[str, Any]:
    """List discovered Power BI Desktop instances without forcing a connection."""
    instances = manager.list_instances()
    return ok(
        "Power BI Desktop instances listed successfully.",
        instances=instances,
        count=len(instances),
    )


def pbi_operation_history_tool(manager: Any, *, last_n: int = 20) -> dict[str, Any]:
    """Return the most-recent N tool operations recorded by the connection
    manager (newest first).

    Each entry: ``{ts, op, kind: "read"|"write", duration_ms, ok}`` plus
    ``error_type``/``error_code``/``error_message`` when ``ok`` is False. Use
    this to self-diagnose what just happened — e.g. an LLM can pull the last
    5 calls after a failure to see which writes already landed.
    """
    if not isinstance(last_n, int) or last_n < 1:
        raise PowerBIValidationError("last_n must be a positive integer.", details={"last_n": last_n})
    history = manager.operation_history(last_n=last_n)
    return ok(
        f"Returned {len(history)} most-recent operation(s).",
        operations=history,
        count=len(history),
        buffer_capacity=int(getattr(getattr(manager, "_operation_log", None), "maxlen", 0) or 0),
    )


def pbi_system_health_tool(manager: Any) -> dict[str, Any]:
    """Single-call self-diagnostic for stability and dependency status.

    Read-only. Skips any TOM/ADOMD probe gracefully when no connection is
    active. Use as a preflight from any LLM agent: one call answers "can I
    talk to Power BI right now?" without juggling connect / list / model_info.
    """
    snapshot: dict[str, Any] = {
        "connected": False,
        "port": None,
        "port_open": None,
        "pid_match": None,
        "tom_available": False,
        "adomd_available": False,
        "model_loaded": False,
        "model_name": None,
        "table_count": None,
        "measure_count": None,
        "cache": {
            "write_generation": int(getattr(manager, "_write_generation", 0) or 0),
            "entries": len(getattr(manager, "_read_cache", {}) or {}),
        },
        "last_operation_ts": None,
    }

    state = getattr(manager, "_state", None)
    if state is not None:
        snapshot["connected"] = True
        instance = getattr(state, "instance", None)
        if instance is not None:
            snapshot["port"] = getattr(instance, "port", None)
        try:
            snapshot["port_open"] = bool(manager._is_port_open(snapshot["port"])) if snapshot["port"] else None
        except Exception:
            snapshot["port_open"] = None
        try:
            cached_pid = getattr(instance, "pid", None) if instance else None
            current_pid = manager._pid_for_port(snapshot["port"]) if snapshot["port"] else None
            snapshot["pid_match"] = (
                cached_pid is not None and current_pid is not None and cached_pid == current_pid
            ) or None
        except Exception:
            snapshot["pid_match"] = None
        snapshot["tom_available"] = getattr(state, "tom_server", None) is not None
        snapshot["adomd_available"] = bool(getattr(state, "adomd_available", False))

        try:
            database = getattr(state, "database", None)
            if database is not None:
                snapshot["model_loaded"] = True
                snapshot["model_name"] = serialize_value(getattr(database, "Name", None))
                model = getattr(database, "Model", None)
                if model is not None:
                    snapshot["table_count"] = int(model.Tables.Count) if hasattr(model.Tables, "Count") else None
                    measure_total = 0
                    for table in model.Tables:
                        try:
                            measure_total += int(table.Measures.Count) if hasattr(table.Measures, "Count") else 0
                        except Exception:
                            pass
                    snapshot["measure_count"] = measure_total
        except Exception:
            # Live model probe failed — leave fields as None; the caller already knows we're connected.
            pass

    history = manager.operation_history(last_n=1)
    if history:
        snapshot["last_operation_ts"] = history[0].get("ts")

    deps: dict[str, Any] = {}
    for module_name in ("mcp", "pythonnet", "pyadomd", "pbi_pyadomd"):
        try:
            __import__(module_name.replace("-", "_"))
            deps[module_name] = "available"
        except Exception:
            deps[module_name] = "missing"
    snapshot["dependencies"] = deps

    return ok("System health snapshot collected.", **snapshot)


def pbi_refresh_metadata_tool(manager: Any) -> dict[str, Any]:
    """Reload cached TOM schema from the server (cheaper than full reconnect)."""
    payload = manager.refresh_metadata()
    return ok(
        "Metadata cache refreshed.",
        changed=payload["changed"],
        previous_version=payload["previous_version"],
        current_version=payload["current_version"],
        database=payload["database"],
    )


def pbi_list_tables_tool(
    manager: Any,
    *,
    include_hidden: bool = False,
    include_row_counts: bool = False,
) -> dict[str, Any]:
    """List model tables and columns."""

    def _reader(state: Any) -> dict[str, Any]:
        tables = []
        for table in state.database.Model.Tables:
            is_hidden = bool(getattr(table, "IsHidden", False))
            if is_hidden and not include_hidden:
                continue

            columns = []
            for column in table.Columns:
                column_hidden = bool(getattr(column, "IsHidden", False))
                if column_hidden and not include_hidden:
                    continue
                columns.append(
                    {
                        "name": str(column.Name),
                        "data_type": serialize_value(getattr(column, "DataType", None)),
                        "column_type": type(column).__name__,
                        "is_hidden": column_hidden,
                        "format_string": serialize_value(getattr(column, "FormatString", "")),
                        "expression": redact_sensitive_data(serialize_value(getattr(column, "Expression", None))),
                    }
                )

            tables.append(
                {
                    "name": str(table.Name),
                    "description": serialize_value(getattr(table, "Description", "")),
                    "is_hidden": is_hidden,
                    "table_type": type(table).__name__,
                    "partitions": [str(partition.Name) for partition in table.Partitions],
                    "columns": columns,
                    "row_count": None,
                    "row_count_error": None,
                }
            )
        return {"tables": tables, "connection": state.snapshot()}

    payload = manager.cached_run_read(f"list_tables:h{include_hidden}", "list_tables", _reader)

    if include_row_counts:
        for table_payload in payload["tables"]:
            try:
                query = f'EVALUATE ROW("__RowCount", COUNTROWS({dax_quote_table_name(table_payload["name"])}))'
                result = manager.run_adomd_query(query, max_rows=1)
                rows = result.get("rows", [])
                if rows:
                    table_payload["row_count"] = rows[0].get("__RowCount")
            except Exception as exc:
                table_payload["row_count_error"] = str(exc)

    return ok(
        "Tables listed successfully.",
        tables=payload["tables"],
        connection=payload["connection"],
    )


def pbi_model_info_tool(
    manager: Any,
    *,
    include_hidden: bool = False,
    include_row_counts: bool = False,
) -> dict[str, Any]:
    """Return a full model snapshot in a single call."""
    from .measures import pbi_list_measures_tool
    from .relationships import pbi_list_relationships_tool

    tables = pbi_list_tables_tool(
        manager,
        include_hidden=include_hidden,
        include_row_counts=include_row_counts,
    )
    measures = pbi_list_measures_tool(manager, include_hidden=include_hidden)
    relationships = pbi_list_relationships_tool(manager)
    return ok(
        "Model snapshot collected successfully.",
        connection=tables["connection"],
        tables=tables["tables"],
        measures=measures["measures"],
        relationships=relationships["relationships"],
    )


def pbi_export_model_tool(
    manager: Any,
    *,
    path: str | None = None,
    include_hidden: bool = False,
    include_row_counts: bool = False,
) -> dict[str, Any]:
    """Export the full model as JSON, optionally writing it to disk."""
    snapshot = pbi_model_info_tool(
        manager,
        include_hidden=include_hidden,
        include_row_counts=include_row_counts,
    )
    model_json = redact_sensitive_data(
        {
            "tables": snapshot["tables"],
            "measures": snapshot["measures"],
            "relationships": snapshot["relationships"],
        }
    )
    written_path = None
    if path:
        output_path = resolve_local_path(path, must_exist=False, allowed_extensions={".json"})
        output_path.parent.mkdir(parents=True, exist_ok=True)
        output_path.write_text(
            json.dumps(model_json, indent=2, ensure_ascii=False),
            encoding="utf-8",
        )
        written_path = str(output_path)

    return ok(
        "Model export completed successfully.",
        connection=snapshot["connection"],
        model=model_json,
        written_path=written_path,
    )


def pbi_create_table_tool(
    manager: Any,
    *,
    name: str,
    expression: str,
    is_hidden: bool = False,
    overwrite: bool = False,
    refresh_after_create: bool = True,
) -> dict[str, Any]:
    """Create or update a calculated table."""
    validate_model_object_name(name)
    validate_model_expression(expression, kind="calculated table expression")

    def _mutator(state: Any, database: Any, model: Any) -> dict[str, Any]:
        tom = manager.tom
        existing = find_named(model.Tables, name)
        action = "created"
        if existing is not None and not overwrite:
            raise PowerBIDuplicateError(
                f"Table '{name}' already exists.",
                details={"table": name},
            )

        if existing is None:
            table = tom.Table()
            table.Name = name
            model.Tables.Add(table)
        else:
            table = existing
            action = "updated"
            if int(table.Partitions.Count) > 1:
                raise PowerBIValidationError(
                    f"Table '{name}' has multiple partitions. Refusing to overwrite it automatically.",
                    details={"table": name, "partition_count": int(table.Partitions.Count)},
                )
            if table.Partitions.Count > 0:
                source = table.Partitions[0].Source
                if type(source).__name__ != "CalculatedPartitionSource":
                    raise PowerBIValidationError(
                        f"Table '{name}' exists but is not a calculated table. Refusing to overwrite it.",
                        details={"table": name},
                    )

        table.IsHidden = is_hidden
        if table.Partitions.Count == 0:
            partition = tom.Partition()
            partition.Name = name
            table.Partitions.Add(partition)
        else:
            partition = table.Partitions[0]

        partition.Name = name
        source = tom.CalculatedPartitionSource()
        source.Expression = expression
        partition.Source = source

        if refresh_after_create:
            table.RequestRefresh(tom.RefreshType.Calculate)

        return {
            "table": {
                "name": name,
                "expression": redact_sensitive_data(expression),
                "is_hidden": is_hidden,
            },
            "action": action,
        }

    payload = manager.execute_write("create_table", _mutator)
    return ok(
        f"Calculated table '{name}' {payload['action']} successfully.",
        table=payload["table"],
        action=payload["action"],
        save_result=payload["save_result"],
        persistence=payload.get("persistence"),
        connection=payload["connection"],
    )


def pbi_delete_table_tool(manager: Any, *, name: str) -> dict[str, Any]:
    """Delete a table. Removes associated relationships and measures."""
    validate_model_object_name(name)

    def _mutator(state: Any, database: Any, model: Any) -> dict[str, Any]:
        table = find_named(model.Tables, name)
        if table is None:
            raise PowerBINotFoundError(f"Table '{name}' was not found.", details={"table": name})
        model.Tables.Remove(table)
        return {"deleted_table": {"name": name}}

    payload = manager.execute_write("delete_table", _mutator)
    return ok(
        f"Table '{name}' deleted successfully.",
        deleted_table=payload["deleted_table"],
        save_result=payload["save_result"],
        persistence=payload.get("persistence"),
        connection=payload["connection"],
    )


def pbi_delete_column_tool(manager: Any, *, table: str, name: str) -> dict[str, Any]:
    """Delete a column from a table."""
    validate_model_object_name(table)
    validate_model_object_name(name)

    def _mutator(state: Any, database: Any, model: Any) -> dict[str, Any]:
        target_table = find_named(model.Tables, table)
        if target_table is None:
            raise PowerBINotFoundError(f"Table '{table}' was not found.", details={"table": table})
        column = find_named(target_table.Columns, name)
        if column is None:
            raise PowerBINotFoundError(
                f"Column '{table}[{name}]' was not found.",
                details={"table": table, "column": name},
            )
        target_table.Columns.Remove(column)
        return {"deleted_column": {"table": table, "name": name}}

    payload = manager.execute_write("delete_column", _mutator)
    return ok(
        f"Column '{table}[{name}]' deleted successfully.",
        deleted_column=payload["deleted_column"],
        save_result=payload["save_result"],
        persistence=payload.get("persistence"),
        connection=payload["connection"],
    )


def pbi_rename_table_tool(manager: Any, *, name: str, new_name: str) -> dict[str, Any]:
    """Rename a table. Callers are responsible for updating dependent DAX expressions."""
    validate_model_object_name(name)
    validate_model_object_name(new_name)

    def _mutator(state: Any, database: Any, model: Any) -> dict[str, Any]:
        table = find_named(model.Tables, name)
        if table is None:
            raise PowerBINotFoundError(f"Table '{name}' was not found.", details={"table": name})
        if find_named(model.Tables, new_name) is not None and new_name.casefold() != name.casefold():
            raise PowerBIDuplicateError(
                f"A table named '{new_name}' already exists.",
                details={"new_name": new_name},
            )
        table.Name = new_name
        return {"rename": {"table_old_name": name, "table_new_name": new_name}}

    payload = manager.execute_write("rename_table", _mutator)
    return ok(
        f"Table '{name}' renamed to '{new_name}'.",
        rename=payload["rename"],
        save_result=payload["save_result"],
        persistence=payload.get("persistence"),
        connection=payload["connection"],
    )


def pbi_rename_column_tool(manager: Any, *, table: str, name: str, new_name: str) -> dict[str, Any]:
    """Rename a column. Callers are responsible for updating dependent DAX."""
    validate_model_object_name(table)
    validate_model_object_name(name)
    validate_model_object_name(new_name)

    def _mutator(state: Any, database: Any, model: Any) -> dict[str, Any]:
        target_table = find_named(model.Tables, table)
        if target_table is None:
            raise PowerBINotFoundError(f"Table '{table}' was not found.", details={"table": table})
        column = find_named(target_table.Columns, name)
        if column is None:
            raise PowerBINotFoundError(
                f"Column '{table}[{name}]' was not found.",
                details={"table": table, "column": name},
            )
        if find_named(target_table.Columns, new_name) is not None and new_name.casefold() != name.casefold():
            raise PowerBIDuplicateError(
                f"Column '{table}[{new_name}]' already exists.",
                details={"table": table, "new_name": new_name},
            )
        column.Name = new_name
        return {"rename": {"table": table, "column_old_name": name, "column_new_name": new_name}}

    payload = manager.execute_write("rename_column", _mutator)
    return ok(
        f"Column '{table}[{name}]' renamed to '{table}[{new_name}]'.",
        rename=payload["rename"],
        save_result=payload["save_result"],
        persistence=payload.get("persistence"),
        connection=payload["connection"],
    )


def pbi_create_column_tool(
    manager: Any,
    *,
    table: str,
    name: str,
    expression: str,
    data_type: str | None = None,
    format_string: str = "",
    display_folder: str = "",
    is_hidden: bool = False,
    overwrite: bool = False,
) -> dict[str, Any]:
    """Create or update a calculated column."""
    validate_model_object_name(table)
    validate_model_object_name(name)
    validate_model_expression(expression, kind="calculated column expression")

    def _mutator(state: Any, database: Any, model: Any) -> dict[str, Any]:
        tom = manager.tom
        target_table = find_named(model.Tables, table)
        if target_table is None:
            raise PowerBINotFoundError(f"Table '{table}' was not found.", details={"table": table})

        existing = find_named(target_table.Columns, name)
        action = "created"
        if existing is not None and not overwrite:
            raise PowerBIDuplicateError(
                f"Column '{table}[{name}]' already exists.",
                details={"table": table, "column": name},
            )

        if existing is None:
            column = tom.CalculatedColumn()
            column.Name = name
            target_table.Columns.Add(column)
        else:
            column = existing
            action = "updated"
            if type(column).__name__ != "CalculatedColumn":
                raise PowerBIValidationError(
                    f"Column '{table}[{name}]' exists but is not a calculated column. Refusing to overwrite it.",
                    details={"table": table, "column": name},
                )

        column.Expression = expression
        column.IsHidden = is_hidden
        if data_type:
            column.DataType = map_enum(tom.DataType, data_type)
        if format_string:
            column.FormatString = format_string
        if display_folder:
            column.DisplayFolder = display_folder

        return {
            "column": {
                "table": table,
                "name": name,
                "expression": redact_sensitive_data(expression),
                "data_type": data_type,
                "format_string": format_string or None,
                "display_folder": display_folder or None,
                "is_hidden": is_hidden,
            },
            "action": action,
        }

    payload = manager.execute_write("create_column", _mutator)
    return ok(
        f"Calculated column '{table}[{name}]' {payload['action']} successfully.",
        column=payload["column"],
        action=payload["action"],
        save_result=payload["save_result"],
        persistence=payload.get("persistence"),
        connection=payload["connection"],
    )


def pbi_generate_dax_context_prompt_tool(
    manager: Any,
    *,
    include_hidden: bool = False,
    include_dax: bool = True,
    include_relationships: bool = True,
    max_chars: int = 12000,
) -> dict[str, Any]:
    """Render a compact markdown snapshot of the model — tables, columns,
    measures, relationships — ready to paste into an LLM system prompt so the
    LLM can author DAX with full schema context in one round-trip.

    Sections:
    - ``# Model: <name>``
    - one ``## <Table>`` per table with ``Columns:`` and (optionally) DAX
      ``Measures:`` lines
    - ``## Relationships`` with ``A[X] → B[Y] (cardinality, active)`` rows

    Output is truncated to ``max_chars`` (default 12 000) with a trailing
    note when truncation kicks in. Set ``include_dax=False`` to omit the
    measure expressions and stay terse.
    """
    if max_chars < 1024:
        raise PowerBIValidationError("max_chars must be >= 1024.", details={"max_chars": max_chars})
    info = pbi_model_info_tool(manager, include_hidden=include_hidden, include_row_counts=False)
    if not info.get("ok"):
        return info

    db_name = info.get("database_name") or info.get("database") or "(unknown)"
    tables = info.get("tables", []) or []
    measures = info.get("measures", []) or []
    relationships = info.get("relationships", []) or []

    lines: list[str] = []
    lines.append(f"# Model: {db_name}")
    lines.append(f"_Tables: {len(tables)} · Measures: {len(measures)} · Relationships: {len(relationships)}_")
    lines.append("")

    measures_by_table: dict[str, list[dict[str, Any]]] = {}
    for measure in measures:
        measures_by_table.setdefault(str(measure.get("table", "")), []).append(measure)

    for table in sorted(tables, key=lambda t: str(t.get("name", "")).casefold()):
        table_name = str(table.get("name", ""))
        if not table_name:
            continue
        lines.append(f"## {table_name}")
        columns = table.get("columns", []) or []
        if columns:
            col_strs: list[str] = []
            for column in columns:
                name = str(column.get("name", ""))
                dtype = str(column.get("data_type", "?"))
                col_strs.append(f"{name} ({dtype})")
            lines.append("**Columns:** " + ", ".join(col_strs))
        table_measures = measures_by_table.get(table_name, [])
        if table_measures:
            lines.append("**Measures:**")
            for measure in sorted(table_measures, key=lambda m: str(m.get("name", "")).casefold()):
                m_name = str(measure.get("name", ""))
                m_format = str(measure.get("format_string", "") or "")
                if include_dax:
                    expr = str(measure.get("expression", "") or "").strip().replace("\r", " ").replace("\n", " ")
                    suffix = f" — `{m_format}`" if m_format else ""
                    lines.append(f"- `{m_name}`{suffix}: `{expr}`")
                else:
                    suffix = f" — `{m_format}`" if m_format else ""
                    lines.append(f"- `{m_name}`{suffix}")
        lines.append("")

    if include_relationships and relationships:
        lines.append("## Relationships")
        for rel in relationships:
            from_table = str(rel.get("from_table", ""))
            from_column = str(rel.get("from_column", ""))
            to_table = str(rel.get("to_table", ""))
            to_column = str(rel.get("to_column", ""))
            cardinality = str(rel.get("cardinality", ""))
            active = bool(rel.get("is_active", True))
            lines.append(
                f"- `{from_table}[{from_column}]` → `{to_table}[{to_column}]` "
                f"({cardinality}, {'active' if active else 'inactive'})"
            )

    output = "\n".join(lines)
    truncated = False
    if len(output) > max_chars:
        # Try to cut on the previous newline so we don't slice mid-line.
        cutoff = output.rfind("\n", 0, max_chars - 80)
        if cutoff < 0:
            cutoff = max_chars - 80
        output = output[:cutoff] + "\n\n_… truncated for max_chars; pass a larger max_chars or set include_dax=False._"
        truncated = True

    return ok(
        f"DAX context prompt ready ({len(output)} chars).",
        prompt=output,
        char_count=len(output),
        max_chars=max_chars,
        truncated=truncated,
        sections={
            "table_count": len(tables),
            "measure_count": len(measures),
            "relationship_count": len(relationships),
        },
    )


def pbi_set_column_data_type_tool(
    manager: Any,
    *,
    table: str,
    column: str,
    data_type: str,
    format_string: str | None = None,
) -> dict[str, Any]:
    """Set the DataType (and optionally FormatString) of an existing column.

    Works for any column kind (source, calculated, calculated table column).
    Use when Power Query type hints (``Int64.Type`` etc.) are overridden by
    PBI's downstream inference and the column ends up as the wrong type.

    ``data_type`` accepts standard TOM names (``Int64``, ``Decimal``,
    ``Double``, ``String``, ``DateTime``, ``Boolean``, ``Currency``,
    ``Binary``, ``Variant``).
    """
    validate_model_object_name(table)
    validate_model_object_name(column)
    if not data_type or not str(data_type).strip():
        raise PowerBIValidationError("data_type must be a non-empty string.", details={"data_type": data_type})

    def _mutator(state: Any, database: Any, model: Any) -> dict[str, Any]:
        tom = manager.tom
        target_table = find_named(model.Tables, table)
        if target_table is None:
            raise PowerBINotFoundError(f"Table '{table}' was not found.", details={"table": table})
        target_column = find_named(target_table.Columns, column)
        if target_column is None:
            raise PowerBINotFoundError(
                f"Column '{table}[{column}]' was not found.",
                details={"table": table, "column": column},
            )
        before = {
            "data_type": str(target_column.DataType),
            "format_string": target_column.FormatString,
        }
        target_column.DataType = map_enum(tom.DataType, data_type)
        if format_string is not None:
            target_column.FormatString = format_string
        after = {
            "data_type": str(target_column.DataType),
            "format_string": target_column.FormatString,
        }
        return {
            "column": {"table": table, "name": column},
            "before": before,
            "after": after,
        }

    payload = manager.execute_write("set_column_data_type", _mutator)
    return ok(
        f"Column '{table}[{column}]' DataType set to {payload['after']['data_type']}.",
        column=payload["column"],
        before=payload["before"],
        after=payload["after"],
        save_result=payload["save_result"],
        persistence=payload.get("persistence"),
        connection=payload["connection"],
    )


def pbi_validate_model_tool(
    manager: Any,
    *,
    include_warnings: bool = True,
) -> dict[str, Any]:
    """Audit the model for common issues: empty expressions, missing format strings, orphan tables, duplicate measure names."""

    def _reader(state: Any) -> dict[str, Any]:
        model = state.database.Model
        issues: list[dict[str, Any]] = []
        warnings: list[dict[str, Any]] = []

        tables_with_rels: set[str] = set()
        for rel in model.Relationships:
            try:
                tables_with_rels.add(str(rel.FromTable.Name))
                tables_with_rels.add(str(rel.ToTable.Name))
            except Exception:
                pass

        measure_name_index: dict[str, list[str]] = {}

        for table in model.Tables:
            table_name = str(table.Name)
            is_hidden = bool(getattr(table, "IsHidden", False))
            is_calc_group = bool(getattr(table, "CalculationGroup", None))
            has_measures = False

            for measure in table.Measures:
                has_measures = True
                m_name = str(measure.Name)
                m_expr = str(measure.Expression or "").strip()
                m_format = str(getattr(measure, "FormatString", "") or "").strip()
                m_hidden = bool(getattr(measure, "IsHidden", False))

                if not m_expr:
                    issues.append(
                        {
                            "type": "empty_expression",
                            "object": f"{table_name}[{m_name}]",
                            "message": "Measure has an empty DAX expression.",
                        }
                    )

                if not m_format and not m_hidden and not is_hidden and include_warnings:
                    warnings.append(
                        {
                            "type": "missing_format_string",
                            "object": f"{table_name}[{m_name}]",
                            "message": "Visible measure has no format string set.",
                        }
                    )

                measure_name_index.setdefault(m_name.casefold(), []).append(table_name)

            if (
                not is_hidden
                and not is_calc_group
                and table_name not in tables_with_rels
                and not has_measures
                and include_warnings
            ):
                warnings.append(
                    {
                        "type": "orphan_table",
                        "object": table_name,
                        "message": "Table has no relationships and no measures — may be unused.",
                    }
                )

        for m_name_lower, tables_list in measure_name_index.items():
            if len(tables_list) > 1:
                issues.append(
                    {
                        "type": "duplicate_measure_name",
                        "object": m_name_lower,
                        "message": f"Measure name exists in multiple tables: {', '.join(tables_list)}.",
                    }
                )

        return {
            "issues": issues,
            "warnings": warnings,
            "connection": state.snapshot(),
        }

    payload = manager.run_read("validate_model", _reader)
    issue_count = len(payload["issues"])
    warning_count = len(payload["warnings"])
    return ok(
        f"Model audit: {issue_count} issue(s), {warning_count} warning(s).",
        issues=payload["issues"],
        issue_count=issue_count,
        warnings=payload["warnings"],
        warning_count=warning_count,
        connection=payload["connection"],
    )

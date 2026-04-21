"""Model inspection and table/column operations for the Power BI MCP server."""

from __future__ import annotations

import json
from pathlib import Path
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

    payload = manager.run_read("list_tables", _reader)

    if include_row_counts:
        for table_payload in payload["tables"]:
            try:
                query = (
                    "EVALUATE "
                    f"ROW(\"__RowCount\", COUNTROWS({dax_quote_table_name(table_payload['name'])}))"
                )
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
        connection=payload["connection"],
    )

"""Calculation Group CRUD for Power BI Tabular models (compatibility level >= 1470)."""

from __future__ import annotations

from typing import Any

from pbi_connection import (
    PowerBIDuplicateError,
    PowerBINotFoundError,
    PowerBIValidationError,
    find_named,
    ok,
    serialize_value,
)
from security import (
    redact_sensitive_data,
    validate_model_expression,
    validate_model_object_name,
)


def _get_calc_group_tables(model: Any) -> list[Any]:
    return [table for table in model.Tables if getattr(table, "CalculationGroup", None) is not None]


def pbi_list_calc_groups_tool(manager: Any) -> dict[str, Any]:
    """List calculation groups and their items."""

    def _reader(state: Any) -> dict[str, Any]:
        groups: list[dict[str, Any]] = []
        for table in _get_calc_group_tables(state.database.Model):
            group = table.CalculationGroup
            items = []
            for item in group.CalculationItems:
                items.append(
                    {
                        "name": str(item.Name),
                        "expression": redact_sensitive_data(str(getattr(item, "Expression", ""))),
                        "ordinal": int(getattr(item, "Ordinal", -1)),
                        "format_string_expression": redact_sensitive_data(
                            str(getattr(item, "FormatStringDefinition", ""))
                        ),
                    }
                )
            items.sort(key=lambda value: (value["ordinal"], value["name"].casefold()))
            # Calculation groups expose a single column (Name column). Name varies per model.
            column_name = "Name"
            for column in table.Columns:
                column_name = str(column.Name)
                break
            groups.append(
                {
                    "table": str(table.Name),
                    "precedence": int(getattr(group, "Precedence", 0) or 0),
                    "description": serialize_value(getattr(group, "Description", "")),
                    "column_name": column_name,
                    "items": items,
                }
            )
        return {"calc_groups": groups, "connection": state.snapshot()}

    payload = manager.run_read("list_calc_groups", _reader)
    return ok(
        "Calculation groups listed successfully.",
        calc_groups=payload["calc_groups"],
        connection=payload["connection"],
    )


def pbi_create_calc_group_tool(
    manager: Any,
    *,
    table_name: str,
    column_name: str = "Name",
    precedence: int = 0,
    items: list[dict[str, Any]] | None = None,
    overwrite: bool = False,
) -> dict[str, Any]:
    """Create a calculation group table with its Name column and optional calculation items.

    items: list of {name, expression, format_string_expression?, ordinal?}.
    """
    validate_model_object_name(table_name)
    validate_model_object_name(column_name)
    items = items or []
    for item in items:
        if "name" not in item or "expression" not in item:
            raise PowerBIValidationError(
                "Each calc group item requires 'name' and 'expression'.",
                details={"item": item},
            )
        validate_model_object_name(item["name"])
        validate_model_expression(item["expression"], kind="calculation item expression")
        fmt_expr = item.get("format_string_expression")
        if fmt_expr:
            validate_model_expression(fmt_expr, kind="calculation item format")

    def _mutator(state: Any, database: Any, model: Any) -> dict[str, Any]:
        tom = manager.tom
        # Calculation groups require DiscourageImplicitMeasures=True at the model level.
        try:
            if not bool(getattr(model, "DiscourageImplicitMeasures", False)):
                model.DiscourageImplicitMeasures = True
        except Exception:
            pass
        existing = find_named(model.Tables, table_name)
        action = "created"
        if existing is not None:
            if not overwrite:
                raise PowerBIDuplicateError(
                    f"Table '{table_name}' already exists.",
                    details={"table": table_name},
                )
            if getattr(existing, "CalculationGroup", None) is None:
                raise PowerBIValidationError(
                    f"Table '{table_name}' exists and is not a calculation group; refusing to overwrite.",
                    details={"table": table_name},
                )
            action = "updated"
            table = existing
        else:
            table = tom.Table()
            table.Name = table_name
            model.Tables.Add(table)

        # Ensure calculation group + Name column
        if getattr(table, "CalculationGroup", None) is None:
            table.CalculationGroup = tom.CalculationGroup()
        group = table.CalculationGroup
        group.Precedence = precedence

        # Calculation-group tables require exactly one String column whose
        # SourceColumn is the literal "Name". Only the displayed Name may vary.
        if table.Columns.Count == 0:
            col = tom.DataColumn()
            col.Name = column_name
            col.DataType = tom.DataType.String
            col.SourceColumn = "Name"
            table.Columns.Add(col)
        else:
            existing_column = table.Columns[0]
            existing_column.Name = column_name
            existing_column.DataType = tom.DataType.String
            existing_column.SourceColumn = "Name"

        if table.Partitions.Count == 0:
            partition = tom.Partition()
            partition.Name = table_name
            source = tom.CalculationGroupSource()
            partition.Source = source
            table.Partitions.Add(partition)

        # Replace items
        while group.CalculationItems.Count:
            group.CalculationItems.RemoveAt(0)

        created_items = []
        for idx, item in enumerate(items):
            calc_item = tom.CalculationItem()
            calc_item.Name = item["name"]
            calc_item.Expression = item["expression"]
            if item.get("format_string_expression"):
                fmt = tom.FormatStringDefinition()
                fmt.Expression = item["format_string_expression"]
                calc_item.FormatStringDefinition = fmt
            ordinal = item.get("ordinal", idx)
            try:
                calc_item.Ordinal = int(ordinal)
            except Exception:
                pass
            group.CalculationItems.Add(calc_item)
            created_items.append({"name": item["name"], "ordinal": int(ordinal)})

        return {
            "calc_group": {
                "table": table_name,
                "column_name": column_name,
                "precedence": precedence,
                "items": created_items,
            },
            "action": action,
        }

    payload = manager.execute_write("create_calc_group", _mutator)
    return ok(
        f"Calculation group '{table_name}' {payload['action']} successfully.",
        calc_group=payload["calc_group"],
        action=payload["action"],
        save_result=payload["save_result"],
        connection=payload["connection"],
    )


def pbi_delete_calc_group_tool(manager: Any, *, table_name: str) -> dict[str, Any]:
    """Delete a calculation group table."""
    validate_model_object_name(table_name)

    def _mutator(state: Any, database: Any, model: Any) -> dict[str, Any]:
        table = find_named(model.Tables, table_name)
        if table is None:
            raise PowerBINotFoundError(f"Table '{table_name}' was not found.", details={"table": table_name})
        if getattr(table, "CalculationGroup", None) is None:
            raise PowerBIValidationError(
                f"Table '{table_name}' is not a calculation group.",
                details={"table": table_name},
            )
        model.Tables.Remove(table)
        return {"deleted_calc_group": {"table": table_name}}

    payload = manager.execute_write("delete_calc_group", _mutator)
    return ok(
        f"Calculation group '{table_name}' deleted successfully.",
        deleted_calc_group=payload["deleted_calc_group"],
        save_result=payload["save_result"],
        connection=payload["connection"],
    )


__all__ = [
    "pbi_list_calc_groups_tool",
    "pbi_create_calc_group_tool",
    "pbi_delete_calc_group_tool",
]

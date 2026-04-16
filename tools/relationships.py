"""Relationship inspection and management for the Power BI MCP server."""

from __future__ import annotations

from typing import Any

from pbi_connection import (
    PowerBIDuplicateError,
    PowerBINotFoundError,
    PowerBIValidationError,
    find_named,
    map_enum,
    ok,
    serialize_value,
)


def pbi_list_relationships_tool(manager: Any) -> dict[str, Any]:
    """List model relationships."""

    def _reader(state: Any) -> dict[str, Any]:
        relationships = []
        for relationship in state.database.Model.Relationships:
            from_column = relationship.FromColumn
            to_column = relationship.ToColumn
            relationships.append(
                {
                    "name": str(relationship.Name),
                    "from_table": str(from_column.Table.Name),
                    "from_column": str(from_column.Name),
                    "to_table": str(to_column.Table.Name),
                    "to_column": str(to_column.Name),
                    "cardinality": f"{serialize_value(relationship.FromCardinality)}To{serialize_value(relationship.ToCardinality)}",
                    "direction": serialize_value(getattr(relationship, "CrossFilteringBehavior", None)),
                    "is_active": bool(getattr(relationship, "IsActive", True)),
                    "relationship_type": type(relationship).__name__,
                }
            )
        relationships.sort(key=lambda item: item["name"].casefold())
        return {"relationships": relationships, "connection": state.snapshot()}

    payload = manager.run_read("list_relationships", _reader)
    return ok(
        "Relationships listed successfully.",
        relationships=payload["relationships"],
        connection=payload["connection"],
    )


def pbi_create_relationship_tool(
    manager: Any,
    *,
    from_table: str,
    from_column: str,
    to_table: str,
    to_column: str,
    cardinality: str = "oneToMany",
    direction: str = "oneDirection",
    is_active: bool = True,
    relationship_name: str | None = None,
) -> dict[str, Any]:
    """Create a single-column relationship."""

    def _mutator(state: Any, database: Any, model: Any) -> dict[str, Any]:
        tom = manager.tom
        source_table = find_named(model.Tables, from_table)
        target_table = find_named(model.Tables, to_table)
        if source_table is None:
            raise PowerBINotFoundError(f"Table '{from_table}' was not found.", details={"table": from_table})
        if target_table is None:
            raise PowerBINotFoundError(f"Table '{to_table}' was not found.", details={"table": to_table})

        source_column = find_named(source_table.Columns, from_column)
        target_column = find_named(target_table.Columns, to_column)
        if source_column is None:
            raise PowerBINotFoundError(
                f"Column '{from_table}[{from_column}]' was not found.",
                details={"table": from_table, "column": from_column},
            )
        if target_column is None:
            raise PowerBINotFoundError(
                f"Column '{to_table}[{to_column}]' was not found.",
                details={"table": to_table, "column": to_column},
            )

        for existing in model.Relationships:
            existing_from = existing.FromColumn
            existing_to = existing.ToColumn
            same_direction = (
                str(existing_from.Table.Name).casefold() == from_table.casefold()
                and str(existing_from.Name).casefold() == from_column.casefold()
                and str(existing_to.Table.Name).casefold() == to_table.casefold()
                and str(existing_to.Name).casefold() == to_column.casefold()
            )
            reverse_direction = (
                str(existing_from.Table.Name).casefold() == to_table.casefold()
                and str(existing_from.Name).casefold() == to_column.casefold()
                and str(existing_to.Table.Name).casefold() == from_table.casefold()
                and str(existing_to.Name).casefold() == from_column.casefold()
            )
            if same_direction or reverse_direction:
                raise PowerBIDuplicateError(
                    f"A relationship between '{from_table}[{from_column}]' and '{to_table}[{to_column}]' already exists.",
                    details={"existing_relationship": str(existing.Name)},
                )

        relationship = tom.SingleColumnRelationship()
        relationship.Name = relationship_name or f"{from_table}_{from_column}__{to_table}_{to_column}"
        relationship.FromColumn = source_column
        relationship.ToColumn = target_column
        relationship.IsActive = is_active

        from_cardinality, to_cardinality = _map_cardinality(tom, cardinality)
        relationship.FromCardinality = from_cardinality
        relationship.ToCardinality = to_cardinality
        relationship.CrossFilteringBehavior = _map_direction(tom, direction)

        model.Relationships.Add(relationship)
        return {
            "relationship": {
                "name": relationship.Name,
                "from_table": from_table,
                "from_column": from_column,
                "to_table": to_table,
                "to_column": to_column,
                "cardinality": cardinality,
                "direction": direction,
                "is_active": is_active,
            }
        }

    payload = manager.execute_write("create_relationship", _mutator)
    return ok(
        f"Relationship '{payload['relationship']['name']}' created successfully.",
        relationship=payload["relationship"],
        save_result=payload["save_result"],
        connection=payload["connection"],
    )


def _map_cardinality(tom: Any, cardinality: str) -> tuple[Any, Any]:
    token = cardinality.strip().casefold()
    enum_cls = tom.RelationshipEndCardinality
    if token == "onetomany":
        return enum_cls.One, enum_cls.Many
    if token == "manytoone":
        return enum_cls.Many, enum_cls.One
    if token == "onetoone":
        return enum_cls.One, enum_cls.One
    if token == "manytomany":
        many = getattr(enum_cls, "Many")
        return many, many
    raise PowerBIValidationError(
        "cardinality must be one of: oneToMany, manyToOne, oneToOne, manyToMany.",
        details={"cardinality": cardinality},
    )


def _map_direction(tom: Any, direction: str) -> Any:
    token = direction.strip().casefold()
    if token in {"onedirection", "single", "singledirection"}:
        return map_enum(tom.CrossFilteringBehavior, "OneDirection")
    if token in {"bothdirections", "both", "bidirectional"}:
        return map_enum(tom.CrossFilteringBehavior, "BothDirections")
    if token == "automatic":
        return map_enum(tom.CrossFilteringBehavior, "Automatic")
    raise PowerBIValidationError(
        "direction must be one of: oneDirection, bothDirections, automatic.",
        details={"direction": direction},
    )


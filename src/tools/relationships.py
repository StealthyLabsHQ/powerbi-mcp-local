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
from security import validate_model_object_name


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
    validate_model_object_name(from_table)
    validate_model_object_name(from_column)
    validate_model_object_name(to_table)
    validate_model_object_name(to_column)
    if relationship_name:
        validate_model_object_name(relationship_name)

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
        # 1:1 relationships must use BothDirections (SSAS hard requirement).
        effective_direction = direction
        if cardinality.strip().casefold() == "onetoone" and direction.strip().casefold() in {"onedirection", "single", "singledirection"}:
            effective_direction = "bothDirections"
        relationship.CrossFilteringBehavior = _map_direction(tom, effective_direction)

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


def _find_relationship(model: Any, *, name: str | None, from_table: str | None, from_column: str | None, to_table: str | None, to_column: str | None) -> Any:
    for relationship in model.Relationships:
        if name and str(relationship.Name).casefold() == name.casefold():
            return relationship
        if from_table and from_column and to_table and to_column:
            fc = relationship.FromColumn
            tc = relationship.ToColumn
            if (
                str(fc.Table.Name).casefold() == from_table.casefold()
                and str(fc.Name).casefold() == from_column.casefold()
                and str(tc.Table.Name).casefold() == to_table.casefold()
                and str(tc.Name).casefold() == to_column.casefold()
            ):
                return relationship
    return None


def pbi_delete_relationship_tool(
    manager: Any,
    *,
    name: str | None = None,
    from_table: str | None = None,
    from_column: str | None = None,
    to_table: str | None = None,
    to_column: str | None = None,
) -> dict[str, Any]:
    """Delete a relationship by name or by endpoint columns."""
    if not name and not (from_table and from_column and to_table and to_column):
        raise PowerBIValidationError(
            "Provide either 'name' or all of from_table/from_column/to_table/to_column.",
        )
    for value in (name, from_table, from_column, to_table, to_column):
        if value:
            validate_model_object_name(value)

    def _mutator(state: Any, database: Any, model: Any) -> dict[str, Any]:
        relationship = _find_relationship(
            model,
            name=name,
            from_table=from_table,
            from_column=from_column,
            to_table=to_table,
            to_column=to_column,
        )
        if relationship is None:
            raise PowerBINotFoundError(
                "Relationship not found.",
                details={
                    "name": name,
                    "from_table": from_table,
                    "from_column": from_column,
                    "to_table": to_table,
                    "to_column": to_column,
                },
            )
        deleted_name = str(relationship.Name)
        model.Relationships.Remove(relationship)
        return {"deleted_relationship": {"name": deleted_name}}

    payload = manager.execute_write("delete_relationship", _mutator)
    return ok(
        f"Relationship '{payload['deleted_relationship']['name']}' deleted successfully.",
        deleted_relationship=payload["deleted_relationship"],
        save_result=payload["save_result"],
        connection=payload["connection"],
    )


def pbi_update_relationship_tool(
    manager: Any,
    *,
    name: str | None = None,
    from_table: str | None = None,
    from_column: str | None = None,
    to_table: str | None = None,
    to_column: str | None = None,
    cardinality: str | None = None,
    direction: str | None = None,
    is_active: bool | None = None,
    new_name: str | None = None,
) -> dict[str, Any]:
    """Update properties of an existing relationship (cardinality, direction, is_active, name)."""
    if not name and not (from_table and from_column and to_table and to_column):
        raise PowerBIValidationError(
            "Provide either 'name' or all of from_table/from_column/to_table/to_column.",
        )
    for value in (name, from_table, from_column, to_table, to_column, new_name):
        if value:
            validate_model_object_name(value)
    if cardinality is None and direction is None and is_active is None and new_name is None:
        raise PowerBIValidationError(
            "Specify at least one of cardinality, direction, is_active, new_name.",
        )

    def _mutator(state: Any, database: Any, model: Any) -> dict[str, Any]:
        tom = manager.tom
        relationship = _find_relationship(
            model,
            name=name,
            from_table=from_table,
            from_column=from_column,
            to_table=to_table,
            to_column=to_column,
        )
        if relationship is None:
            raise PowerBINotFoundError("Relationship not found.", details={"name": name})

        if cardinality is not None:
            from_card, to_card = _map_cardinality(tom, cardinality)
            relationship.FromCardinality = from_card
            relationship.ToCardinality = to_card
        if direction is not None:
            relationship.CrossFilteringBehavior = _map_direction(tom, direction)
        if is_active is not None:
            relationship.IsActive = is_active
        if new_name is not None:
            relationship.Name = new_name

        return {
            "relationship": {
                "name": str(relationship.Name),
                "cardinality": cardinality,
                "direction": direction,
                "is_active": bool(relationship.IsActive),
            }
        }

    payload = manager.execute_write("update_relationship", _mutator)
    return ok(
        f"Relationship '{payload['relationship']['name']}' updated successfully.",
        relationship=payload["relationship"],
        save_result=payload["save_result"],
        connection=payload["connection"],
    )


def _map_cardinality(tom: Any, cardinality: str) -> tuple[Any, Any]:
    token = cardinality.strip().casefold()
    enum_cls = tom.RelationshipEndCardinality
    # Tabular convention: the "from" side is always the FK ("many"); the "to"
    # side is the lookup ("one"). Both user-facing labels map to the same pair.
    if token in {"onetomany", "manytoone"}:
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

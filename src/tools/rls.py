"""Row-Level Security (RLS) CRUD operations for Power BI models."""

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
from security import (
    redact_sensitive_data,
    validate_model_expression,
    validate_model_object_name,
)


_PERMISSION_ALIASES = {
    "none": "None",
    "read": "Read",
    "readrefresh": "ReadRefresh",
    "refresh": "Refresh",
    "administrator": "Administrator",
    "admin": "Administrator",
}


def _map_permission(tom: Any, permission: str) -> Any:
    token = permission.strip().casefold()
    canonical = _PERMISSION_ALIASES.get(token)
    if canonical is None:
        raise PowerBIValidationError(
            "permission must be one of: None, Read, ReadRefresh, Refresh, Administrator.",
            details={"permission": permission},
        )
    return map_enum(tom.ModelPermission, canonical)


def _serialize_role(role: Any) -> dict[str, Any]:
    members: list[dict[str, Any]] = []
    for member in role.Members:
        members.append(
            {
                "name": serialize_value(getattr(member, "MemberName", getattr(member, "Name", ""))),
                "type": type(member).__name__,
                "identity_provider": serialize_value(getattr(member, "IdentityProvider", None)),
            }
        )
    filters: list[dict[str, Any]] = []
    for permission in role.TablePermissions:
        filters.append(
            {
                "table": str(permission.Table.Name),
                "filter_expression": redact_sensitive_data(str(getattr(permission, "FilterExpression", "") or "")),
            }
        )
    return {
        "name": str(role.Name),
        "description": serialize_value(getattr(role, "Description", "")),
        "model_permission": serialize_value(getattr(role, "ModelPermission", None)),
        "members": members,
        "filters": filters,
    }


def pbi_list_roles_tool(manager: Any) -> dict[str, Any]:
    """List all model roles with their members and table filters."""

    def _reader(state: Any) -> dict[str, Any]:
        roles = [_serialize_role(role) for role in state.database.Model.Roles]
        return {"roles": roles, "connection": state.snapshot()}

    payload = manager.run_read("list_roles", _reader)
    return ok(
        "Roles listed successfully.",
        roles=payload["roles"],
        connection=payload["connection"],
    )


def pbi_create_role_tool(
    manager: Any,
    *,
    name: str,
    permission: str = "Read",
    description: str = "",
    overwrite: bool = False,
) -> dict[str, Any]:
    """Create or update a model role."""
    validate_model_object_name(name)

    def _mutator(state: Any, database: Any, model: Any) -> dict[str, Any]:
        tom = manager.tom
        existing = find_named(model.Roles, name)
        action = "created"
        if existing is not None and not overwrite:
            raise PowerBIDuplicateError(
                f"Role '{name}' already exists.",
                details={"role": name},
            )
        if existing is None:
            role = tom.ModelRole()
            role.Name = name
            model.Roles.Add(role)
        else:
            role = existing
            action = "updated"
        role.ModelPermission = _map_permission(tom, permission)
        if description:
            role.Description = description
        return {"role": {"name": name, "permission": permission}, "action": action}

    payload = manager.execute_write("create_role", _mutator)
    return ok(
        f"Role '{name}' {payload['action']} successfully.",
        role=payload["role"],
        action=payload["action"],
        save_result=payload["save_result"],
        connection=payload["connection"],
    )


def pbi_delete_role_tool(manager: Any, *, name: str) -> dict[str, Any]:
    """Delete a model role."""
    validate_model_object_name(name)

    def _mutator(state: Any, database: Any, model: Any) -> dict[str, Any]:
        role = find_named(model.Roles, name)
        if role is None:
            raise PowerBINotFoundError(f"Role '{name}' was not found.", details={"role": name})
        model.Roles.Remove(role)
        return {"deleted_role": {"name": name}}

    payload = manager.execute_write("delete_role", _mutator)
    return ok(
        f"Role '{name}' deleted successfully.",
        deleted_role=payload["deleted_role"],
        save_result=payload["save_result"],
        connection=payload["connection"],
    )


def pbi_set_role_filter_tool(
    manager: Any,
    *,
    role: str,
    table: str,
    filter_expression: str | None,
) -> dict[str, Any]:
    """Set (or remove when filter_expression is None/empty) the RLS DAX filter on a table for a role."""
    validate_model_object_name(role)
    validate_model_object_name(table)
    if filter_expression:
        validate_model_expression(filter_expression, kind="RLS filter expression")

    def _mutator(state: Any, database: Any, model: Any) -> dict[str, Any]:
        tom = manager.tom
        role_obj = find_named(model.Roles, role)
        if role_obj is None:
            raise PowerBINotFoundError(f"Role '{role}' was not found.", details={"role": role})
        table_obj = find_named(model.Tables, table)
        if table_obj is None:
            raise PowerBINotFoundError(f"Table '{table}' was not found.", details={"table": table})

        existing = None
        for perm in role_obj.TablePermissions:
            if str(perm.Table.Name).casefold() == table.casefold():
                existing = perm
                break

        if not filter_expression:
            if existing is None:
                return {"filter": {"role": role, "table": table, "removed": False}}
            role_obj.TablePermissions.Remove(existing)
            return {"filter": {"role": role, "table": table, "removed": True}}

        if existing is None:
            perm = tom.TablePermission()
            perm.Table = table_obj
            perm.FilterExpression = filter_expression
            role_obj.TablePermissions.Add(perm)
        else:
            existing.FilterExpression = filter_expression
        return {
            "filter": {
                "role": role,
                "table": table,
                "filter_expression": redact_sensitive_data(filter_expression),
            }
        }

    payload = manager.execute_write("set_role_filter", _mutator)
    return ok(
        f"RLS filter applied on '{role}' for table '{table}'.",
        filter=payload["filter"],
        save_result=payload["save_result"],
        connection=payload["connection"],
    )


def pbi_add_role_member_tool(
    manager: Any,
    *,
    role: str,
    member_name: str,
    member_type: str = "external",
    identity_provider: str = "AzureAD",
) -> dict[str, Any]:
    """Add a member (user/group) to a role.

    member_type:
        'external' (default, ExternalModelRoleMember — recommended for Power BI service)
        'windows'  (WindowsModelRoleMember — domain principal SID)
    """
    validate_model_object_name(role)
    if not member_name or not member_name.strip():
        raise PowerBIValidationError("member_name is required.")
    token = member_type.strip().casefold()
    if token not in {"external", "windows"}:
        raise PowerBIValidationError(
            "member_type must be 'external' or 'windows'.",
            details={"member_type": member_type},
        )

    def _mutator(state: Any, database: Any, model: Any) -> dict[str, Any]:
        tom = manager.tom
        role_obj = find_named(model.Roles, role)
        if role_obj is None:
            raise PowerBINotFoundError(f"Role '{role}' was not found.", details={"role": role})

        if token == "external":
            member = tom.ExternalModelRoleMember()
            member.MemberName = member_name
            member.IdentityProvider = identity_provider
        else:
            member = tom.WindowsModelRoleMember()
            member.MemberName = member_name

        role_obj.Members.Add(member)
        return {
            "member": {
                "role": role,
                "name": member_name,
                "type": token,
                "identity_provider": identity_provider if token == "external" else None,
            }
        }

    payload = manager.execute_write("add_role_member", _mutator)
    return ok(
        f"Member '{member_name}' added to role '{role}'.",
        member=payload["member"],
        save_result=payload["save_result"],
        connection=payload["connection"],
    )


def pbi_remove_role_member_tool(manager: Any, *, role: str, member_name: str) -> dict[str, Any]:
    """Remove a member from a role, matching by MemberName (case-insensitive)."""
    validate_model_object_name(role)
    if not member_name or not member_name.strip():
        raise PowerBIValidationError("member_name is required.")

    def _mutator(state: Any, database: Any, model: Any) -> dict[str, Any]:
        role_obj = find_named(model.Roles, role)
        if role_obj is None:
            raise PowerBINotFoundError(f"Role '{role}' was not found.", details={"role": role})
        target = None
        for member in role_obj.Members:
            existing_name = str(getattr(member, "MemberName", getattr(member, "Name", "")))
            if existing_name.casefold() == member_name.casefold():
                target = member
                break
        if target is None:
            raise PowerBINotFoundError(
                f"Member '{member_name}' was not found in role '{role}'.",
                details={"role": role, "member_name": member_name},
            )
        role_obj.Members.Remove(target)
        return {"removed_member": {"role": role, "name": member_name}}

    payload = manager.execute_write("remove_role_member", _mutator)
    return ok(
        f"Member '{member_name}' removed from role '{role}'.",
        removed_member=payload["removed_member"],
        save_result=payload["save_result"],
        connection=payload["connection"],
    )


__all__ = [
    "pbi_list_roles_tool",
    "pbi_create_role_tool",
    "pbi_delete_role_tool",
    "pbi_set_role_filter_tool",
    "pbi_add_role_member_tool",
    "pbi_remove_role_member_tool",
]

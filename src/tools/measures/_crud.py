"""CRUD measure operations: create/rename/delete/list measures and format strings."""

from __future__ import annotations

from typing import Any

from pbi_connection import (
    PowerBIDuplicateError,
    PowerBINotFoundError,
    PowerBIValidationError,
    error_payload,
    find_named,
    ok,
    serialize_value,
)
from security import (
    redact_sensitive_data,
    validate_measure_name,
    validate_model_expression,
    validate_model_object_name,
)


def pbi_list_measures_tool(
    manager: Any,
    *,
    include_hidden: bool = False,
) -> dict[str, Any]:
    """List all model measures."""

    def _reader(state: Any) -> dict[str, Any]:
        measures = []
        for table in state.database.Model.Tables:
            for measure in table.Measures:
                is_hidden = bool(getattr(measure, "IsHidden", False))
                if is_hidden and not include_hidden:
                    continue
                measures.append(
                    {
                        "name": str(measure.Name),
                        "table": str(table.Name),
                        "expression": redact_sensitive_data(str(measure.Expression)),
                        "format_string": serialize_value(getattr(measure, "FormatString", "")),
                        "display_folder": serialize_value(getattr(measure, "DisplayFolder", "")),
                        "description": serialize_value(getattr(measure, "Description", "")),
                        "is_hidden": is_hidden,
                    }
                )
        measures.sort(key=lambda item: (item["table"].casefold(), item["name"].casefold()))
        return {"measures": measures, "connection": state.snapshot()}

    payload = manager.cached_run_read(f"list_measures:h{include_hidden}", "list_measures", _reader)
    return ok(
        "Measures listed successfully.",
        measures=payload["measures"],
        connection=payload["connection"],
    )


def pbi_create_measure_tool(
    manager: Any,
    *,
    table: str,
    name: str,
    expression: str,
    format_string: str = "",
    description: str = "",
    display_folder: str = "",
    is_hidden: bool = False,
    overwrite: bool = True,
) -> dict[str, Any]:
    """Create or update a DAX measure."""
    validate_model_object_name(table)
    validate_measure_name(name)
    validate_model_expression(expression, kind="measure expression")

    def _mutator(state: Any, database: Any, model: Any) -> dict[str, Any]:
        target_table = find_named(model.Tables, table)
        if target_table is None:
            raise PowerBINotFoundError(f"Table '{table}' was not found.", details={"table": table})

        existing = find_named(target_table.Measures, name)
        action = "created"
        if existing is not None and not overwrite:
            raise PowerBIDuplicateError(
                f"Measure '{table}[{name}]' already exists.",
                details={"table": table, "measure": name},
            )

        if existing is None:
            measure = manager.tom.Measure()
            measure.Name = name
            target_table.Measures.Add(measure)
        else:
            measure = existing
            action = "updated"

        measure.Expression = expression
        if format_string:
            measure.FormatString = format_string
        if description:
            measure.Description = description
        if display_folder:
            measure.DisplayFolder = display_folder
        measure.IsHidden = is_hidden

        return {
            "measure": {
                "table": table,
                "name": name,
                "expression": redact_sensitive_data(expression),
                "format_string": format_string or None,
                "description": description or None,
                "display_folder": display_folder or None,
                "is_hidden": is_hidden,
            },
            "action": action,
        }

    payload = manager.execute_write("create_measure", _mutator)
    return ok(
        f"Measure '{table}[{name}]' {payload['action']} successfully.",
        measure=payload["measure"],
        action=payload["action"],
        save_result=payload["save_result"],
        persistence=payload.get("persistence"),
        connection=payload["connection"],
    )


def pbi_delete_measure_tool(manager: Any, *, table: str, name: str) -> dict[str, Any]:
    """Delete a DAX measure."""
    validate_model_object_name(table)
    validate_measure_name(name)

    def _mutator(state: Any, database: Any, model: Any) -> dict[str, Any]:
        target_table = find_named(model.Tables, table)
        if target_table is None:
            raise PowerBINotFoundError(f"Table '{table}' was not found.", details={"table": table})

        measure = find_named(target_table.Measures, name)
        if measure is None:
            raise PowerBINotFoundError(
                f"Measure '{table}[{name}]' was not found.",
                details={"table": table, "measure": name},
            )

        target_table.Measures.Remove(measure)
        return {
            "deleted_measure": {"table": table, "name": name},
        }

    payload = manager.execute_write("delete_measure", _mutator)
    return ok(
        f"Measure '{table}[{name}]' deleted successfully.",
        deleted_measure=payload["deleted_measure"],
        save_result=payload["save_result"],
        persistence=payload.get("persistence"),
        connection=payload["connection"],
    )


def pbi_rename_measure_tool(
    manager: Any,
    *,
    table: str,
    name: str,
    new_name: str,
) -> dict[str, Any]:
    """Rename a DAX measure. Callers must update downstream DAX expressions themselves."""
    validate_model_object_name(table)
    validate_measure_name(name)
    validate_measure_name(new_name)

    def _mutator(state: Any, database: Any, model: Any) -> dict[str, Any]:
        target_table = find_named(model.Tables, table)
        if target_table is None:
            raise PowerBINotFoundError(f"Table '{table}' was not found.", details={"table": table})
        measure = find_named(target_table.Measures, name)
        if measure is None:
            raise PowerBINotFoundError(
                f"Measure '{table}[{name}]' was not found.",
                details={"table": table, "measure": name},
            )
        if new_name.casefold() != name.casefold():
            for candidate_table in model.Tables:
                if find_named(candidate_table.Measures, new_name) is not None:
                    raise PowerBIDuplicateError(
                        f"Measure '{new_name}' already exists in table '{candidate_table.Name}'.",
                        details={"new_name": new_name, "conflict_table": str(candidate_table.Name)},
                    )
        measure.Name = new_name
        return {"rename": {"table": table, "measure_old_name": name, "measure_new_name": new_name}}

    payload = manager.execute_write("rename_measure", _mutator)
    return ok(
        f"Measure '{table}[{name}]' renamed to '{table}[{new_name}]'.",
        rename=payload["rename"],
        save_result=payload["save_result"],
        persistence=payload.get("persistence"),
        connection=payload["connection"],
    )


def pbi_set_format_tool(
    manager: Any,
    *,
    table: str,
    names: list[str],
    format_string: str,
    object_type: str = "measure",
) -> dict[str, Any]:
    """Batch-apply format strings to measures or columns."""
    if not names:
        raise PowerBIValidationError("At least one object name is required.")
    validate_model_object_name(table)
    for object_name in names:
        validate_model_object_name(object_name)
    normalized_type = object_type.strip().casefold()
    if normalized_type not in {"measure", "column"}:
        raise PowerBIValidationError(
            "object_type must be either 'measure' or 'column'.",
            details={"object_type": object_type},
        )

    def _mutator(state: Any, database: Any, model: Any) -> dict[str, Any]:
        target_table = find_named(model.Tables, table)
        if target_table is None:
            raise PowerBINotFoundError(f"Table '{table}' was not found.", details={"table": table})

        collection = target_table.Measures if normalized_type == "measure" else target_table.Columns
        updated = []
        missing = []
        for object_name in names:
            obj = find_named(collection, object_name)
            if obj is None:
                missing.append(object_name)
                continue
            obj.FormatString = format_string
            updated.append(object_name)

        if not updated:
            raise PowerBINotFoundError(
                f"No {normalized_type}s were updated in table '{table}'.",
                details={"table": table, "names": names},
            )

        return {
            "updated": updated,
            "missing": missing,
            "object_type": normalized_type,
            "table": table,
            "format_string": format_string,
        }

    payload = manager.execute_write("set_format", _mutator)
    return ok(
        f"Format string applied to {len(payload['updated'])} {payload['object_type']}(s).",
        updated=payload["updated"],
        missing=payload["missing"],
        object_type=payload["object_type"],
        table=payload["table"],
        format_string=payload["format_string"],
        save_result=payload["save_result"],
        persistence=payload.get("persistence"),
        connection=payload["connection"],
    )


def pbi_create_measures_tool(
    manager: Any,
    *,
    table: str,
    measures: list[dict[str, Any]],
    overwrite: bool = True,
    stop_on_error: bool = False,
    dry_run: bool = False,
) -> dict[str, Any]:
    """Batch-create or update multiple DAX measures with a single SaveChanges call.

    Each item in *measures* must have at minimum: name, expression.
    Optional keys: format_string, description, display_folder, is_hidden.

    With ``dry_run=True``, every measure is name + expression validated and the
    per-item ``planned_action`` is reported (``would_create``/``would_update``/
    ``would_fail``), but no model mutation happens — ``SaveChanges`` is never
    called. Use this for an LLM preflight before committing the batch.
    """
    validate_model_object_name(table)
    if not measures:
        raise PowerBIValidationError("At least one measure is required.")

    for item in measures:
        validate_measure_name(item.get("name", ""))
        validate_model_expression(item.get("expression", ""), kind="measure expression")

    if dry_run:
        # Read-only preflight: walk the live model and report what would happen
        # without touching it. SaveChanges is never invoked.
        def _planner(state: Any) -> dict[str, Any]:
            target_table = find_named(state.database.Model.Tables, table)
            if target_table is None:
                raise PowerBINotFoundError(f"Table '{table}' was not found.", details={"table": table})
            seen_names: set[str] = set()
            plan: list[dict[str, Any]] = []
            would_create = would_update = would_fail = 0
            for item in measures:
                name = str(item.get("name", ""))
                duplicate_in_batch = name.casefold() in seen_names
                seen_names.add(name.casefold())
                existing = find_named(target_table.Measures, name)
                if duplicate_in_batch:
                    plan.append({"name": name, "planned_action": "would_fail", "reason": "duplicate name in batch"})
                    would_fail += 1
                    continue
                if existing is not None and not overwrite:
                    plan.append({"name": name, "planned_action": "would_fail", "reason": "exists and overwrite=False"})
                    would_fail += 1
                    continue
                if existing is not None:
                    plan.append({"name": name, "planned_action": "would_update"})
                    would_update += 1
                else:
                    plan.append({"name": name, "planned_action": "would_create"})
                    would_create += 1
            return {
                "table": table,
                "dry_run": True,
                "plan": plan,
                "would_create": would_create,
                "would_update": would_update,
                "would_fail": would_fail,
            }

        payload = manager.run_read("create_measures_batch_dry_run", _planner)
        return ok(
            f"Dry run: would create {payload['would_create']}, update {payload['would_update']}, "
            f"fail {payload['would_fail']} measure(s) in '{table}'.",
            table=payload["table"],
            dry_run=True,
            plan=payload["plan"],
            would_create=payload["would_create"],
            would_update=payload["would_update"],
            would_fail=payload["would_fail"],
        )

    def _mutator(state: Any, database: Any, model: Any) -> dict[str, Any]:
        target_table = find_named(model.Tables, table)
        if target_table is None:
            raise PowerBINotFoundError(f"Table '{table}' was not found.", details={"table": table})

        results: list[dict[str, Any]] = []
        created = 0
        updated = 0
        failed = 0

        for item in measures:
            name = item.get("name", "")
            expression = item.get("expression", "")
            format_string = str(item.get("format_string", "") or "")
            description = str(item.get("description", "") or "")
            display_folder = str(item.get("display_folder", "") or "")
            is_hidden = bool(item.get("is_hidden", False))

            try:
                existing = find_named(target_table.Measures, name)
                action = "created"
                if existing is not None and not overwrite:
                    raise PowerBIDuplicateError(
                        f"Measure '{table}[{name}]' already exists.",
                        details={"table": table, "measure": name},
                    )
                if existing is None:
                    measure = manager.tom.Measure()
                    measure.Name = name
                    target_table.Measures.Add(measure)
                else:
                    measure = existing
                    action = "updated"

                measure.Expression = expression
                if format_string:
                    measure.FormatString = format_string
                if description:
                    measure.Description = description
                if display_folder:
                    measure.DisplayFolder = display_folder
                measure.IsHidden = is_hidden

                if action == "created":
                    created += 1
                else:
                    updated += 1
                results.append({"name": name, "ok": True, "action": action})
            except Exception as exc:
                failed += 1
                results.append({"name": name, "ok": False, "error": error_payload(exc)["error"]})
                if stop_on_error:
                    break

        return {
            "results": results,
            "created": created,
            "updated": updated,
            "failed": failed,
            "table": table,
        }

    payload = manager.execute_write("create_measures_batch", _mutator)
    return ok(
        f"Batch: {payload['created']} created, {payload['updated']} updated, {payload['failed']} failed in '{table}'.",
        table=payload["table"],
        results=payload["results"],
        created=payload["created"],
        updated=payload["updated"],
        failed=payload["failed"],
        save_result=payload["save_result"],
        persistence=payload.get("persistence"),
        connection=payload["connection"],
    )

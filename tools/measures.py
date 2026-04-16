"""Measure operations for the Power BI MCP server."""

from __future__ import annotations

import re
import textwrap
from pathlib import Path
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
                        "expression": str(measure.Expression),
                        "format_string": serialize_value(getattr(measure, "FormatString", "")),
                        "display_folder": serialize_value(getattr(measure, "DisplayFolder", "")),
                        "description": serialize_value(getattr(measure, "Description", "")),
                        "is_hidden": is_hidden,
                    }
                )
        measures.sort(key=lambda item: (item["table"].casefold(), item["name"].casefold()))
        return {"measures": measures, "connection": state.snapshot()}

    payload = manager.run_read("list_measures", _reader)
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
    if not name.strip():
        raise PowerBIValidationError("Measure name cannot be empty.")
    if not expression.strip():
        raise PowerBIValidationError("Measure expression cannot be empty.")

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
                "expression": expression,
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
        connection=payload["connection"],
    )


def pbi_delete_measure_tool(manager: Any, *, table: str, name: str) -> dict[str, Any]:
    """Delete a DAX measure."""

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
        connection=payload["connection"],
    )


def pbi_import_dax_file_tool(
    manager: Any,
    *,
    path: str,
    table: str = "Measures",
    overwrite: bool = True,
    default_format_string: str = "",
    default_display_folder: str = "",
    stop_on_error: bool = False,
) -> dict[str, Any]:
    """Parse a .dax file and bulk-create measures."""
    measures = _parse_dax_file(Path(path))
    results = []
    created = 0
    updated = 0
    failed = 0

    for measure in measures:
        try:
            response = pbi_create_measure_tool(
                manager,
                table=table,
                name=measure["name"],
                expression=measure["expression"],
                format_string=default_format_string,
                display_folder=default_display_folder,
                overwrite=overwrite,
            )
            action = response["action"]
            if action == "created":
                created += 1
            elif action == "updated":
                updated += 1
            results.append(
                {
                    "name": measure["name"],
                    "ok": True,
                    "action": action,
                }
            )
        except Exception as exc:
            failed += 1
            results.append(
                {
                    "name": measure["name"],
                    "ok": False,
                    "error": error_payload(exc)["error"],
                }
            )
            if stop_on_error:
                break

    return ok(
        f"Imported {created + updated} measure(s) from '{path}'.",
        table=table,
        source_path=str(Path(path).expanduser()),
        parsed_count=len(measures),
        created=created,
        updated=updated,
        failed=failed,
        results=results,
    )


def _parse_dax_file(path: Path) -> list[dict[str, str]]:
    if not path.exists():
        raise PowerBINotFoundError(f"DAX file '{path}' was not found.", details={"path": str(path)})

    raw_text = path.read_text(encoding="utf-8")
    blocks = [block.strip() for block in re.split(r"(?:\r?\n){2,}", raw_text) if block.strip()]
    if not blocks:
        raise PowerBIValidationError(f"DAX file '{path}' is empty.", details={"path": str(path)})

    parsed: list[dict[str, str]] = []
    for index, block in enumerate(blocks, start=1):
        lines = block.splitlines()
        header = lines[0]
        match = re.match(r"^\s*(?P<name>[^=]+?)\s*=\s*(?P<inline>.*)$", header)
        if not match:
            raise PowerBIValidationError(
                f"Invalid measure header in block {index}: '{header}'. Expected 'MeasureName ='",
                details={"path": str(path), "block": index},
            )

        name = match.group("name").strip()
        inline_expression = match.group("inline").strip()
        expression_lines = []
        if inline_expression:
            expression_lines.append(inline_expression)
        expression_lines.extend(lines[1:])
        expression = textwrap.dedent("\n".join(expression_lines)).strip()

        if not name:
            raise PowerBIValidationError(
                f"Block {index} is missing a measure name.",
                details={"path": str(path), "block": index},
            )
        if not expression:
            raise PowerBIValidationError(
                f"Block {index} is missing a DAX expression for measure '{name}'.",
                details={"path": str(path), "block": index, "measure": name},
            )

        parsed.append({"name": name, "expression": expression})

    return parsed


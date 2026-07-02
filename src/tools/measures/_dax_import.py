"""Bulk measure import from .dax files: parsing, comment stripping, creation loop."""

from __future__ import annotations

import re
import textwrap
from pathlib import Path
from typing import Any

from pbi_connection import (
    PowerBINotFoundError,
    PowerBIValidationError,
    error_payload,
    ok,
)
from security import (
    resolve_local_path,
    validate_measure_name,
    validate_model_expression,
    validate_model_object_name,
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
    validate_model_object_name(table)
    resolved_path = resolve_local_path(path, must_exist=True, allowed_extensions={".dax"})
    measures = _parse_dax_file(resolved_path)
    results = []
    created = 0
    updated = 0
    failed = 0

    from . import pbi_create_measure_tool  # late binding: keeps tools.measures patchable

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
        source_path=str(resolved_path),
        parsed_count=len(measures),
        created=created,
        updated=updated,
        failed=failed,
        results=results,
    )


def _parse_dax_file(path: str | Path) -> list[dict[str, str]]:
    resolved = resolve_local_path(str(path), must_exist=True, allowed_extensions={".dax"})
    if not resolved.exists():
        raise PowerBINotFoundError(f"DAX file '{resolved}' was not found.", details={"path": str(resolved)})

    raw_text = resolved.read_text(encoding="utf-8")
    cleaned_text = _strip_dax_comments(raw_text)
    normalized_text = "\n".join(line.rstrip() for line in cleaned_text.splitlines())
    blocks = [block.strip() for block in re.split(r"(?:\n\s*){2,}", normalized_text) if block.strip()]
    if not blocks:
        raise PowerBIValidationError(f"DAX file '{resolved}' is empty.", details={"path": str(resolved)})

    parsed: list[dict[str, str]] = []
    for index, block in enumerate(blocks, start=1):
        lines = block.splitlines()
        header = lines[0]
        match = re.match(r"^\s*(?P<name>[^=]+?)\s*=\s*(?P<inline>.*)$", header)
        if not match:
            raise PowerBIValidationError(
                f"Invalid measure header in block {index}: '{header}'. Expected 'MeasureName ='",
                details={"path": str(resolved), "block": index},
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
                details={"path": str(resolved), "block": index},
            )
        if not expression:
            raise PowerBIValidationError(
                f"Block {index} is missing a DAX expression for measure '{name}'.",
                details={"path": str(resolved), "block": index, "measure": name},
            )
        validate_measure_name(name)
        validate_model_expression(expression, kind="measure expression")

        parsed.append({"name": name, "expression": expression})

    return parsed


def _strip_dax_comments(text: str) -> str:
    """Remove // and /* */ comments while preserving text inside string literals."""
    output: list[str] = []
    index = 0
    in_string = False
    in_line_comment = False
    in_block_comment = False

    while index < len(text):
        char = text[index]
        next_char = text[index + 1] if index + 1 < len(text) else ""

        if in_line_comment:
            if char in "\r\n":
                in_line_comment = False
                output.append(char)
            index += 1
            continue

        if in_block_comment:
            if char == "*" and next_char == "/":
                in_block_comment = False
                index += 2
                continue
            if char in "\r\n":
                output.append(char)
            index += 1
            continue

        if in_string:
            output.append(char)
            if char == '"':
                if next_char == '"':
                    output.append(next_char)
                    index += 2
                    continue
                in_string = False
            index += 1
            continue

        if char == '"':
            in_string = True
            output.append(char)
            index += 1
            continue
        if char == "/" and next_char == "/":
            in_line_comment = True
            index += 2
            continue
        if char == "/" and next_char == "*":
            in_block_comment = True
            index += 2
            continue

        output.append(char)
        index += 1

    return "".join(output)

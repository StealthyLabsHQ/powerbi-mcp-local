"""Persistent PBIX report builder helpers."""

from __future__ import annotations

import os
import subprocess
from pathlib import Path
from typing import Any

from pbi_connection import PowerBIConfigurationError, PowerBIValidationError, ok
from security import resolve_local_path, validate_model_object_name

DATA_TYPE_ALIASES = {
    "string": "String",
    "text": "String",
    "int": "Int64",
    "integer": "Int64",
    "int64": "Int64",
    "whole": "Int64",
    "double": "Double",
    "float": "Double",
    "number": "Double",
    "decimal": "Decimal",
    "datetime": "DateTime",
    "date": "DateTime",
    "boolean": "Boolean",
    "bool": "Boolean",
}
PBIX_DATA_TYPES = {"String", "Int64", "Double", "DateTime", "Decimal", "Boolean"}

# v0.13.1: Vertipaq DBCC string-store hardening. Empty Import tables with
# String columns leave the dictionary un-primed on disk → DBCC fails on
# reopen with "Database consistency checks failed while checking the
# string store". A single sentinel row per affected table primes the
# dictionary and the segments so DBCC's dict ↔ segment check passes.
_SENTINEL_VALUES_BY_TYPE: dict[str, Any] = {
    "String": "_seed_",
    "Int64": 0,
    "Double": 0.0,
    "Decimal": 0.0,
    "DateTime": "1899-12-30T00:00:00",
    "Boolean": False,
}


def _load_pbix_builder() -> Any:
    try:
        from pbix_mcp.builder import PBIXBuilder
    except Exception as exc:  # pragma: no cover - exercised via tests with monkeypatch
        raise PowerBIConfigurationError(
            "Persistent PBIX builder is unavailable. Install optional dependency in Python 3.11: pip install pbix-mcp==0.9.2",
            details={"dependency": "pbix-mcp==0.9.2", "reason": str(exc)},
        ) from exc
    return PBIXBuilder


def _require_mapping(value: Any, field: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        raise PowerBIValidationError(f"{field} must be an object.")
    return value


def _require_list(value: Any, field: str) -> list[Any]:
    if not isinstance(value, list):
        raise PowerBIValidationError(f"{field} must be a list.")
    return value


def _validate_columns(table_name: str, columns: Any) -> list[dict[str, Any]]:
    result: list[dict[str, Any]] = []
    for index, item in enumerate(_require_list(columns, f"tables[{table_name}].columns")):
        column = _require_mapping(item, f"tables[{table_name}].columns[{index}]")
        name = str(column.get("name", "")).strip()
        raw_data_type = str(column.get("data_type", column.get("type", ""))).strip()
        data_type = DATA_TYPE_ALIASES.get(raw_data_type.casefold(), raw_data_type)
        validate_model_object_name(name)
        if data_type not in PBIX_DATA_TYPES:
            raise PowerBIValidationError(
                "Column data_type must be one of String, Int64, Double, DateTime, Decimal, Boolean.",
                details={"table": table_name, "column": name, "data_type": raw_data_type},
            )
        normalized = dict(column)
        normalized["name"] = name
        normalized["data_type"] = data_type
        result.append(normalized)
    if not result:
        raise PowerBIValidationError("Tables must define at least one column.", details={"table": table_name})
    return result


def _validate_tables(tables: Any) -> list[dict[str, Any]]:
    result: list[dict[str, Any]] = []
    for index, item in enumerate(_require_list(tables, "tables")):
        table = _require_mapping(item, f"tables[{index}]")
        name = str(table.get("name", "")).strip()
        validate_model_object_name(name)
        rows = table.get("rows", [])
        if rows is not None and not isinstance(rows, list):
            raise PowerBIValidationError("Table rows must be a list.", details={"table": name})
        normalized = dict(table)
        normalized["name"] = name
        normalized["columns"] = _validate_columns(name, table.get("columns", []))
        normalized["rows"] = rows or []
        if normalized.get("source_csv"):
            normalized["source_csv"] = str(
                resolve_local_path(str(normalized["source_csv"]), must_exist=True, allowed_extensions={".csv"})
            )
        if normalized.get("mode", "import") not in {"import", "directquery"}:
            raise PowerBIValidationError("Table mode must be 'import' or 'directquery'.", details={"table": name})
        result.append(normalized)
    if not result:
        raise PowerBIValidationError("At least one table is required.")
    return result


def _validate_measures(measures: Any) -> list[dict[str, Any]]:
    result: list[dict[str, Any]] = []
    for index, item in enumerate(_require_list(measures or [], "measures")):
        measure = _require_mapping(item, f"measures[{index}]")
        table = str(measure.get("table", "")).strip()
        name = str(measure.get("name", "")).strip()
        expression = str(measure.get("expression", "")).strip()
        validate_model_object_name(table)
        validate_model_object_name(name)
        if not expression:
            raise PowerBIValidationError("Measure expression cannot be empty.", details={"measure": name})
        normalized = {"table": table, "name": name, "expression": expression}
        if measure.get("format_string") is not None:
            normalized["format_string"] = str(measure["format_string"])
        result.append(normalized)
    return result


def _validate_relationships(relationships: Any) -> list[dict[str, Any]]:
    result: list[dict[str, Any]] = []
    required = ("from_table", "from_column", "to_table", "to_column")
    for index, item in enumerate(_require_list(relationships or [], "relationships")):
        relationship = _require_mapping(item, f"relationships[{index}]")
        normalized = {key: str(relationship.get(key, "")).strip() for key in required}
        for key, value in normalized.items():
            validate_model_object_name(value)
            if not value:
                raise PowerBIValidationError(f"Relationship {key} cannot be empty.", details={"relationship": index})
        result.append(normalized)
    return result


def _validate_pages(pages: Any) -> list[dict[str, Any]]:
    result: list[dict[str, Any]] = []
    for index, item in enumerate(_require_list(pages or [], "pages")):
        page = _require_mapping(item, f"pages[{index}]")
        name = str(page.get("name", f"Page {index + 1}")).strip()
        validate_model_object_name(name)
        visuals = _require_list(page.get("visuals", []), f"pages[{index}].visuals")
        result.append({"name": name, "visuals": visuals})
    return result


def _validation_issue_dict(issue: Any) -> dict[str, Any]:
    if isinstance(issue, dict):
        return issue
    level = getattr(issue, "level", None)
    message = getattr(issue, "message", None)
    return {"level": str(level or ""), "message": str(message or issue)}


def _prime_string_store(table: dict[str, Any]) -> dict[str, Any]:
    """Inject a single sentinel row when a table is at DBCC risk.

    The risk is: ``mode='import'`` + ``rows=[]`` + no ``source_csv`` /
    ``source_db`` + at least one ``String`` column. Vertipaq's
    consistency check rejects the resulting un-primed dictionary on
    reopen. Adding one typed sentinel row is the minimal mutation that
    primes the dictionary while leaving the column count, names, and
    types intact.
    """
    if table.get("rows"):
        return table
    if str(table.get("mode", "import")).lower() != "import":
        return table
    if table.get("source_csv") or table.get("source_db"):
        return table
    columns = table.get("columns") or []
    has_string = any(
        isinstance(col, dict) and str(col.get("data_type", "")) == "String" for col in columns
    )
    if not has_string:
        return table
    sentinel_row: list[Any] = []
    for col in columns:
        data_type = str(col.get("data_type", "")) if isinstance(col, dict) else ""
        sentinel_row.append(_SENTINEL_VALUES_BY_TYPE.get(data_type, None))
    primed = dict(table)
    primed["rows"] = [sentinel_row]
    primed["__primed_string_store__"] = True
    return primed


def pbi_create_persistent_report_tool(
    output_path: str,
    tables: list[dict[str, Any]],
    measures: list[dict[str, Any]] | None = None,
    relationships: list[dict[str, Any]] | None = None,
    pages: list[dict[str, Any]] | None = None,
    open_after_create: bool = False,
    prime_string_store: bool = True,
    style_preset: str | None = None,
    style_wallpaper_path: str | None = None,
    style_custom_spec: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """Create a persistent PBIX with DataModel, DAX measures, relationships, pages, and native visuals.

    ``prime_string_store`` (default ``True``) injects a one-row sentinel
    into every empty Import table whose schema declares a String column,
    so Vertipaq's DBCC string-store consistency check passes on reopen.
    Disable only when you know the table will be populated by Power BI
    Desktop's first refresh (e.g. via a DirectQuery binding or a Power
    Query expression set after the PBIX is created).

    ``style_preset`` (optional) names a built-in styling preset
    (``glassmorph_dark``, ``glassmorph_light``, ``neon_cyber``,
    ``minimal_corporate``, ``dark_pro``, or ``custom``). When set, the
    matching preset is applied to the file via
    ``pbi_apply_style_preset_tool`` right after save, so the produced
    PBIX opens with wallpaper, card chrome, accent colours, and the
    preset's theme already wired up. ``style_wallpaper_path`` overrides
    the preset's default wallpaper; ``style_custom_spec`` is required
    when ``style_preset='custom'``.
    """
    output = resolve_local_path(output_path, must_exist=False, allowed_extensions={".pbix"})
    output.parent.mkdir(parents=True, exist_ok=True)

    normalized_tables = _validate_tables(tables)
    if prime_string_store:
        normalized_tables = [_prime_string_store(table) for table in normalized_tables]
    primed_table_names = [t["name"] for t in normalized_tables if t.get("__primed_string_store__")]
    normalized_measures = _validate_measures(measures)
    normalized_relationships = _validate_relationships(relationships)
    normalized_pages = _validate_pages(pages)

    builder_class = _load_pbix_builder()
    builder = builder_class()

    for table in normalized_tables:
        builder.add_table(
            table["name"],
            table["columns"],
            table["rows"],
            source_csv=table.get("source_csv"),
            source_db=table.get("source_db"),
            mode=table.get("mode", "import"),
        )
    for measure in normalized_measures:
        # Prefer the clean keyword-only ``format_string`` parameter when
        # the installed pbix-mcp build supports it (≥0.9.3 / upstream
        # patched 0.9.2). Fall back to the legacy post-hoc dict mutation
        # so older deployments still pick up the format string.
        try:
            builder.add_measure(
                measure["table"],
                measure["name"],
                measure["expression"],
                format_string=measure.get("format_string"),
            )
        except TypeError:
            builder.add_measure(measure["table"], measure["name"], measure["expression"])
            builder_measures = getattr(builder, "_measures", None)
            if isinstance(builder_measures, list) and builder_measures:
                builder_measures[-1]["format_string"] = measure.get("format_string")
    for relationship in normalized_relationships:
        builder.add_relationship(
            relationship["from_table"],
            relationship["from_column"],
            relationship["to_table"],
            relationship["to_column"],
        )
    for page in normalized_pages:
        builder.add_page(page["name"], page["visuals"])

    pre_build_issues = []
    pre_build_check = getattr(builder, "_pre_build_checks", None)
    if callable(pre_build_check):
        pre_build_issues = [_validation_issue_dict(issue) for issue in pre_build_check()]

    try:
        builder.save(str(output))
    except Exception as exc:
        raise PowerBIConfigurationError(
            "Persistent PBIX builder failed to save the report.",
            details={"output_path": str(output), "reason": str(exc)},
        ) from exc

    validation_issues = []
    validate = getattr(builder, "validate", None)
    if callable(validate):
        validation_issues = [_validation_issue_dict(issue) for issue in validate()]

    opened = False
    if open_after_create:
        if os.name != "nt":
            raise PowerBIConfigurationError("open_after_create is only supported on Windows.")
        subprocess.Popen(["cmd", "/c", "start", "", str(output)], shell=False)
        opened = True

    styling_result: dict[str, Any] | None = None
    if style_preset:
        # Lazy import to avoid bringing the styling layer into module
        # import time for callers that don't use it.
        from .styling import pbi_apply_style_preset_tool

        styling_result = pbi_apply_style_preset_tool(
            pbix_path=str(output),
            preset=style_preset,
            wallpaper_path=style_wallpaper_path,
            custom_spec=style_custom_spec,
        )

    return ok(
        "Persistent PBIX report created successfully.",
        output_path=str(Path(output)),
        size_bytes=output.stat().st_size,
        table_count=len(normalized_tables),
        measure_count=len(normalized_measures),
        relationship_count=len(normalized_relationships),
        page_count=len(normalized_pages),
        pre_build_issues=pre_build_issues,
        validation_issues=validation_issues,
        opened=opened,
        primed_string_store_tables=primed_table_names,
        styling=styling_result,
    )


__all__ = ["pbi_create_persistent_report_tool"]

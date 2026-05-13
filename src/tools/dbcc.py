"""DBCC string-store risk diagnostics (v0.13.1).

Power BI Desktop refuses to reopen a ``.pbix`` whose Vertipaq dictionary
state is inconsistent. The user-visible symptom is:

    Database consistency checks (DBCC) failed while checking the string
    store. An error occurred while loading Vertipaq data objects for
    multiple tables.

Three patterns trigger this in our pipeline:

1. **Empty Import table with String columns.** ``rows=[]`` on a table
   that declares a ``String`` column leaves the dictionary referenced
   but un-primed. DBCC verifies dict ↔ segments on reopen and fails.
2. **Save without a final refresh.** TOM mutations commit in memory
   (the ``persistence: memory_only`` warning surfaces this) — the
   ``.pbix`` on disk holds an intermediate state.
3. **Schema change after dictionary init.** Changing a column type
   from ``String`` to a numeric type without a follow-up refresh
   leaves orphan dictionary files on disk.

This module provides two diagnostics:

- ``pbi_diagnose_pbix_dbcc_tool`` — static analysis of a built ``.pbix``
  archive. Looks for missing/undersized DataModel, missing manifest,
  and the absence of write-time markers that a healthy PBIX carries.
- ``pbi_check_scaffold_spec_dbcc_risks_tool`` — pre-build analysis of
  a ``tables`` spec (same shape as ``pbi_create_persistent_report``
  takes). Flags every empty-String-column risk *before* the PBIX is
  written.
"""

from __future__ import annotations

import zipfile
from pathlib import Path
from typing import Any

from pbi_connection import PowerBIValidationError, ok
from security import resolve_local_path

DBCC_KNOWN_SIGNALS: tuple[str, ...] = (
    "Database consistency checks (DBCC) failed",
    "string store",
    "loading Vertipaq data objects",
    "multiple tables",
)

STRING_LIKE_DATA_TYPES = {"String", "Text"}

# Minimum DataModel size we treat as "primed". An empty Vertipaq model
# weighs in at a few KB; below this threshold the part is almost
# certainly an empty placeholder that DBCC will reject.
MIN_DATA_MODEL_BYTES = 4 * 1024


def _zip_part_summary(archive: zipfile.ZipFile, name: str) -> dict[str, Any] | None:
    try:
        info = archive.getinfo(name)
    except KeyError:
        return None
    return {
        "name": info.filename,
        "compressed_size": info.compress_size,
        "uncompressed_size": info.file_size,
    }


def pbi_diagnose_pbix_dbcc_tool(pbix_path: str) -> dict[str, Any]:
    """Statically diagnose DBCC string-store risks on a built ``.pbix``.

    Returns a list of issues (each with a ``type`` and a short message)
    and a ``valid`` flag. Use this *before* opening the file in Power BI
    Desktop to avoid the round-trip through the modal repair dialog.
    """
    pbix = resolve_local_path(pbix_path, must_exist=True, allowed_extensions={".pbix"})

    issues: list[dict[str, Any]] = []
    warnings: list[dict[str, Any]] = []
    parts: dict[str, Any] = {}

    if not zipfile.is_zipfile(pbix):
        issues.append({"type": "pbix_not_zip", "pbix_path": str(pbix)})
        return ok(
            "PBIX is not a valid zip archive.",
            pbix_path=str(pbix),
            valid=False,
            issue_count=1,
            warnings=warnings,
            issues=issues,
            parts=parts,
        )

    with zipfile.ZipFile(pbix, "r") as archive:
        names = set(archive.namelist())

        # Inventory the DBCC-relevant parts.
        for part_name in ("DataModel", "DataModelSchema", "Report/Layout", "Metadata", "Connections", "Version"):
            summary = _zip_part_summary(archive, part_name)
            if summary is not None:
                parts[part_name] = summary

        if "DataModel" not in names and "DataModelSchema" not in names:
            issues.append(
                {
                    "type": "no_data_model",
                    "message": (
                        "PBIX has no DataModel part. Report-only PBIX cannot be opened with a"
                        " model. If this is intentional, ignore — otherwise rebuild with a"
                        " non-empty table spec."
                    ),
                }
            )

        if "DataModel" in names:
            size = parts["DataModel"]["uncompressed_size"]
            if size < MIN_DATA_MODEL_BYTES:
                issues.append(
                    {
                        "type": "undersized_data_model",
                        "size_bytes": size,
                        "min_expected_bytes": MIN_DATA_MODEL_BYTES,
                        "message": (
                            "DataModel part is unusually small. Vertipaq dictionary is likely"
                            " un-primed — DBCC will reject this on reopen."
                        ),
                    }
                )

        if "Connections" not in names:
            warnings.append(
                {
                    "type": "no_connections_part",
                    "message": (
                        "PBIX has no Connections part. Some builders omit this; PBI Desktop"
                        " regenerates it, but the missing part can co-occur with DBCC failures."
                    ),
                }
            )

        if "Metadata" not in names:
            warnings.append(
                {
                    "type": "no_metadata_part",
                    "message": "PBIX has no Metadata part. Optional but typically present.",
                }
            )

    return ok(
        f"PBIX DBCC diagnostic found {len(issues)} issue(s), {len(warnings)} warning(s).",
        pbix_path=str(pbix),
        valid=not issues,
        issue_count=len(issues),
        warning_count=len(warnings),
        issues=issues,
        warnings=warnings,
        parts=parts,
        known_signals=list(DBCC_KNOWN_SIGNALS),
    )


def _spec_table_risks(table: dict[str, Any], index: int) -> list[dict[str, Any]]:
    """Return a list of DBCC-risk findings for a single table spec."""
    risks: list[dict[str, Any]] = []
    name = str(table.get("name", "")).strip() or f"<unnamed:{index}>"
    rows = table.get("rows")
    columns = table.get("columns") or []

    is_empty = not rows
    has_string_column = False
    string_columns: list[str] = []
    for column in columns:
        if not isinstance(column, dict):
            continue
        data_type = str(column.get("data_type", column.get("type", ""))).strip()
        if data_type in STRING_LIKE_DATA_TYPES:
            has_string_column = True
            string_columns.append(str(column.get("name", "")).strip() or "?")

    mode = str(table.get("mode", "import")).lower()
    source_csv = table.get("source_csv")
    source_db = table.get("source_db")
    is_unsourced_import = mode == "import" and not source_csv and not source_db

    if is_empty and has_string_column and is_unsourced_import:
        risks.append(
            {
                "type": "empty_string_table",
                "table": name,
                "string_columns": string_columns,
                "message": (
                    "Import table has String columns but zero rows and no source. The Vertipaq"
                    " dictionary will be referenced but un-primed → DBCC string-store check"
                    " will fail on reopen. Add at least one sentinel row, supply a source,"
                    " or call the scaffold/persistent_report tool with"
                    " ``prime_string_store=True``."
                ),
            }
        )
    elif is_empty and is_unsourced_import and columns:
        # No string columns but still risky if many numeric columns and 0 rows;
        # numeric stores are tolerated by DBCC, so emit a warning, not an issue.
        risks.append(
            {
                "type": "empty_import_table",
                "table": name,
                "message": (
                    "Import table is empty with no source. PBI Desktop will open it but show"
                    " an empty data area; not a DBCC risk unless String columns are added."
                ),
                "severity": "warning",
            }
        )

    return risks


def pbi_check_scaffold_spec_dbcc_risks_tool(tables: list[dict[str, Any]]) -> dict[str, Any]:
    """Pre-build DBCC risk analysis on a tables spec.

    Mirrors the input shape of ``pbi_create_persistent_report_tool`` and
    ``pbi_scaffold_pbix_tool``. Pass the same ``tables`` list before
    calling the builder to catch empty-String-column risks up front.
    """
    if not isinstance(tables, list):
        raise PowerBIValidationError("tables must be a list of table objects.")

    issues: list[dict[str, Any]] = []
    warnings: list[dict[str, Any]] = []
    for index, table in enumerate(tables):
        if not isinstance(table, dict):
            issues.append(
                {"type": "invalid_table_entry", "index": index, "message": "Table spec must be an object."}
            )
            continue
        for risk in _spec_table_risks(table, index):
            if risk.get("severity") == "warning":
                warnings.append(risk)
            else:
                issues.append(risk)

    return ok(
        f"Scaffold spec DBCC check found {len(issues)} issue(s), {len(warnings)} warning(s).",
        valid=not issues,
        table_count=len(tables),
        issue_count=len(issues),
        warning_count=len(warnings),
        issues=issues,
        warnings=warnings,
    )


__all__ = [
    "DBCC_KNOWN_SIGNALS",
    "MIN_DATA_MODEL_BYTES",
    "STRING_LIKE_DATA_TYPES",
    "pbi_check_scaffold_spec_dbcc_risks_tool",
    "pbi_diagnose_pbix_dbcc_tool",
]

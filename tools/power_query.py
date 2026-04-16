"""Power Query (M language) tools for the Power BI MCP server.

These tools read and write the M expressions on table partitions,
enabling programmatic data source setup without the PBI Desktop UI.
"""

from __future__ import annotations

from typing import Any

from pbi_connection import PowerBINotFoundError, find_named, ok


# ── Helpers ──────────────────────────────────────────────────────────


def _get_partition(model: Any, table_name: str) -> tuple[Any, Any]:
    """Return (table, first_partition) or raise PowerBINotFoundError."""
    table = find_named(model.Tables, table_name)
    if table is None:
        raise PowerBINotFoundError(
            f"Table '{table_name}' was not found.",
            details={"table": table_name},
        )
    if table.Partitions.Count == 0:
        raise PowerBINotFoundError(
            f"Table '{table_name}' has no partitions.",
            details={"table": table_name},
        )
    return table, table.Partitions[0]


def _build_excel_m(
    excel_path: str,
    sheet_name: str,
    promote_headers: bool = True,
) -> str:
    """Generate M code to import a sheet from an Excel workbook.

    Uses ``Excel.Workbook`` and optionally promotes the first row to headers.
    Backslashes in *excel_path* are doubled for the M string literal.
    """
    safe_path = excel_path.replace("\\", "\\\\")
    steps = [
        f'    Source = Excel.Workbook(File.Contents("{safe_path}"), null, true)',
        f'    Sheet = Source{{[Item="{sheet_name}",Kind="Sheet"]}}[Data]',
    ]
    if promote_headers:
        steps.append(
            "    Promoted = Table.PromoteHeaders(Sheet, [PromoteAllScalars=true])"
        )
        final = "Promoted"
    else:
        final = "Sheet"

    return "let\n" + ",\n".join(steps) + f"\nin\n    {final}"


# ── Tool implementations ─────────────────────────────────────────────


def pbi_get_power_query_tool(
    manager: Any,
    *,
    table: str,
) -> dict[str, Any]:
    """Read the Power Query (M) expression for a table's partition."""

    def _reader(state: Any, database: Any, model: Any) -> dict[str, Any]:
        tbl, partition = _get_partition(model, table)
        expr = str(partition.Expression) if partition.Expression else ""
        source_type = str(partition.SourceType) if hasattr(partition, "SourceType") else "unknown"
        return {
            "table": str(tbl.Name),
            "partition": str(partition.Name),
            "source_type": source_type,
            "m_expression": expr,
        }

    payload = manager.execute_read(_reader)
    return ok(
        f"Power Query expression retrieved for table '{table}'.",
        **payload,
    )


def pbi_set_power_query_tool(
    manager: Any,
    *,
    table: str,
    m_expression: str,
    refresh_after: bool = False,
) -> dict[str, Any]:
    """Write or update the Power Query (M) expression for a table.

    If *refresh_after* is ``True`` a full refresh is requested after the
    expression is saved (equivalent to clicking "Refresh" in the UI).
    """

    def _mutator(state: Any, database: Any, model: Any) -> dict[str, Any]:
        tbl, partition = _get_partition(model, table)
        old_expr = str(partition.Expression) if partition.Expression else ""
        partition.Expression = m_expression

        if refresh_after and hasattr(model, "RequestRefresh"):
            tom = manager.tom
            if hasattr(tom, "RefreshType"):
                tbl.RequestRefresh(tom.RefreshType.Full)

        return {
            "table": str(tbl.Name),
            "partition": str(partition.Name),
            "old_expression_length": len(old_expr),
            "new_expression_length": len(m_expression),
            "refresh_requested": refresh_after,
        }

    payload = manager.execute_write("set_power_query", _mutator)
    return ok(
        f"Power Query expression updated for table '{table}'.",
        **payload,
    )


def pbi_create_import_query_tool(
    manager: Any,
    *,
    table: str,
    excel_path: str,
    sheet_name: str,
    promote_headers: bool = True,
    refresh_after: bool = True,
) -> dict[str, Any]:
    """Generate and inject an Excel import Power Query for a table.

    Builds the M expression automatically from *excel_path* and *sheet_name*,
    sets it on the table's partition, and optionally triggers a refresh.
    """
    m_expression = _build_excel_m(excel_path, sheet_name, promote_headers)

    def _mutator(state: Any, database: Any, model: Any) -> dict[str, Any]:
        tbl, partition = _get_partition(model, table)
        partition.Expression = m_expression

        if refresh_after and hasattr(model, "RequestRefresh"):
            tom = manager.tom
            if hasattr(tom, "RefreshType"):
                tbl.RequestRefresh(tom.RefreshType.Full)

        return {
            "table": str(tbl.Name),
            "partition": str(partition.Name),
            "excel_path": excel_path,
            "sheet_name": sheet_name,
            "promote_headers": promote_headers,
            "refresh_requested": refresh_after,
            "m_expression": m_expression,
        }

    payload = manager.execute_write("create_import_query", _mutator)
    return ok(
        f"Excel import query created for table '{table}' from sheet '{sheet_name}'.",
        **payload,
    )


def pbi_bulk_import_excel_tool(
    manager: Any,
    *,
    excel_path: str,
    sheet_table_map: dict[str, str] | None = None,
    promote_headers: bool = True,
    refresh_after: bool = True,
) -> dict[str, Any]:
    """Bulk-create Excel import queries for multiple tables at once.

    *sheet_table_map* maps sheet names to PBI table names, e.g.:
    ``{"FaitsCA": "FaitsCA", "Dim_Temps": "Dim_Temps"}``.

    If *sheet_table_map* is ``None``, maps each sheet to a PBI table with
    the same name (assumes sheet names == table names).
    """

    def _mutator(state: Any, database: Any, model: Any) -> dict[str, Any]:
        mapping = sheet_table_map
        if mapping is None:
            # Auto-map: find all PBI tables and try to match sheet names
            mapping = {}
            for tbl in model.Tables:
                mapping[str(tbl.Name)] = str(tbl.Name)

        results = []
        for sheet, tbl_name in mapping.items():
            tbl = find_named(model.Tables, tbl_name)
            if tbl is None:
                results.append({
                    "table": tbl_name,
                    "sheet": sheet,
                    "status": "skipped",
                    "reason": f"Table '{tbl_name}' not found in model",
                })
                continue

            if tbl.Partitions.Count == 0:
                results.append({
                    "table": tbl_name,
                    "sheet": sheet,
                    "status": "skipped",
                    "reason": f"Table '{tbl_name}' has no partitions",
                })
                continue

            m_code = _build_excel_m(excel_path, sheet, promote_headers)
            tbl.Partitions[0].Expression = m_code

            if refresh_after:
                tom = manager.tom
                if hasattr(tom, "RefreshType"):
                    tbl.RequestRefresh(tom.RefreshType.Full)

            results.append({
                "table": tbl_name,
                "sheet": sheet,
                "status": "created",
            })

        return {
            "excel_path": excel_path,
            "results": results,
            "created": sum(1 for r in results if r["status"] == "created"),
            "skipped": sum(1 for r in results if r["status"] == "skipped"),
            "refresh_requested": refresh_after,
        }

    payload = manager.execute_write("bulk_import_excel", _mutator)
    created = payload.get("created", 0)
    skipped = payload.get("skipped", 0)
    return ok(
        f"Bulk import done: {created} queries created, {skipped} skipped.",
        **payload,
    )

"""Power Query (M language) tools for the Power BI MCP server."""

from __future__ import annotations

import os
import re
from pathlib import Path
from typing import Any

from pbi_connection import (
    PowerBIConfigurationError,
    PowerBINotFoundError,
    PowerBIValidationError,
    find_named,
    normalize_token,
    ok,
    serialize_value,
)
from security import (
    inspect_excel_archive,
    redact_sensitive_data,
    resolve_local_path,
    validate_expression_text,
    validate_model_object_name,
)


def _m_string(value: str) -> str:
    return '"' + str(value).replace('"', '""') + '"'


# ── M expression security ────────────────────────────────────────────

# M functions that can make network calls or access external databases.
# Blocked by default to prevent SSRF and data exfiltration.
_M_BLOCKED_FUNCTIONS = [
    re.compile(r"\bWeb\.Contents\b", re.IGNORECASE),
    re.compile(r"\bWeb\.Page\b", re.IGNORECASE),
    re.compile(r"\bWeb\.BrowserContents\b", re.IGNORECASE),
    re.compile(r"\bExpression\.Evaluate\b", re.IGNORECASE),
    re.compile(r"\bValue\.NativeQuery\b", re.IGNORECASE),
    re.compile(r"\bOData\.Feed\b", re.IGNORECASE),
    re.compile(r"\bSql\.Database\b", re.IGNORECASE),
    re.compile(r"\bSql\.Databases\b", re.IGNORECASE),
    re.compile(r"\bOracle\.Database\b", re.IGNORECASE),
    re.compile(r"\bPostgreSQL\.Database\b", re.IGNORECASE),
    re.compile(r"\bMySQL\.Database\b", re.IGNORECASE),
    re.compile(r"\bOdbc\.DataSource\b", re.IGNORECASE),
    re.compile(r"\bOdbc\.Query\b", re.IGNORECASE),
    re.compile(r"\bOleDb\.DataSource\b", re.IGNORECASE),
    re.compile(r"\bOleDb\.Query\b", re.IGNORECASE),
    re.compile(r"\bSharePoint\.\w+", re.IGNORECASE),
    re.compile(r"\bActiveDirectory\.\w+", re.IGNORECASE),
    re.compile(r"\bAzureStorage\.\w+", re.IGNORECASE),
    re.compile(r"#shared\b", re.IGNORECASE),
]

def _validate_m_expression(expression: str) -> None:
    validate_expression_text(expression)
    text = expression.strip()
    if not text:
        raise PowerBIValidationError("m_expression cannot be empty.")

    # Syntax check: balanced delimiters
    stack: list[tuple[str, int]] = []
    pairs = {"(": ")", "[": "]", "{": "}"}
    in_string = False
    index = 0
    while index < len(text):
        char = text[index]
        if char == '"':
            if in_string and index + 1 < len(text) and text[index + 1] == '"':
                index += 2
                continue
            in_string = not in_string
        elif not in_string and char in pairs:
            stack.append((char, index))
        elif not in_string and char in pairs.values():
            if not stack or pairs[stack[-1][0]] != char:
                raise PowerBIValidationError(
                    "M expression has unbalanced delimiters.",
                    details={"position": index, "character": char},
                )
            stack.pop()
        index += 1
    if in_string:
        raise PowerBIValidationError("M expression contains an unterminated string literal.")
    if stack:
        opener, position = stack[-1]
        raise PowerBIValidationError(
            "M expression has unbalanced delimiters.",
            details={"position": position, "character": opener},
        )
    if re.match(r"^\s*let\b", text, flags=re.IGNORECASE) and re.search(r"\bin\b", text, flags=re.IGNORECASE) is None:
        raise PowerBIValidationError("M expression starts with 'let' but has no matching 'in' clause.")

    # Security check: block external/network M functions
    if os.environ.get("PBI_MCP_ALLOW_EXTERNAL_M", "0") != "1":
        for pattern in _M_BLOCKED_FUNCTIONS:
            match = pattern.search(text)
            if match:
                raise PowerBIValidationError(
                    f"M expression contains blocked function '{match.group()}'. "
                    f"Network/external database access is disabled by default. "
                    f"Set PBI_MCP_ALLOW_EXTERNAL_M=1 to allow.",
                    details={"blocked_function": match.group(), "position": match.start()},
                )


def _source_type_token(partition: Any) -> str:
    source = getattr(partition, "Source", None)
    source_name = type(source).__name__ if source is not None else ""
    raw = str(getattr(partition, "SourceType") or "") if hasattr(partition, "SourceType") else ""
    token = normalize_token(source_name or raw)
    if token in {"m", "none", "query", "entity", "calculated", "calculationgroup"}:
        return {
            "m": "m",
            "none": "none",
            "query": "query",
            "entity": "entity",
            "calculated": "calculated",
            "calculationgroup": "calculation_group",
        }[token]
    if token.endswith("mpartitionsource"):
        return "m"
    if token.endswith("calculatedpartitionsource"):
        return "calculated"
    if token.endswith("querypartitionsource"):
        return "query"
    if token.endswith("entitypartitionsource"):
        return "entity"
    if token.endswith("policyrangepartitionsource") or token == "policyrange":
        return "policy_range"
    return "unknown"


def _partition_expression(partition: Any) -> str:
    source = getattr(partition, "Source", None)
    token = _source_type_token(partition)
    if source is not None:
        if token in {"m", "calculated", "policy_range"} and hasattr(source, "Expression"):
            return str(source.Expression or "")
        if token == "query" and hasattr(source, "Query"):
            return str(source.Query or "")
        if hasattr(source, "Expression"):
            return str(source.Expression or "")
    if hasattr(partition, "Expression"):
        return str(getattr(partition, "Expression") or "")
    return ""


def _partition_payload(table: Any, partition: Any) -> dict[str, Any]:
    expression = _partition_expression(partition)
    return {
        "table": str(table.Name),
        "partition": str(partition.Name),
        "source_type": _source_type_token(partition),
        "source_type_raw": serialize_value(getattr(partition, "SourceType", None)),
        "m_expression": redact_sensitive_data(expression),
        "expression_length": len(expression),
    }


def _get_target_partition(model: Any, table_name: str, partition_name: str | None = None) -> tuple[Any, Any]:
    table = find_named(model.Tables, table_name)
    if table is None:
        raise PowerBINotFoundError(f"Table '{table_name}' was not found.", details={"table": table_name})
    count = int(table.Partitions.Count)
    if count == 0:
        raise PowerBINotFoundError(f"Table '{table_name}' has no partitions.", details={"table": table_name})
    if partition_name:
        partition = find_named(table.Partitions, partition_name)
        if partition is None:
            raise PowerBINotFoundError(
                f"Partition '{partition_name}' was not found on table '{table_name}'.",
                details={"table": table_name, "partition": partition_name},
            )
        return table, partition
    if count > 1:
        raise PowerBIValidationError(
            f"Table '{table_name}' has multiple partitions. Specify partition_name explicitly.",
            details={"table": table_name, "partitions": [str(item.Name) for item in table.Partitions]},
        )
    return table, table.Partitions[0]


def _ensure_m_supported(database: Any) -> None:
    compatibility = getattr(database, "CompatibilityLevel", None)
    if compatibility is not None and int(compatibility) < 1400:
        raise PowerBIValidationError(
            "Power Query partition injection requires compatibility level 1400 or higher.",
            details={"compatibility_level": compatibility},
        )


def _set_partition_m_expression(manager: Any, database: Any, partition: Any, expression: str) -> str:
    _ensure_m_supported(database)
    source_type = _source_type_token(partition)
    if source_type == "calculated":
        raise PowerBIValidationError(
            f"Partition '{partition.Name}' is calculated and cannot be overwritten with an M expression.",
            details={"partition": str(partition.Name), "source_type": source_type},
        )
    source = getattr(partition, "Source", None)
    if source_type != "m" or source is None or not hasattr(source, "Expression"):
        if not hasattr(manager.tom, "MPartitionSource"):
            raise PowerBIConfigurationError("This TOM build does not expose MPartitionSource.")
        source = manager.tom.MPartitionSource()
        partition.Source = source
    source.Expression = expression
    return _source_type_token(partition)


def _request_refresh(manager: Any, table: Any, refresh_after: bool) -> None:
    if refresh_after and hasattr(manager.tom, "RefreshType"):
        table.RequestRefresh(manager.tom.RefreshType.Full)


def _load_excel_sheet_names(excel_path: str) -> list[str]:
    try:
        from openpyxl import load_workbook
    except ImportError as exc:  # pragma: no cover - dependency guard
        raise PowerBIConfigurationError("openpyxl is required for Excel import query helpers.") from exc
    path = inspect_excel_archive(excel_path)
    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        return list(workbook.sheetnames)
    finally:
        close = getattr(workbook, "close", None)
        if callable(close):
            close()


def _ensure_file(path_value: str, *, kind: str, allowed_extensions: set[str] | None = None) -> str:
    path = resolve_local_path(path_value, must_exist=True, allowed_extensions=allowed_extensions)
    if not path.is_file():
        raise PowerBINotFoundError(f"{kind} '{path}' was not found.", details={"path": str(path)})
    return str(path)


def _ensure_folder(path_value: str) -> str:
    path = resolve_local_path(path_value, must_exist=True)
    if not path.is_dir():
        raise PowerBINotFoundError(f"Folder '{path}' was not found.", details={"path": str(path)})
    return str(path)


def _build_excel_m(excel_path: str, sheet_name: str, promote_headers: bool = True) -> str:
    final_step = "Promoted" if promote_headers else "Sheet"
    steps = [
        f"    Source = Excel.Workbook(File.Contents({_m_string(excel_path)}), null, true)",
        f"    Sheet = Source{{[Item={_m_string(sheet_name)},Kind=\"Sheet\"]}}[Data]",
    ]
    if promote_headers:
        steps.append("    Promoted = Table.PromoteHeaders(Sheet, [PromoteAllScalars=true])")
    return "let\n" + ",\n".join(steps) + f"\nin\n    {final_step}"


def _build_csv_m(
    csv_path: str,
    *,
    delimiter: str = ",",
    encoding: int = 65001,
    quote_style: str = "csv",
    promote_headers: bool = True,
) -> str:
    token = normalize_token(quote_style)
    if token not in {"csv", "none"}:
        raise PowerBIValidationError(
            "quote_style must be 'csv' or 'none'.",
            details={"quote_style": quote_style},
        )
    final_step = "Promoted" if promote_headers else "Source"
    steps = [
        "    Source = Csv.Document(",
        f"        File.Contents({_m_string(csv_path)}),",
        f"        [Delimiter={_m_string(delimiter)}, Encoding={encoding}, QuoteStyle=QuoteStyle.{token.title()}]",
        "    )",
    ]
    if promote_headers:
        steps.append("    Promoted = Table.PromoteHeaders(Source, [PromoteAllScalars=true])")
    return "let\n" + ",\n".join(steps) + f"\nin\n    {final_step}"


def _build_folder_m(
    folder_path: str,
    *,
    extension_filter: str | None = None,
    include_hidden_files: bool = False,
) -> str:
    final_step = "FilteredExtension" if extension_filter else "VisibleFiles"
    extension = extension_filter if not extension_filter or extension_filter.startswith(".") else "." + extension_filter
    steps = [f"    Source = Folder.Files({_m_string(folder_path)})"]
    if include_hidden_files:
        steps.append("    VisibleFiles = Source")
    else:
        steps.append("    VisibleFiles = Table.SelectRows(Source, each [Attributes]?[Hidden]? <> true)")
    if extension:
        steps.append(
            f"    FilteredExtension = Table.SelectRows(VisibleFiles, each Text.Lower([Extension]) = {_m_string(extension.lower())})"
        )
    return "let\n" + ",\n".join(steps) + f"\nin\n    {final_step}"


def _build_auto_sheet_map(model: Any, sheet_names: list[str]) -> dict[str, str]:
    mapping: dict[str, str] = {}
    for sheet_name in sheet_names:
        table = find_named(model.Tables, sheet_name)
        if table is None or bool(getattr(table, "IsHidden", False)):
            continue
        mapping[sheet_name] = str(table.Name)
    return mapping


def pbi_get_power_query_tool(manager: Any, *, table: str, partition_name: str | None = None) -> dict[str, Any]:
    """Read the M expression for a specific table partition."""
    validate_model_object_name(table)
    if partition_name:
        validate_model_object_name(partition_name)

    def _reader(state: Any) -> dict[str, Any]:
        tbl, partition = _get_target_partition(state.database.Model, table, partition_name)
        return {"query": _partition_payload(tbl, partition), "connection": state.snapshot()}

    payload = manager.run_read("get_power_query", _reader)
    return ok(
        f"Power Query expression retrieved for table '{table}'.",
        query=payload["query"],
        connection=payload["connection"],
    )


def pbi_list_power_queries_tool(manager: Any, *, include_hidden: bool = False) -> dict[str, Any]:
    """List table partitions and their current source expressions."""

    def _reader(state: Any) -> dict[str, Any]:
        queries = []
        for table in state.database.Model.Tables:
            is_hidden = bool(getattr(table, "IsHidden", False))
            if is_hidden and not include_hidden:
                continue
            partitions = [_partition_payload(table, partition) for partition in table.Partitions]
            queries.append(
                {
                    "table": str(table.Name),
                    "is_hidden": is_hidden,
                    "partition_count": len(partitions),
                    "partitions": partitions,
                }
            )
        queries.sort(key=lambda item: item["table"].casefold())
        return {"queries": queries, "connection": state.snapshot()}

    payload = manager.run_read("list_power_queries", _reader)
    return ok(
        "Power Query expressions listed successfully.",
        queries=payload["queries"],
        connection=payload["connection"],
    )


def pbi_set_power_query_tool(
    manager: Any,
    *,
    table: str,
    m_expression: str,
    partition_name: str | None = None,
    refresh_after: bool = False,
) -> dict[str, Any]:
    """Write or update an M expression on a table partition."""
    validate_model_object_name(table)
    if partition_name:
        validate_model_object_name(partition_name)
    _validate_m_expression(m_expression)

    def _mutator(state: Any, database: Any, model: Any) -> dict[str, Any]:
        tbl, partition = _get_target_partition(model, table, partition_name)
        previous = _partition_payload(tbl, partition)
        new_source_type = _set_partition_m_expression(manager, database, partition, m_expression)
        _request_refresh(manager, tbl, refresh_after)
        return {
            "query": {
                "table": str(tbl.Name),
                "partition": str(partition.Name),
                "previous_source_type": previous["source_type"],
                "source_type": new_source_type,
                "previous_expression_length": previous["expression_length"],
                "expression_length": len(m_expression),
                "m_expression": redact_sensitive_data(m_expression),
                "refresh_requested": refresh_after,
            }
        }

    payload = manager.execute_write("set_power_query", _mutator)
    return ok(
        f"Power Query expression updated for table '{table}'.",
        query=payload["query"],
        save_result=payload["save_result"],
        connection=payload["connection"],
    )


def pbi_create_import_query_tool(
    manager: Any,
    *,
    table: str,
    excel_path: str,
    sheet_name: str,
    partition_name: str | None = None,
    promote_headers: bool = True,
    refresh_after: bool = True,
) -> dict[str, Any]:
    """Generate and inject an Excel import Power Query for a table."""
    validate_model_object_name(table)
    if partition_name:
        validate_model_object_name(partition_name)
    workbook_path = str(inspect_excel_archive(excel_path))
    available_sheets = _load_excel_sheet_names(workbook_path)
    if sheet_name not in available_sheets:
        raise PowerBINotFoundError(
            f"Sheet '{sheet_name}' was not found in workbook '{workbook_path}'.",
            details={"path": workbook_path, "sheet": sheet_name, "available_sheets": available_sheets},
        )
    m_expression = _build_excel_m(workbook_path, sheet_name, promote_headers)
    response = pbi_set_power_query_tool(
        manager,
        table=table,
        m_expression=m_expression,
        partition_name=partition_name,
        refresh_after=refresh_after,
    )
    if response.get("ok"):
        response["message"] = f"Excel import query created for table '{table}' from sheet '{sheet_name}'."
    return response


def pbi_create_csv_import_query_tool(
    manager: Any,
    *,
    table: str,
    csv_path: str,
    partition_name: str | None = None,
    delimiter: str = ",",
    encoding: int = 65001,
    quote_style: str = "csv",
    promote_headers: bool = True,
    refresh_after: bool = True,
) -> dict[str, Any]:
    """Generate and inject a CSV import Power Query for a table."""
    validate_model_object_name(table)
    if partition_name:
        validate_model_object_name(partition_name)
    csv_file = _ensure_file(csv_path, kind="CSV file", allowed_extensions={".csv", ".txt"})
    m_expression = _build_csv_m(
        csv_file,
        delimiter=delimiter,
        encoding=encoding,
        quote_style=quote_style,
        promote_headers=promote_headers,
    )
    response = pbi_set_power_query_tool(
        manager,
        table=table,
        m_expression=m_expression,
        partition_name=partition_name,
        refresh_after=refresh_after,
    )
    if response.get("ok"):
        response["message"] = f"CSV import query created for table '{table}' from '{csv_file}'."
    return response


def pbi_create_folder_import_query_tool(
    manager: Any,
    *,
    table: str,
    folder_path: str,
    partition_name: str | None = None,
    extension_filter: str | None = None,
    include_hidden_files: bool = False,
    refresh_after: bool = True,
) -> dict[str, Any]:
    """Generate and inject a folder import Power Query for a table."""
    validate_model_object_name(table)
    if partition_name:
        validate_model_object_name(partition_name)
    folder = _ensure_folder(folder_path)
    m_expression = _build_folder_m(
        folder,
        extension_filter=extension_filter,
        include_hidden_files=include_hidden_files,
    )
    response = pbi_set_power_query_tool(
        manager,
        table=table,
        m_expression=m_expression,
        partition_name=partition_name,
        refresh_after=refresh_after,
    )
    if response.get("ok"):
        response["message"] = f"Folder import query created for table '{table}' from '{folder}'."
    return response


def pbi_bulk_import_excel_tool(
    manager: Any,
    *,
    excel_path: str,
    sheet_table_map: dict[str, str] | None = None,
    promote_headers: bool = True,
    refresh_after: bool = True,
) -> dict[str, Any]:
    """Bulk-create Excel import queries for multiple tables."""
    workbook_path = str(inspect_excel_archive(excel_path))
    available_sheets = _load_excel_sheet_names(workbook_path)

    def _mutator(state: Any, database: Any, model: Any) -> dict[str, Any]:
        mapping = dict(sheet_table_map or _build_auto_sheet_map(model, available_sheets))
        available_sheet_set = set(available_sheets)
        results = []
        created = 0
        for sheet_name, table_name in mapping.items():
            table = find_named(model.Tables, table_name)
            if sheet_name not in available_sheet_set:
                results.append({"table": table_name, "sheet": sheet_name, "status": "skipped", "reason": "sheet_not_found"})
                continue
            if table is None:
                results.append({"table": table_name, "sheet": sheet_name, "status": "skipped", "reason": "table_not_found"})
                continue
            if bool(getattr(table, "IsHidden", False)):
                results.append({"table": table_name, "sheet": sheet_name, "status": "skipped", "reason": "table_hidden"})
                continue
            if int(table.Partitions.Count) == 0:
                results.append({"table": table_name, "sheet": sheet_name, "status": "skipped", "reason": "no_partitions"})
                continue
            if int(table.Partitions.Count) > 1:
                results.append(
                    {
                        "table": table_name,
                        "sheet": sheet_name,
                        "status": "skipped",
                        "reason": "multiple_partitions",
                        "partitions": [str(item.Name) for item in table.Partitions],
                    }
                )
                continue
            partition = table.Partitions[0]
            try:
                m_expression = _build_excel_m(workbook_path, sheet_name, promote_headers)
                _validate_m_expression(m_expression)
                new_source_type = _set_partition_m_expression(
                    manager,
                    database,
                    partition,
                    m_expression,
                )
                _request_refresh(manager, table, refresh_after)
            except Exception as exc:
                results.append(
                    {
                        "table": table_name,
                        "sheet": sheet_name,
                        "status": "skipped",
                        "reason": getattr(exc, "message", str(exc)),
                        "error_code": getattr(exc, "code", "internal_error"),
                    }
                )
                continue
            created += 1
            results.append(
                {
                    "table": str(table.Name),
                    "sheet": sheet_name,
                    "partition": str(partition.Name),
                    "status": "created",
                    "source_type": new_source_type,
                }
            )
        return {
            "excel_path": workbook_path,
            "sheet_table_map": mapping,
            "results": results,
            "created": created,
            "skipped": len(results) - created,
            "refresh_requested": refresh_after,
        }

    payload = manager.execute_write("bulk_import_excel", _mutator)
    return ok(
        f"Bulk import done: {payload['created']} queries created, {payload['skipped']} skipped.",
        excel_path=payload["excel_path"],
        sheet_table_map=payload["sheet_table_map"],
        results=payload["results"],
        created=payload["created"],
        skipped=payload["skipped"],
        refresh_requested=payload["refresh_requested"],
        save_result=payload["save_result"],
        connection=payload["connection"],
    )


__all__ = [
    "_build_csv_m",
    "_build_excel_m",
    "_build_folder_m",
    "_validate_m_expression",
    "pbi_bulk_import_excel_tool",
    "pbi_create_csv_import_query_tool",
    "pbi_create_folder_import_query_tool",
    "pbi_create_import_query_tool",
    "pbi_get_power_query_tool",
    "pbi_list_power_queries_tool",
    "pbi_set_power_query_tool",
]

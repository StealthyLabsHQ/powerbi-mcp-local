"""MCP wrappers — domain: power_query."""

from __future__ import annotations

from typing import Any

from mcp_core import CONNECTION_MANAGER, _run, mcp
from tools import (
    pbi_bulk_import_excel_tool,
    pbi_create_csv_import_query_tool,
    pbi_create_folder_import_query_tool,
    pbi_create_import_query_tool,
    pbi_excel_import_workflow_tool,
    pbi_get_power_query_tool,
    pbi_import_excel_workbook_tool,
    pbi_list_power_queries_tool,
    pbi_parameterize_data_source_tool,
    pbi_relocate_data_source_tool,
    pbi_set_power_query_tool,
)


@mcp.tool()
def pbi_get_power_query(table: str, partition_name: str | None = None) -> dict[str, Any]:
    """Read the Power Query (M) expression for a table."""
    return _run(
        "pbi_get_power_query",
        pbi_get_power_query_tool,
        CONNECTION_MANAGER,
        table=table,
        partition_name=partition_name,
    )


@mcp.tool()
def pbi_list_power_queries(include_hidden: bool = False) -> dict[str, Any]:
    """List table partitions with their current source expressions."""
    return _run(
        "pbi_list_power_queries",
        pbi_list_power_queries_tool,
        CONNECTION_MANAGER,
        include_hidden=include_hidden,
    )


@mcp.tool()
def pbi_set_power_query(
    table: str,
    m_expression: str,
    partition_name: str | None = None,
    refresh_after: bool = False,
) -> dict[str, Any]:
    """Write or update the Power Query (M) expression for a table."""
    return _run(
        "pbi_set_power_query",
        pbi_set_power_query_tool,
        CONNECTION_MANAGER,
        table=table,
        m_expression=m_expression,
        partition_name=partition_name,
        refresh_after=refresh_after,
    )


@mcp.tool()
def pbi_parameterize_data_source(
    parameter_name: str,
    file_path: str,
    partitions: list[str] | None = None,
    dry_run: bool = False,
    refresh_after: bool = False,
) -> dict[str, Any]:
    """Make a workbook path portable via a Power Query parameter.

    Creates an M parameter (default value = file_path) and rewrites every
    matching M partition to call File.Contents(<parameter_name>) instead of the
    hardcoded path. Collaborators can then change the path in one place via
    Power BI Desktop's "Manage parameters" UI.
    """
    return _run(
        "pbi_parameterize_data_source",
        pbi_parameterize_data_source_tool,
        CONNECTION_MANAGER,
        parameter_name=parameter_name,
        file_path=file_path,
        partitions=partitions,
        dry_run=dry_run,
        refresh_after=refresh_after,
    )


@mcp.tool()
def pbi_relocate_data_source(
    old_path: str,
    new_path: str,
    case_sensitive: bool = False,
    dry_run: bool = False,
    refresh_after: bool = False,
) -> dict[str, Any]:
    """Bulk-rewrite a hardcoded file/folder path inside every M partition.

    Use this when an Excel/CSV/folder data source moves and queries fail with
    DataSource.NotFound. Pass `dry_run=True` first to preview the matches.
    """
    return _run(
        "pbi_relocate_data_source",
        pbi_relocate_data_source_tool,
        CONNECTION_MANAGER,
        old_path=old_path,
        new_path=new_path,
        case_sensitive=case_sensitive,
        dry_run=dry_run,
        refresh_after=refresh_after,
    )


@mcp.tool()
def pbi_create_import_query(
    table: str,
    excel_path: str,
    sheet_name: str,
    partition_name: str | None = None,
    promote_headers: bool = True,
    refresh_after: bool = True,
) -> dict[str, Any]:
    """Generate and inject an Excel import Power Query for a table."""
    return _run(
        "pbi_create_import_query",
        pbi_create_import_query_tool,
        CONNECTION_MANAGER,
        table=table,
        excel_path=excel_path,
        sheet_name=sheet_name,
        partition_name=partition_name,
        promote_headers=promote_headers,
        refresh_after=refresh_after,
    )


@mcp.tool()
def pbi_create_csv_import_query(
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
    return _run(
        "pbi_create_csv_import_query",
        pbi_create_csv_import_query_tool,
        CONNECTION_MANAGER,
        table=table,
        csv_path=csv_path,
        partition_name=partition_name,
        delimiter=delimiter,
        encoding=encoding,
        quote_style=quote_style,
        promote_headers=promote_headers,
        refresh_after=refresh_after,
    )


@mcp.tool()
def pbi_create_folder_import_query(
    table: str,
    folder_path: str,
    partition_name: str | None = None,
    extension_filter: str | None = None,
    include_hidden_files: bool = False,
    refresh_after: bool = True,
) -> dict[str, Any]:
    """Generate and inject a folder import Power Query for a table."""
    return _run(
        "pbi_create_folder_import_query",
        pbi_create_folder_import_query_tool,
        CONNECTION_MANAGER,
        table=table,
        folder_path=folder_path,
        partition_name=partition_name,
        extension_filter=extension_filter,
        include_hidden_files=include_hidden_files,
        refresh_after=refresh_after,
    )


@mcp.tool()
def pbi_bulk_import_excel(
    excel_path: str,
    sheet_table_map: dict[str, str] | None = None,
    promote_headers: bool = True,
    refresh_after: bool = True,
) -> dict[str, Any]:
    """Bulk-create Excel import queries for multiple tables at once."""
    return _run(
        "pbi_bulk_import_excel",
        pbi_bulk_import_excel_tool,
        CONNECTION_MANAGER,
        excel_path=excel_path,
        sheet_table_map=sheet_table_map,
        promote_headers=promote_headers,
        refresh_after=refresh_after,
    )


@mcp.tool()
def pbi_import_excel_workbook(
    excel_path: str,
    sheet_table_map: dict[str, str] | None = None,
    promote_headers: bool = True,
    refresh_after: bool = True,
) -> dict[str, Any]:
    """Import an Excel workbook into Power BI tables in one call."""
    return _run(
        "pbi_import_excel_workbook",
        pbi_import_excel_workbook_tool,
        CONNECTION_MANAGER,
        excel_path=excel_path,
        sheet_table_map=sheet_table_map,
        promote_headers=promote_headers,
        refresh_after=refresh_after,
    )


@mcp.tool()
def pbi_excel_import_workflow(
    excel_path: str,
    sheet_table_map: dict[str, str] | None = None,
    promote_headers: bool = True,
    refresh_after: bool = True,
    apply: bool = False,
) -> dict[str, Any]:
    """Plan or run Excel workbook import into Power BI tables."""
    return _run(
        "pbi_excel_import_workflow",
        pbi_excel_import_workflow_tool,
        CONNECTION_MANAGER,
        excel_path=excel_path,
        sheet_table_map=sheet_table_map,
        promote_headers=promote_headers,
        refresh_after=refresh_after,
        apply=apply,
    )

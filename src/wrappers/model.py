"""MCP wrappers — domain: model."""

from __future__ import annotations

from typing import Any

from mcp_core import CONNECTION_MANAGER, _run, mcp
from tools import (
    pbi_create_column_tool,
    pbi_create_table_tool,
    pbi_delete_column_tool,
    pbi_delete_table_tool,
    pbi_list_tables_tool,
    pbi_model_info_tool,
    pbi_rename_column_tool,
    pbi_rename_table_tool,
    pbi_set_column_data_type_tool,
    pbi_validate_model_tool,
)


@mcp.tool()
def pbi_list_tables(
    include_hidden: bool = False,
    include_row_counts: bool = False,
) -> dict[str, Any]:
    """List tables and columns in the active Power BI model."""
    return _run(
        "pbi_list_tables",
        pbi_list_tables_tool,
        CONNECTION_MANAGER,
        include_hidden=include_hidden,
        include_row_counts=include_row_counts,
    )


@mcp.tool()
def pbi_delete_table(name: str) -> dict[str, Any]:
    """Delete a table from the model."""
    return _run("pbi_delete_table", pbi_delete_table_tool, CONNECTION_MANAGER, name=name)


@mcp.tool()
def pbi_delete_column(table: str, name: str) -> dict[str, Any]:
    """Delete a column from a table."""
    return _run(
        "pbi_delete_column",
        pbi_delete_column_tool,
        CONNECTION_MANAGER,
        table=table,
        name=name,
    )


@mcp.tool()
def pbi_rename_table(name: str, new_name: str) -> dict[str, Any]:
    """Rename a table. Dependent DAX expressions must be updated separately."""
    return _run(
        "pbi_rename_table",
        pbi_rename_table_tool,
        CONNECTION_MANAGER,
        name=name,
        new_name=new_name,
    )


@mcp.tool()
def pbi_rename_column(table: str, name: str, new_name: str) -> dict[str, Any]:
    """Rename a column. Dependent DAX expressions must be updated separately."""
    return _run(
        "pbi_rename_column",
        pbi_rename_column_tool,
        CONNECTION_MANAGER,
        table=table,
        name=name,
        new_name=new_name,
    )


@mcp.tool()
def pbi_model_info(
    include_hidden: bool = False,
    include_row_counts: bool = False,
) -> dict[str, Any]:
    """Return a full model snapshot."""
    return _run(
        "pbi_model_info",
        pbi_model_info_tool,
        CONNECTION_MANAGER,
        include_hidden=include_hidden,
        include_row_counts=include_row_counts,
    )


@mcp.tool()
def pbi_create_table(
    name: str,
    expression: str,
    is_hidden: bool = False,
    overwrite: bool = False,
    refresh_after_create: bool = True,
) -> dict[str, Any]:
    """Create or update a calculated table."""
    return _run(
        "pbi_create_table",
        pbi_create_table_tool,
        CONNECTION_MANAGER,
        name=name,
        expression=expression,
        is_hidden=is_hidden,
        overwrite=overwrite,
        refresh_after_create=refresh_after_create,
    )


@mcp.tool()
def pbi_create_column(
    table: str,
    name: str,
    expression: str,
    data_type: str | None = None,
    format_string: str = "",
    display_folder: str = "",
    is_hidden: bool = False,
    overwrite: bool = False,
) -> dict[str, Any]:
    """Create or update a calculated column."""
    return _run(
        "pbi_create_column",
        pbi_create_column_tool,
        CONNECTION_MANAGER,
        table=table,
        name=name,
        expression=expression,
        data_type=data_type,
        format_string=format_string,
        display_folder=display_folder,
        is_hidden=is_hidden,
        overwrite=overwrite,
    )


@mcp.tool()
def pbi_set_column_data_type(
    table: str,
    column: str,
    data_type: str,
    format_string: str | None = None,
) -> dict[str, Any]:
    """Set an existing column's DataType (and optionally FormatString).

    Use when Power Query type hints (Int64.Type, etc.) are overridden by PBI's
    inference and a column ends up as the wrong type. data_type accepts standard
    TOM names: Int64, Decimal, Double, String, DateTime, Boolean, Currency.
    """
    return _run(
        "pbi_set_column_data_type",
        pbi_set_column_data_type_tool,
        CONNECTION_MANAGER,
        table=table,
        column=column,
        data_type=data_type,
        format_string=format_string,
    )


@mcp.tool()
def pbi_validate_model(include_warnings: bool = True) -> dict[str, Any]:
    """Audit the model for issues: empty expressions, missing format strings, orphan tables, duplicate measure names."""
    return _run(
        "pbi_validate_model",
        pbi_validate_model_tool,
        CONNECTION_MANAGER,
        include_warnings=include_warnings,
    )

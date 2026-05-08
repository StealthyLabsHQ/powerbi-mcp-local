"""MCP wrappers — domain: tmdl."""

from __future__ import annotations

from typing import Any

from mcp_core import _run, mcp
from tools import (
    pbi_list_tmdl_files_tool,
    pbi_patch_tmdl_measure_tool,
    pbi_read_tmdl_file_tool,
    pbi_write_tmdl_file_tool,
)


@mcp.tool()
def pbi_list_tmdl_files(project_path: str) -> dict[str, Any]:
    """List TMDL files in a Power BI Project semantic model definition folder."""
    return _run("pbi_list_tmdl_files", pbi_list_tmdl_files_tool, project_path=project_path)


@mcp.tool()
def pbi_read_tmdl_file(project_path: str, relative_file: str) -> dict[str, Any]:
    """Read one TMDL file from a Power BI Project definition folder."""
    return _run(
        "pbi_read_tmdl_file",
        pbi_read_tmdl_file_tool,
        project_path=project_path,
        relative_file=relative_file,
    )


@mcp.tool()
def pbi_write_tmdl_file(
    project_path: str,
    relative_file: str,
    content: str,
    create: bool = False,
) -> dict[str, Any]:
    """Create or overwrite one TMDL file inside a Power BI Project definition folder."""
    return _run(
        "pbi_write_tmdl_file",
        pbi_write_tmdl_file_tool,
        project_path=project_path,
        relative_file=relative_file,
        content=content,
        create=create,
    )


@mcp.tool()
def pbi_patch_tmdl_measure(
    project_path: str,
    table_file: str,
    measure_name: str,
    expression: str,
    format_string: str = "",
    display_folder: str = "",
    overwrite: bool = True,
) -> dict[str, Any]:
    """Create or replace a measure block in one table TMDL file."""
    return _run(
        "pbi_patch_tmdl_measure",
        pbi_patch_tmdl_measure_tool,
        project_path=project_path,
        table_file=table_file,
        measure_name=measure_name,
        expression=expression,
        format_string=format_string,
        display_folder=display_folder,
        overwrite=overwrite,
    )

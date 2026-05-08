"""MCP wrappers — domain: tmdl."""

from __future__ import annotations

from tools import (
    pbi_list_tmdl_files_tool,
    pbi_patch_tmdl_measure_tool,
    pbi_read_tmdl_file_tool,
    pbi_write_tmdl_file_tool,
)

from ._helpers import register_tool

register_tool(pbi_list_tmdl_files_tool)
register_tool(pbi_read_tmdl_file_tool)
register_tool(pbi_write_tmdl_file_tool)
register_tool(pbi_patch_tmdl_measure_tool)

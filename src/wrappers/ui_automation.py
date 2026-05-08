"""MCP wrappers — domain: UI automation (Windows-only, opt-in)."""

from __future__ import annotations

from tools import pbi_persist_now_tool

from ._helpers import register_tool

register_tool(pbi_persist_now_tool)

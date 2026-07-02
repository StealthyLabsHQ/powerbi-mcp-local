"""MCP @mcp.tool() wrapper registration — single source of truth.

Every public ``*_tool`` callable exported by :mod:`tools` (its ``__all__``)
is registered as an MCP tool through :func:`wrappers._helpers.register_tool`,
which generates a pass-through wrapper (manager injection + audit logging +
security policy + response-size cap via ``mcp_core._run``).

Historically each domain had a hand-maintained ``wrappers/<domain>.py``
listing imports + ``register_tool`` calls — the same tool name repeated in
four places. ``register_all()`` derives the registry from ``tools.__all__``
alone: adding a tool now means implementing it and exporting it from
``tools/__init__.py``, nothing else.
"""

from __future__ import annotations

from mcp_core import mcp

from ._helpers import register_tool


def register_all() -> int:
    """Register every ``*_tool`` in ``tools.__all__`` as an MCP tool.

    Idempotent: names already present in the FastMCP registry are skipped,
    so importing ``server`` twice (tests) cannot double-register.
    Returns the number of tools registered by this call.
    """
    import tools

    manager = getattr(mcp, "_tool_manager", None)
    tools_map = getattr(manager, "_tools", {}) if manager is not None else {}
    registered = 0
    for name in tools.__all__:
        if not name.endswith("_tool"):
            continue
        mcp_name = name[: -len("_tool")]
        if mcp_name in tools_map:
            continue
        register_tool(getattr(tools, name))
        registered += 1
    return registered


__all__ = ["register_all", "register_tool"]

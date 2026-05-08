"""Google Antigravity-compatible MCP entry point.

Antigravity's bundled MCP client closes the connection during the
``resources/list`` exchange when talking to FastMCP 1.27.x:

    connection closed: calling "resources/list": client is closing: EOF

The default ``src/server.py`` works fine for Claude Desktop, Cursor, and
the Anthropic CLI. This module is a *separate* entry point that keeps the
main server untouched and applies Antigravity-specific compatibility
tweaks before running stdio:

1. **stdio hygiene** — UTF-8, line-buffered, ``\\n`` line endings, no BOM.
   On Windows, Python's default text mode emits ``\\r\\n`` which breaks
   strict JSON-RPC framing. Stdout is reserved exclusively for JSON-RPC;
   logging goes to stderr.
2. **Quiet logging** — every logger redirected to stderr, FastMCP's own
   logger silenced to ``ERROR``. A stray ``INFO:`` line on stdout would
   poison the next JSON-RPC frame the client tries to read.
3. **Minimal capabilities** — only ``tools`` is advertised. No
   ``prompts``, no ``resources``, no ``experimental`` block, no
   ``listChanged`` notifications. Antigravity's strict validator rejects
   capability shapes it doesn't recognize, which manifests as the EOF
   above.

Tools, resources, and prompts registered on the FastMCP instance are
still available through the underlying ``tools/call`` JSON-RPC method —
they're simply not advertised at the capability level.
"""

from __future__ import annotations

import logging
import os
import sys


def _harden_stdio() -> None:
    """UTF-8, line-buffered, ``\\n`` line endings on stdout/stderr.

    Must run *before* any third-party import that might write to stdout
    (mcp logs, pythonnet warnings, etc.). The MCP client uses stdout as
    a JSON-RPC byte stream — any non-JSON byte makes the next frame
    unparseable.
    """
    os.environ.setdefault("PYTHONIOENCODING", "utf-8")
    os.environ.setdefault("PYTHONUTF8", "1")
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", newline="\n", line_buffering=True)
        except (AttributeError, OSError):
            # On exotic stream types (pytest capture, frozen builds) the
            # reconfigure call may not exist; the env vars above still
            # cover the common case.
            pass


def _silence_loggers() -> None:
    """Route every log record to stderr at ERROR level.

    FastMCP and pythonnet both attach handlers at import time. Without
    this, an ``INFO`` message printed during startup would land on
    stdout (FastMCP's default in some configurations) and corrupt the
    JSON-RPC stream.
    """
    root = logging.getLogger()
    for handler in list(root.handlers):
        root.removeHandler(handler)
    handler = logging.StreamHandler(sys.stderr)
    handler.setFormatter(logging.Formatter("[%(levelname)s] %(name)s: %(message)s"))
    root.addHandler(handler)
    root.setLevel(logging.ERROR)
    # Known noisy library loggers:
    for name in ("FastMCP", "mcp", "uvicorn", "asyncio", "pythonnet", "powerbi_mcp"):
        logging.getLogger(name).setLevel(logging.ERROR)


def _parse_args(argv: list[str] | None = None) -> object:
    """Mirror the relevant ``--readonly`` / ``--profile`` flags from the
    main entry point so the same launcher invocation works either way.

    The Antigravity adapter only uses stdio, so transport flags are
    irrelevant — we accept them silently for argv-passthrough simplicity
    in case the launcher ever forwards them.
    """
    import argparse

    parser = argparse.ArgumentParser(
        prog="powerbi-mcp-antigravity",
        description="Antigravity-compatible stdio MCP server (subset of src/server.py flags).",
    )
    parser.add_argument(
        "--readonly",
        action="store_true",
        help="Disable write and destructive tools for this server process.",
    )
    parser.add_argument(
        "--profile",
        choices=["readonly", "write", "all", "grading"],
        default="all",
        help="Filter exposed tool surface (same semantics as src/server.py).",
    )
    return parser.parse_args(argv)


async def _run_minimal_stdio() -> None:
    """Replicate FastMCP.run_stdio_async with a stripped capabilities payload.

    The default ``mcp.run(transport="stdio")`` calls
    ``create_initialization_options`` which advertises every capability
    derived from registered handlers — including ``prompts``,
    ``resources/listChanged``, and an empty ``experimental`` block.
    Antigravity rejects shapes it doesn't recognize and closes the
    connection. We build an InitializationOptions that exposes only
    ``tools`` and feed it directly to the low-level server loop.
    """
    from mcp.server.models import InitializationOptions
    from mcp.server.stdio import stdio_server
    from mcp.types import ServerCapabilities, ToolsCapability

    # Late import: server.py performs side-effect registration of every
    # @mcp.tool() wrapper, so it must run after stdio + logging are
    # already hardened.
    from mcp_core import _acquire_single_instance_lock, _start_parent_watcher
    from server import mcp

    # Same stdio lifecycle hooks as the default entry point.
    _acquire_single_instance_lock()
    _start_parent_watcher()

    server_version = getattr(mcp, "version", None)
    if not server_version:
        try:
            from importlib.metadata import version as _pkg_version

            server_version = _pkg_version("powerbi-mcp-local")
        except Exception:
            server_version = "0.0.0"

    init_options = InitializationOptions(
        server_name=mcp.name,
        server_version=server_version,
        capabilities=ServerCapabilities(
            tools=ToolsCapability(listChanged=False),
        ),
        instructions=mcp.instructions,
    )
    async with stdio_server() as (read_stream, write_stream):
        await mcp._mcp_server.run(read_stream, write_stream, init_options)


def main() -> None:
    """Entry point — Antigravity-compatible stdio MCP server.

    Equivalent to ``python src/server.py`` for the tool surface, but
    with the stdio/capability tweaks documented at the top of this
    module. Supports ``--readonly`` and ``--profile`` for parity with
    the main entry point.
    """
    _harden_stdio()
    _silence_loggers()

    args = _parse_args()

    # Late import: defer touching SECURITY / mcp_core until after stdio
    # + logging are hardened.
    from mcp_core import _apply_profile
    from security import SECURITY

    SECURITY.policy(reload=True)  # honors operator CWD security_policy.json
    if args.readonly or args.profile == "readonly":
        SECURITY.set_runtime_readonly(True)

    # _apply_profile mutates the registered FastMCP tool surface. It must
    # run BEFORE _run_minimal_stdio so the pruned set is what gets
    # advertised through tools/list.
    from server import mcp  # noqa: F401  -- side-effect registration

    _apply_profile(args.profile)

    import anyio

    anyio.run(_run_minimal_stdio)


if __name__ == "__main__":
    main()

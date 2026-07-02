"""Shared MCP runtime singletons and process-lifecycle helpers.

Extracted from ``server.py`` so wrappers and lifecycle code do not have to
live in the same 3500-line module. ``server.py`` keeps tool registration
and CLI entry; this module owns the FastMCP instance, the connection
manager, and the single-instance + parent-watcher infrastructure.
"""

from __future__ import annotations

import atexit
import os
import signal
import sys
import tempfile
import threading
import time
from pathlib import Path
from typing import Any

from mcp.server.fastmcp import FastMCP

from pbi_connection import PowerBIConnectionManager, error_payload, logger
from security import SECURITY

mcp = FastMCP(
    "powerbi-desktop",
    instructions=(
        "Connects to the local Power BI Desktop Analysis Services instance, "
        "lets clients inspect the semantic model, manage measures and "
        "relationships, run DAX queries, trigger model refreshes, manage "
        "Power Query partitions, and read or write Excel workbooks used in "
        "the Power BI pipeline. It can also extract, modify, and compile "
        "report layouts for page and visual automation."
    ),
    json_response=True,
    log_level="INFO",
)


CONNECTION_MANAGER = PowerBIConnectionManager(logger)


def _run(tool_name: str, callback: Any, *args: Any, **kwargs: Any) -> dict[str, Any]:
    """Execute a tool callback with audit logging and error normalization."""
    safe_kwargs = {
        key: SECURITY.sanitize_for_logging(value)
        for key, value in kwargs.items()
        if key != "manager" and not key.startswith("_")
    }
    logger.info("TOOL_CALL tool=%s params=%s", tool_name, safe_kwargs)
    try:
        policy = SECURITY.validate_tool_call(tool_name, kwargs)
        result = callback(*args, **kwargs)
        status = ("ok" if result.get("ok") else "error") if isinstance(result, dict) else "ok"
        logger.info("TOOL_OK tool=%s status=%s", tool_name, status)
        return _enforce_response_size(tool_name, result, policy)
    except Exception as exc:
        logger.warning("TOOL_FAIL tool=%s error=%s", tool_name, str(exc)[:300])
        logger.exception("Tool '%s' failed", tool_name)
        return error_payload(exc)


def _enforce_response_size(tool_name: str, result: Any, policy: Any) -> Any:
    """Reject responses larger than ``policy.max_response_bytes``.

    The runaway case is ``pbi_export_model`` against a multi-GB model, or
    a DAX query that bypassed ``max_rows_for_dax`` via metadata side-loads.
    A hard cap protects the LLM client process from OOM.
    """
    cap = getattr(policy, "max_response_bytes", 0) or 0
    if cap <= 0 or not isinstance(result, dict):
        return result
    try:
        import json as _json

        size = len(_json.dumps(result, ensure_ascii=False, default=str).encode("utf-8"))
    except Exception:
        return result
    if size <= cap:
        return result
    logger.warning("TOOL_TRUNCATED tool=%s response_bytes=%d cap=%d", tool_name, size, cap)
    return {
        "ok": False,
        "error": {
            "code": "response_too_large",
            "retryable": False,
            "message": (
                f"Tool '{tool_name}' response is {size} bytes, exceeding the "
                f"configured cap of {cap} bytes. Narrow the query or set "
                f"max_response_bytes higher in security_policy.json."
            ),
            "details": {"response_bytes": size, "limit": cap},
        },
    }


# --------------------------------------------------------------------------- #
# Single-instance PID lock + parent watcher (v0.10.3 stability infra)         #
# --------------------------------------------------------------------------- #

_PID_LOCK_PATH = Path(tempfile.gettempdir()) / "powerbi-mcp.pid"


def _release_pid_lock() -> None:
    try:
        if _PID_LOCK_PATH.exists():
            content = _PID_LOCK_PATH.read_text(encoding="utf-8").strip()
            if content == str(os.getpid()):
                _PID_LOCK_PATH.unlink()
    except Exception:
        pass


def _pid_alive(pid: int) -> bool:
    try:
        import psutil

        return psutil.pid_exists(pid) and psutil.Process(pid).is_running()
    except Exception:
        try:
            os.kill(pid, 0)
            return True
        except OSError:
            return False


def _acquire_single_instance_lock() -> None:
    """Force single-instance: kill any prior server holding the PID file, then claim it."""
    try:
        if _PID_LOCK_PATH.exists():
            try:
                old_pid = int(_PID_LOCK_PATH.read_text(encoding="utf-8").strip())
            except (OSError, ValueError):
                old_pid = 0
            if old_pid and old_pid != os.getpid() and _pid_alive(old_pid):
                logger.info("Single-instance: killing prior server PID %d", old_pid)
                try:
                    import psutil

                    psutil.Process(old_pid).kill()
                except Exception as exc:
                    logger.warning("Single-instance: could not kill PID %d: %s", old_pid, exc)
                else:
                    for _ in range(20):
                        if not _pid_alive(old_pid):
                            break
                        time.sleep(0.1)
        _PID_LOCK_PATH.write_text(str(os.getpid()), encoding="utf-8")
    except Exception as exc:
        logger.warning("Single-instance lock acquire failed: %s", exc)
        return

    atexit.register(_release_pid_lock)
    for sig in (signal.SIGINT, signal.SIGTERM):
        try:
            signal.signal(sig, lambda *_: sys.exit(0))
        except (ValueError, OSError):
            pass


def _start_parent_watcher() -> None:
    """Daemon thread: if the LLM parent process disappears, exit so atexit fires."""
    try:
        import psutil
    except ImportError:
        return
    parent_pid = os.getppid() if hasattr(os, "getppid") else None
    if not parent_pid or parent_pid <= 1:
        return

    def _watch() -> None:
        try:
            parent = psutil.Process(parent_pid)
        except Exception:
            return
        while True:
            time.sleep(2.0)
            try:
                if not parent.is_running() or parent.status() == psutil.STATUS_ZOMBIE:
                    logger.info("Parent PID %d exited — shutting down", parent_pid)
                    _release_pid_lock()
                    os._exit(0)
            except psutil.NoSuchProcess:
                logger.info("Parent PID %d gone — shutting down", parent_pid)
                _release_pid_lock()
                os._exit(0)
            except Exception:
                continue

    t = threading.Thread(target=_watch, daemon=True, name="parent-watcher")
    t.start()


# --------------------------------------------------------------------------- #
# Tool registry audit + profile filter                                        #
# --------------------------------------------------------------------------- #


def _audit_tool_registry(strict: bool = False) -> dict[str, Any]:
    """Verify every public ``pbi_*_tool`` from the tools package has an
    @mcp.tool() wrapper registered.

    Returns ``{orphan_implementations, unknown_wrappers, registered_count,
    implementation_count}``. Logs warnings on drift; if ``strict`` is True
    and orphans exist, raises RuntimeError so CI pre-flight fails.
    """
    import tools as _tools

    exported = set(getattr(_tools, "__all__", ()))
    pbi_impls = {name[: -len("_tool")] for name in exported if name.startswith("pbi_") and name.endswith("_tool")}

    manager = getattr(mcp, "_tool_manager", None)
    tools_map = getattr(manager, "_tools", None)
    registered = set(tools_map.keys()) if isinstance(tools_map, dict) else set()
    pbi_registered = {name for name in registered if name.startswith("pbi_")}

    orphan_implementations = sorted(pbi_impls - pbi_registered)
    unknown_wrappers = sorted(pbi_registered - pbi_impls)

    if orphan_implementations:
        logger.warning(
            "tool_registry: %d implementation(s) not wrapped as @mcp.tool(): %s",
            len(orphan_implementations),
            ", ".join(orphan_implementations),
        )
    if unknown_wrappers:
        logger.info(
            "tool_registry: %d wrapper(s) without a matching pbi_*_tool: %s",
            len(unknown_wrappers),
            ", ".join(unknown_wrappers),
        )

    if strict and orphan_implementations:
        raise RuntimeError(
            "tool_registry strict check failed: "
            f"{len(orphan_implementations)} unwrapped tools: {orphan_implementations}"
        )

    return {
        "orphan_implementations": orphan_implementations,
        "unknown_wrappers": unknown_wrappers,
        "registered_count": len(pbi_registered),
        "implementation_count": len(pbi_impls),
    }


def _apply_profile(profile: str) -> None:
    """Prune FastMCP's registered tools based on the selected profile."""
    if profile == "all":
        return
    from security import GRADING_TOOLS, READ_TOOLS, WRITE_TOOLS

    if profile == "readonly":
        allowed = set(READ_TOOLS)
    elif profile == "write":
        allowed = set(READ_TOOLS) | set(WRITE_TOOLS)
    elif profile == "grading":
        allowed = set(GRADING_TOOLS)
    else:
        return

    manager = getattr(mcp, "_tool_manager", None)
    tools_map = getattr(manager, "_tools", None)
    if tools_map is None:
        logger.warning("profile filter skipped: FastMCP tool registry not accessible")
        return
    removed = [name for name in list(tools_map.keys()) if name not in allowed]
    for name in removed:
        tools_map.pop(name, None)
    logger.info("PROFILE %s applied: removed %d tools, exposing %d", profile, len(removed), len(tools_map))


__all__ = [
    "mcp",
    "CONNECTION_MANAGER",
    "logger",
    "_run",
    "_acquire_single_instance_lock",
    "_start_parent_watcher",
    "_release_pid_lock",
    "_audit_tool_registry",
    "_apply_profile",
]

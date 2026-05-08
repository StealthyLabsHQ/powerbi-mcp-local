"""Connection and serialization utilities for the Power BI Desktop MCP server.

This module is intentionally Windows-first. Importing it on macOS/Linux is safe,
but establishing a connection will fail with a clear, JSON-serializable error.
"""

from __future__ import annotations

import collections
import importlib
import json
import logging
import math
import os
import socket
import sys
import threading
import time
from contextlib import contextmanager
from dataclasses import dataclass, field
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path
from typing import Any, Callable, Iterator

try:
    import psutil
except ImportError:  # pragma: no cover - optional until installed on Windows
    psutil = None  # type: ignore[assignment]


LOGGER_NAME = "powerbi_mcp"
DEFAULT_LOG_LEVEL = os.getenv("PBI_MCP_LOG_LEVEL", "INFO").upper()


def configure_logging(level: str = DEFAULT_LOG_LEVEL) -> logging.Logger:
    """Configure and return the shared logger."""
    logger = logging.getLogger(LOGGER_NAME)
    if not logger.handlers:
        handler = logging.StreamHandler(sys.stderr)
        handler.setFormatter(
            logging.Formatter(
                "%(asctime)s %(levelname)s %(name)s - %(message)s",
                datefmt="%Y-%m-%d %H:%M:%S",
            )
        )
        logger.addHandler(handler)
    logger.setLevel(level)
    return logger


logger = configure_logging()


class PowerBIError(Exception):
    """Base class for user-facing Power BI errors."""

    code = "powerbi_error"
    retryable = False

    def __init__(self, message: str, *, details: dict[str, Any] | None = None):
        super().__init__(message)
        self.message = message
        self.details = details or {}


class UnsupportedPlatformError(PowerBIError):
    code = "unsupported_platform"


class PowerBINotRunningError(PowerBIError):
    code = "powerbi_not_running"
    retryable = True


class PowerBIConnectionError(PowerBIError):
    code = "connection_error"
    retryable = True


class PowerBIConfigurationError(PowerBIError):
    code = "configuration_error"


class PowerBIValidationError(PowerBIError):
    code = "validation_error"


class PowerBIDuplicateError(PowerBIError):
    code = "duplicate_object"


class PowerBINotFoundError(PowerBIError):
    code = "not_found"


class PowerBIQueryError(PowerBIError):
    code = "query_error"


class PowerBIWriteError(PowerBIError):
    code = "write_error"


@dataclass
class DiscoveredInstance:
    """A Power BI Desktop-backed Analysis Services instance."""

    port: int
    workspace_path: str | None = None
    discovered_via: set[str] = field(default_factory=set)
    port_file: str | None = None
    modified_time: float | None = None
    pid: int | None = None
    process_name: str | None = None
    process_exe: str | None = None
    process_started_at: float | None = None

    def sort_key(self) -> tuple[float, float, int]:
        return (
            self.modified_time or 0.0,
            self.process_started_at or 0.0,
            self.port,
        )

    def to_dict(self) -> dict[str, Any]:
        return {
            "port": self.port,
            "workspace_path": self.workspace_path,
            "port_file": self.port_file,
            "discovered_via": sorted(self.discovered_via),
            "pid": self.pid,
            "process_name": self.process_name,
            "process_exe": self.process_exe,
            "modified_time": self.modified_time,
            "process_started_at": self.process_started_at,
        }


@dataclass
class ConnectionState:
    """In-memory connection state shared by MCP tools."""

    instance: DiscoveredInstance
    tom_server: Any
    database: Any
    adomd_connection: Any | None
    adomd_backend: str | None
    adomd_available: bool
    dll_directory: str | None
    connected_at: str
    warnings: list[str] = field(default_factory=list)

    def snapshot(self) -> dict[str, Any]:
        return {
            "connected": True,
            "port": self.instance.port,
            "workspace_path": self.instance.workspace_path,
            "database": safe_getattr(self.database, "Name"),
            "dll_directory": self.dll_directory,
            "adomd_backend": self.adomd_backend,
            "adomd_available": self.adomd_available,
            "warnings": list(self.warnings),
            "connected_at": self.connected_at,
            "instance": self.instance.to_dict(),
        }


def ok(message: str, **data: Any) -> dict[str, Any]:
    """Standard JSON response for successful operations."""
    payload = {"ok": True, "message": message}
    payload.update({key: serialize_value(value) for key, value in data.items()})
    return payload


def error_payload(exc: Exception | str, *, code: str | None = None) -> dict[str, Any]:
    """Standard JSON response for failed operations.

    For PowerBIError subclasses, returns the structured ``code``/``message``/
    ``details`` exactly as raised. For other exceptions, the response surfaces:
    - ``message`` — the full chained text via ``flatten_exception_message`` so
      callers see the underlying cause (.NET InnerException, ``__cause__``,
      ``__context__``) instead of only the topmost frame.
    - ``details.cause_chain`` — list of `{type, message}` for each link in the
      chain (top-most first), useful for programmatic error analysis.
    """
    if isinstance(exc, PowerBIError):
        details = serialize_value(exc.details) if exc.details else {}
        if isinstance(details, dict):
            chain = _exception_chain_summary(exc)
            if len(chain) > 1:
                details = {**details, "cause_chain": chain[1:]}
        return {
            "ok": False,
            "error": {
                "code": exc.code,
                "message": exc.message,
                "retryable": exc.retryable,
                "details": details,
            },
        }
    if isinstance(exc, Exception):
        return {
            "ok": False,
            "error": {
                "code": code or "internal_error",
                "message": flatten_exception_message(exc),
                "retryable": False,
                "details": {"cause_chain": _exception_chain_summary(exc)},
            },
        }
    return {
        "ok": False,
        "error": {
            "code": code or "internal_error",
            "message": str(exc),
            "retryable": False,
            "details": {},
        },
    }


def _exception_chain_summary(exc: BaseException) -> list[dict[str, str]]:
    """Walk the exception chain (Python ``__cause__``/``__context__`` and .NET
    ``InnerException``) and return one ``{type, message}`` dict per link."""
    out: list[dict[str, str]] = []
    seen: set[int] = set()
    current: Any = exc
    while current is not None and id(current) not in seen:
        seen.add(id(current))
        message = str(current).strip()
        out.append({
            "type": type(current).__name__,
            "message": message,
        })
        nxt = getattr(current, "InnerException", None)
        if nxt is None:
            nxt = getattr(current, "__cause__", None) or getattr(current, "__context__", None)
        current = nxt
    return out


def serialize_value(value: Any) -> Any:
    """Convert Python and pythonnet values into JSON-serializable structures."""
    if value is None:
        return None
    if isinstance(value, (str, bool, int)):
        return value
    if isinstance(value, float):
        if math.isfinite(value):
            return value
        return None
    if isinstance(value, Decimal):
        integral = value.to_integral_value()
        return int(value) if value == integral else float(value)
    if isinstance(value, (datetime, date)):
        return value.isoformat()
    if isinstance(value, Path):
        return str(value)
    if isinstance(value, dict):
        return {str(key): serialize_value(item) for key, item in value.items()}
    if isinstance(value, (list, tuple, set)):
        return [serialize_value(item) for item in value]

    if type(value).__name__ == "DBNull":
        return None

    if hasattr(value, "ToString"):
        try:
            text = str(value.ToString())
            if text in {"Infinity", "-Infinity", "NaN"}:
                return None
            return text
        except Exception:
            pass

    return str(value)


def safe_getattr(obj: Any, name: str, default: Any = None) -> Any:
    """Read a possibly .NET-backed attribute without raising into callers."""
    try:
        value = getattr(obj, name)
    except Exception:
        return default
    return serialize_value(value)


def flatten_exception_message(exc: BaseException) -> str:
    """Flatten nested Python and .NET exception chains into one readable string."""
    parts: list[str] = []
    seen: set[int] = set()
    current: Any = exc
    while current is not None and id(current) not in seen:
        seen.add(id(current))
        text = str(current).strip()
        if text and text not in parts:
            parts.append(text)

        inner = getattr(current, "InnerException", None)
        if inner is not None:
            current = inner
            continue

        if getattr(current, "__cause__", None) is not None:
            current = current.__cause__
            continue

        if getattr(current, "__context__", None) is not None:
            current = current.__context__
            continue

        current = None

    return " | ".join(parts) or exc.__class__.__name__


def normalize_token(value: str) -> str:
    """Normalize a free-form token into a comparison-friendly slug."""
    return "".join(ch for ch in value.lower() if ch.isalnum())


def map_enum(enum_cls: Any, token: str) -> Any:
    """Map a case-insensitive token to a .NET enum member."""
    wanted = normalize_token(token)
    for name in dir(enum_cls):
        if name.startswith("_"):
            continue
        if normalize_token(name) == wanted:
            return getattr(enum_cls, name)
    raise PowerBIValidationError(
        f"Unsupported value '{token}' for enum {enum_cls.__name__}.",
        details={"enum": enum_cls.__name__, "value": token},
    )


def find_named(collection: Any, name: str) -> Any | None:
    """Find an object by Name in a TOM collection."""
    try:
        item = collection.Find(name)
        if item is not None:
            return item
    except Exception:
        pass

    lowered = name.casefold()
    for item in collection:
        try:
            if str(item.Name).casefold() == lowered:
                return item
        except Exception:
            continue
    return None


def dax_quote_table_name(table_name: str) -> str:
    """Quote a DAX table identifier."""
    return "'" + table_name.replace("'", "''") + "'"


def ensure_windows() -> None:
    """Raise when the module is used on a non-Windows host."""
    if os.name != "nt":
        raise UnsupportedPlatformError(
            "This MCP server is Windows-only because Power BI Desktop and ADOMD/TOM "
            "automation are Windows-only."
        )


class PowerBIConnectionManager:
    """Centralized Power BI Desktop connection manager.

    The skeleton in CLAUDE.md uses two unsynchronized globals. That is unsafe:
    `_connection` and `_server` can point at different ports if Power BI is
    reopened between calls. This manager keeps both under one lock and reconnects
    atomically on port changes.
    """

    def __init__(self, logger_: logging.Logger | None = None) -> None:
        self._logger = logger_ or logger
        self._lock = threading.RLock()
        self._state: ConnectionState | None = None
        self._clr: Any | None = None
        self._tom: Any | None = None
        self._adomd_client: Any | None = None
        self._dll_search_paths: set[str] = set()
        self._dll_directory_handles: list[Any] = []
        self._write_generation: int = 0
        self._read_cache: dict[str, tuple[int, Any]] = {}
        # Ring buffer of recent operations for self-diagnostic by tools/LLMs.
        self._operation_log: collections.deque[dict[str, Any]] = collections.deque(maxlen=50)
        # Short-lived cache for instance discovery (TTL 5s) to avoid rescanning on
        # every tool call that needs a connection.
        self._instance_cache: tuple[float, list[DiscoveredInstance]] | None = None

    def _record_operation(self, *, op: str, kind: str, started: float, ok_: bool, error: BaseException | None = None) -> None:
        """Append an entry to the operation history ring buffer."""
        entry: dict[str, Any] = {
            "ts": datetime.utcnow().isoformat() + "Z",
            "op": op,
            "kind": kind,  # "read" | "write"
            "duration_ms": int((time.perf_counter() - started) * 1000),
            "ok": ok_,
        }
        if not ok_ and error is not None:
            entry["error_message"] = flatten_exception_message(error)
            entry["error_type"] = type(error).__name__
            code = getattr(error, "code", None)
            if code:
                entry["error_code"] = str(code)
        self._operation_log.append(entry)

    def operation_history(self, last_n: int = 20) -> list[dict[str, Any]]:
        """Return the most recent operations (newest first)."""
        n = max(1, min(int(last_n), len(self._operation_log)))
        if n == 0:
            return []
        items = list(self._operation_log)[-n:]
        items.reverse()
        return items

    def list_instances(self) -> list[dict[str, Any]]:
        """Return discovered Power BI Desktop instances."""
        return [item.to_dict() for item in self._discover_instances()]

    def connect(
        self,
        *,
        preferred_port: int | None = None,
        force_reconnect: bool = False,
    ) -> dict[str, Any]:
        """Connect or reconnect to the active Power BI Desktop instance."""
        with self._lock:
            ensure_windows()
            instance = self._select_instance(preferred_port=preferred_port)
            if (
                not force_reconnect
                and self._state is not None
                and self._state.instance.port == instance.port
                and self._is_current_state_usable_locked()
            ):
                snapshot = self._state.snapshot()
                snapshot["instances"] = self.list_instances()
                return snapshot

            self._disconnect_locked()
            self._logger.info("Connecting to Power BI Desktop on port %s", instance.port)
            tom_server, database, dll_directory = self._open_tom_locked(instance)
            adomd_connection, adomd_backend, warnings = self._open_adomd_locked(
                instance.port,
                str(database.Name),
                instance=instance,
            )

            self._state = ConnectionState(
                instance=instance,
                tom_server=tom_server,
                database=database,
                adomd_connection=adomd_connection,
                adomd_backend=adomd_backend,
                adomd_available=adomd_connection is not None,
                dll_directory=dll_directory,
                warnings=warnings,
                connected_at=datetime.utcnow().isoformat() + "Z",
            )
            snapshot = self._state.snapshot()
            snapshot["instances"] = self.list_instances()
            return snapshot

    def disconnect(self) -> dict[str, Any]:
        """Close cached connections."""
        with self._lock:
            self._disconnect_locked()
        return ok("Disconnected from Power BI Desktop.")

    def refresh_metadata(self) -> dict[str, Any]:
        """Reload TOM schema from the server without dropping the connection.

        Returns the new database version (if exposed by TOM) so callers can detect stale caches.
        """
        with self._lock:
            self._ensure_connected_locked()
            assert self._state is not None
            server = self._state.tom_server
            database = self._state.database
            previous_version = serialize_value(getattr(database, "Version", None))
            try:
                server.Refresh(database)
            except TypeError:
                # Some TOM versions expect enum flags — fall back to simple Refresh()
                database.Refresh()
            except Exception as exc:
                raise self._translate_exception(exc, "refresh_metadata") from exc
            current_version = serialize_value(getattr(database, "Version", None))
            changed = previous_version != current_version
            self._write_generation += 1
            self._read_cache.clear()
            return {
                "changed": changed,
                "previous_version": previous_version,
                "current_version": current_version,
                "database": str(database.Name),
            }

    @property
    def tom(self) -> Any:
        """Expose the TOM namespace to tool modules."""
        if self._tom is None:
            raise PowerBIConfigurationError(
                "TOM assemblies are not loaded. Call pbi_connect first."
            )
        return self._tom

    @contextmanager
    def read_state(self) -> Iterator[ConnectionState]:
        """Yield the active connection state under the manager lock."""
        with self._lock:
            self._ensure_connected_locked()
            assert self._state is not None
            yield self._state

    def cached_run_read(
        self,
        cache_key: str,
        operation_name: str,
        reader: Callable[[ConnectionState], Any],
    ) -> Any:
        """Like run_read but caches results until the next write, reconnect, or metadata refresh."""
        with self._lock:
            cached = self._read_cache.get(cache_key)
            if cached is not None and cached[0] == self._write_generation:
                self._logger.debug("Cache hit: %s (gen=%d)", cache_key, self._write_generation)
                return cached[1]
        result = self.run_read(operation_name, reader)
        with self._lock:
            self._read_cache[cache_key] = (self._write_generation, result)
        return result

    def run_read(self, operation_name: str, reader: Callable[[ConnectionState], Any]) -> Any:
        """Run a read operation with one automatic reconnect on connection loss."""
        started = time.perf_counter()
        last_error: PowerBIError | None = None
        for attempt in range(2):
            with self._lock:
                self._ensure_connected_locked(force_reconnect=attempt == 1)
                assert self._state is not None
                try:
                    result = reader(self._state)
                    self._record_operation(op=operation_name, kind="read", started=started, ok_=True)
                    return result
                except Exception as exc:  # pragma: no cover - exercised on Windows
                    translated = self._translate_exception(exc, operation_name)
                    if translated.retryable and attempt == 0:
                        self._logger.warning(
                            "Retrying read operation '%s' after reconnect: %s",
                            operation_name,
                            translated.message,
                        )
                        self._disconnect_locked()
                        last_error = translated
                        continue
                    self._record_operation(op=operation_name, kind="read", started=started, ok_=False, error=translated)
                    raise translated from exc
        if last_error is not None:
            self._record_operation(op=operation_name, kind="read", started=started, ok_=False, error=last_error)
            raise last_error
        raise PowerBIConnectionError(f"Read operation '{operation_name}' failed unexpectedly.")

    def execute_read(
        self,
        operation_name: str | Callable[[ConnectionState, Any, Any], Any],
        reader: Callable[[ConnectionState, Any, Any], Any] | None = None,
    ) -> Any:
        """Compatibility wrapper mirroring execute_write for read operations."""
        if reader is None:
            reader_fn = operation_name
            name = getattr(reader_fn, "__name__", "read")
        else:
            reader_fn = reader
            name = operation_name
        return self.run_read(
            str(name),
            lambda state: reader_fn(state, state.database, state.database.Model),
        )

    def execute_write(
        self,
        operation_name: str,
        mutator: Callable[[ConnectionState, Any, Any], dict[str, Any]],
    ) -> dict[str, Any]:
        """Run a TOM write operation, save changes, and reset state on failure."""
        started = time.perf_counter()
        with self._lock:
            self._ensure_connected_locked()
            assert self._state is not None
            database = self._state.database
            model = database.Model
            try:
                payload = mutator(self._state, database, model)
                save_result = model.SaveChanges()
                self._write_generation += 1
                self._read_cache.clear()
                payload["save_result"] = serialize_value(save_result)
                payload["connection"] = self._state.snapshot()
                # The TOM SaveChanges call only commits to the in-memory AS engine.
                # The on-disk .pbix is NOT updated until the user runs Ctrl+S in
                # Power BI Desktop (or a future pbi_persist_now tool). Surface this
                # so callers don't assume durability.
                payload["persistence"] = {
                    "scope": "memory_only",
                    "hint": (
                        "Change committed to the AS engine in memory. The .pbix on "
                        "disk is unchanged until Power BI Desktop saves the file. "
                        "Closing Desktop without saving will lose this change."
                    ),
                }
                self._record_operation(op=operation_name, kind="write", started=started, ok_=True)
                return payload
            except Exception as exc:  # pragma: no cover - exercised on Windows
                translated = self._translate_exception(exc, operation_name)
                self._logger.exception("Write operation '%s' failed", operation_name)
                self._record_operation(op=operation_name, kind="write", started=started, ok_=False, error=translated)
                self._disconnect_locked()
                raise translated from exc

    def run_adomd_query(
        self,
        query: str,
        *,
        max_rows: int = 1000,
        timeout_seconds: int | None = None,
    ) -> dict[str, Any]:
        """Execute a DAX or DMV query through ADOMD and return JSON rows."""
        if max_rows < 1:
            raise PowerBIValidationError("max_rows must be >= 1", details={"max_rows": max_rows})
        if timeout_seconds is not None and timeout_seconds < 0:
            raise PowerBIValidationError(
                "timeout_seconds must be >= 0 (0 disables the timeout).",
                details={"timeout_seconds": timeout_seconds},
            )

        def _execute(state: ConnectionState) -> dict[str, Any]:
            if not state.adomd_available or state.adomd_connection is None:
                raise PowerBIConfigurationError(
                    "ADOMD query support is unavailable. Install pythonnet and a supported "
                    "pyadomd backend, then reconnect.",
                    details={"warnings": state.warnings},
                )

            backend = state.adomd_backend or "unknown"
            if backend.startswith("pyadomd"):
                return self._query_with_pyadomd(
                    state.adomd_connection, query, max_rows, timeout_seconds=timeout_seconds
                )
            return self._query_with_pythonnet(
                state.adomd_connection, query, max_rows, timeout_seconds=timeout_seconds
            )

        return self.run_read("execute_dax", _execute)

    def _ensure_connected_locked(self, *, force_reconnect: bool = False) -> None:
        if force_reconnect or self._state is None or not self._is_current_state_usable_locked():
            self.connect(force_reconnect=True)

    def _is_current_state_usable_locked(self) -> bool:
        if self._state is None:
            return False
        cached = self._state.instance
        port = cached.port
        if not self._is_port_open(port):
            return False
        try:
            _ = self._state.tom_server.Connected
        except Exception:
            return False
        # Detect "port reused by a different msmdsrv process": if the running
        # process on this port has a different PID than what we cached, treat
        # the connection as stale so we reopen against the new instance.
        cached_pid = getattr(cached, "pid", None)
        if cached_pid is not None:
            current_pid = self._pid_for_port(port)
            if current_pid is not None and current_pid != cached_pid:
                return False
        return True

    def _pid_for_port(self, port: int) -> int | None:
        """Return the PID of the msmdsrv process currently bound to ``port``.

        Best-effort: returns None if discovery fails. Used to detect port reuse
        across PBI Desktop restarts.
        """
        try:
            for instance in self._discover_process_instances():
                if instance.port == port:
                    return getattr(instance, "pid", None)
        except Exception:
            return None
        return None

    def _discover_instances(self) -> list[DiscoveredInstance]:
        ensure_windows()

        now = time.monotonic()
        if self._instance_cache is not None and now - self._instance_cache[0] < 5.0:
            return list(self._instance_cache[1])

        merged: dict[int, DiscoveredInstance] = {}

        for instance in self._discover_workspace_instances():
            current = merged.get(instance.port)
            if current is None:
                merged[instance.port] = instance
                continue
            current.discovered_via |= instance.discovered_via
            current.workspace_path = current.workspace_path or instance.workspace_path
            current.port_file = current.port_file or instance.port_file
            current.modified_time = max(current.modified_time or 0.0, instance.modified_time or 0.0)

        # Skip the expensive psutil scan when workspace files already located every
        # instance (port_file present → process is definitely running).
        if not merged or not all(i.port_file for i in merged.values()):
            for instance in self._discover_process_instances():
                current = merged.get(instance.port)
                if current is None:
                    merged[instance.port] = instance
                    continue
                current.discovered_via |= instance.discovered_via
                current.pid = current.pid or instance.pid
                current.process_name = current.process_name or instance.process_name
                current.process_exe = current.process_exe or instance.process_exe
                current.process_started_at = max(
                    current.process_started_at or 0.0,
                    instance.process_started_at or 0.0,
                )
                current.workspace_path = current.workspace_path or instance.workspace_path

        instances = sorted(merged.values(), key=lambda item: item.sort_key(), reverse=True)
        if not instances:
            raise PowerBINotRunningError(
                "No running Power BI Desktop Analysis Services instance was found.",
                details={"workspace_roots": [str(path) for path in self._workspace_roots()]},
            )
        self._instance_cache = (now, instances)
        return instances

    def _select_instance(self, *, preferred_port: int | None = None) -> DiscoveredInstance:
        instances = self._discover_instances()
        if preferred_port is None:
            return instances[0]
        for instance in instances:
            if instance.port == preferred_port:
                return instance
        raise PowerBINotFoundError(
            f"No Power BI Desktop instance is listening on port {preferred_port}.",
            details={"available_ports": [item.port for item in instances]},
        )

    def _workspace_roots(self) -> list[Path]:
        localappdata = os.getenv("LOCALAPPDATA")
        userprofile = os.getenv("USERPROFILE")
        candidates: list[Path] = []

        if localappdata:
            candidates.append(
                Path(localappdata)
                / "Microsoft"
                / "Power BI Desktop"
                / "AnalysisServicesWorkspaces"
            )
            packages_root = Path(localappdata) / "Packages"
            if packages_root.exists():
                for package_dir in packages_root.glob("Microsoft.MicrosoftPowerBIDesktop*"):
                    candidates.append(
                        package_dir
                        / "LocalCache"
                        / "Local"
                        / "Microsoft"
                        / "Power BI Desktop"
                        / "AnalysisServicesWorkspaces"
                    )

        if userprofile:
            candidates.append(
                Path(userprofile)
                / "Microsoft"
                / "Power BI Desktop Store App"
                / "AnalysisServicesWorkspaces"
            )

        extra_roots = os.getenv("PBI_WORKSPACE_ROOTS")
        if extra_roots:
            for item in extra_roots.split(os.pathsep):
                if item.strip():
                    candidates.append(Path(item.strip()))

        deduped: list[Path] = []
        seen: set[str] = set()
        for candidate in candidates:
            key = str(candidate)
            if key not in seen:
                deduped.append(candidate)
                seen.add(key)
        return deduped

    @staticmethod
    def _bounded_glob(root: Path, name: str, max_depth: int) -> Iterator[Path]:
        """Walk *root* up to *max_depth* levels deep, yielding files named *name*.

        Replaces rglob to avoid scanning the entire Packages hierarchy (100+ dirs).
        Power BI port files live at depth ≤4 from the Packages root.
        """
        stack: list[tuple[Path, int]] = [(root, 0)]
        while stack:
            directory, depth = stack.pop()
            try:
                for entry in directory.iterdir():
                    if entry.is_file() and entry.name == name:
                        yield entry
                    elif entry.is_dir() and depth < max_depth:
                        stack.append((entry, depth + 1))
            except OSError:
                continue

    def _discover_workspace_instances(self) -> list[DiscoveredInstance]:
        instances: list[DiscoveredInstance] = []
        for root in self._workspace_roots():
            if not root.exists():
                continue
            for port_file in self._bounded_glob(root, "msmdsrv.port.txt", max_depth=5):
                try:
                    port = int(port_file.read_text(encoding="utf-8").strip())
                except (OSError, ValueError):
                    continue

                workspace_path = port_file.parent.parent if port_file.parent.name.lower() == "data" else port_file.parent
                stat = port_file.stat()
                instances.append(
                    DiscoveredInstance(
                        port=port,
                        workspace_path=str(workspace_path),
                        discovered_via={"workspace"},
                        port_file=str(port_file),
                        modified_time=stat.st_mtime,
                    )
                )
        return instances

    def _discover_process_instances(self) -> list[DiscoveredInstance]:
        if psutil is None:
            return []

        instances: list[DiscoveredInstance] = []
        for proc in psutil.process_iter(["pid", "name", "create_time", "exe"]):
            info = proc.info
            name = str(info.get("name") or "").lower()
            exe = str(info.get("exe") or "")
            exe_lower = exe.lower()
            if name != "msmdsrv.exe":
                continue
            if "analysisservicesworkspaces" not in exe_lower and "power bi desktop" not in exe_lower:
                continue

            try:
                connections = proc.net_connections(kind="tcp")
            except Exception:
                continue

            for conn in connections:
                if conn.status != getattr(psutil, "CONN_LISTEN", "LISTEN"):
                    continue
                port = getattr(conn.laddr, "port", None)
                if port is None:
                    continue
                workspace_path = None
                if "analysisservicesworkspaces" in exe_lower:
                    workspace_path = str(Path(exe).parent.parent)
                instances.append(
                    DiscoveredInstance(
                        port=int(port),
                        workspace_path=workspace_path,
                        discovered_via={"process"},
                        pid=info.get("pid"),
                        process_name=info.get("name"),
                        process_exe=exe or None,
                        process_started_at=info.get("create_time"),
                    )
                )
        return instances

    def _open_tom_locked(self, instance: DiscoveredInstance) -> tuple[Any, Any, str | None]:
        dll_directory = self._load_analysis_services_assemblies_locked(instance)
        try:
            server = self.tom.Server()
            server.Connect(f"localhost:{instance.port}")
        except Exception as exc:  # pragma: no cover - exercised on Windows
            raise PowerBIConnectionError(
                f"Unable to connect TOM server to localhost:{instance.port}",
                details={"reason": flatten_exception_message(exc), "port": instance.port},
            ) from exc

        database = self._select_database(server)
        return server, database, dll_directory

    def _open_adomd_locked(
        self,
        port: int,
        database_name: str,
        *,
        instance: DiscoveredInstance,
    ) -> tuple[Any | None, str | None, list[str]]:
        warnings: list[str] = []
        conn_str = (
            "Provider=MSOLAP;"
            f"Data Source=localhost:{port};"
            f"Initial Catalog={database_name};"
            "Integrated Security=SSPI;"
        )

        for module_name in ("pyadomd", "pbi_pyadomd"):
            try:
                module = importlib.import_module(module_name)
                conn = module.Pyadomd(conn_str)
                conn.open()
                return conn, module_name, warnings
            except Exception as exc:
                warnings.append(f"{module_name} unavailable: {flatten_exception_message(exc)}")

        try:
            if self._adomd_client is None:
                from Microsoft.AnalysisServices import AdomdClient  # type: ignore

                self._adomd_client = AdomdClient
            conn = self._adomd_client.AdomdConnection(conn_str)
            conn.Open()
            return conn, "pythonnet", warnings
        except Exception as exc:  # pragma: no cover - exercised on Windows
            warnings.append(
                "pythonnet ADOMD unavailable: " + flatten_exception_message(exc)
            )
            self._logger.warning(
                "ADOMD query backend unavailable for port %s: %s",
                instance.port,
                warnings[-1],
            )
            return None, None, warnings

    def _load_analysis_services_assemblies_locked(
        self, instance: DiscoveredInstance
    ) -> str | None:
        if self._tom is not None and self._clr is not None:
            return next(iter(self._dll_search_paths), None)

        try:
            import clr  # type: ignore
        except Exception as exc:  # pragma: no cover - exercised on Windows
            raise PowerBIConfigurationError(
                "pythonnet is required for TOM-based model operations.",
                details={"reason": flatten_exception_message(exc)},
            ) from exc

        self._clr = clr
        candidate_dirs = self._candidate_dll_directories(instance)
        for directory in candidate_dirs:
            self._add_dll_search_path(directory)

        # Search for each DLL independently across all candidate dirs
        tabular_names = [
            "Microsoft.AnalysisServices.Tabular",
            "Microsoft.AnalysisServices.Server.Tabular",
        ]
        adomd_names = [
            "Microsoft.AnalysisServices.AdomdClient",
            "Microsoft.PowerBI.AdomdClient",
        ]

        tabular_loaded = False
        adomd_loaded = False
        loaded_dir: str | None = None

        for directory in candidate_dirs:
            if not tabular_loaded:
                for name in tabular_names:
                    if self._try_add_reference(name, directory):
                        tabular_loaded = True
                        loaded_dir = directory
                        break
            if not adomd_loaded:
                for name in adomd_names:
                    if self._try_add_reference(name, directory):
                        adomd_loaded = True
                        if loaded_dir is None:
                            loaded_dir = directory
                        break
            if tabular_loaded and adomd_loaded:
                break

        if not tabular_loaded or not adomd_loaded:
            raise PowerBIConfigurationError(
                "Could not load Microsoft.AnalysisServices assemblies. Set PBI_DESKTOP_BIN "
                "to the Power BI Desktop bin directory if auto-discovery fails.",
                details={"searched_directories": candidate_dirs},
            )

        try:
            from Microsoft.AnalysisServices import Tabular  # type: ignore
        except ImportError:
            try:
                import Microsoft.AnalysisServices.Tabular as Tabular  # type: ignore
            except ImportError:
                import Microsoft.AnalysisServices as _as_mod  # type: ignore
                Tabular = getattr(_as_mod, "Tabular", _as_mod)

        try:
            from Microsoft.AnalysisServices import AdomdClient  # type: ignore
        except ImportError:
            import Microsoft.AnalysisServices.AdomdClient as AdomdClient  # type: ignore

        self._tom = Tabular
        self._adomd_client = AdomdClient
        return loaded_dir

    def _candidate_dll_directories(self, instance: DiscoveredInstance) -> list[str]:
        dirs: list[str] = []

        for env_name in ("PBI_DESKTOP_BIN", "PBI_DLL_DIR"):
            value = os.getenv(env_name)
            if value:
                candidate = Path(value).expanduser()
                if candidate.is_absolute():
                    dirs.append(str(candidate))
                else:
                    self._logger.warning(
                        "Ignoring %s because it is not an absolute path: %s",
                        env_name,
                        value,
                    )

        if instance.process_exe:
            dirs.append(str(Path(instance.process_exe).parent))

        if psutil is not None:
            process_names = {"pbidesktop.exe", "powerbi.exe", "pbidesktoprs.exe"}
            for proc in psutil.process_iter(["name", "exe"]):
                name = str(proc.info.get("name") or "").lower()
                exe = str(proc.info.get("exe") or "")
                if name in process_names and exe:
                    dirs.append(str(Path(exe).parent))

        # PATH lookup (covers user installs exposing pbidesktop.exe via PATH)
        import shutil
        for exe_name in ("PBIDesktop.exe", "pbidesktoprs.exe"):
            located = shutil.which(exe_name)
            if located:
                dirs.append(str(Path(located).parent))

        # Windows registry lookup (HKLM + HKCU, App Paths and PBI install key)
        dirs.extend(self._registry_pbi_dirs())

        program_files = os.getenv("ProgramFiles", r"C:\Program Files")
        program_files_x86 = os.getenv("ProgramFiles(x86)", r"C:\Program Files (x86)")
        local_app = os.getenv("LOCALAPPDATA", "")
        dirs.extend(
            [
                os.path.join(program_files, "Microsoft Power BI Desktop", "bin"),
                os.path.join(program_files_x86, "Microsoft Power BI Desktop", "bin"),
                os.path.join(program_files, "Microsoft Power BI Desktop RS", "bin"),
                os.path.join(program_files_x86, "Microsoft Power BI Desktop RS", "bin"),
            ]
        )
        # Microsoft Store install (WindowsApps is ACL-restricted but readable for DLL load)
        store_roots = [
            os.path.join(program_files, "WindowsApps"),
            os.path.join(local_app, "Microsoft", "WindowsApps") if local_app else "",
        ]
        for store_root in store_roots:
            if not store_root or not os.path.isdir(store_root):
                continue
            try:
                for entry in os.listdir(store_root):
                    if entry.startswith("Microsoft.MicrosoftPowerBIDesktop"):
                        dirs.append(os.path.join(store_root, entry, "bin"))
            except (PermissionError, OSError):
                continue
        # User-scope install (winget/Scoop often target LocalAppData\Programs)
        if local_app:
            dirs.append(os.path.join(local_app, "Programs", "Microsoft Power BI Desktop", "bin"))
        # ADOMD.NET client libraries installed separately
        for adomd_ver in ("160", "150", "140"):
            dirs.append(os.path.join(program_files, "Microsoft.NET", "ADOMD.NET", adomd_ver))
            dirs.append(os.path.join(program_files_x86, "Microsoft.NET", "ADOMD.NET", adomd_ver))

        deduped: list[str] = []
        seen: set[str] = set()
        for directory in dirs:
            if not directory:
                continue
            resolved = str(Path(directory).expanduser().resolve())
            if resolved not in seen and Path(resolved).is_dir():
                deduped.append(resolved)
                seen.add(resolved)
        return deduped

    def _registry_pbi_dirs(self) -> list[str]:
        """Query Windows registry for Power BI Desktop install locations."""
        if os.name != "nt":
            return []
        try:
            import winreg  # type: ignore[import-not-found]
        except ImportError:
            return []

        found: list[str] = []
        probes = [
            (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Microsoft Power BI Desktop", "InstallLocation"),
            (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\Microsoft\Microsoft Power BI Desktop", "InstallLocation"),
            (winreg.HKEY_CURRENT_USER, r"SOFTWARE\Microsoft\Microsoft Power BI Desktop", "InstallLocation"),
            (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\PBIDesktop.exe", None),
            (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\PBIDesktop.exe", None),
        ]
        for hive, subkey, value_name in probes:
            try:
                with winreg.OpenKey(hive, subkey) as key:
                    value, _ = winreg.QueryValueEx(key, value_name) if value_name else winreg.QueryValueEx(key, "")
            except (FileNotFoundError, OSError):
                continue
            if not value:
                continue
            candidate = Path(str(value).strip('"'))
            if candidate.suffix.lower() == ".exe":
                candidate = candidate.parent
            found.append(str(candidate))
            bin_dir = candidate / "bin"
            if bin_dir.is_dir():
                found.append(str(bin_dir))
        return found

    def _add_dll_search_path(self, directory: str) -> None:
        if directory in self._dll_search_paths:
            return
        self._dll_search_paths.add(directory)
        add_dll_directory = getattr(os, "add_dll_directory", None)
        if callable(add_dll_directory):
            try:
                handle = add_dll_directory(directory)
            except OSError:
                handle = None
            if handle is not None:
                self._dll_directory_handles.append(handle)

    def _try_add_reference(self, assembly_name: str, directory: str) -> bool:
        assert self._clr is not None
        dll_path = os.path.join(directory, f"{assembly_name}.dll")
        if os.path.exists(dll_path):
            self._clr.AddReference(dll_path)
            return True
        return False

    def _select_database(self, server: Any) -> Any:
        candidates = []
        for database in server.Databases:
            try:
                if database.Model is None:
                    continue
            except Exception:
                continue
            name = str(database.Name)
            if name.startswith("$"):
                continue
            candidates.append(database)

        if not candidates:
            raise PowerBIConnectionError("No user model database was found on the Power BI instance.")
        return candidates[0]

    def _query_with_pyadomd(
        self,
        connection: Any,
        query: str,
        max_rows: int,
        *,
        timeout_seconds: int | None = None,
    ) -> dict[str, Any]:
        cursor = connection.cursor()
        if timeout_seconds is not None:
            inner_cmd = getattr(cursor, "command", None) or getattr(cursor, "_command", None)
            if inner_cmd is not None and hasattr(inner_cmd, "CommandTimeout"):
                try:
                    inner_cmd.CommandTimeout = int(timeout_seconds)
                except Exception:
                    pass
        try:
            cursor.execute(query)
            columns = []
            if cursor.description:
                for column in cursor.description:
                    columns.append(getattr(column, "name", column[0]))

            rows: list[dict[str, Any]] = []
            truncated = False
            if hasattr(cursor, "fetchone"):
                while True:
                    row = cursor.fetchone()
                    if row is None:
                        break
                    if len(rows) >= max_rows:
                        truncated = True
                        break
                    rows.append(
                        {
                            columns[index]: serialize_value(value)
                            for index, value in enumerate(row)
                        }
                    )
            else:
                raw_rows = cursor.fetchall()
                truncated = len(raw_rows) > max_rows
                for row in raw_rows[:max_rows]:
                    rows.append(
                        {
                            columns[index]: serialize_value(value)
                            for index, value in enumerate(row)
                        }
                    )
            return {
                "columns": columns,
                "rows": rows,
                "row_count": len(rows),
                "truncated": truncated,
            }
        finally:
            try:
                cursor.close()
            except Exception:
                pass

    def _query_with_pythonnet(
        self,
        connection: Any,
        query: str,
        max_rows: int,
        *,
        timeout_seconds: int | None = None,
    ) -> dict[str, Any]:
        assert self._adomd_client is not None
        command = self._adomd_client.AdomdCommand(query, connection)
        if timeout_seconds is not None:
            try:
                command.CommandTimeout = int(timeout_seconds)
            except Exception:
                pass
        reader = None
        try:
            reader = command.ExecuteReader()
            columns = [str(reader.GetName(index)) for index in range(reader.FieldCount)]
            rows: list[dict[str, Any]] = []
            truncated = False
            while reader.Read():
                if len(rows) >= max_rows:
                    truncated = True
                    break
                row = {
                    columns[index]: serialize_value(reader.GetValue(index))
                    for index in range(reader.FieldCount)
                }
                rows.append(row)
            return {
                "columns": columns,
                "rows": rows,
                "row_count": len(rows),
                "truncated": truncated,
            }
        finally:
            if reader is not None:
                try:
                    reader.Close()
                except Exception:
                    pass
            try:
                command.Dispose()
            except Exception:
                pass

    def _translate_exception(self, exc: Exception, operation_name: str) -> PowerBIError:
        if isinstance(exc, PowerBIError):
            return exc

        message = flatten_exception_message(exc)
        lowered = message.casefold()
        details = {"operation": operation_name, "reason": message}

        if any(token in lowered for token in ("syntax", "parse", "token", "dax")) and "error" in lowered:
            return PowerBIValidationError(message, details=details)
        if any(token in lowered for token in ("already exists", "duplicate", "conflict")):
            return PowerBIDuplicateError(message, details=details)
        if any(token in lowered for token in ("not found", "does not exist", "cannot find")):
            return PowerBINotFoundError(message, details=details)
        if any(
            token in lowered
            for token in (
                "transport",
                "connection",
                "server is not running",
                "target machine actively refused",
                "a connection attempt failed",
                "the connection is either not usable or closed",
                "broken pipe",
            )
        ):
            return PowerBIConnectionError(message, details=details)
        if operation_name in {
            "create_measure",
            "delete_measure",
            "create_relationship",
            "refresh",
            "create_table",
            "create_column",
            "set_format",
            "import_dax_file",
        }:
            return PowerBIWriteError(message, details=details)
        return PowerBIQueryError(message, details=details)

    def _disconnect_locked(self) -> None:
        if self._state is None:
            return

        try:
            if self._state.adomd_connection is not None:
                close = getattr(self._state.adomd_connection, "Close", None)
                if callable(close):
                    close()
                else:
                    close = getattr(self._state.adomd_connection, "close", None)
                    if callable(close):
                        close()
        except Exception:
            pass

        try:
            disconnect = getattr(self._state.tom_server, "Disconnect", None)
            if callable(disconnect):
                disconnect()
        except Exception:
            pass

        self._state = None
        self._write_generation += 1
        self._read_cache.clear()
        self._instance_cache = None

    def _is_port_open(self, port: int) -> bool:
        try:
            with socket.create_connection(("127.0.0.1", port), timeout=0.5):
                return True
        except OSError:
            return False

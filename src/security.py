"""Centralized security policy and validation helpers for the MCP server."""

from __future__ import annotations

import json
import logging
import os
import re
import threading
import time
import zipfile
from collections import deque
from dataclasses import dataclass, field
from functools import wraps
from pathlib import Path
from typing import Any, Callable

from pbi_connection import PowerBIValidationError, serialize_value


READ_TOOLS = {
    "pbi_connect",
    "pbi_list_instances",
    "pbi_list_tables",
    "pbi_list_measures",
    "pbi_list_relationships",
    "pbi_list_roles",
    "pbi_list_calc_groups",
    "pbi_model_info",
    "pbi_execute_dax",
    "pbi_execute_dax_as_role",
    "pbi_trace_query",
    "pbi_validate_dax",
    "pbi_measure_dependencies",
    "pbi_refresh_metadata",
    "pbi_get_power_query",
    "pbi_list_power_queries",
    "excel_list_sheets",
    "excel_read_sheet",
    "excel_read_cell",
    "excel_search",
    "excel_workbook_info",
    "excel_to_pbi_check",
    "pbi_list_pages",
    "pbi_get_page",
}

WRITE_TOOLS = {
    "pbi_create_measure",
    "pbi_create_relationship",
    "pbi_update_relationship",
    "pbi_rename_table",
    "pbi_rename_column",
    "pbi_rename_measure",
    "pbi_set_format",
    "pbi_refresh",
    "pbi_import_dax_file",
    "pbi_create_table",
    "pbi_create_column",
    "pbi_export_model",
    "pbi_create_role",
    "pbi_set_role_filter",
    "pbi_add_role_member",
    "pbi_remove_role_member",
    "pbi_create_calc_group",
    "pbi_add_visual",
    "pbi_set_power_query",
    "pbi_create_import_query",
    "pbi_create_csv_import_query",
    "pbi_create_folder_import_query",
    "pbi_bulk_import_excel",
    "excel_write_cell",
    "excel_write_range",
    "excel_create_sheet",
    "excel_format_range",
    "excel_auto_width",
    "excel_create_workbook",
    "pbi_extract_report",
    "pbi_compile_report",
    "pbi_patch_layout",
    "pbi_create_page",
    "pbi_set_page_size",
    "pbi_add_card",
    "pbi_add_bar_chart",
    "pbi_add_line_chart",
    "pbi_add_donut_chart",
    "pbi_add_table_visual",
    "pbi_add_waterfall",
    "pbi_add_slicer",
    "pbi_add_gauge",
    "pbi_add_text_box",
    "pbi_move_visual",
    "pbi_apply_design",
    "pbi_apply_theme",
    "pbi_build_dashboard",
}

DESTRUCTIVE_TOOLS = {
    "pbi_delete_measure",
    "pbi_delete_relationship",
    "pbi_delete_table",
    "pbi_delete_column",
    "pbi_delete_role",
    "pbi_delete_calc_group",
    "excel_delete_sheet",
    "pbi_delete_page",
    "pbi_remove_visual",
}

EXCEL_TOOLS = {
    "excel_list_sheets",
    "excel_read_sheet",
    "excel_read_cell",
    "excel_search",
    "excel_write_cell",
    "excel_write_range",
    "excel_create_sheet",
    "excel_delete_sheet",
    "excel_format_range",
    "excel_auto_width",
    "excel_create_workbook",
    "excel_workbook_info",
    "excel_to_pbi_check",
}

MODEL_NAME_PARAM_KEYS = {
    "name",
    "new_name",
    "table",
    "table_name",
    "from_table",
    "to_table",
    "from_column",
    "to_column",
    "column",
    "column_name",
    "partition_name",
    "relationship_name",
    "role",
    "username",
    "sheet_name",
    "sheet",
}
MEASURE_NAME_PARAM_KEYS = {"name", "new_name"}
PATH_PARAM_KEYS = {
    "path",
    "file_path",
    "excel_path",
    "csv_path",
    "folder_path",
    "pbix_path",
    "extract_folder",
    "output_path",
    "theme_json_path",
}
EXPRESSION_PARAM_KEYS = {"expression", "m_expression"}
QUERY_PARAM_KEYS = {"query"}
DESCRIPTION_PARAM_KEYS = {"description", "display_folder", "format_string", "delimiter", "quote_style"}
JSON_EXPORT_EXTENSIONS = {".json"}
DAX_FILE_EXTENSIONS = {".dax"}
DEFAULT_EXCEL_EXTENSIONS = {".xlsx", ".xlsm"}
PBIX_FILE_EXTENSIONS = {".pbix"}
SENSITIVE_KEYWORDS = {
    "password",
    "pwd",
    "clientsecret",
    "secret",
    "token",
    "accountkey",
    "sharedaccesssignature",
    "connectionstring",
    "connection_string",
}

CONNECTION_STRING_PATTERNS = [
    re.compile(r"(?i)\b(password|pwd|accountkey|sharedaccesssignature|clientsecret|secret|token)\s*=\s*([^;\n]+)"),
    re.compile(r"(?i)\b(user id|uid)\s*=\s*([^;\n]+)"),
    re.compile(r"(?i)([a-z][a-z0-9+.-]*://)([^:/\s]+):([^@/\s]+)@"),
]

QUERY_ONLY_DAX_PATTERNS = [
    re.compile(r"^\s*evaluate\b", re.IGNORECASE),
    re.compile(r"^\s*define\b", re.IGNORECASE),
    re.compile(r"^\s*order\s+by\b", re.IGNORECASE),
    re.compile(r"^\s*start\s+at\b", re.IGNORECASE),
]


class SecurityPolicyError(PowerBIValidationError):
    code = "security_policy_violation"


@dataclass
class SecurityPolicy:
    allow_read: bool = True
    allow_write: bool = True
    allow_destructive: bool = True
    disabled_tools: set[str] = field(default_factory=set)
    enabled_tools: set[str] | None = None
    readonly: bool = False
    max_string_length: int = 8192
    max_name_length: int = 256
    max_expression_length: int = 200000
    max_query_length: int = 100000
    max_path_length: int = 4096
    max_rows_for_dax: int = 5000
    allowed_excel_extensions: set[str] = field(default_factory=lambda: set(DEFAULT_EXCEL_EXTENSIONS))
    max_excel_zip_uncompressed_bytes: int = 100 * 1024 * 1024
    max_excel_zip_members: int = 10000
    max_excel_zip_compression_ratio: float = 250.0
    max_excel_cells_scanned: int = 200000
    warn_after_calls_per_minute: int = 120
    rate_limit_calls_per_minute: int | None = None
    allowed_base_dirs: list[str] = field(default_factory=list)

    @classmethod
    def from_mapping(cls, mapping: dict[str, Any] | None) -> "SecurityPolicy":
        data = dict(mapping or {})
        enabled = data.get("enabled_tools")
        disabled = data.get("disabled_tools", [])
        allowed_extensions = data.get("allowed_file_extensions") or data.get("allowed_excel_extensions")
        allow_categories = {str(item).casefold() for item in data.get("allow_categories", ["read", "write", "destructive"])}
        deny_categories = {str(item).casefold() for item in data.get("deny_categories", [])}
        return cls(
            allow_read="read" in allow_categories and "read" not in deny_categories,
            allow_write="write" in allow_categories and "write" not in deny_categories,
            allow_destructive="destructive" in allow_categories and "destructive" not in deny_categories,
            disabled_tools={str(item) for item in disabled},
            enabled_tools={str(item) for item in enabled} if enabled else None,
            readonly=bool(data.get("readonly", False) or os.getenv("PBI_MCP_READONLY", "0") == "1"),
            max_string_length=int(data.get("max_string_length", 8192)),
            max_name_length=int(data.get("max_name_length", 256)),
            max_expression_length=int(data.get("max_expression_length", 200000)),
            max_query_length=int(data.get("max_query_length", 100000)),
            max_path_length=int(data.get("max_path_length", 4096)),
            max_rows_for_dax=int(data.get("max_dax_rows", 5000)),
            allowed_excel_extensions={str(item).lower() for item in (allowed_extensions or DEFAULT_EXCEL_EXTENSIONS)},
            max_excel_zip_uncompressed_bytes=int(data.get("max_excel_zip_uncompressed_bytes", 100 * 1024 * 1024)),
            max_excel_zip_members=int(data.get("max_excel_zip_members", 10000)),
            max_excel_zip_compression_ratio=float(data.get("max_excel_zip_compression_ratio", 250.0)),
            max_excel_cells_scanned=int(data.get("max_excel_cells_scanned", 200000)),
            warn_after_calls_per_minute=int(data.get("warn_after_calls_per_minute", 120)),
            rate_limit_calls_per_minute=int(data["rate_limit_calls_per_minute"]) if data.get("rate_limit_calls_per_minute") is not None else None,
            allowed_base_dirs=[str(item) for item in data.get("allowed_base_dirs", [])],
        )


def _load_policy_mapping(cwd: Path | None = None) -> dict[str, Any]:
    env_value = os.getenv("PBI_MCP_SECURITY_POLICY", "").strip()
    if env_value:
        if env_value.startswith("{"):
            loaded = json.loads(env_value)
            if not isinstance(loaded, dict):
                raise SecurityPolicyError("PBI_MCP_SECURITY_POLICY must resolve to a JSON object.")
            return loaded
        candidate = Path(env_value).expanduser()
        if candidate.exists():
            return json.loads(candidate.read_text(encoding="utf-8"))
        raise SecurityPolicyError(
            "PBI_MCP_SECURITY_POLICY must be a JSON object or a path to a JSON file.",
            details={"value": env_value},
        )
    policy_path = (cwd or Path.cwd()) / "security_policy.json"
    if policy_path.exists():
        return json.loads(policy_path.read_text(encoding="utf-8"))
    return {}


class SecurityManager:
    def __init__(self, logger_: logging.Logger | None = None) -> None:
        self._logger = logger_ or logging.getLogger("powerbi_mcp.security")
        self._lock = threading.RLock()
        self._policy: SecurityPolicy | None = None
        self._policy_cwd: Path | None = None
        self._runtime_readonly: bool | None = None
        self._allowed_dirs_override: list[Path] = []
        self._calls: deque[float] = deque()

    def policy(self, *, reload: bool = False, cwd: Path | None = None) -> SecurityPolicy:
        with self._lock:
            if reload or self._policy is None or (cwd is not None and cwd != self._policy_cwd):
                self._policy_cwd = cwd or Path.cwd()
                self._policy = SecurityPolicy.from_mapping(_load_policy_mapping(self._policy_cwd))
            if self._runtime_readonly is not None and self._policy is not None:
                self._policy.readonly = self._runtime_readonly
            return self._policy

    def set_runtime_readonly(self, readonly: bool) -> None:
        with self._lock:
            self._runtime_readonly = readonly
            if self._policy is not None:
                self._policy.readonly = readonly

    def configure_allowed_dirs(self, dirs: list[str]) -> None:
        with self._lock:
            self._allowed_dirs_override.clear()
            self._allowed_dirs_override.extend(
                Path(item).expanduser().resolve() for item in dirs if str(item).strip()
            )

    def allowed_base_dirs(self, policy: SecurityPolicy | None = None) -> list[Path]:
        if self._allowed_dirs_override:
            return list(self._allowed_dirs_override)
        chosen = policy or self.policy()
        if chosen.allowed_base_dirs:
            return [Path(item).expanduser().resolve() for item in chosen.allowed_base_dirs]
        env_dirs = os.getenv("PBI_MCP_ALLOWED_DIRS", "").strip()
        if env_dirs:
            return [Path(item.strip()).expanduser().resolve() for item in env_dirs.split(";") if item.strip()]
        return [Path.cwd().resolve()]

    def validate_directory(self, path_value: str, *, must_exist: bool = True) -> Path:
        resolved = resolve_local_path(path_value, must_exist=must_exist, policy=self.policy())
        if must_exist and not resolved.is_dir():
            raise SecurityPolicyError("Path must point to a directory.", details={"path": str(resolved)})
        return resolved

    def sanitize_for_logging(self, value: Any, *, max_chars: int = 200) -> Any:
        value = redact_sensitive_data(serialize_value(value))
        if isinstance(value, str) and len(value) > max_chars:
            return f"{value[:max_chars]}... ({len(value)} chars)"
        if isinstance(value, list):
            return [self.sanitize_for_logging(item, max_chars=max_chars) for item in value[:50]]
        if isinstance(value, dict):
            return {
                key: self.sanitize_for_logging(item, max_chars=max_chars)
                for key, item in list(value.items())[:50]
            }
        return value

    def validate_tool_call(self, tool_name: str, params: dict[str, Any]) -> SecurityPolicy:
        policy = self.policy()
        category = tool_category(tool_name, params)
        if policy.enabled_tools is not None and tool_name not in policy.enabled_tools:
            raise SecurityPolicyError(f"Tool '{tool_name}' is not enabled by the active security policy.")
        if tool_name in policy.disabled_tools:
            raise SecurityPolicyError(f"Tool '{tool_name}' is disabled by the active security policy.")
        if policy.readonly and category in {"write", "destructive"}:
            raise SecurityPolicyError(
                f"Readonly mode blocks tool '{tool_name}'.",
                details={"tool": tool_name, "category": category},
            )
        if category == "read" and not policy.allow_read:
            raise SecurityPolicyError(f"Read tools are disabled by the active security policy.")
        if category == "write" and not policy.allow_write:
            raise SecurityPolicyError(f"Write tools are disabled by the active security policy.")
        if category == "destructive" and not policy.allow_destructive:
            raise SecurityPolicyError(f"Destructive tools are disabled by the active security policy.")
        self._record_call(tool_name, policy)
        _validate_params(tool_name, params, policy)
        return policy

    def _record_call(self, tool_name: str, policy: SecurityPolicy) -> None:
        now = time.time()
        with self._lock:
            while self._calls and now - self._calls[0] > 60:
                self._calls.popleft()
            self._calls.append(now)
            count = len(self._calls)
        if policy.warn_after_calls_per_minute and count > policy.warn_after_calls_per_minute:
            self._logger.warning("SECURITY: high MCP tool call volume detected (%d calls/minute, latest=%s)", count, tool_name)
        if policy.rate_limit_calls_per_minute is not None and count > policy.rate_limit_calls_per_minute:
            raise SecurityPolicyError(
                "MCP tool call rate limit exceeded.",
                details={"count_last_minute": count, "limit": policy.rate_limit_calls_per_minute},
            )


def tool_category(tool_name: str, params: dict[str, Any] | None = None) -> str:
    if tool_name == "pbi_export_model" and params and not params.get("path"):
        return "read"
    if tool_name in DESTRUCTIVE_TOOLS:
        return "destructive"
    if tool_name in WRITE_TOOLS:
        return "write"
    return "read"


def _validate_params(tool_name: str, params: dict[str, Any], policy: SecurityPolicy) -> None:
    for key, value in params.items():
        if key == "manager" or key.startswith("_"):
            continue
        _validate_value(tool_name, key, value, policy)
    if tool_name == "pbi_execute_dax":
        max_rows = params.get("max_rows", 1000)
        if int(max_rows) > policy.max_rows_for_dax:
            raise SecurityPolicyError(
                f"max_rows {max_rows} exceeds the configured limit of {policy.max_rows_for_dax}.",
                details={"max_rows": max_rows, "limit": policy.max_rows_for_dax},
            )


def _validate_value(tool_name: str, key: str, value: Any, policy: SecurityPolicy) -> None:
    if isinstance(value, str):
        _validate_string(tool_name, key, value, policy)
        return
    if isinstance(value, dict):
        if len(value) > 1000:
            raise SecurityPolicyError(f"Parameter '{key}' contains too many entries.", details={"count": len(value)})
        for child_key, child_value in value.items():
            _validate_value(tool_name, f"{key}.{child_key}", str(child_key), policy)
            _validate_value(tool_name, f"{key}.{child_key}", child_value, policy)
        return
    if isinstance(value, (list, tuple, set)):
        if len(value) > 10000:
            raise SecurityPolicyError(f"Parameter '{key}' contains too many items.", details={"count": len(value)})
        for index, item in enumerate(value):
            _validate_value(tool_name, f"{key}[{index}]", item, policy)


def _validate_string(tool_name: str, key: str, value: str, policy: SecurityPolicy) -> None:
    limit = policy.max_string_length
    lowered = key.split(".")[-1]
    if lowered in QUERY_PARAM_KEYS:
        limit = policy.max_query_length
    elif lowered in EXPRESSION_PARAM_KEYS:
        limit = policy.max_expression_length
    elif lowered in MODEL_NAME_PARAM_KEYS or lowered in MEASURE_NAME_PARAM_KEYS:
        limit = policy.max_name_length
    elif lowered in PATH_PARAM_KEYS:
        limit = policy.max_path_length
    if len(value) > limit:
        raise SecurityPolicyError(
            f"Parameter '{key}' exceeds the maximum allowed length.",
            details={"length": len(value), "limit": limit},
        )
    if lowered in PATH_PARAM_KEYS:
        allowed_extensions = None
        if tool_name in EXCEL_TOOLS or lowered in {"file_path", "excel_path"}:
            allowed_extensions = policy.allowed_excel_extensions
        elif tool_name == "pbi_import_dax_file":
            allowed_extensions = DAX_FILE_EXTENSIONS
        elif tool_name == "pbi_export_model":
            allowed_extensions = JSON_EXPORT_EXTENSIONS
        elif lowered in {"pbix_path", "output_path"}:
            allowed_extensions = PBIX_FILE_EXTENSIONS
        elif lowered == "theme_json_path":
            allowed_extensions = JSON_EXPORT_EXTENSIONS
        resolve_local_path(value, must_exist=False, allowed_extensions=allowed_extensions, policy=policy)
    if lowered in MODEL_NAME_PARAM_KEYS:
        validate_model_object_name(value, max_length=policy.max_name_length)
    if lowered in MEASURE_NAME_PARAM_KEYS and tool_name in {"pbi_create_measure"}:
        validate_measure_name(value, max_length=policy.max_name_length)
    if lowered in QUERY_PARAM_KEYS:
        validate_query_text(value, max_length=policy.max_query_length)
    if lowered in EXPRESSION_PARAM_KEYS:
        validate_expression_text(value, max_length=policy.max_expression_length)
    if lowered in DESCRIPTION_PARAM_KEYS and any(ord(ch) < 32 and ch not in "\r\n\t" for ch in value):
        raise SecurityPolicyError(f"Parameter '{key}' contains control characters.")


def validate_model_object_name(name: str, *, max_length: int = 256) -> None:
    value = str(name).strip()
    if not value:
        raise SecurityPolicyError("Object names cannot be empty.")
    if len(value) > max_length:
        raise SecurityPolicyError("Object name exceeds the maximum allowed length.")
    if any(ord(ch) < 32 or ord(ch) == 127 for ch in value):
        raise SecurityPolicyError("Object names cannot contain control characters.")


def validate_measure_name(name: str, *, max_length: int = 256) -> None:
    validate_model_object_name(name, max_length=max_length)
    if any(ch in name for ch in "[]'\""):
        raise SecurityPolicyError(
            "Measure names cannot contain brackets or quotes.",
            details={"name": name},
        )


def validate_connection_string_property_value(value: str, *, field: str) -> None:
    text = str(value).strip()
    if not text:
        raise SecurityPolicyError(f"{field} cannot be empty.")
    if ";" in text:
        raise SecurityPolicyError(
            f"{field} cannot contain ';'.",
            details={"field": field},
        )


def validate_expression_text(expression: str, *, max_length: int = 200000) -> None:
    text = str(expression).strip()
    if not text:
        raise SecurityPolicyError("Expressions cannot be empty.")
    if len(text) > max_length:
        raise SecurityPolicyError("Expression exceeds the maximum allowed length.", details={"length": len(text), "limit": max_length})


def validate_query_text(query: str, *, max_length: int = 100000) -> None:
    text = str(query).strip()
    if not text:
        raise SecurityPolicyError("Query cannot be empty.")
    if len(text) > max_length:
        raise SecurityPolicyError("Query exceeds the maximum allowed length.", details={"length": len(text), "limit": max_length})


def validate_model_expression(expression: str, *, kind: str = "expression", max_length: int = 200000) -> None:
    validate_expression_text(expression, max_length=max_length)
    text = str(expression).strip()
    for pattern in QUERY_ONLY_DAX_PATTERNS:
        if pattern.search(text):
            raise SecurityPolicyError(
                f"{kind} cannot start with query-only DAX syntax.",
                details={"pattern": pattern.pattern},
            )


def resolve_local_path(
    path_value: str,
    *,
    must_exist: bool = False,
    allowed_extensions: set[str] | None = None,
    policy: SecurityPolicy | None = None,
) -> Path:
    chosen = policy or SECURITY.policy()
    text = str(path_value).strip()
    if not text:
        raise SecurityPolicyError("Path cannot be empty.")
    if len(text) > chosen.max_path_length:
        raise SecurityPolicyError("Path exceeds the maximum allowed length.")
    normalized = os.path.normpath(text.replace("\\", os.sep))
    path = Path(normalized).expanduser()
    if not path.is_absolute():
        path = Path.cwd() / path
    _reject_symlink_path(path)
    if path.exists():
        resolved = path.resolve(strict=True)
    else:
        resolved = path.resolve(strict=False)
    if must_exist and not resolved.exists():
        raise SecurityPolicyError("Path does not exist.", details={"path": str(resolved)})
    allowed_bases = SECURITY.allowed_base_dirs(chosen)
    for base in allowed_bases:
        try:
            resolved.relative_to(base)
            break
        except ValueError:
            continue
    else:
        raise SecurityPolicyError(
            f"Path traversal blocked: '{resolved}' is outside allowed directories.",
            details={"path": str(resolved), "allowed": [str(item) for item in allowed_bases]},
        )
    if allowed_extensions is not None and resolved.suffix.lower() not in {item.lower() for item in allowed_extensions}:
        raise SecurityPolicyError(
            "File extension is not allowed by the active security policy.",
            details={"path": str(resolved), "allowed_extensions": sorted(allowed_extensions)},
        )
    return resolved


def _reject_symlink_path(path: Path) -> None:
    if path.exists() and path.is_symlink():
        raise SecurityPolicyError(
            "Symlink paths are blocked by the active security policy.",
            details={"path": str(path), "symlink": str(path)},
        )


def inspect_excel_archive(
    path_value: str | Path,
    *,
    policy: SecurityPolicy | None = None,
    max_uncompressed_bytes: int | None = None,
    max_members: int | None = None,
    max_ratio: float | None = None,
) -> Path:
    chosen = policy or SECURITY.policy()
    resolved = resolve_local_path(
        str(path_value),
        must_exist=True,
        allowed_extensions=chosen.allowed_excel_extensions,
        policy=chosen,
    )
    if not zipfile.is_zipfile(resolved):
        raise SecurityPolicyError("Excel workbook is not a valid Open XML ZIP archive.", details={"path": str(resolved)})
    uncompressed_limit = max_uncompressed_bytes or chosen.max_excel_zip_uncompressed_bytes
    member_limit = max_members or chosen.max_excel_zip_members
    ratio_limit = max_ratio or chosen.max_excel_zip_compression_ratio
    with zipfile.ZipFile(resolved) as archive:
        infos = archive.infolist()
        if len(infos) > member_limit:
            raise SecurityPolicyError(
                "Excel workbook exceeds the maximum number of ZIP members.",
                details={"members": len(infos), "limit": member_limit},
            )
        total_uncompressed = 0
        for info in infos:
            total_uncompressed += int(info.file_size)
            if total_uncompressed > uncompressed_limit:
                raise SecurityPolicyError(
                    "Excel workbook exceeds the maximum decompressed size.",
                    details={"bytes": total_uncompressed, "limit": uncompressed_limit},
                )
            compressed = max(int(info.compress_size), 1)
            if info.file_size and (info.file_size / compressed) > ratio_limit:
                raise SecurityPolicyError(
                    "Excel workbook looks like a ZIP bomb.",
                    details={"member": info.filename, "ratio": round(info.file_size / compressed, 2), "limit": ratio_limit},
                )
    return resolved


def redact_sensitive_text(text: str) -> str:
    redacted = text
    for pattern in CONNECTION_STRING_PATTERNS:
        if "://" in pattern.pattern:
            redacted = pattern.sub(r"\1[REDACTED]:[REDACTED]@", redacted)
        else:
            redacted = pattern.sub(lambda match: f"{match.group(1)}=[REDACTED]", redacted)
    return redacted


def redact_sensitive_data(value: Any) -> Any:
    serialized = serialize_value(value)
    if isinstance(serialized, str):
        return redact_sensitive_text(serialized)
    if isinstance(serialized, list):
        return [redact_sensitive_data(item) for item in serialized]
    if isinstance(serialized, dict):
        redacted: dict[str, Any] = {}
        for key, item in serialized.items():
            if any(token in str(key).casefold().replace(" ", "") for token in SENSITIVE_KEYWORDS):
                redacted[str(key)] = "[REDACTED]"
            else:
                redacted[str(key)] = redact_sensitive_data(item)
        return redacted
    return serialized


def secure_tool(tool_name: str) -> Callable[[Callable[..., Any]], Callable[..., Any]]:
    def decorator(func: Callable[..., Any]) -> Callable[..., Any]:
        @wraps(func)
        def wrapper(*args: Any, **kwargs: Any) -> Any:
            SECURITY.validate_tool_call(tool_name, kwargs)
            return func(*args, **kwargs)

        return wrapper

    return decorator


SECURITY = SecurityManager()
ALLOWED_BASE_DIRS = SECURITY._allowed_dirs_override
configure_allowed_dirs = SECURITY.configure_allowed_dirs


__all__ = [
    "ALLOWED_BASE_DIRS",
    "DESTRUCTIVE_TOOLS",
    "EXCEL_TOOLS",
    "READ_TOOLS",
    "SECURITY",
    "SecurityManager",
    "SecurityPolicy",
    "SecurityPolicyError",
    "WRITE_TOOLS",
    "configure_allowed_dirs",
    "inspect_excel_archive",
    "redact_sensitive_data",
    "redact_sensitive_text",
    "resolve_local_path",
    "secure_tool",
    "tool_category",
    "validate_expression_text",
    "validate_connection_string_property_value",
    "validate_measure_name",
    "validate_model_expression",
    "validate_model_object_name",
    "validate_query_text",
]

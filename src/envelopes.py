"""Response envelopes and pure serialization/lookup helpers.

Extracted from ``pbi_connection`` so tool modules can build ``ok``/error
payloads without importing the connection manager. ``pbi_connection``
re-exports every name here, so ``from pbi_connection import ok`` keeps
working unchanged.

This module may import from ``pbi_errors`` only — never from
``pbi_connection`` — to stay circular-import free.
"""

from __future__ import annotations

import math
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path
from typing import Any

from pbi_errors import PowerBIError, PowerBIValidationError


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
        out.append(
            {
                "type": type(current).__name__,
                "message": message,
            }
        )
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

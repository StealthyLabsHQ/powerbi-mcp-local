"""User-facing Power BI error hierarchy.

Extracted from ``pbi_connection`` so pure modules can raise structured errors
without importing the connection manager. ``pbi_connection`` re-exports every
name here, so ``from pbi_connection import PowerBIValidationError`` keeps
working unchanged.

This module must not import anything from ``pbi_connection`` (or ``envelopes``)
to stay circular-import free.
"""

from __future__ import annotations

from typing import Any


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

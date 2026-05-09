"""Wrapper-registration helpers.

The vast majority of MCP wrappers in :mod:`wrappers` are pure pass-throughs:
they take the underlying ``pbi_*_tool``'s parameters, optionally inject the
shared ``CONNECTION_MANAGER`` for ``manager``, and pipe the result through
:func:`mcp_core._run` for audit logging + error normalisation.

``register_tool`` generates such wrappers from the underlying tool's
signature so wrapper modules don't have to repeat the boilerplate. The
generated wrapper preserves the original tool's parameter schema (via
``__signature__``) so FastMCP's ``@mcp.tool()`` decorator picks up the
right JSON schema for clients.

Manager-injection is auto-detected:

- ``manager`` is the first positional parameter → inject ``CONNECTION_MANAGER``
  positionally; drop it from the wrapper's external signature.
- ``manager`` is a keyword-only parameter with a default of ``None`` (typical
  for visual tools) → inject ``manager=CONNECTION_MANAGER`` as a keyword
  argument; drop it from the external signature.
- No ``manager`` parameter → pure pass-through, no injection.

A handful of wrappers need extra logic (e.g. ``pbi_add_visual`` adds a
``dry_run`` flag, ``pbi_validate_report_fields`` injects manager only if the
caller didn't already pass one). Those keep their hand-written form.
"""

from __future__ import annotations

import functools
import inspect
from collections.abc import Callable
from typing import Any

from mcp_core import CONNECTION_MANAGER, _run, mcp


def register_tool(
    tool_fn: Callable[..., dict[str, Any]],
    *,
    name: str | None = None,
    inject_manager: bool | None = None,
    docstring: str | None = None,
) -> Callable[..., dict[str, Any]]:
    """Generate a pass-through ``@mcp.tool()`` wrapper for ``tool_fn``.

    Parameters
    ----------
    tool_fn:
        The underlying ``pbi_*_tool`` (or ``excel_*_tool``) function.
    name:
        Override the registered MCP tool name. Defaults to ``tool_fn.__name__``
        with a trailing ``_tool`` stripped.
    inject_manager:
        Force-on or force-off the manager injection. ``None`` (default)
        auto-detects from the signature.
    docstring:
        Override the wrapper's docstring. Defaults to the underlying tool's
        ``__doc__``.

    Returns
    -------
    The decorated wrapper (also registered with FastMCP as a side effect).
    """
    tool_name = name or _strip_tool_suffix(tool_fn.__name__)
    sig = inspect.signature(tool_fn)
    params = list(sig.parameters.values())

    inject_kind = _detect_manager_injection(params) if inject_manager is None else inject_manager
    public_params = [p for p in params if p.name != "manager"] if inject_kind else params
    public_sig = sig.replace(parameters=public_params)

    if inject_kind:
        # Always inject ``manager`` as a keyword argument. The previous
        # positional path bound CONNECTION_MANAGER to ``*args[0]`` of the
        # underlying tool, which only worked when ``manager`` was the FIRST
        # positional parameter. Tools like ``pbi_patch_layout_tool`` that
        # declare ``manager`` later in the signature would receive the
        # connection manager as ``extract_folder`` and then collide with the
        # actual ``extract_folder=`` kwarg from the MCP client (``TypeError:
        # multiple values for argument 'extract_folder'``).
        def _wrapper(*args: Any, **kwargs: Any) -> dict[str, Any]:
            kwargs.setdefault("manager", CONNECTION_MANAGER)
            return _run(tool_name, tool_fn, *args, **kwargs)

    else:

        def _wrapper(*args: Any, **kwargs: Any) -> dict[str, Any]:
            return _run(tool_name, tool_fn, *args, **kwargs)

    functools.update_wrapper(_wrapper, tool_fn)
    _wrapper.__signature__ = public_sig
    _wrapper.__name__ = tool_name
    _wrapper.__qualname__ = tool_name
    _wrapper.__doc__ = docstring if docstring is not None else (tool_fn.__doc__ or "")

    mcp.tool()(_wrapper)
    return _wrapper


def _strip_tool_suffix(name: str) -> str:
    return name[: -len("_tool")] if name.endswith("_tool") else name


def _detect_manager_injection(params: list[inspect.Parameter]) -> str | None:
    """Return ``"positional"``, ``"keyword"``, or ``None``."""
    for param in params:
        if param.name != "manager":
            continue
        if param.kind in (
            inspect.Parameter.POSITIONAL_ONLY,
            inspect.Parameter.POSITIONAL_OR_KEYWORD,
        ):
            return "positional"
        if param.kind == inspect.Parameter.KEYWORD_ONLY:
            return "keyword"
    return None

"""Layout I/O — load, atomic write with .bak fallback, dry-run interception,
embedded JSON helpers, and page utilities."""

from __future__ import annotations

import json
import logging
import threading
from collections.abc import Iterator
from contextlib import contextmanager
from pathlib import Path
from typing import Any

from atomic_io import atomic_write_bytes

from ._base import (
    DEFAULT_PAGE_HEIGHT,
    DEFAULT_PAGE_WIDTH,
    PageNotFoundError,
    ReportLayoutError,
)
from ._paths import _layout_path, _resolve_extract_folder

logger = logging.getLogger("tools.visuals._layout")


def _load_layout(extract_folder: str | Path) -> tuple[Path, dict[str, Any]]:
    folder = _resolve_extract_folder(str(extract_folder), must_exist=True)
    if not folder.is_dir():
        raise ReportLayoutError("Extract folder does not exist or is not a directory.", details={"path": str(folder)})
    layout_path = _layout_path(folder)
    if not layout_path.exists():
        raise ReportLayoutError(
            "Report/Layout file was not found in the extract folder.", details={"path": str(layout_path)}
        )
    try:
        layout = json.loads(layout_path.read_text(encoding="utf-16-le"))
    except UnicodeDecodeError as exc:
        raise ReportLayoutError(
            "Report/Layout could not be decoded as UTF-16-LE.", details={"path": str(layout_path)}
        ) from exc
    except json.JSONDecodeError as exc:
        raise ReportLayoutError(
            "Report/Layout is not valid JSON.", details={"path": str(layout_path), "line": exc.lineno}
        ) from exc
    if not isinstance(layout, dict):
        raise ReportLayoutError("Report/Layout root must be a JSON object.", details={"path": str(layout_path)})
    layout.setdefault("sections", [])
    return folder, layout


_LAYOUT_WRITE_TL = threading.local()


def _is_dry_run() -> bool:
    return bool(getattr(_LAYOUT_WRITE_TL, "active", False))


def _record_dry_run_write(folder: Path, layout: dict[str, Any]) -> None:
    log = getattr(_LAYOUT_WRITE_TL, "log", None)
    if log is None:
        return
    sections = layout.get("sections", []) or []
    log.append(
        {
            "folder": str(folder),
            "section_count": len(sections),
            "visual_count": sum(len(s.get("visualContainers", []) or []) for s in sections if isinstance(s, dict)),
        }
    )


@contextmanager
def dry_run_layout_writes() -> Iterator[list[dict[str, Any]]]:
    """While active, ``_save_layout`` records the would-be write instead of
    flushing to disk. Yields the list that captures one entry per intercepted
    write — callers can return it as a "preview" payload.
    """
    _LAYOUT_WRITE_TL.active = True
    _LAYOUT_WRITE_TL.log = []
    try:
        yield _LAYOUT_WRITE_TL.log
    finally:
        _LAYOUT_WRITE_TL.active = False
        _LAYOUT_WRITE_TL.log = None


def _save_layout(extract_folder: Path, layout: dict[str, Any]) -> None:
    """Atomic layout write with .bak fallback.

    1. Serialise to a sibling temp file (``Layout.tmp.<pid>``).
    2. Copy the existing Layout (if any) to ``Layout.bak`` so a previous
       known-good version is recoverable after a crash mid-write.
    3. ``os.replace`` the temp onto Layout — atomic on Windows + POSIX.

    On any exception during steps 1/2, the original Layout is untouched.
    """
    if _is_dry_run():
        _record_dry_run_write(extract_folder, layout)
        return
    layout_path = _layout_path(extract_folder)
    encoded = json.dumps(layout, ensure_ascii=False, indent=2).encode("utf-16-le")
    atomic_write_bytes(layout_path, encoded, backup=True)


def _parse_embedded_json(value: Any, default: Any) -> Any:
    if value in (None, ""):
        return default
    if isinstance(value, (dict, list)):
        return value
    if not isinstance(value, str):
        return default
    try:
        return json.loads(value)
    except json.JSONDecodeError:
        return default


def _dump_embedded_json(value: Any) -> str:
    return json.dumps(value, ensure_ascii=False, separators=(",", ":"))


def _normalize_page_name(display_name: str) -> str:
    cleaned = "".join(char for char in display_name if char.isalnum())
    return cleaned or "Page"


def _next_page_name(layout: dict[str, Any], display_name: str) -> str:
    existing = {str(section.get("name", "")) for section in layout.get("sections", [])}
    base = f"ReportSection{_normalize_page_name(display_name)}"
    if base not in existing:
        return base
    index = 1
    while f"{base}{index}" in existing:
        index += 1
    return f"{base}{index}"


def _find_page(layout: dict[str, Any], page: str) -> dict[str, Any]:
    wanted = page.casefold()
    for section in layout.get("sections", []):
        name = str(section.get("name", ""))
        display_name = str(section.get("displayName", ""))
        if name.casefold() == wanted or display_name.casefold() == wanted:
            return section
    raise PageNotFoundError(
        f"Page '{page}' was not found.",
        details={
            "page": page,
            "available_pages": [
                str(item.get("displayName") or item.get("name")) for item in layout.get("sections", [])
            ],
        },
    )


def _page_summary(section: dict[str, Any]) -> dict[str, Any]:
    visuals = section.get("visualContainers", []) or []
    return {
        "name": str(section.get("name", "")),
        "display_name": str(section.get("displayName", "")),
        "width": int(section.get("width", DEFAULT_PAGE_WIDTH)),
        "height": int(section.get("height", DEFAULT_PAGE_HEIGHT)),
        "visual_count": len(visuals),
    }

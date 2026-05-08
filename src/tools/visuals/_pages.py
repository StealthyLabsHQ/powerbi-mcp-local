"""Page-level tools: list, get, describe, create, delete, set page size."""

from __future__ import annotations

from typing import Any

from pbi_connection import PowerBIValidationError, ok

from ._base import DEFAULT_PAGE_HEIGHT, DEFAULT_PAGE_WIDTH
from ._bindings import _live_model_field_index, _visual_binding_issues
from ._containers import _validate_dimensions, _visual_payload
from ._home_tables import _scan_measure_home_tables
from ._layout import (
    _find_page,
    _load_layout,
    _next_page_name,
    _page_summary,
    _parse_embedded_json,
    _save_layout,
)
from ._paths import _resolve_extract_folder


def _run(callback):
    from pbi_connection import error_payload

    try:
        return callback()
    except Exception as exc:
        return error_payload(exc)


def pbi_list_pages_tool(extract_folder: str) -> dict[str, Any]:
    def _impl() -> dict[str, Any]:
        _, layout = _load_layout(extract_folder)
        pages = [_page_summary(section) for section in layout.get("sections", [])]
        return ok(
            "Pages listed successfully.",
            extract_folder=str(_resolve_extract_folder(extract_folder, must_exist=True)),
            pages=pages,
        )

    return _run(_impl)


def pbi_get_page_tool(extract_folder: str, page: str) -> dict[str, Any]:
    def _impl() -> dict[str, Any]:
        _, layout = _load_layout(extract_folder)
        section = _find_page(layout, page)
        visuals = [_visual_payload(container) for container in section.get("visualContainers", []) or []]
        payload = _page_summary(section)
        payload["visuals"] = visuals
        return ok(
            "Page retrieved successfully.",
            extract_folder=str(_resolve_extract_folder(extract_folder, must_exist=True)),
            page=payload,
        )

    return _run(_impl)


def pbi_describe_page_tool(
    extract_folder: str,
    page: str,
    *,
    manager: Any | None = None,
) -> dict[str, Any]:
    """Return a structured, LLM-friendly snapshot of a report page.

    One entry per visual with id/type/position, role-keyed ``bindings``,
    ``formatting`` (title + axis titles + label_display_units), and a
    ``binding_health`` rollup (``ok`` | ``missing_field`` | ``wrong_role``)
    based on live-model validation when ``manager`` is supplied.
    """

    def _impl() -> dict[str, Any]:
        folder, layout = _load_layout(extract_folder)
        section = _find_page(layout, page)
        measure_home_map = _scan_measure_home_tables(folder)
        model_fields, _ = (
            _live_model_field_index(manager, include_hidden=False) if manager else (None, {"status": "unavailable"})
        )

        visuals: list[dict[str, Any]] = []
        for container in section.get("visualContainers", []) or []:
            cfg = _parse_embedded_json(container.get("config"), {})
            sv = cfg.get("singleVisual", {}) if isinstance(cfg, dict) else {}
            visual_type = str(sv.get("visualType") or "unknown")
            projections = sv.get("projections", {}) or {}
            bindings: dict[str, list[str]] = {}
            for role, items in projections.items():
                refs: list[str] = []
                if isinstance(items, list):
                    for item in items:
                        ref = item.get("queryRef") if isinstance(item, dict) else None
                        if isinstance(ref, str) and ref:
                            refs.append(ref)
                bindings[role] = refs

            objects = sv.get("objects", {}) if isinstance(sv, dict) else {}

            def _extract_literal_text(obj_name: str, prop: str) -> str | None:
                entries = objects.get(obj_name)
                if not isinstance(entries, list) or not entries:
                    return None
                props = entries[0].get("properties", {}) if isinstance(entries[0], dict) else {}
                value = props.get(prop, {})
                literal = value.get("expr", {}).get("Literal", {}).get("Value") if isinstance(value, dict) else None
                if not isinstance(literal, str):
                    return None
                if len(literal) >= 2 and literal[0] == "'" and literal[-1] == "'":
                    return literal[1:-1].replace("''", "'")
                return literal

            formatting: dict[str, Any] = {}
            title_text = _extract_literal_text("title", "text")
            if title_text is not None:
                formatting["title"] = title_text
            x_axis_title = _extract_literal_text("categoryAxis", "titleText") or _extract_literal_text(
                "categoryAxis", "axisTitle"
            )
            if x_axis_title is not None:
                formatting["x_axis_title"] = x_axis_title
            y_axis_title = _extract_literal_text("valueAxis", "titleText") or _extract_literal_text(
                "valueAxis", "axisTitle"
            )
            if y_axis_title is not None:
                formatting["y_axis_title"] = y_axis_title
            labels = objects.get("labels")
            if isinstance(labels, list) and labels:
                lu_value = (
                    labels[0].get("properties", {}).get("labelDisplayUnits", {}) if isinstance(labels[0], dict) else {}
                )
                lu_literal = (
                    lu_value.get("expr", {}).get("Literal", {}).get("Value") if isinstance(lu_value, dict) else None
                )
                if lu_literal is not None:
                    formatting["label_display_units"] = lu_literal

            issues, _ = _visual_binding_issues(
                container, str(section.get("displayName") or section.get("name", "")), measure_home_map, model_fields
            )
            if not issues:
                health = "ok"
            else:
                kinds = {item.get("issue") for item in issues}
                if "live_model_missing" in kinds or "live_model_unknown_field" in kinds:
                    health = "missing_field"
                elif any(k and "role" in k for k in kinds):
                    health = "wrong_role"
                else:
                    health = "issues"

            visuals.append(
                {
                    "id": str(cfg.get("name", "")),
                    "type": visual_type,
                    "position": {
                        "x": int(container.get("x", 0)),
                        "y": int(container.get("y", 0)),
                        "width": int(container.get("width", 0)),
                        "height": int(container.get("height", 0)),
                    },
                    "bindings": bindings,
                    "formatting": formatting,
                    "binding_health": health,
                    "issues": issues,
                }
            )

        return ok(
            f"Page '{section.get('displayName')}' described — {len(visuals)} visual(s).",
            extract_folder=str(folder),
            page={
                "name": str(section.get("name", "")),
                "display_name": str(section.get("displayName", "")),
                "width": int(section.get("width", DEFAULT_PAGE_WIDTH)),
                "height": int(section.get("height", DEFAULT_PAGE_HEIGHT)),
                "visual_count": len(visuals),
            },
            visuals=visuals,
        )

    return _run(_impl)


def pbi_create_page_tool(
    extract_folder: str, display_name: str, width: int = DEFAULT_PAGE_WIDTH, height: int = DEFAULT_PAGE_HEIGHT
) -> dict[str, Any]:
    def _impl() -> dict[str, Any]:
        _validate_dimensions(0, 0, width, height)
        folder, layout = _load_layout(extract_folder)
        section = {
            "name": _next_page_name(layout, display_name),
            "displayName": display_name,
            "displayOption": 0,
            "width": width,
            "height": height,
            "visualContainers": [],
            "filters": "[]",
        }
        if any("ordinal" in item for item in layout.get("sections", [])):
            section["ordinal"] = len(layout.get("sections", []))
        layout.setdefault("sections", []).append(section)
        _save_layout(folder, layout)
        return ok("Page created successfully.", extract_folder=str(folder), page=_page_summary(section))

    return _run(_impl)


def pbi_delete_page_tool(extract_folder: str, page: str) -> dict[str, Any]:
    def _impl() -> dict[str, Any]:
        folder, layout = _load_layout(extract_folder)
        sections = layout.get("sections", [])
        if len(sections) <= 1:
            raise PowerBIValidationError("Cannot delete the last remaining page.")
        section = _find_page(layout, page)
        layout["sections"] = [item for item in sections if item is not section]
        _save_layout(folder, layout)
        return ok(
            "Page deleted successfully.",
            extract_folder=str(folder),
            deleted_page=str(section.get("displayName") or section.get("name")),
        )

    return _run(_impl)


def pbi_set_page_size_tool(extract_folder: str, page: str, width: int, height: int) -> dict[str, Any]:
    def _impl() -> dict[str, Any]:
        _validate_dimensions(0, 0, width, height)
        folder, layout = _load_layout(extract_folder)
        section = _find_page(layout, page)
        section["width"] = width
        section["height"] = height
        _save_layout(folder, layout)
        return ok("Page size updated successfully.", extract_folder=str(folder), page=_page_summary(section))

    return _run(_impl)

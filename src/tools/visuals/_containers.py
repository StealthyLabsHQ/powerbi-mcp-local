"""Visual container builders.

Wraps the per-visual config + binding into the JSON shape Power BI's
``visualContainers`` array expects, plus helpers to find/extract a visual
from a page section and to atomically append a new one.
"""

from __future__ import annotations

import uuid
from collections.abc import Callable
from typing import Any

from pbi_connection import PowerBIValidationError, ok

from ._base import VisualNotFoundError
from ._bindings import (
    _assert_container_bindings,
    _build_prototype_query,
    _validate_projection_roles,
)
from ._formatting import _title_objects
from ._layout import (
    _dump_embedded_json,
    _find_page,
    _load_layout,
    _page_summary,
    _parse_embedded_json,
    _save_layout,
)


def _unique_visual_id() -> str:
    return uuid.uuid4().hex[:20]


def _validate_dimensions(x: int, y: int, width: int, height: int) -> None:
    if min(x, y) < 0:
        raise PowerBIValidationError("x and y must be >= 0.", details={"x": x, "y": y})
    if width <= 0 or height <= 0:
        raise PowerBIValidationError("width and height must be > 0.", details={"width": width, "height": height})


def _page_next_z(section: dict[str, Any]) -> int:
    z_values = [
        int(container.get("z", 0)) for container in section.get("visualContainers", []) if isinstance(container, dict)
    ]
    return (max(z_values) + 1) if z_values else 0


def _base_visual_config(
    *,
    visual_id: str,
    visual_type: str,
    x: int,
    y: int,
    width: int,
    height: int,
    references: list[str] | None = None,
    measure_home_map: dict[str, str] | None = None,
    projections: dict[str, list[dict[str, str]]] | None = None,
    title: str | None = None,
    extra_single_visual: dict[str, Any] | None = None,
) -> tuple[dict[str, Any], dict[str, Any]]:
    position = {"x": x, "y": y, "width": width, "height": height}
    single_visual = {
        "visualType": visual_type,
        "projections": projections or {},
        "prototypeQuery": _build_prototype_query(references or [], measure_home_map),
        "objects": {},
    }
    if title:
        single_visual["objects"].update(_title_objects(title))
    if extra_single_visual:
        extra_objects = extra_single_visual.get("objects")
        if isinstance(extra_objects, dict):
            single_visual["objects"].update(extra_objects)
        for key, val in extra_single_visual.items():
            if key == "objects":
                continue
            single_visual[key] = val
    config = {
        "name": visual_id,
        "layouts": [{"id": 0, "position": position}],
        "singleVisual": single_visual,
    }
    query = {
        "Commands": [
            {
                "SemanticQueryDataShapeCommand": {
                    "Query": single_visual["prototypeQuery"],
                }
            }
        ]
    }
    return config, query


def _make_visual_container(
    *,
    section: dict[str, Any],
    visual_type: str,
    x: int,
    y: int,
    width: int,
    height: int,
    references: list[str] | None = None,
    measure_home_map: dict[str, str] | None = None,
    projections: dict[str, list[dict[str, str]]] | None = None,
    title: str | None = None,
    filters: Any | None = None,
    extra_single_visual: dict[str, Any] | None = None,
) -> dict[str, Any]:
    _validate_dimensions(x, y, width, height)
    visual_id = _unique_visual_id()
    config, query = _base_visual_config(
        visual_id=visual_id,
        visual_type=visual_type,
        x=x,
        y=y,
        width=width,
        height=height,
        references=references,
        measure_home_map=measure_home_map,
        projections=projections,
        title=title,
        extra_single_visual=extra_single_visual,
    )
    return {
        "x": x,
        "y": y,
        "z": _page_next_z(section),
        "width": width,
        "height": height,
        "config": _dump_embedded_json(config),
        "filters": _dump_embedded_json(filters if filters is not None else []),
        "query": _dump_embedded_json(query),
        "dataTransforms": _dump_embedded_json({}),
    }


def _visual_payload(container: dict[str, Any]) -> dict[str, Any]:
    config = _parse_embedded_json(container.get("config"), {})
    single_visual = config.get("singleVisual", {}) if isinstance(config, dict) else {}
    title = None
    text = None
    objects = single_visual.get("objects", {}) if isinstance(single_visual, dict) else {}
    title_entries = objects.get("title", [])
    if title_entries:
        title = title_entries[0].get("properties", {}).get("text", {}).get("expr", {}).get("Literal", {}).get("Value")
    if isinstance(single_visual, dict) and "textContent" in single_visual:
        text = single_visual.get("textContent")
    return {
        "id": str(config.get("name") or ""),
        "type": str(single_visual.get("visualType") or "unknown"),
        "x": int(container.get("x", 0)),
        "y": int(container.get("y", 0)),
        "z": int(container.get("z", 0)),
        "width": int(container.get("width", 0)),
        "height": int(container.get("height", 0)),
        "data": {
            "title": title,
            "text": text,
            "projections": single_visual.get("projections", {}),
        },
    }


def _find_visual(section: dict[str, Any], visual_id: str) -> tuple[int, dict[str, Any], dict[str, Any]]:
    for index, container in enumerate(section.get("visualContainers", []) or []):
        config = _parse_embedded_json(container.get("config"), {})
        if str(config.get("name", "")).casefold() == visual_id.casefold():
            return index, container, config
    raise VisualNotFoundError(
        f"Visual '{visual_id}' was not found on page '{section.get('displayName') or section.get('name')}'.",
        details={"visual_id": visual_id},
    )


def _append_visual(
    extract_folder: str,
    page: str,
    factory: Callable[[dict[str, Any], dict[str, str]], dict[str, Any]],
    measure_home_map: dict[str, str],
) -> dict[str, Any]:
    folder, layout = _load_layout(extract_folder)
    section = _find_page(layout, page)
    section.setdefault("visualContainers", [])
    container = factory(section, measure_home_map)
    _assert_container_bindings(container, measure_home_map)
    section["visualContainers"].append(container)
    _save_layout(folder, layout)
    visual = _visual_payload(container)
    return ok(
        f"Visual '{visual['id']}' added to page '{section.get('displayName')}'.",
        page=_page_summary(section),
        visual=visual,
    )


def _create_chart_container(
    section: dict[str, Any],
    *,
    visual_type: str,
    x: int,
    y: int,
    width: int,
    height: int,
    title: str | None,
    projections: dict[str, list[dict[str, str]]],
    references: list[str],
    measure_home_map: dict[str, str] | None = None,
    extra_single_visual: dict[str, Any] | None = None,
    manager: Any | None = None,
) -> dict[str, Any]:
    # Pre-flight role validation: catches "role not allowed for this visual
    # type" and (with manager) "wrong reference kind in role" before we ever
    # write a layout that PBI Desktop would refuse to render.
    _validate_projection_roles(visual_type, projections, manager=manager)
    return _make_visual_container(
        section=section,
        visual_type=visual_type,
        x=x,
        y=y,
        width=width,
        height=height,
        title=title,
        projections=projections,
        references=references,
        measure_home_map=measure_home_map,
        extra_single_visual=extra_single_visual,
    )

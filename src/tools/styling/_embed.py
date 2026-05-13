"""PBIX zip read / patch / write for the styling layer.

The PBIX format is a ZIP archive with a small fixed part list. For the
styling tool we touch:

- ``Report/Layout``                                  — patched in place
- ``StaticResources/RegisteredResources/<name>.png`` — new entry
- ``StaticResources/SharedResources/BaseThemes/<theme>.json``
                                                     — new entry

The functions in this module are pure: they take bytes / dicts and
return bytes / dicts. The orchestrator in ``_apply.py`` handles the
file-system + path-validation concerns.
"""

from __future__ import annotations

import hashlib
import io
import json
import re
import zipfile
from pathlib import Path
from typing import Any

LAYOUT_PART = "Report/Layout"
WALLPAPER_DIR = "StaticResources/RegisteredResources"
THEMES_DIR = "StaticResources/SharedResources/BaseThemes"


_SANITIZE_RE = re.compile(r"[^A-Za-z0-9._-]+")


def sanitize_resource_name(name: str) -> str:
    """Strip a filename down to PBIX-safe characters."""
    stripped = _SANITIZE_RE.sub("_", name).strip("._-")
    return stripped or "resource"


def sha1_short(data: bytes) -> str:
    """SHA-1 of ``data``, hex, first 12 chars. Used to dedupe wallpapers
    embedded under the same preset across multiple apply runs.
    """
    return hashlib.sha1(data).hexdigest()[:12]


def _decode_layout(raw: bytes) -> dict[str, Any]:
    return json.loads(raw.decode("utf-16-le"))


def _encode_layout(layout: dict[str, Any]) -> bytes:
    return json.dumps(layout, ensure_ascii=False).encode("utf-16-le")


def _decode_section_config(section: dict[str, Any]) -> dict[str, Any]:
    raw = section.get("config")
    if not raw:
        return {}
    if isinstance(raw, dict):
        return raw
    try:
        return json.loads(raw)
    except (json.JSONDecodeError, TypeError):
        return {}


def _encode_section_config(cfg: dict[str, Any]) -> str:
    return json.dumps(cfg, ensure_ascii=False)


def patch_layout_for_wallpaper(
    layout: dict[str, Any],
    *,
    resource_name: str,
    fit: str = "Fit",
    transparency: int = 0,
    page_filter: set[str] | None = None,
    page_background_color: str | None = None,
    page_background_transparency: int = 0,
) -> tuple[dict[str, Any], list[str]]:
    """Inject the canonical ``objects.background`` block into every
    targeted section's config.

    ``resource_name`` is the relative path under
    ``StaticResources/RegisteredResources``. Returns the patched layout
    and the list of page display names that were touched (so the caller
    can report which pages received the wallpaper).
    """
    if fit not in ("Fit", "Fill", "Normal"):
        raise ValueError(f"fit must be Fit / Fill / Normal, got '{fit}'")

    touched: list[str] = []
    for section in layout.get("sections", []) or []:
        if not isinstance(section, dict):
            continue
        display = str(section.get("displayName") or section.get("name") or "")
        if page_filter and display not in page_filter:
            continue
        cfg = _decode_section_config(section)
        objects = cfg.setdefault("objects", {})
        background_entry: dict[str, Any] = {
            "properties": {
                "image": {
                    "name": {"expr": {"Literal": {"Value": f"'{resource_name}'"}}},
                    "url": {"expr": {"Literal": {"Value": f"'{resource_name}'"}}},
                    "scaling": {"expr": {"Literal": {"Value": f"'{fit}'"}}},
                },
                "transparency": {"expr": {"Literal": {"Value": f"{transparency}D"}}},
            }
        }
        objects["background"] = [background_entry]
        if page_background_color:
            objects["outspace"] = [
                {
                    "properties": {
                        "color": {
                            "solid": {
                                "color": {"expr": {"Literal": {"Value": f"'{page_background_color}'"}}}
                            }
                        },
                        "transparency": {
                            "expr": {"Literal": {"Value": f"{page_background_transparency}D"}}
                        },
                    }
                }
            ]
        section["config"] = _encode_section_config(cfg)
        touched.append(display)

    return layout, touched


def _literal(value: Any) -> dict[str, Any]:
    if isinstance(value, bool):
        return {"expr": {"Literal": {"Value": "true" if value else "false"}}}
    if isinstance(value, (int, float)):
        return {"expr": {"Literal": {"Value": f"{value}D"}}}
    return {"expr": {"Literal": {"Value": f"'{value}'"}}}


def _solid_color(hex_color: str) -> dict[str, Any]:
    return {"solid": {"color": _literal(hex_color)}}


def _vc_objects_for_card(
    cards_spec: dict[str, Any], accent_color: str | None
) -> dict[str, Any]:
    """Compose vcObjects (container chrome) for a card-like visual."""
    bg = cards_spec.get("background", {})
    border = cards_spec.get("border", {})
    shadow = cards_spec.get("dropShadow", {})
    vc: dict[str, Any] = {
        "background": [
            {
                "properties": {
                    "show": _literal(True),
                    "color": _solid_color(bg.get("color", "#FFFFFF")),
                    "transparency": _literal(int(bg.get("transparency", 0))),
                }
            }
        ],
        "border": [
            {
                "properties": {
                    "show": _literal(True),
                    "color": _solid_color(accent_color or border.get("color", "#FFFFFF")),
                    "radius": _literal(int(border.get("radius", 0))),
                    "weight": _literal(float(border.get("weight", 1.0))),
                    "transparency": _literal(int(border.get("transparency", 0))),
                }
            }
        ],
    }
    if shadow and shadow.get("blur", 0) > 0:
        vc["dropShadow"] = [
            {
                "properties": {
                    "show": _literal(True),
                    "color": _solid_color(shadow.get("color", "#000000")),
                    "blur": _literal(int(shadow.get("blur", 16))),
                    "angle": _literal(int(shadow.get("angle", 90))),
                    "transparency": _literal(int(shadow.get("transparency", 70))),
                }
            }
        ]
    return vc


def patch_layout_visuals(
    layout: dict[str, Any],
    *,
    preset: dict[str, Any],
    accent_picker,
    page_filter: set[str] | None = None,
) -> tuple[dict[str, Any], dict[str, int]]:
    """Apply the preset's card / chart styling to every visual.

    ``accent_picker(measure_name)`` returns a hex colour. The styler
    introspects each container's bound measure (via the projection
    name) and picks an accent based on it.
    """
    cards_spec = preset.get("cards", {})
    charts_spec = preset.get("charts", {})
    accent_map = cards_spec.get("accentMap", {})

    counts = {"cards_styled": 0, "charts_styled": 0}
    for section in layout.get("sections", []) or []:
        if not isinstance(section, dict):
            continue
        display = str(section.get("displayName") or section.get("name") or "")
        if page_filter and display not in page_filter:
            continue
        for container in section.get("visualContainers", []) or []:
            if not isinstance(container, dict):
                continue
            raw_cfg = container.get("config")
            try:
                cfg = json.loads(raw_cfg) if isinstance(raw_cfg, str) else (raw_cfg or {})
            except json.JSONDecodeError:
                continue
            sv = cfg.get("singleVisual")
            if not isinstance(sv, dict):
                continue
            visual_type = str(sv.get("visualType", "")).lower()
            if visual_type in ("card", "multirowcard", "kpi", "labelledcard"):
                measure_name = _first_measure_name(sv)
                accent = accent_picker(measure_name, accent_map) if accent_picker else None
                sv["vcObjects"] = _vc_objects_for_card(cards_spec, accent)
                counts["cards_styled"] += 1
            elif visual_type:
                # Apply chart chrome (background + border) without accent.
                sv["vcObjects"] = _vc_objects_for_card(charts_spec, None)
                counts["charts_styled"] += 1
            container["config"] = json.dumps(cfg, ensure_ascii=False)
    return layout, counts


def _first_measure_name(single_visual: dict[str, Any]) -> str | None:
    projections = single_visual.get("projections") or {}
    if not isinstance(projections, dict):
        return None
    # Prefer Values then Y then any.
    for key in ("Values", "Y", "Indicator"):
        slot = projections.get(key)
        if isinstance(slot, list) and slot:
            ref = slot[0].get("queryRef") or ""
            if "." in ref:
                return ref.split(".", 1)[1]
    return None


def repack_pbix(
    source_pbix: bytes,
    *,
    new_layout: bytes,
    new_resources: dict[str, bytes],
) -> bytes:
    """Return a new PBIX zip with the patched ``Report/Layout`` and any
    extra resources merged in. Original parts are copied verbatim.
    """
    out_buf = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(source_pbix), "r") as src, zipfile.ZipFile(
        out_buf, "w", zipfile.ZIP_DEFLATED
    ) as dst:
        replaced = {LAYOUT_PART, *new_resources.keys()}
        for info in src.infolist():
            if info.filename in replaced:
                continue
            dst.writestr(info, src.read(info.filename))
        dst.writestr(LAYOUT_PART, new_layout)
        for name, data in new_resources.items():
            dst.writestr(name, data)
    return out_buf.getvalue()


__all__ = [
    "LAYOUT_PART",
    "WALLPAPER_DIR",
    "THEMES_DIR",
    "patch_layout_for_wallpaper",
    "patch_layout_visuals",
    "repack_pbix",
    "sanitize_resource_name",
    "sha1_short",
]

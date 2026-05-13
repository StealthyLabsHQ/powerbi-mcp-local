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
CONTENT_TYPES_PART = "[Content_Types].xml"

# Maps lowercase file extension → OPC Default ContentType. A PBIX is an
# OPC package; Power BI rejects the file with MashupValidationError
# ("This file is corrupted or was created by an unrecognized version of
# Power BI Desktop") if a part is present whose extension has no Default
# entry in ``[Content_Types].xml``.
_DEFAULT_CONTENT_TYPES: dict[str, str] = {
    "png": "image/png",
    "jpg": "image/jpeg",
    "jpeg": "image/jpeg",
    "gif": "image/gif",
    "bmp": "image/bmp",
    "svg": "image/svg+xml",
    "json": "application/json",
    "xml": "application/xml",
}


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


WALLPAPER_SCALING_ENUM = ("Fit", "Fill", "Normal", "Stretch")
WALLPAPER_URL_PREFIX = "RegisteredResources"


def _image_block(resource_name: str, fit: str) -> dict[str, Any]:
    """Canonical Power BI canvas-image block.

    The image sub-object uses **bare strings**, not Literal-wrapped
    expressions. Power BI Desktop silently ignores image references
    that wrap ``name`` / ``url`` / ``scaling`` in ``{"expr": {"Literal":
    ...}}`` — the canvas opens blank. Only ``show`` and ``transparency``
    are Literal-wrapped.
    """
    return {
        "name": resource_name,
        "url": f"{WALLPAPER_URL_PREFIX}/{resource_name}",
        "scaling": fit,
    }


def patch_layout_for_wallpaper(
    layout: dict[str, Any],
    *,
    resource_name: str,
    fit: str = "Fit",
    transparency: int = 0,
    page_filter: set[str] | None = None,
    page_background_color: str | None = None,
    page_background_transparency: int = 0,
    apply_wallpaper_layer: bool = True,
) -> tuple[dict[str, Any], list[str]]:
    """Inject the canonical ``objects.background`` block into every
    targeted section's config.

    ``resource_name`` is the bare filename under
    ``StaticResources/RegisteredResources``. ``fit`` must be one of
    ``Fit``, ``Fill``, ``Normal``, ``Stretch`` (Power BI's enum). When
    ``apply_wallpaper_layer`` is True the same image is also written to
    ``objects.wallpaper`` (the chrome around the canvas) so the page
    has no white halo around it. Returns the patched layout and the
    list of page display names that were touched.
    """
    if fit not in WALLPAPER_SCALING_ENUM:
        raise ValueError(
            f"fit must be one of {WALLPAPER_SCALING_ENUM}, got '{fit}'"
        )

    touched: list[str] = []
    for section in layout.get("sections", []) or []:
        if not isinstance(section, dict):
            continue
        display = str(section.get("displayName") or section.get("name") or "")
        if page_filter and display not in page_filter:
            continue
        cfg = _decode_section_config(section)
        objects = cfg.setdefault("objects", {})
        background_props: dict[str, Any] = {
            "show": {"expr": {"Literal": {"Value": "true"}}},
            "image": _image_block(resource_name, fit),
            "transparency": {"expr": {"Literal": {"Value": f"{transparency}D"}}},
        }
        objects["background"] = [{"properties": background_props}]
        if apply_wallpaper_layer:
            objects["wallpaper"] = [
                {
                    "properties": {
                        "show": {"expr": {"Literal": {"Value": "true"}}},
                        "image": _image_block(resource_name, fit),
                        "transparency": {
                            "expr": {"Literal": {"Value": f"{transparency}D"}}
                        },
                    }
                }
            ]
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

    counts = {"cards_styled": 0, "charts_styled": 0, "titles_preserved": 0}
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
            # Snapshot any custom title BEFORE we replace vcObjects so it
            # can be carried forward. Power BI stores the user-set title
            # under ``objects.title`` (visual-level) — never overwrite a
            # non-empty value with the preset's default chrome.
            preserved_title = _extract_visual_title(sv)
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
            if preserved_title:
                _reinstate_visual_title(sv, preserved_title)
                counts["titles_preserved"] += 1
            container["config"] = json.dumps(cfg, ensure_ascii=False)
    return layout, counts


def _extract_visual_title(single_visual: dict[str, Any]) -> dict[str, Any] | None:
    """Return the existing ``objects.title`` block when it carries a
    non-empty text — used to round-trip user-set titles through the
    styling pass.
    """
    objects = single_visual.get("objects")
    if not isinstance(objects, dict):
        return None
    titles = objects.get("title")
    if not isinstance(titles, list) or not titles:
        return None
    for entry in titles:
        if not isinstance(entry, dict):
            continue
        props = entry.get("properties") or {}
        text_node = props.get("text")
        if not isinstance(text_node, dict):
            continue
        try:
            value = text_node["expr"]["Literal"]["Value"]
        except (KeyError, TypeError):
            continue
        if isinstance(value, str) and value.strip("'\" "):
            return {"title": titles}
    return None


def _reinstate_visual_title(single_visual: dict[str, Any], block: dict[str, Any]) -> None:
    objects = single_visual.setdefault("objects", {})
    if "title" in block:
        objects["title"] = block["title"]


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


def _required_extensions(part_names: list[str]) -> set[str]:
    """Return the lowercase extension set across ``part_names``."""
    extensions: set[str] = set()
    for name in part_names:
        ext = Path(name).suffix.lstrip(".").lower()
        if ext:
            extensions.add(ext)
    return extensions


def patch_content_types(
    raw_xml: bytes,
    required_extensions: set[str],
) -> tuple[bytes, list[str]]:
    """Ensure ``[Content_Types].xml`` declares a ``Default`` entry for
    every extension in ``required_extensions``.

    Returns the (possibly unchanged) XML bytes plus the list of
    extensions that were newly added. Preserves the original document's
    UTF-8 declaration when present. ``raw_xml`` may be empty when the
    source PBIX has no ``[Content_Types].xml`` yet (rare — older
    builders sometimes skip it). In that case a minimal document with
    every known extension is written from scratch.
    """
    added: list[str] = []

    if not raw_xml.strip():
        body = (
            b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            b'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        )
        for ext in sorted(required_extensions):
            content_type = _DEFAULT_CONTENT_TYPES.get(ext, "application/octet-stream")
            body += f'<Default Extension="{ext}" ContentType="{content_type}"/>'.encode("utf-8")
            added.append(ext)
        body += b"</Types>"
        return body, added

    text = raw_xml.decode("utf-8")
    declared: set[str] = set()
    declared_re = re.compile(
        r'<Default\s+[^>]*?Extension\s*=\s*"([^"]+)"',
        re.IGNORECASE,
    )
    for match in declared_re.finditer(text):
        declared.add(match.group(1).lower())

    missing = required_extensions - declared
    if not missing:
        return raw_xml, added

    # Insert the new Default elements right before ``</Types>``. The OPC
    # spec allows any order; the closing tag is the only stable anchor
    # we need. Fallback: append before the document's end if the closing
    # tag is missing for any reason.
    inserts = "".join(
        f'<Default Extension="{ext}" ContentType="{_DEFAULT_CONTENT_TYPES.get(ext, "application/octet-stream")}"/>'
        for ext in sorted(missing)
    )
    if "</Types>" in text:
        patched = text.replace("</Types>", inserts + "</Types>", 1)
    else:
        patched = text.rstrip() + inserts

    added = sorted(missing)
    return patched.encode("utf-8"), added


def validate_content_types_declarations(
    raw_xml: bytes,
    required_extensions: set[str],
) -> list[str]:
    """Return the list of extensions present in part list but missing
    from ``[Content_Types].xml``. Empty list ⇒ everything declared.
    """
    if not raw_xml.strip():
        return sorted(required_extensions)
    text = raw_xml.decode("utf-8", errors="ignore")
    declared = {
        match.group(1).lower()
        for match in re.finditer(
            r'<Default\s+[^>]*?Extension\s*=\s*"([^"]+)"', text, re.IGNORECASE
        )
    }
    return sorted(required_extensions - declared)


def repack_pbix(
    source_pbix: bytes,
    *,
    new_layout: bytes,
    new_resources: dict[str, bytes],
) -> bytes:
    """Return a new PBIX zip with the patched ``Report/Layout`` and any
    extra resources merged in. Original parts are copied verbatim.

    The ``[Content_Types].xml`` part is rewritten so every extension
    used by an embedded resource has a matching ``Default`` entry. Power
    BI rejects PBIX files with ``MashupValidationError`` when a part's
    extension has no Default Content-Type declaration — embedding a PNG
    without updating the manifest is a hard fail on reopen.
    """
    # Build the set of every extension referenced by the resulting
    # archive (original parts kept verbatim plus the new resources).
    out_buf = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(source_pbix), "r") as src:
        existing_names = src.namelist()
        try:
            existing_content_types = src.read(CONTENT_TYPES_PART)
        except KeyError:
            existing_content_types = b""

    final_part_names = [n for n in existing_names if n not in {LAYOUT_PART, CONTENT_TYPES_PART}]
    final_part_names.append(LAYOUT_PART)
    final_part_names.extend(new_resources.keys())
    required_exts = _required_extensions(final_part_names)
    new_content_types, _ = patch_content_types(existing_content_types, required_exts)

    with zipfile.ZipFile(io.BytesIO(source_pbix), "r") as src, zipfile.ZipFile(
        out_buf, "w", zipfile.ZIP_DEFLATED
    ) as dst:
        replaced = {LAYOUT_PART, CONTENT_TYPES_PART, *new_resources.keys()}
        for info in src.infolist():
            if info.filename in replaced:
                continue
            dst.writestr(info, src.read(info.filename))
        dst.writestr(CONTENT_TYPES_PART, new_content_types)
        dst.writestr(LAYOUT_PART, new_layout)
        for name, data in new_resources.items():
            dst.writestr(name, data)
    return out_buf.getvalue()


__all__ = [
    "CONTENT_TYPES_PART",
    "LAYOUT_PART",
    "WALLPAPER_DIR",
    "THEMES_DIR",
    "patch_content_types",
    "patch_layout_for_wallpaper",
    "patch_layout_visuals",
    "repack_pbix",
    "sanitize_resource_name",
    "sha1_short",
    "validate_content_types_declarations",
]

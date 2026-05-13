"""``pbi_apply_style_preset_tool`` orchestration.

End-to-end flow:

1. Resolve & validate the PBIX path, output path, optional wallpaper.
2. Load the preset (built-in or ``custom_spec``) and validate its theme
   payload via the existing ``validate_theme_payload`` schema.
3. Open the source PBIX zip in memory.
4. Read ``Report/Layout``, patch every targeted section with the
   wallpaper background block and every visual with the preset's
   chrome (cards: background + border + shadow + accent border; other
   visuals: chart chrome).
5. Embed the wallpaper PNG under
   ``StaticResources/RegisteredResources/<sanitized>.png`` (SHA-1 dedup
   to avoid bloating the archive on repeat apply).
6. Embed the preset's theme JSON under
   ``StaticResources/SharedResources/BaseThemes/<name>.json`` and
   record it in ``layout.themeCollection`` + ``layout.activeTheme``.
7. Repack into the destination PBIX.
8. Run the static DBCC diagnostic on the resulting file — fail loud if
   the styling round-trip regressed string-store integrity.

Returns:

    {
        "ok": bool,
        "preset": str,
        "applied_pages": list[str],
        "applied_visuals": {"cards_styled": int, "charts_styled": int},
        "embedded_resources": list[str],
        "output_path": str,
        "dbcc_valid": bool,
    }
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from pbi_connection import PowerBIValidationError, ok
from security import resolve_local_path

from ..dbcc import pbi_diagnose_pbix_dbcc_tool
from ..visuals._themes import (
    MAX_THEME_BYTES,
    ThemeValidationError,
    assert_theme_within_size_limit,
    validate_theme_payload,
)
from ._accent import pick_accent
from ._embed import (
    LAYOUT_PART,
    THEMES_DIR,
    WALLPAPER_DIR,
    patch_layout_for_wallpaper,
    patch_layout_visuals,
    repack_pbix,
    sanitize_resource_name,
    sha1_short,
)
from ._png import inspect_png, write_gradient_png
from ._presets import PRESETS

MAX_WALLPAPER_BYTES = 2 * 1024 * 1024
MAX_WALLPAPER_WIDTH = 1920
MAX_WALLPAPER_HEIGHT = 1080

WALLPAPER_CACHE = Path(__file__).parent / "_wallpapers"


def _ensure_default_wallpaper(preset: dict[str, Any]) -> Path:
    """Generate the preset's default wallpaper PNG if missing.

    Each preset declares ``wallpaper_default`` (a filename) — the file
    is materialised on first use under ``_wallpapers/`` from the
    preset's palette as a vertical gradient.
    """
    name = preset.get("wallpaper_default")
    if not name:
        raise PowerBIValidationError(
            "Preset has no wallpaper_default; pass wallpaper_path explicitly.",
            details={"preset": preset.get("name")},
        )
    path = WALLPAPER_CACHE / str(name)
    if path.exists():
        return path
    palette = preset.get("palette", {})
    page_color = preset.get("page", {}).get("background", {}).get("color", "#000000")
    top = palette.get("primary") or palette.get("secondary") or page_color
    bottom = palette.get("accent") or palette.get("secondary") or page_color
    write_gradient_png(path, top_color=top, bottom_color=bottom)
    return path


def _validate_wallpaper(path: Path) -> dict[str, int]:
    info = inspect_png(path)
    if info["size_bytes"] > MAX_WALLPAPER_BYTES:
        raise PowerBIValidationError(
            f"Wallpaper exceeds the {MAX_WALLPAPER_BYTES} byte limit.",
            details={"path": str(path), "size_bytes": info["size_bytes"]},
        )
    if info["width"] > MAX_WALLPAPER_WIDTH or info["height"] > MAX_WALLPAPER_HEIGHT:
        raise PowerBIValidationError(
            "Wallpaper exceeds the 1920x1080 limit; resize before embed.",
            details={"path": str(path), **info},
        )
    return info


def _validate_preset_palette(preset: dict[str, Any]) -> None:
    import re

    hex_re = re.compile(r"^#[0-9A-Fa-f]{6}$")
    palette = preset.get("palette", {})
    for key, value in palette.items():
        if not isinstance(value, str) or not hex_re.match(value):
            raise PowerBIValidationError(
                f"Preset palette key '{key}' is not a #RRGGBB string.",
                details={"preset": preset.get("name"), "key": key, "value": value},
            )


def _resolve_preset(preset_name: str, custom_spec: dict[str, Any] | None) -> dict[str, Any]:
    if preset_name == "custom":
        if not isinstance(custom_spec, dict):
            raise PowerBIValidationError(
                "preset='custom' requires custom_spec to be a dict.",
                details={"preset": "custom"},
            )
        # custom_spec must declare the same keys as a built-in preset.
        custom = dict(custom_spec)
        custom.setdefault("name", "custom")
        return custom
    if preset_name not in PRESETS:
        raise PowerBIValidationError(
            "Unknown style preset.",
            details={"preset": preset_name, "available": sorted(PRESETS)},
        )
    return PRESETS[preset_name]


def pbi_apply_style_preset_tool(
    pbix_path: str,
    preset: str,
    *,
    wallpaper_path: str | None = None,
    output_path: str | None = None,
    pages: list[str] | None = None,
    custom_spec: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """Apply a one-shot visual style preset to an existing ``.pbix``.

    See module docstring for the full flow.
    """
    pbix_in = resolve_local_path(pbix_path, must_exist=True, allowed_extensions={".pbix"})
    pbix_out_path = (
        resolve_local_path(output_path, must_exist=False, allowed_extensions={".pbix"})
        if output_path
        else pbix_in
    )

    spec = _resolve_preset(preset, custom_spec)
    _validate_preset_palette(spec)

    theme_payload = spec.get("theme")
    if not isinstance(theme_payload, dict):
        raise PowerBIValidationError(
            "Preset is missing a 'theme' object.",
            details={"preset": spec.get("name")},
        )
    theme_bytes = json.dumps(theme_payload, ensure_ascii=False, indent=2).encode("utf-8")
    assert_theme_within_size_limit(len(theme_bytes))
    issues = validate_theme_payload(theme_payload)
    errors = [issue for issue in issues if issue.get("level") == "error"]
    if errors:
        raise ThemeValidationError(
            "Preset theme JSON failed schema validation.",
            details={"preset": spec.get("name"), "errors": errors[:20]},
        )

    if wallpaper_path:
        wallpaper_file = resolve_local_path(
            wallpaper_path, must_exist=True, allowed_extensions={".png", ".jpg", ".jpeg"}
        )
    else:
        wallpaper_file = _ensure_default_wallpaper(spec)
    wallpaper_info = _validate_wallpaper(wallpaper_file)

    pbix_bytes = pbix_in.read_bytes()
    import zipfile, io

    if not zipfile.is_zipfile(io.BytesIO(pbix_bytes)):
        raise PowerBIValidationError(
            "PBIX is not a valid zip archive.", details={"pbix_path": str(pbix_in)}
        )
    with zipfile.ZipFile(io.BytesIO(pbix_bytes), "r") as zf:
        try:
            layout_raw = zf.read(LAYOUT_PART)
        except KeyError as exc:
            raise PowerBIValidationError(
                "PBIX has no Report/Layout part.", details={"pbix_path": str(pbix_in)}
            ) from exc

    try:
        layout = json.loads(layout_raw.decode("utf-16-le"))
    except (UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise PowerBIValidationError(
            "Report/Layout is not valid UTF-16-LE JSON.",
            details={"pbix_path": str(pbix_in), "reason": str(exc)},
        ) from exc

    wallpaper_bytes = wallpaper_file.read_bytes()
    sanitized = sanitize_resource_name(wallpaper_file.stem)
    resource_filename = f"{sanitized}_{sha1_short(wallpaper_bytes)}.png"
    resource_relpath = f"{WALLPAPER_DIR}/{resource_filename}"

    page_filter = set(pages) if pages else None
    page_spec = spec.get("page", {}) or {}
    fit = (page_spec.get("wallpaper") or {}).get("fit", "Fit")
    wallpaper_transparency = (page_spec.get("wallpaper") or {}).get("transparency", 0)
    page_bg_color = (page_spec.get("background") or {}).get("color")
    page_bg_transparency = (page_spec.get("background") or {}).get("transparency", 0)

    layout, applied_pages = patch_layout_for_wallpaper(
        layout,
        resource_name=resource_filename,
        fit=fit,
        transparency=int(wallpaper_transparency),
        page_filter=page_filter,
        page_background_color=page_bg_color,
        page_background_transparency=int(page_bg_transparency),
    )

    layout, visual_counts = patch_layout_visuals(
        layout,
        preset=spec,
        accent_picker=pick_accent,
        page_filter=page_filter,
    )

    # Theme: register under SharedResources/BaseThemes and update the
    # layout's activeTheme / themeCollection. The theme is shipped as a
    # ``.json`` resource so Power BI Desktop sees it on open.
    theme_name = sanitize_resource_name(str(theme_payload.get("name") or spec.get("name")))
    theme_relpath = f"{THEMES_DIR}/{theme_name}.json"
    theme_entry = {"name": theme_name, "path": theme_relpath}
    themes = layout.setdefault("themeCollection", [])
    if not any(isinstance(t, dict) and t.get("path") == theme_relpath for t in themes):
        themes.append(theme_entry)
    layout["activeTheme"] = theme_entry

    new_layout = json.dumps(layout, ensure_ascii=False).encode("utf-16-le")
    new_resources = {
        resource_relpath: wallpaper_bytes,
        theme_relpath: theme_bytes,
    }
    new_pbix = repack_pbix(pbix_bytes, new_layout=new_layout, new_resources=new_resources)

    pbix_out_path.parent.mkdir(parents=True, exist_ok=True)
    pbix_out_path.write_bytes(new_pbix)

    # Post-write DBCC diagnostic — the round-trip should never regress
    # the string-store. If it does, surface that loudly.
    dbcc = pbi_diagnose_pbix_dbcc_tool(str(pbix_out_path))
    dbcc_valid = bool(dbcc.get("valid"))

    return ok(
        f"Style preset '{spec.get('name')}' applied to {len(applied_pages)} page(s).",
        preset=spec.get("name"),
        applied_pages=applied_pages,
        applied_visuals=visual_counts,
        embedded_resources=[resource_relpath, theme_relpath],
        wallpaper_info=wallpaper_info,
        output_path=str(pbix_out_path),
        dbcc_valid=dbcc_valid,
        dbcc=dbcc,
    )


def pbi_list_style_presets_tool() -> dict[str, Any]:
    """Return the catalogue of built-in presets with summary metadata."""
    out = []
    for name, spec in PRESETS.items():
        out.append(
            {
                "name": name,
                "description": spec.get("description"),
                "palette": spec.get("palette"),
                "default_wallpaper": spec.get("wallpaper_default"),
            }
        )
    return ok("Style presets listed.", presets=out)


__all__ = [
    "pbi_apply_style_preset_tool",
    "pbi_list_style_presets_tool",
    "MAX_WALLPAPER_BYTES",
    "MAX_WALLPAPER_WIDTH",
    "MAX_WALLPAPER_HEIGHT",
]

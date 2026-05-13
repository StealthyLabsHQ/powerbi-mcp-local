"""One-shot visual styling for PBIX files (v0.13.3).

Public surface:

- :func:`pbi_apply_style_preset_tool` — apply a built-in or custom
  preset to an existing ``.pbix``: wallpaper, page chrome, card +
  chart chrome, theme. Output PBIX opens with zero clicks.
- :func:`pbi_list_style_presets_tool` — catalogue listing.
- :data:`PRESETS` — built-in preset dict by name.
"""

from __future__ import annotations

from ._accent import infer_accent_key, pick_accent
from ._apply import (
    MAX_WALLPAPER_BYTES,
    MAX_WALLPAPER_HEIGHT,
    MAX_WALLPAPER_WIDTH,
    pbi_apply_style_preset_tool,
    pbi_list_style_presets_tool,
)
from ._embed import (
    CONTENT_TYPES_PART,
    LAYOUT_PART,
    THEMES_DIR,
    WALLPAPER_DIR,
    patch_content_types,
    patch_layout_for_wallpaper,
    patch_layout_visuals,
    repack_pbix,
    sanitize_resource_name,
    sha1_short,
    validate_content_types_declarations,
)
from ._png import inspect_png, write_gradient_png
from ._presets import PRESETS

__all__ = [
    "PRESETS",
    "pbi_apply_style_preset_tool",
    "pbi_list_style_presets_tool",
    "infer_accent_key",
    "pick_accent",
    "patch_content_types",
    "patch_layout_for_wallpaper",
    "patch_layout_visuals",
    "repack_pbix",
    "sanitize_resource_name",
    "sha1_short",
    "validate_content_types_declarations",
    "inspect_png",
    "write_gradient_png",
    "CONTENT_TYPES_PART",
    "LAYOUT_PART",
    "WALLPAPER_DIR",
    "THEMES_DIR",
    "MAX_WALLPAPER_BYTES",
    "MAX_WALLPAPER_WIDTH",
    "MAX_WALLPAPER_HEIGHT",
]

"""Preset catalogue for v0.13.3 styling.

Each submodule exports a single ``PRESET`` dict with the shape

    {
        "name": str,
        "description": str,
        "palette": dict[str, str],     # logical accent name -> #RRGGBB
        "theme": dict,                 # Power BI theme JSON (validated)
        "page": {
            "background": {"color": "#RRGGBB", "transparency": int},
            "wallpaper": {"fit": "Fit"|"Fill"|"Normal", "transparency": int},
        },
        "cards": {
            "background": {"color": "#RRGGBB", "transparency": int},
            "border":     {"color": "#RRGGBB", "radius": int, "weight": float, "transparency": int},
            "dropShadow": {"blur": int, "transparency": int, "angle": int, "color": "#RRGGBB"},
            "labelColor": "#RRGGBB",
            "valueColor": "#RRGGBB",
            "accentMap":  {"positive": "#RRGGBB", "warning": "#RRGGBB",
                            "info": "#RRGGBB", "neutral": "#RRGGBB"},
        },
        "charts": {
            "background": {"color": "#RRGGBB", "transparency": int},
            "border":     {"color": "#RRGGBB", "radius": int, "weight": float, "transparency": int},
            "labels":     {"color": "#RRGGBB", "fontSize": int},
        },
        "wallpaper_default": "<filename in _wallpapers/>",
    }
"""

from __future__ import annotations

from .glassmorph_dark import PRESET as GLASSMORPH_DARK
from .glassmorph_light import PRESET as GLASSMORPH_LIGHT
from .neon_cyber import PRESET as NEON_CYBER
from .minimal_corporate import PRESET as MINIMAL_CORPORATE
from .dark_pro import PRESET as DARK_PRO

PRESETS: dict[str, dict] = {
    GLASSMORPH_DARK["name"]: GLASSMORPH_DARK,
    GLASSMORPH_LIGHT["name"]: GLASSMORPH_LIGHT,
    NEON_CYBER["name"]: NEON_CYBER,
    MINIMAL_CORPORATE["name"]: MINIMAL_CORPORATE,
    DARK_PRO["name"]: DARK_PRO,
}

__all__ = ["PRESETS"]

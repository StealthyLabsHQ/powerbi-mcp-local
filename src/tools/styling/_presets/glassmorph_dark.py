"""glassmorph_dark — frosted navy glass on a dark gradient."""

from __future__ import annotations

PRESET: dict = {
    "name": "glassmorph_dark",
    "description": "Frosted navy glass on dark gradient. KPI cards translucent over a deep-blue wallpaper, soft white borders, drop shadows for depth.",
    "palette": {
        "primary": "#1F4E78",
        "secondary": "#4F81BD",
        "accent": "#5B9BD5",
        "positive": "#70AD47",
        "warning": "#E07B00",
        "danger": "#C00000",
        "neutral": "#FFFFFF",
        "ink": "#FFFFFF",
    },
    "theme": {
        "name": "Glassmorph Dark",
        "dataColors": ["#5B9BD5", "#70AD47", "#E07B00", "#C00000", "#9DC3E6", "#264478", "#A6A6A6", "#4F81BD"],
        "foreground": "#FFFFFF",
        "foregroundNeutralSecondary": "#CBD5E1",
        "foregroundNeutralTertiary": "#94A3B8",
        "background": "#0F2235",
        "backgroundLight": "#1E3A5F",
        "backgroundNeutral": "#243B53",
        "tableAccent": "#5B9BD5",
        "good": "#70AD47",
        "neutral": "#E07B00",
        "bad": "#C00000",
        "maximum": "#5B9BD5",
        "center": "#4F81BD",
        "minimum": "#1E3A5F",
        "hyperlink": "#5B9BD5",
        "visitedHyperlink": "#9DC3E6",
        "textClasses": {
            "callout": {"fontSize": 32, "fontFace": "Segoe UI Semibold", "color": "#FFFFFF"},
            "title": {"fontSize": 14, "fontFace": "Segoe UI Semibold", "color": "#FFFFFF"},
            "header": {"fontSize": 12, "fontFace": "Segoe UI Semibold", "color": "#FFFFFF"},
            "label": {"fontSize": 11, "fontFace": "Segoe UI", "color": "#CBD5E1"},
        },
    },
    "page": {
        "background": {"color": "#0F2235", "transparency": 0},
        "wallpaper": {"fit": "Stretch", "transparency": 0},
    },
    "cards": {
        "background": {"color": "#1F4E78", "transparency": 50},
        "border": {"color": "#FFFFFF", "radius": 16, "weight": 1.5, "transparency": 60},
        "dropShadow": {"blur": 32, "transparency": 70, "angle": 90, "color": "#000000"},
        "labelColor": "#FFFFFF",
        "valueColor": "#FFFFFF",
        "accentMap": {
            "positive": "#70AD47",
            "warning": "#E07B00",
            "info": "#5B9BD5",
            "neutral": "#FFFFFF",
        },
    },
    "charts": {
        "background": {"color": "#1F4E78", "transparency": 60},
        "border": {"color": "#FFFFFF", "radius": 12, "weight": 1.0, "transparency": 70},
        "labels": {"color": "#FFFFFF", "fontSize": 11},
    },
    "wallpaper_default": "bg_glassmorph_dark.png",
}

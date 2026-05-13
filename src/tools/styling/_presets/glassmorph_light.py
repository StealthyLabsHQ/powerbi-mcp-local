"""glassmorph_light — frosted glass on a pale sky gradient."""

from __future__ import annotations

PRESET: dict = {
    "name": "glassmorph_light",
    "description": "Translucent white panels on a sky-blue gradient. Light theme, ink on white, gentle shadows.",
    "palette": {
        "primary": "#1F4E78",
        "secondary": "#4F81BD",
        "accent": "#5B9BD5",
        "positive": "#70AD47",
        "warning": "#E07B00",
        "danger": "#C00000",
        "neutral": "#1F4E78",
        "ink": "#1F2937",
    },
    "theme": {
        "name": "Glassmorph Light",
        "dataColors": ["#1F4E78", "#4F81BD", "#70AD47", "#E07B00", "#C00000", "#7030A0", "#A6A6A6", "#264478"],
        "foreground": "#1F2937",
        "foregroundNeutralSecondary": "#475569",
        "foregroundNeutralTertiary": "#94A3B8",
        "background": "#F3F8FE",
        "backgroundLight": "#FFFFFF",
        "backgroundNeutral": "#E5EBF5",
        "tableAccent": "#1F4E78",
        "good": "#70AD47",
        "neutral": "#E07B00",
        "bad": "#C00000",
        "maximum": "#1F4E78",
        "center": "#4F81BD",
        "minimum": "#A6A6A6",
        "hyperlink": "#1F4E78",
        "visitedHyperlink": "#264478",
        "textClasses": {
            "callout": {"fontSize": 32, "fontFace": "Segoe UI Semibold", "color": "#1F4E78"},
            "title": {"fontSize": 14, "fontFace": "Segoe UI Semibold", "color": "#1F4E78"},
            "header": {"fontSize": 12, "fontFace": "Segoe UI Semibold", "color": "#1F2937"},
            "label": {"fontSize": 11, "fontFace": "Segoe UI", "color": "#475569"},
        },
    },
    "page": {
        "background": {"color": "#F3F8FE", "transparency": 0},
        "wallpaper": {"fit": "Stretch", "transparency": 0},
    },
    "cards": {
        "background": {"color": "#FFFFFF", "transparency": 30},
        "border": {"color": "#1F4E78", "radius": 14, "weight": 1.0, "transparency": 70},
        "dropShadow": {"blur": 24, "transparency": 80, "angle": 90, "color": "#1F4E78"},
        "labelColor": "#475569",
        "valueColor": "#1F4E78",
        "accentMap": {
            "positive": "#70AD47",
            "warning": "#E07B00",
            "info": "#1F4E78",
            "neutral": "#475569",
        },
    },
    "charts": {
        "background": {"color": "#FFFFFF", "transparency": 40},
        "border": {"color": "#1F4E78", "radius": 12, "weight": 1.0, "transparency": 75},
        "labels": {"color": "#1F2937", "fontSize": 11},
    },
    "wallpaper_default": "bg_glassmorph_light.png",
}

"""neon_cyber — synthwave neon on near-black."""

from __future__ import annotations

PRESET: dict = {
    "name": "neon_cyber",
    "description": "Synthwave neon palette (magenta / cyan / lime) on a near-black background. High contrast, glow borders, aggressive accents.",
    "palette": {
        "primary": "#FF00A8",
        "secondary": "#00E5FF",
        "accent": "#B8FF00",
        "positive": "#B8FF00",
        "warning": "#FFD600",
        "danger": "#FF3D00",
        "neutral": "#00E5FF",
        "ink": "#E0F7FF",
    },
    "theme": {
        "name": "Neon Cyber",
        "dataColors": ["#FF00A8", "#00E5FF", "#B8FF00", "#FFD600", "#FF3D00", "#7C4DFF", "#00BFA5", "#FF6E40"],
        "foreground": "#E0F7FF",
        "foregroundNeutralSecondary": "#9CCFE5",
        "foregroundNeutralTertiary": "#64748B",
        "background": "#0A0014",
        "backgroundLight": "#1A0033",
        "backgroundNeutral": "#0F001F",
        "tableAccent": "#FF00A8",
        "good": "#B8FF00",
        "neutral": "#FFD600",
        "bad": "#FF3D00",
        "maximum": "#FF00A8",
        "center": "#7C4DFF",
        "minimum": "#1A0033",
        "hyperlink": "#00E5FF",
        "visitedHyperlink": "#7C4DFF",
        "textClasses": {
            "callout": {"fontSize": 34, "fontFace": "Segoe UI Semibold", "color": "#00E5FF"},
            "title": {"fontSize": 14, "fontFace": "Segoe UI Semibold", "color": "#FF00A8"},
            "header": {"fontSize": 12, "fontFace": "Segoe UI Semibold", "color": "#E0F7FF"},
            "label": {"fontSize": 11, "fontFace": "Segoe UI", "color": "#9CCFE5"},
        },
    },
    "page": {
        "background": {"color": "#0A0014", "transparency": 0},
        "wallpaper": {"fit": "Fit", "transparency": 0},
    },
    "cards": {
        "background": {"color": "#1A0033", "transparency": 40},
        "border": {"color": "#FF00A8", "radius": 12, "weight": 2.0, "transparency": 30},
        "dropShadow": {"blur": 40, "transparency": 60, "angle": 90, "color": "#FF00A8"},
        "labelColor": "#9CCFE5",
        "valueColor": "#00E5FF",
        "accentMap": {
            "positive": "#B8FF00",
            "warning": "#FFD600",
            "info": "#00E5FF",
            "neutral": "#E0F7FF",
        },
    },
    "charts": {
        "background": {"color": "#1A0033", "transparency": 50},
        "border": {"color": "#00E5FF", "radius": 10, "weight": 1.5, "transparency": 50},
        "labels": {"color": "#E0F7FF", "fontSize": 11},
    },
    "wallpaper_default": "bg_neon_cyber.png",
}

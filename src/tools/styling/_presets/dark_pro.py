"""dark_pro — saturated dark dashboard, no glass."""

from __future__ import annotations

PRESET: dict = {
    "name": "dark_pro",
    "description": "Saturated dark dashboard with opaque cards, sharp accents, no glass. Aimed at TV / kiosk display where translucency hurts readability.",
    "palette": {
        "primary": "#2563EB",
        "secondary": "#0EA5E9",
        "accent": "#22D3EE",
        "positive": "#22C55E",
        "warning": "#F59E0B",
        "danger": "#EF4444",
        "neutral": "#94A3B8",
        "ink": "#F1F5F9",
    },
    "theme": {
        "name": "Dark Pro",
        "dataColors": ["#2563EB", "#22C55E", "#F59E0B", "#EF4444", "#22D3EE", "#A855F7", "#94A3B8", "#0EA5E9"],
        "foreground": "#F1F5F9",
        "foregroundNeutralSecondary": "#CBD5E1",
        "foregroundNeutralTertiary": "#64748B",
        "background": "#0B1220",
        "backgroundLight": "#111827",
        "backgroundNeutral": "#1E293B",
        "tableAccent": "#2563EB",
        "good": "#22C55E",
        "neutral": "#F59E0B",
        "bad": "#EF4444",
        "maximum": "#2563EB",
        "center": "#0EA5E9",
        "minimum": "#1E293B",
        "hyperlink": "#22D3EE",
        "visitedHyperlink": "#A855F7",
        "textClasses": {
            "callout": {"fontSize": 32, "fontFace": "Segoe UI Semibold", "color": "#F1F5F9"},
            "title": {"fontSize": 14, "fontFace": "Segoe UI Semibold", "color": "#F1F5F9"},
            "header": {"fontSize": 12, "fontFace": "Segoe UI Semibold", "color": "#F1F5F9"},
            "label": {"fontSize": 11, "fontFace": "Segoe UI", "color": "#CBD5E1"},
        },
    },
    "page": {
        "background": {"color": "#0B1220", "transparency": 0},
        "wallpaper": {"fit": "Fit", "transparency": 0},
    },
    "cards": {
        "background": {"color": "#111827", "transparency": 0},
        "border": {"color": "#1E293B", "radius": 8, "weight": 1.0, "transparency": 0},
        "dropShadow": {"blur": 16, "transparency": 70, "angle": 90, "color": "#000000"},
        "labelColor": "#CBD5E1",
        "valueColor": "#F1F5F9",
        "accentMap": {
            "positive": "#22C55E",
            "warning": "#F59E0B",
            "info": "#22D3EE",
            "neutral": "#94A3B8",
        },
    },
    "charts": {
        "background": {"color": "#111827", "transparency": 0},
        "border": {"color": "#1E293B", "radius": 8, "weight": 1.0, "transparency": 0},
        "labels": {"color": "#F1F5F9", "fontSize": 11},
    },
    "wallpaper_default": "bg_dark_pro.png",
}

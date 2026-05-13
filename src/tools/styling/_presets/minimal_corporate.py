"""minimal_corporate — white canvas, subtle slate accents, no chrome."""

from __future__ import annotations

PRESET: dict = {
    "name": "minimal_corporate",
    "description": "Boardroom-ready white canvas with subtle slate accents. No shadows, hairline borders, calm typography.",
    "palette": {
        "primary": "#1F2937",
        "secondary": "#475569",
        "accent": "#0F766E",
        "positive": "#15803D",
        "warning": "#B45309",
        "danger": "#B91C1C",
        "neutral": "#475569",
        "ink": "#1F2937",
    },
    "theme": {
        "name": "Minimal Corporate",
        "dataColors": ["#0F766E", "#1F2937", "#B45309", "#B91C1C", "#475569", "#0EA5E9", "#7C3AED", "#65A30D"],
        "foreground": "#1F2937",
        "foregroundNeutralSecondary": "#475569",
        "foregroundNeutralTertiary": "#94A3B8",
        "background": "#FFFFFF",
        "backgroundLight": "#F8FAFC",
        "backgroundNeutral": "#E2E8F0",
        "tableAccent": "#0F766E",
        "good": "#15803D",
        "neutral": "#B45309",
        "bad": "#B91C1C",
        "maximum": "#0F766E",
        "center": "#475569",
        "minimum": "#E2E8F0",
        "hyperlink": "#0F766E",
        "visitedHyperlink": "#1F2937",
        "textClasses": {
            "callout": {"fontSize": 30, "fontFace": "Segoe UI", "color": "#1F2937"},
            "title": {"fontSize": 13, "fontFace": "Segoe UI Semibold", "color": "#1F2937"},
            "header": {"fontSize": 12, "fontFace": "Segoe UI Semibold", "color": "#1F2937"},
            "label": {"fontSize": 11, "fontFace": "Segoe UI", "color": "#475569"},
        },
    },
    "page": {
        "background": {"color": "#FFFFFF", "transparency": 0},
        "wallpaper": {"fit": "Fit", "transparency": 0},
    },
    "cards": {
        "background": {"color": "#FFFFFF", "transparency": 0},
        "border": {"color": "#E2E8F0", "radius": 4, "weight": 1.0, "transparency": 0},
        "dropShadow": {"blur": 0, "transparency": 100, "angle": 90, "color": "#000000"},
        "labelColor": "#475569",
        "valueColor": "#1F2937",
        "accentMap": {
            "positive": "#15803D",
            "warning": "#B45309",
            "info": "#0F766E",
            "neutral": "#475569",
        },
    },
    "charts": {
        "background": {"color": "#FFFFFF", "transparency": 0},
        "border": {"color": "#E2E8F0", "radius": 4, "weight": 1.0, "transparency": 0},
        "labels": {"color": "#1F2937", "fontSize": 11},
    },
    "wallpaper_default": "bg_minimal_corporate.png",
}

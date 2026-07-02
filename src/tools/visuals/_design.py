"""Theme + design preset application + dashboard composition.

Bundles the named ``DESIGN_PRESETS`` catalogue, theme JSON application,
the holistic ``pbi_apply_design_tool`` (preset + page background + card
styling), and the ``pbi_build_dashboard_tool`` multi-visual composer
(driven by its own ``_create_visual_from_spec`` mini-dispatcher).
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from pbi_connection import PowerBIValidationError, ok
from security import SECURITY

from ._base import (
    DEFAULT_VISUAL_SIZES,
    DESIGN_THEME_RELATIVE_PATH,
    HEX_COLOR_RE,
    THEMES_RELATIVE_DIR,
    _run,
)
from ._bindings import _assert_container_bindings
from ._containers import _create_chart_container, _make_visual_container, _visual_payload
from ._home_tables import _scan_measure_home_tables
from ._layout import (
    _dump_embedded_json,
    _find_page,
    _load_layout,
    _page_summary,
    _parse_embedded_json,
    _save_layout,
)
from ._paths import _resolve_theme_path
from ._refs import _projection
from ._themes import (
    MAX_THEME_BYTES,
    ThemeValidationError,
    assert_theme_within_size_limit,
    validate_theme_payload,
)

DESIGN_PRESETS: dict[str, dict[str, Any]] = {
    "powerbi-navy-pro": {
        "name": "Power BI Navy Pro",
        "dataColors": ["#1E40AF", "#0EA5E9", "#059669", "#D97706", "#7C3AED", "#DB2777", "#0891B2", "#EA580C"],
        "foreground": "#1E293B",
        "foregroundNeutralSecondary": "#475569",
        "foregroundNeutralTertiary": "#94A3B8",
        "background": "#FFFFFF",
        "backgroundLight": "#F1F5F9",
        "backgroundNeutral": "#CBD5E0",
        "tableAccent": "#1E40AF",
        "good": "#059669",
        "neutral": "#D97706",
        "bad": "#DC2626",
        "maximum": "#1E40AF",
        "center": "#D97706",
        "minimum": "#DBEAFE",
        "hyperlink": "#1E40AF",
        "visitedHyperlink": "#1E40AF",
        "textClasses": {
            "callout": {"fontSize": 28, "fontFace": "Segoe UI Semibold", "color": "#1E293B"},
            "title": {"fontSize": 13, "fontFace": "Segoe UI Semibold", "color": "#1E40AF"},
            "header": {"fontSize": 12, "fontFace": "Segoe UI Semibold", "color": "#1E293B"},
            "label": {"fontSize": 10, "fontFace": "Segoe UI", "color": "#475569"},
        },
        "visualStyles": {
            "*": {
                "*": {
                    "background": [{"show": True, "color": {"solid": {"color": "#FFFFFF"}}, "transparency": 0}],
                    "border": [{"show": True, "color": {"solid": {"color": "#DBEAFE"}}, "radius": 8}],
                    "shadow": [{"show": True}],
                    "title": [
                        {
                            "show": True,
                            "fontColor": {"solid": {"color": "#1E40AF"}},
                            "background": {"solid": {"color": "#FFFFFF"}},
                            "fontSize": 12,
                            "fontFamily": "Segoe UI Semibold",
                        }
                    ],
                    "lineStyles": [{"strokeWidth": 3}],
                    "categoryAxis": [
                        {
                            "showAxisTitle": False,
                            "gridlineStyle": "dotted",
                            "gridlineColor": {"solid": {"color": "#E2E8F0"}},
                        }
                    ],
                    "valueAxis": [
                        {
                            "showAxisTitle": False,
                            "gridlineStyle": "dotted",
                            "gridlineColor": {"solid": {"color": "#E2E8F0"}},
                        }
                    ],
                }
            },
            "card": {
                "*": {
                    "labels": [
                        {
                            "color": {"solid": {"color": "#1E293B"}},
                            "fontSize": 22,
                            "fontBold": True,
                            "fontFamily": "Segoe UI Semibold",
                        }
                    ],
                    "categoryLabels": [
                        {"color": {"solid": {"color": "#475569"}}, "fontSize": 11, "fontFamily": "Segoe UI"}
                    ],
                    "outline": [{"show": True, "color": {"solid": {"color": "#BFDBFE"}}, "weight": 2}],
                    "background": [{"show": True, "color": {"solid": {"color": "#FFFFFF"}}, "transparency": 0}],
                    "border": [{"show": True, "color": {"solid": {"color": "#BFDBFE"}}, "radius": 8}],
                    "shadow": [{"show": True}],
                    "title": [{"show": False}],
                }
            },
            "slicer": {
                "*": {
                    "background": [{"show": True, "color": {"solid": {"color": "#FFFFFF"}}, "transparency": 0}],
                    "border": [{"show": True, "color": {"solid": {"color": "#BFDBFE"}}, "radius": 8}],
                    "title": [{"show": True, "fontColor": {"solid": {"color": "#1E40AF"}}, "fontSize": 12}],
                }
            },
            "gauge": {
                "*": {
                    "calloutValue": [{"color": {"solid": {"color": "#1E293B"}}, "fontSize": 20, "fontBold": True}],
                    "background": [{"show": True, "color": {"solid": {"color": "#FFFFFF"}}, "transparency": 0}],
                    "border": [{"show": True, "color": {"solid": {"color": "#DBEAFE"}}, "radius": 8}],
                    "shadow": [{"show": True}],
                }
            },
            "tableEx": {
                "*": {
                    "background": [{"show": True, "color": {"solid": {"color": "#FFFFFF"}}, "transparency": 0}],
                    "border": [{"show": True, "color": {"solid": {"color": "#DBEAFE"}}, "radius": 8}],
                    "shadow": [{"show": True}],
                    "columnHeaders": [
                        {
                            "fontColor": {"solid": {"color": "#1E40AF"}},
                            "backColor": {"solid": {"color": "#EFF6FF"}},
                            "fontSize": 11,
                            "fontBold": True,
                        }
                    ],
                    "values": [
                        {
                            "fontColor": {"solid": {"color": "#1E293B"}},
                            "backColor": {"solid": {"color": "#FFFFFF"}},
                            "altBackColor": {"solid": {"color": "#F8FAFC"}},
                            "fontSize": 10,
                        }
                    ],
                }
            },
        },
    }
}


def _validate_hex_color(value: str, *, field: str) -> None:
    if not HEX_COLOR_RE.match(value):
        raise PowerBIValidationError(
            f"{field} must match '#RRGGBB'.",
            details={"field": field, "value": value},
        )


def _validate_preset_hex_colors(value: Any, *, field: str) -> None:
    if isinstance(value, str):
        if value.startswith("#"):
            _validate_hex_color(value, field=field)
        return
    if isinstance(value, list):
        for index, item in enumerate(value):
            _validate_preset_hex_colors(item, field=f"{field}[{index}]")
        return
    if isinstance(value, dict):
        for key, item in value.items():
            _validate_preset_hex_colors(item, field=f"{field}.{key}")


def _card_vc_objects() -> dict[str, Any]:
    return {
        "background": [
            {
                "properties": {
                    "show": {"expr": {"Literal": {"Value": "true"}}},
                    "color": {"solid": {"color": "#FFFFFF"}},
                }
            }
        ],
        "border": [
            {
                "properties": {
                    "show": {"expr": {"Literal": {"Value": "true"}}},
                    "color": {"solid": {"color": "#BFDBFE"}},
                }
            }
        ],
        "shadow": [{"properties": {"show": {"expr": {"Literal": {"Value": "true"}}}}}],
    }


def pbi_apply_theme_tool(extract_folder: str, theme_json_path: str) -> dict[str, Any]:
    def _impl() -> dict[str, Any]:
        folder, layout = _load_layout(extract_folder)
        theme_path = _resolve_theme_path(theme_json_path)
        raw = theme_path.read_bytes()
        assert_theme_within_size_limit(len(raw))
        try:
            theme_payload = json.loads(raw.decode("utf-8"))
        except json.JSONDecodeError as exc:
            raise PowerBIValidationError(
                "Theme JSON is invalid.", details={"path": str(theme_path), "line": exc.lineno}
            ) from exc
        issues = validate_theme_payload(theme_payload)
        errors = [issue for issue in issues if issue.get("level") == "error"]
        if errors:
            raise ThemeValidationError(
                "Theme JSON failed schema validation.",
                details={"path": str(theme_path), "errors": errors[:20]},
            )
        target = folder / THEMES_RELATIVE_DIR / theme_path.name
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_text(json.dumps(theme_payload, ensure_ascii=False, indent=2), encoding="utf-8")
        relative_path = str(target.relative_to(folder)).replace("\\", "/")
        theme_entry = {"name": theme_path.stem, "path": relative_path}
        themes = layout.setdefault("themeCollection", [])
        if not any(str(item.get("path")) == relative_path for item in themes if isinstance(item, dict)):
            themes.append(theme_entry)
        layout["activeTheme"] = theme_entry
        _save_layout(folder, layout)
        return ok(
            "Theme applied successfully.",
            extract_folder=str(folder),
            theme=theme_entry,
            warnings=[issue for issue in issues if issue.get("level") != "error"],
        )

    return _run(_impl)


def pbi_validate_theme_tool(theme_json_path: str) -> dict[str, Any]:
    """Dry-run validation of a user-supplied theme JSON file.

    Returns the parsed payload size, any schema issues (errors +
    warnings), and the list of allowed top-level keys for reference.
    Performs no disk write and does not require an extract folder.
    """

    def _impl() -> dict[str, Any]:
        theme_path = _resolve_theme_path(theme_json_path)
        raw = theme_path.read_bytes()
        size_bytes = len(raw)
        try:
            assert_theme_within_size_limit(size_bytes)
            theme_payload = json.loads(raw.decode("utf-8"))
        except json.JSONDecodeError as exc:
            raise PowerBIValidationError(
                "Theme JSON is invalid.", details={"path": str(theme_path), "line": exc.lineno}
            ) from exc
        issues = validate_theme_payload(theme_payload)
        errors = [issue for issue in issues if issue.get("level") == "error"]
        return ok(
            "Theme validation complete.",
            path=str(theme_path),
            size_bytes=size_bytes,
            size_limit_bytes=MAX_THEME_BYTES,
            valid=not errors,
            error_count=len(errors),
            warning_count=len(issues) - len(errors),
            issues=issues,
        )

    return _run(_impl)


def pbi_export_active_theme_tool(extract_folder: str, output_path: str) -> dict[str, Any]:
    """Export the currently active theme JSON from an extracted report.

    Writes a copy of the theme referenced by ``activeTheme`` (or the
    last entry in ``themeCollection``) to ``output_path``. Useful for
    capturing the baseline theme before customising it.
    """

    def _impl() -> dict[str, Any]:
        from security import resolve_local_path

        folder, layout = _load_layout(extract_folder)
        active = layout.get("activeTheme") or {}
        if not isinstance(active, dict):
            active = {}
        relative_path = active.get("path")
        if not relative_path:
            themes = layout.get("themeCollection") or []
            if isinstance(themes, list):
                for entry in reversed(themes):
                    if isinstance(entry, dict) and entry.get("path"):
                        relative_path = entry["path"]
                        break
        if not relative_path:
            raise PowerBIValidationError(
                "No active theme is referenced by the report layout.",
                details={"extract_folder": str(folder)},
            )
        # Layout stores the theme path with either separator; normalize to
        # forward slashes so the join works on POSIX too (Windows Path
        # accepts "/" natively).
        theme_file = folder / Path(str(relative_path).replace("\\", "/"))
        if not theme_file.exists():
            raise PowerBIValidationError(
                "Active theme file is missing from the extract folder.",
                details={"theme_path": str(theme_file)},
            )
        destination = resolve_local_path(output_path, must_exist=False, allowed_extensions={".json"})
        destination.parent.mkdir(parents=True, exist_ok=True)
        raw = theme_file.read_bytes()
        destination.write_bytes(raw)
        return ok(
            "Active theme exported.",
            source=str(theme_file),
            output_path=str(destination),
            size_bytes=len(raw),
            theme_name=str(active.get("name") or destination.stem),
        )

    return _run(_impl)


def pbi_apply_design_tool(
    extract_folder: str,
    *,
    preset: str = "powerbi-navy-pro",
    page_background: str | None = "#F0F4FB",
    style_cards: bool = True,
) -> dict[str, Any]:
    def _impl() -> dict[str, Any]:
        folder = SECURITY.validate_directory(extract_folder, must_exist=True)
        if preset not in DESIGN_PRESETS:
            raise PowerBIValidationError(
                "Unknown design preset.",
                details={"preset": preset, "available_presets": sorted(DESIGN_PRESETS)},
            )
        if page_background is not None:
            _validate_hex_color(page_background, field="page_background")

        theme_payload = DESIGN_PRESETS[preset]
        _validate_preset_hex_colors(theme_payload, field=f"preset:{preset}")

        _, layout = _load_layout(folder)

        pages_updated = 0
        if page_background is not None:
            for section in layout.get("sections", []):
                if not isinstance(section, dict):
                    continue
                section_config = _parse_embedded_json(section.get("config"), {})
                if not isinstance(section_config, dict):
                    section_config = {}
                section_config["background"] = {
                    "color": {"solid": {"color": page_background}},
                    "transparency": 0,
                }
                section["config"] = _dump_embedded_json(section_config)
                pages_updated += 1

        cards_styled = 0
        if style_cards:
            for section in layout.get("sections", []):
                if not isinstance(section, dict):
                    continue
                for container in section.get("visualContainers", []) or []:
                    if not isinstance(container, dict):
                        continue
                    container_config = _parse_embedded_json(container.get("config"), {})
                    if not isinstance(container_config, dict):
                        continue
                    single_visual = container_config.get("singleVisual")
                    if not isinstance(single_visual, dict):
                        continue
                    if str(single_visual.get("visualType", "")).casefold() != "card":
                        continue
                    single_visual["vcObjects"] = _card_vc_objects()
                    container["config"] = _dump_embedded_json(container_config)
                    cards_styled += 1

        theme_path = folder / DESIGN_THEME_RELATIVE_PATH
        theme_path.parent.mkdir(parents=True, exist_ok=True)
        theme_path.write_text(json.dumps(theme_payload, ensure_ascii=False, indent=2), encoding="utf-8")
        relative_theme_path = str(theme_path.relative_to(folder)).replace("\\", "/")
        theme_entry = {"name": str(theme_payload.get("name") or preset), "path": relative_theme_path}
        themes = layout.setdefault("themeCollection", [])
        if not any(str(item.get("path")) == relative_theme_path for item in themes if isinstance(item, dict)):
            themes.append(theme_entry)
        layout["activeTheme"] = theme_entry

        _save_layout(folder, layout)
        return ok(
            f"Design '{preset}' applied.",
            preset=preset,
            theme_file=str(theme_path),
            pages_updated=pages_updated,
            cards_styled=cards_styled,
            page_background=page_background,
        )

    return _run(_impl)


def _create_visual_from_spec(
    section: dict[str, Any],
    spec: dict[str, Any],
    measure_home_map: dict[str, str] | None = None,
) -> dict[str, Any]:
    visual_type = str(spec.get("type", "")).strip().casefold()
    x = int(spec.get("x", 0))
    y = int(spec.get("y", 0))
    width = int(spec.get("width", DEFAULT_VISUAL_SIZES.get(visual_type, (400, 300))[0]))
    height = int(spec.get("height", DEFAULT_VISUAL_SIZES.get(visual_type, (400, 300))[1]))
    title = spec.get("title")
    if visual_type == "card":
        return _create_chart_container(
            section,
            visual_type="card",
            x=x,
            y=y,
            width=width,
            height=height,
            title=title,
            projections={"Values": [_projection(spec["measure"])]},
            references=[spec["measure"]],
            measure_home_map=measure_home_map,
        )
    if visual_type in {"bar_chart", "bar"}:
        projections = {
            "Category": [_projection(spec["category"])],
            "Y": [_projection(spec["measure"])],
        }
        references = [spec["category"], spec["measure"]]
        if spec.get("legend"):
            projections["Series"] = [_projection(spec["legend"])]
            references.append(spec["legend"])
        return _create_chart_container(
            section,
            visual_type="clusteredBarChart",
            x=x,
            y=y,
            width=width,
            height=height,
            title=title,
            projections=projections,
            references=references,
            measure_home_map=measure_home_map,
        )
    if visual_type in {"line_chart", "line"}:
        measures = list(spec.get("measures") or [spec.get("measure")])
        return _create_chart_container(
            section,
            visual_type="lineChart",
            x=x,
            y=y,
            width=width,
            height=height,
            title=title,
            projections={
                "Category": [_projection(spec["axis"])],
                "Y": [_projection(item) for item in measures],
            },
            references=[spec["axis"], *measures],
            measure_home_map=measure_home_map,
        )
    if visual_type in {"donut", "donut_chart", "pie", "pie_chart"}:
        return _create_chart_container(
            section,
            visual_type="donutChart",
            x=x,
            y=y,
            width=width,
            height=height,
            title=title,
            projections={
                "Category": [_projection(spec["category"])],
                "Y": [_projection(spec["measure"])],
            },
            references=[spec["category"], spec["measure"]],
            measure_home_map=measure_home_map,
        )
    if visual_type in {"table", "table_visual"}:
        return _create_chart_container(
            section,
            visual_type="tableEx",
            x=x,
            y=y,
            width=width,
            height=height,
            title=title,
            projections={"Values": [_projection(item) for item in spec["columns"]]},
            references=list(spec["columns"]),
            measure_home_map=measure_home_map,
        )
    if visual_type == "waterfall":
        return _create_chart_container(
            section,
            visual_type="waterfallChart",
            x=x,
            y=y,
            width=width,
            height=height,
            title=title,
            projections={
                "Category": [_projection(spec["category"])],
                "Y": [_projection(spec["measure"])],
            },
            references=[spec["category"], spec["measure"]],
            measure_home_map=measure_home_map,
        )
    if visual_type == "slicer":
        return _make_visual_container(
            section=section,
            visual_type="slicer",
            x=x,
            y=y,
            width=width,
            height=height,
            projections={"Values": [_projection(spec["column"])]},
            references=[spec["column"]],
            measure_home_map=measure_home_map,
            extra_single_visual={"slicerType": str(spec.get("slicer_type", "dropdown")).casefold()},
        )
    if visual_type in {"text", "text_box"}:
        return _make_visual_container(
            section=section,
            visual_type="textbox",
            x=x,
            y=y,
            width=width,
            height=height,
            measure_home_map=measure_home_map,
            extra_single_visual={
                "textContent": spec["text"],
                "textStyle": {
                    "fontSize": int(spec.get("font_size", 16)),
                    "bold": bool(spec.get("bold", False)),
                    "color": str(spec.get("color", "#222222")),
                },
                "prototypeQuery": {"Version": 2, "From": [], "Select": []},
            },
        )
    if visual_type == "gauge":
        return _create_chart_container(
            section,
            visual_type="gauge",
            x=x,
            y=y,
            width=width,
            height=height,
            title=title,
            projections={"Y": [_projection(spec["measure"])]},
            references=[spec["measure"]],
            measure_home_map=measure_home_map,
        )
    if visual_type == "kpi":
        measures = [spec["measure"]]
        if spec.get("target_measure"):
            measures.append(spec["target_measure"])
        return _create_chart_container(
            section,
            visual_type="kpi",
            x=x,
            y=y,
            width=width,
            height=height,
            title=title,
            projections={
                "Value": [_projection(spec["measure"])],
                "Goal": [_projection(spec["target_measure"])],
            }
            if spec.get("target_measure")
            else {"Value": [_projection(spec["measure"])]},
            references=measures,
            measure_home_map=measure_home_map,
        )
    if visual_type == "map":
        refs = [spec["location"]]
        projections = {"Category": [_projection(spec["location"])]}
        if spec.get("measure"):
            refs.append(spec["measure"])
            projections["Y"] = [_projection(spec["measure"])]
        return _create_chart_container(
            section,
            visual_type="map",
            x=x,
            y=y,
            width=width,
            height=height,
            title=title,
            projections=projections,
            references=refs,
            measure_home_map=measure_home_map,
        )
    raise PowerBIValidationError("Unsupported dashboard visual type.", details={"type": visual_type})


def pbi_build_dashboard_tool(extract_folder: str, page: str, layout: list[dict[str, Any]]) -> dict[str, Any]:
    def _impl() -> dict[str, Any]:
        if not isinstance(layout, list):
            raise PowerBIValidationError("layout must be a list of visual specifications.")
        folder, report_layout = _load_layout(extract_folder)
        measure_home_map = _scan_measure_home_tables(folder)
        section = _find_page(report_layout, page)
        section.setdefault("visualContainers", [])
        created = []
        for item in layout:
            if not isinstance(item, dict):
                raise PowerBIValidationError("Each layout item must be an object.", details={"item": item})
            container = _create_visual_from_spec(section, item, measure_home_map)
            _assert_container_bindings(container, measure_home_map)
            section["visualContainers"].append(container)
            created.append(_visual_payload(container))
        _save_layout(folder, report_layout)
        return ok(
            f"Dashboard page '{section.get('displayName')}' updated successfully.",
            extract_folder=str(folder),
            page=_page_summary(section),
            created_visuals=created,
        )

    return _run(_impl)

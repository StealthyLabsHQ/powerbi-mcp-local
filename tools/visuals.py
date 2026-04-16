"""Report page and visual automation tools using pbi-tools and Layout JSON."""

from __future__ import annotations

import json
import os
import shutil
import subprocess
import uuid
from pathlib import Path
from typing import Any, Callable

from pbi_connection import PowerBIError, PowerBINotFoundError, PowerBIValidationError, error_payload, ok
from security import resolve_local_path

DEFAULT_PAGE_WIDTH = 1280
DEFAULT_PAGE_HEIGHT = 720
LAYOUT_RELATIVE_PATH = Path("Report") / "Layout"
THEMES_RELATIVE_DIR = Path("Report") / "StaticResources" / "Themes"
DEFAULT_VISUAL_SIZES = {
    "card": (200, 120),
    "bar_chart": (400, 300),
    "line_chart": (420, 300),
    "donut": (320, 280),
    "table": (520, 320),
    "waterfall": (420, 300),
    "slicer": (220, 120),
    "text": (280, 80),
    "gauge": (280, 220),
    "kpi": (260, 140),
    "map": (420, 320),
}


class VisualToolError(PowerBIError):
    code = "visual_error"


class PBIToolsNotInstalledError(VisualToolError):
    code = "pbi_tools_not_found"


class ReportLayoutError(VisualToolError):
    code = "report_layout_error"


class PageNotFoundError(VisualToolError):
    code = "report_page_not_found"


class VisualNotFoundError(VisualToolError):
    code = "report_visual_not_found"


def _run(callback: Callable[..., dict[str, Any]], *args: Any, **kwargs: Any) -> dict[str, Any]:
    try:
        return callback(*args, **kwargs)
    except Exception as exc:
        return error_payload(exc)


def _find_pbi_tools() -> str:
    custom = os.environ.get("PBI_TOOLS_PATH", "").strip()
    if custom:
        candidate = Path(custom).expanduser()
        if candidate.exists():
            return str(candidate)
        raise PBIToolsNotInstalledError(
            "PBI_TOOLS_PATH points to a missing executable.",
            details={"path": str(candidate)},
        )
    discovered = shutil.which("pbi-tools") or shutil.which("pbi-tools.exe")
    if discovered:
        return discovered
    raise PBIToolsNotInstalledError(
        "pbi-tools was not found on PATH. Install it with winget or dotnet tool install -g pbi-tools."
    )


def _run_pbi_tools(arguments: list[str]) -> dict[str, Any]:
    executable = _find_pbi_tools()
    try:
        completed = subprocess.run(
            [executable, *arguments],
            capture_output=True,
            text=True,
            check=False,
            shell=False,
        )
    except FileNotFoundError as exc:
        raise PBIToolsNotInstalledError("pbi-tools executable could not be launched.") from exc
    if completed.returncode != 0:
        raise VisualToolError(
            "pbi-tools command failed.",
            details={
                "command": [executable, *arguments],
                "returncode": completed.returncode,
                "stdout": completed.stdout[-2000:],
                "stderr": completed.stderr[-2000:],
            },
        )
    return {
        "stdout": completed.stdout,
        "stderr": completed.stderr,
        "returncode": completed.returncode,
    }


def _resolve_pbix_path(pbix_path: str, *, must_exist: bool) -> Path:
    return resolve_local_path(pbix_path, must_exist=must_exist, allowed_extensions={".pbix"})


def _resolve_extract_folder(extract_folder: str, *, must_exist: bool) -> Path:
    return resolve_local_path(extract_folder, must_exist=must_exist)


def _resolve_theme_path(theme_json_path: str) -> Path:
    return resolve_local_path(theme_json_path, must_exist=True, allowed_extensions={".json"})


def _layout_path(extract_folder: Path) -> Path:
    return extract_folder / LAYOUT_RELATIVE_PATH


def _load_layout(extract_folder: str | Path) -> tuple[Path, dict[str, Any]]:
    folder = _resolve_extract_folder(str(extract_folder), must_exist=True)
    if not folder.is_dir():
        raise ReportLayoutError("Extract folder does not exist or is not a directory.", details={"path": str(folder)})
    layout_path = _layout_path(folder)
    if not layout_path.exists():
        raise ReportLayoutError("Report/Layout file was not found in the extract folder.", details={"path": str(layout_path)})
    try:
        layout = json.loads(layout_path.read_text(encoding="utf-16-le"))
    except UnicodeDecodeError as exc:
        raise ReportLayoutError("Report/Layout could not be decoded as UTF-16-LE.", details={"path": str(layout_path)}) from exc
    except json.JSONDecodeError as exc:
        raise ReportLayoutError("Report/Layout is not valid JSON.", details={"path": str(layout_path), "line": exc.lineno}) from exc
    if not isinstance(layout, dict):
        raise ReportLayoutError("Report/Layout root must be a JSON object.", details={"path": str(layout_path)})
    layout.setdefault("sections", [])
    return folder, layout


def _save_layout(extract_folder: Path, layout: dict[str, Any]) -> None:
    layout_path = _layout_path(extract_folder)
    layout_path.parent.mkdir(parents=True, exist_ok=True)
    layout_path.write_text(json.dumps(layout, ensure_ascii=False, indent=2), encoding="utf-16-le")


def _parse_embedded_json(value: Any, default: Any) -> Any:
    if value in (None, ""):
        return default
    if isinstance(value, (dict, list)):
        return value
    if not isinstance(value, str):
        return default
    try:
        return json.loads(value)
    except json.JSONDecodeError:
        return default


def _dump_embedded_json(value: Any) -> str:
    return json.dumps(value, ensure_ascii=False, separators=(",", ":"))


def _normalize_page_name(display_name: str) -> str:
    cleaned = "".join(char for char in display_name if char.isalnum())
    return cleaned or "Page"


def _next_page_name(layout: dict[str, Any], display_name: str) -> str:
    existing = {str(section.get("name", "")) for section in layout.get("sections", [])}
    base = f"ReportSection{_normalize_page_name(display_name)}"
    if base not in existing:
        return base
    index = 1
    while f"{base}{index}" in existing:
        index += 1
    return f"{base}{index}"


def _find_page(layout: dict[str, Any], page: str) -> dict[str, Any]:
    wanted = page.casefold()
    for section in layout.get("sections", []):
        name = str(section.get("name", ""))
        display_name = str(section.get("displayName", ""))
        if name.casefold() == wanted or display_name.casefold() == wanted:
            return section
    raise PageNotFoundError(
        f"Page '{page}' was not found.",
        details={"page": page, "available_pages": [str(item.get("displayName") or item.get("name")) for item in layout.get("sections", [])]},
    )


def _page_summary(section: dict[str, Any]) -> dict[str, Any]:
    visuals = section.get("visualContainers", []) or []
    return {
        "name": str(section.get("name", "")),
        "display_name": str(section.get("displayName", "")),
        "width": int(section.get("width", DEFAULT_PAGE_WIDTH)),
        "height": int(section.get("height", DEFAULT_PAGE_HEIGHT)),
        "visual_count": len(visuals),
    }


def _split_column_ref(reference: str) -> tuple[str, str]:
    if "." not in reference:
        raise PowerBIValidationError(
            "Column references must use 'TableName.ColumnName' format.",
            details={"reference": reference},
        )
    table, column = reference.rsplit(".", 1)
    if not table.strip() or not column.strip():
        raise PowerBIValidationError(
            "Column references must include both a table and a column name.",
            details={"reference": reference},
        )
    return table.strip(), column.strip()


def _unique_visual_id() -> str:
    return uuid.uuid4().hex[:20]


def _validate_dimensions(x: int, y: int, width: int, height: int) -> None:
    if min(x, y) < 0:
        raise PowerBIValidationError("x and y must be >= 0.", details={"x": x, "y": y})
    if width <= 0 or height <= 0:
        raise PowerBIValidationError("width and height must be > 0.", details={"width": width, "height": height})


def _page_next_z(section: dict[str, Any]) -> int:
    z_values = [int(container.get("z", 0)) for container in section.get("visualContainers", []) if isinstance(container, dict)]
    return (max(z_values) + 1) if z_values else 0


def _build_select_entry(reference: str, aliases: dict[str, str]) -> dict[str, Any]:
    if "." in reference:
        table, column = _split_column_ref(reference)
        alias = aliases.setdefault(table, f"s{len(aliases)}")
        return {
            "Column": {"Expression": {"SourceRef": {"Source": alias}}, "Property": column},
            "Name": reference,
            "NativeReferenceName": column,
        }
    alias = aliases.setdefault("$Measures", f"s{len(aliases)}")
    return {
        "Measure": {"Expression": {"SourceRef": {"Source": alias}}, "Property": reference},
        "Name": reference,
        "NativeReferenceName": reference,
    }


def _build_prototype_query(references: list[str]) -> dict[str, Any]:
    aliases: dict[str, str] = {}
    select = [_build_select_entry(reference, aliases) for reference in references]
    from_entries = [{"Name": alias, "Entity": entity} for entity, alias in aliases.items()]
    return {"Version": 2, "From": from_entries, "Select": select}


def _literal_value(value: Any) -> dict[str, Any]:
    return {"expr": {"Literal": {"Value": json.dumps(value)}}}


def _title_objects(title: str) -> dict[str, Any]:
    return {
        "title": [
            {
                "properties": {
                    "show": _literal_value(True),
                    "text": _literal_value(title),
                }
            }
        ]
    }


def _base_visual_config(
    *,
    visual_id: str,
    visual_type: str,
    x: int,
    y: int,
    width: int,
    height: int,
    references: list[str] | None = None,
    projections: dict[str, list[dict[str, str]]] | None = None,
    title: str | None = None,
    extra_single_visual: dict[str, Any] | None = None,
) -> tuple[dict[str, Any], dict[str, Any]]:
    position = {"x": x, "y": y, "width": width, "height": height}
    single_visual = {
        "visualType": visual_type,
        "projections": projections or {},
        "prototypeQuery": _build_prototype_query(references or []),
        "objects": {},
    }
    if title:
        single_visual["objects"].update(_title_objects(title))
    if extra_single_visual:
        single_visual.update(extra_single_visual)
    config = {
        "name": visual_id,
        "layouts": [{"id": 0, "position": position}],
        "singleVisual": single_visual,
    }
    query = {
        "Commands": [
            {
                "SemanticQueryDataShapeCommand": {
                    "Query": single_visual["prototypeQuery"],
                }
            }
        ]
    }
    return config, query


def _make_visual_container(
    *,
    section: dict[str, Any],
    visual_type: str,
    x: int,
    y: int,
    width: int,
    height: int,
    references: list[str] | None = None,
    projections: dict[str, list[dict[str, str]]] | None = None,
    title: str | None = None,
    filters: Any | None = None,
    extra_single_visual: dict[str, Any] | None = None,
) -> dict[str, Any]:
    _validate_dimensions(x, y, width, height)
    visual_id = _unique_visual_id()
    config, query = _base_visual_config(
        visual_id=visual_id,
        visual_type=visual_type,
        x=x,
        y=y,
        width=width,
        height=height,
        references=references,
        projections=projections,
        title=title,
        extra_single_visual=extra_single_visual,
    )
    return {
        "x": x,
        "y": y,
        "z": _page_next_z(section),
        "width": width,
        "height": height,
        "config": _dump_embedded_json(config),
        "filters": _dump_embedded_json(filters if filters is not None else []),
        "query": _dump_embedded_json(query),
        "dataTransforms": _dump_embedded_json({}),
    }


def _visual_payload(container: dict[str, Any]) -> dict[str, Any]:
    config = _parse_embedded_json(container.get("config"), {})
    single_visual = config.get("singleVisual", {}) if isinstance(config, dict) else {}
    title = None
    text = None
    objects = single_visual.get("objects", {}) if isinstance(single_visual, dict) else {}
    title_entries = objects.get("title", [])
    if title_entries:
        title = (
            title_entries[0]
            .get("properties", {})
            .get("text", {})
            .get("expr", {})
            .get("Literal", {})
            .get("Value")
        )
    if isinstance(single_visual, dict) and "textContent" in single_visual:
        text = single_visual.get("textContent")
    return {
        "id": str(config.get("name") or ""),
        "type": str(single_visual.get("visualType") or "unknown"),
        "x": int(container.get("x", 0)),
        "y": int(container.get("y", 0)),
        "z": int(container.get("z", 0)),
        "width": int(container.get("width", 0)),
        "height": int(container.get("height", 0)),
        "data": {
            "title": title,
            "text": text,
            "projections": single_visual.get("projections", {}),
        },
    }


def _find_visual(section: dict[str, Any], visual_id: str) -> tuple[int, dict[str, Any], dict[str, Any]]:
    for index, container in enumerate(section.get("visualContainers", []) or []):
        config = _parse_embedded_json(container.get("config"), {})
        if str(config.get("name", "")).casefold() == visual_id.casefold():
            return index, container, config
    raise VisualNotFoundError(
        f"Visual '{visual_id}' was not found on page '{section.get('displayName') or section.get('name')}'.",
        details={"visual_id": visual_id},
    )


def _append_visual(extract_folder: str, page: str, factory: Callable[[dict[str, Any]], dict[str, Any]]) -> dict[str, Any]:
    folder, layout = _load_layout(extract_folder)
    section = _find_page(layout, page)
    section.setdefault("visualContainers", [])
    container = factory(section)
    section["visualContainers"].append(container)
    _save_layout(folder, layout)
    visual = _visual_payload(container)
    return ok(
        f"Visual '{visual['id']}' added to page '{section.get('displayName')}'.",
        page=_page_summary(section),
        visual=visual,
    )


def _create_chart_container(
    section: dict[str, Any],
    *,
    visual_type: str,
    x: int,
    y: int,
    width: int,
    height: int,
    title: str | None,
    projections: dict[str, list[dict[str, str]]],
    references: list[str],
    extra_single_visual: dict[str, Any] | None = None,
) -> dict[str, Any]:
    return _make_visual_container(
        section=section,
        visual_type=visual_type,
        x=x,
        y=y,
        width=width,
        height=height,
        title=title,
        projections=projections,
        references=references,
        extra_single_visual=extra_single_visual,
    )


def pbi_extract_report_tool(pbix_path: str, extract_folder: str | None = None) -> dict[str, Any]:
    def _impl() -> dict[str, Any]:
        pbix = _resolve_pbix_path(pbix_path, must_exist=True)
        target = _resolve_extract_folder(str(extract_folder or pbix.with_name(f"{pbix.stem}_extracted")), must_exist=False)
        target.parent.mkdir(parents=True, exist_ok=True)
        _run_pbi_tools(["extract", str(pbix), "-extractFolder", str(target)])
        _, layout = _load_layout(target)
        pages = [_page_summary(section) for section in layout.get("sections", [])]
        return ok(
            "Report extracted successfully.",
            pbix_path=str(pbix),
            extract_folder=str(target),
            pages=pages,
            visual_count=sum(page["visual_count"] for page in pages),
        )

    return _run(_impl)


def pbi_compile_report_tool(extract_folder: str, output_path: str) -> dict[str, Any]:
    def _impl() -> dict[str, Any]:
        folder = _resolve_extract_folder(extract_folder, must_exist=True)
        output = _resolve_pbix_path(output_path, must_exist=False)
        output.parent.mkdir(parents=True, exist_ok=True)
        _run_pbi_tools(["compile", str(folder), "-outPath", str(output)])
        return ok(
            "Report compiled successfully.",
            extract_folder=str(folder),
            output_path=str(output),
            size_bytes=output.stat().st_size if output.exists() else None,
        )

    return _run(_impl)


def pbi_list_pages_tool(extract_folder: str) -> dict[str, Any]:
    def _impl() -> dict[str, Any]:
        _, layout = _load_layout(extract_folder)
        pages = [_page_summary(section) for section in layout.get("sections", [])]
        return ok("Pages listed successfully.", extract_folder=str(_resolve_extract_folder(extract_folder, must_exist=True)), pages=pages)

    return _run(_impl)


def pbi_get_page_tool(extract_folder: str, page: str) -> dict[str, Any]:
    def _impl() -> dict[str, Any]:
        _, layout = _load_layout(extract_folder)
        section = _find_page(layout, page)
        visuals = [_visual_payload(container) for container in section.get("visualContainers", []) or []]
        payload = _page_summary(section)
        payload["visuals"] = visuals
        return ok("Page retrieved successfully.", extract_folder=str(_resolve_extract_folder(extract_folder, must_exist=True)), page=payload)

    return _run(_impl)


def pbi_create_page_tool(extract_folder: str, display_name: str, width: int = DEFAULT_PAGE_WIDTH, height: int = DEFAULT_PAGE_HEIGHT) -> dict[str, Any]:
    def _impl() -> dict[str, Any]:
        _validate_dimensions(0, 0, width, height)
        folder, layout = _load_layout(extract_folder)
        section = {
            "name": _next_page_name(layout, display_name),
            "displayName": display_name,
            "displayOption": 0,
            "width": width,
            "height": height,
            "visualContainers": [],
            "filters": "[]",
        }
        if any("ordinal" in item for item in layout.get("sections", [])):
            section["ordinal"] = len(layout.get("sections", []))
        layout.setdefault("sections", []).append(section)
        _save_layout(folder, layout)
        return ok("Page created successfully.", extract_folder=str(folder), page=_page_summary(section))

    return _run(_impl)


def pbi_delete_page_tool(extract_folder: str, page: str) -> dict[str, Any]:
    def _impl() -> dict[str, Any]:
        folder, layout = _load_layout(extract_folder)
        sections = layout.get("sections", [])
        if len(sections) <= 1:
            raise PowerBIValidationError("Cannot delete the last remaining page.")
        section = _find_page(layout, page)
        layout["sections"] = [item for item in sections if item is not section]
        _save_layout(folder, layout)
        return ok("Page deleted successfully.", extract_folder=str(folder), deleted_page=str(section.get("displayName") or section.get("name")))

    return _run(_impl)


def pbi_set_page_size_tool(extract_folder: str, page: str, width: int, height: int) -> dict[str, Any]:
    def _impl() -> dict[str, Any]:
        _validate_dimensions(0, 0, width, height)
        folder, layout = _load_layout(extract_folder)
        section = _find_page(layout, page)
        section["width"] = width
        section["height"] = height
        _save_layout(folder, layout)
        return ok("Page size updated successfully.", extract_folder=str(folder), page=_page_summary(section))

    return _run(_impl)


def pbi_add_card_tool(extract_folder: str, page: str, measure: str, x: int, y: int, width: int = 200, height: int = 120, title: str = "") -> dict[str, Any]:
    return _append_visual(
        extract_folder,
        page,
        lambda section: _create_chart_container(
            section,
            visual_type="card",
            x=x,
            y=y,
            width=width,
            height=height,
            title=title or None,
            projections={"Values": [{"queryRef": measure}]},
            references=[measure],
        ),
    )


def pbi_add_bar_chart_tool(
    extract_folder: str,
    page: str,
    category_column: str,
    value_measure: str,
    x: int,
    y: int,
    width: int = 400,
    height: int = 300,
    title: str = "",
    legend_column: str | None = None,
) -> dict[str, Any]:
    projections = {"Category": [{"queryRef": category_column}], "Y": [{"queryRef": value_measure}]}
    references = [category_column, value_measure]
    if legend_column:
        projections["Series"] = [{"queryRef": legend_column}]
        references.append(legend_column)
    return _append_visual(
        extract_folder,
        page,
        lambda section: _create_chart_container(
            section,
            visual_type="clusteredBarChart",
            x=x,
            y=y,
            width=width,
            height=height,
            title=title or None,
            projections=projections,
            references=references,
        ),
    )


def pbi_add_line_chart_tool(
    extract_folder: str,
    page: str,
    axis_column: str,
    value_measures: list[str],
    x: int,
    y: int,
    width: int = 420,
    height: int = 300,
    title: str = "",
) -> dict[str, Any]:
    if not value_measures:
        raise PowerBIValidationError("value_measures must contain at least one measure.")
    return _append_visual(
        extract_folder,
        page,
        lambda section: _create_chart_container(
            section,
            visual_type="lineChart",
            x=x,
            y=y,
            width=width,
            height=height,
            title=title or None,
            projections={
                "Category": [{"queryRef": axis_column}],
                "Y": [{"queryRef": measure} for measure in value_measures],
            },
            references=[axis_column, *value_measures],
        ),
    )


def pbi_add_donut_chart_tool(extract_folder: str, page: str, category_column: str, value_measure: str, x: int, y: int, width: int = 320, height: int = 280, title: str = "") -> dict[str, Any]:
    return _append_visual(
        extract_folder,
        page,
        lambda section: _create_chart_container(
            section,
            visual_type="donutChart",
            x=x,
            y=y,
            width=width,
            height=height,
            title=title or None,
            projections={"Category": [{"queryRef": category_column}], "Y": [{"queryRef": value_measure}]},
            references=[category_column, value_measure],
        ),
    )


def pbi_add_table_visual_tool(extract_folder: str, page: str, columns: list[str], x: int, y: int, width: int = 520, height: int = 320, title: str = "") -> dict[str, Any]:
    if not columns:
        raise PowerBIValidationError("columns must contain at least one field or measure.")
    return _append_visual(
        extract_folder,
        page,
        lambda section: _create_chart_container(
            section,
            visual_type="tableEx",
            x=x,
            y=y,
            width=width,
            height=height,
            title=title or None,
            projections={"Values": [{"queryRef": item} for item in columns]},
            references=list(columns),
        ),
    )


def pbi_add_waterfall_tool(extract_folder: str, page: str, category_column: str, value_measure: str, x: int, y: int, width: int = 420, height: int = 300, title: str = "") -> dict[str, Any]:
    return _append_visual(
        extract_folder,
        page,
        lambda section: _create_chart_container(
            section,
            visual_type="waterfallChart",
            x=x,
            y=y,
            width=width,
            height=height,
            title=title or None,
            projections={"Category": [{"queryRef": category_column}], "Y": [{"queryRef": value_measure}]},
            references=[category_column, value_measure],
        ),
    )


def pbi_add_slicer_tool(extract_folder: str, page: str, column: str, x: int, y: int, width: int = 220, height: int = 120, slicer_type: str = "dropdown") -> dict[str, Any]:
    slicer_kind = slicer_type.strip().casefold()
    if slicer_kind not in {"dropdown", "list", "range"}:
        raise PowerBIValidationError("slicer_type must be one of: dropdown, list, range.", details={"slicer_type": slicer_type})
    return _append_visual(
        extract_folder,
        page,
        lambda section: _create_chart_container(
            section,
            visual_type="slicer",
            x=x,
            y=y,
            width=width,
            height=height,
            title=None,
            projections={"Values": [{"queryRef": column}]},
            references=[column],
            extra_single_visual={"slicerType": slicer_kind},
        ),
    )


def pbi_add_gauge_tool(extract_folder: str, page: str, measure: str, x: int, y: int, width: int = 280, height: int = 220, title: str = "", target_measure: str | None = None) -> dict[str, Any]:
    projections = {"Value": [{"queryRef": measure}]}
    references = [measure]
    if target_measure:
        projections["Goal"] = [{"queryRef": target_measure}]
        references.append(target_measure)
    return _append_visual(
        extract_folder,
        page,
        lambda section: _create_chart_container(
            section,
            visual_type="gauge",
            x=x,
            y=y,
            width=width,
            height=height,
            title=title or None,
            projections=projections,
            references=references,
        ),
    )


def pbi_add_text_box_tool(
    extract_folder: str,
    page: str,
    text: str,
    x: int,
    y: int,
    width: int = 280,
    height: int = 80,
    font_size: int = 16,
    bold: bool = False,
    color: str = "#222222",
) -> dict[str, Any]:
    return _append_visual(
        extract_folder,
        page,
        lambda section: _make_visual_container(
            section=section,
            visual_type="textbox",
            x=x,
            y=y,
            width=width,
            height=height,
            references=[],
            projections={},
            extra_single_visual={
                "textContent": text,
                "textStyle": {"fontSize": font_size, "bold": bold, "color": color},
                "prototypeQuery": {"Version": 2, "From": [], "Select": []},
                "objects": {"paragraphs": [{"text": text, "fontSize": font_size, "bold": bold, "color": color}]},
            },
        ),
    )


def pbi_remove_visual_tool(extract_folder: str, page: str, visual_id: str) -> dict[str, Any]:
    def _impl() -> dict[str, Any]:
        folder, layout = _load_layout(extract_folder)
        section = _find_page(layout, page)
        index, _, _ = _find_visual(section, visual_id)
        removed = section["visualContainers"].pop(index)
        _save_layout(folder, layout)
        return ok("Visual removed successfully.", extract_folder=str(folder), page=str(section.get("displayName") or section.get("name")), visual=_visual_payload(removed))

    return _run(_impl)


def pbi_move_visual_tool(extract_folder: str, page: str, visual_id: str, x: int, y: int, width: int | None = None, height: int | None = None) -> dict[str, Any]:
    def _impl() -> dict[str, Any]:
        folder, layout = _load_layout(extract_folder)
        section = _find_page(layout, page)
        _, container, config = _find_visual(section, visual_id)
        new_width = width if width is not None else int(container.get("width", 0))
        new_height = height if height is not None else int(container.get("height", 0))
        _validate_dimensions(x, y, new_width, new_height)
        container.update({"x": x, "y": y, "width": new_width, "height": new_height})
        layouts = config.get("layouts", [])
        if layouts:
            layouts[0].setdefault("position", {})
            layouts[0]["position"].update({"x": x, "y": y, "width": new_width, "height": new_height})
        container["config"] = _dump_embedded_json(config)
        _save_layout(folder, layout)
        return ok("Visual moved successfully.", extract_folder=str(folder), page=str(section.get("displayName") or section.get("name")), visual=_visual_payload(container))

    return _run(_impl)


def pbi_apply_theme_tool(extract_folder: str, theme_json_path: str) -> dict[str, Any]:
    def _impl() -> dict[str, Any]:
        folder, layout = _load_layout(extract_folder)
        theme_path = _resolve_theme_path(theme_json_path)
        try:
            theme_payload = json.loads(theme_path.read_text(encoding="utf-8"))
        except json.JSONDecodeError as exc:
            raise PowerBIValidationError("Theme JSON is invalid.", details={"path": str(theme_path), "line": exc.lineno}) from exc
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
        return ok("Theme applied successfully.", extract_folder=str(folder), theme=theme_entry)

    return _run(_impl)


def _create_visual_from_spec(section: dict[str, Any], spec: dict[str, Any]) -> dict[str, Any]:
    visual_type = str(spec.get("type", "")).strip().casefold()
    x = int(spec.get("x", 0))
    y = int(spec.get("y", 0))
    width = int(spec.get("width", DEFAULT_VISUAL_SIZES.get(visual_type, (400, 300))[0]))
    height = int(spec.get("height", DEFAULT_VISUAL_SIZES.get(visual_type, (400, 300))[1]))
    title = spec.get("title")
    if visual_type == "card":
        return _create_chart_container(section, visual_type="card", x=x, y=y, width=width, height=height, title=title, projections={"Values": [{"queryRef": spec["measure"]}]}, references=[spec["measure"]])
    if visual_type in {"bar_chart", "bar"}:
        projections = {"Category": [{"queryRef": spec["category"]}], "Y": [{"queryRef": spec["measure"]}]}
        references = [spec["category"], spec["measure"]]
        if spec.get("legend"):
            projections["Series"] = [{"queryRef": spec["legend"]}]
            references.append(spec["legend"])
        return _create_chart_container(section, visual_type="clusteredBarChart", x=x, y=y, width=width, height=height, title=title, projections=projections, references=references)
    if visual_type in {"line_chart", "line"}:
        measures = list(spec.get("measures") or [spec.get("measure")])
        return _create_chart_container(section, visual_type="lineChart", x=x, y=y, width=width, height=height, title=title, projections={"Category": [{"queryRef": spec["axis"]}], "Y": [{"queryRef": item} for item in measures]}, references=[spec["axis"], *measures])
    if visual_type in {"donut", "donut_chart", "pie", "pie_chart"}:
        return _create_chart_container(section, visual_type="donutChart", x=x, y=y, width=width, height=height, title=title, projections={"Category": [{"queryRef": spec["category"]}], "Y": [{"queryRef": spec["measure"]}]}, references=[spec["category"], spec["measure"]])
    if visual_type in {"table", "table_visual"}:
        return _create_chart_container(section, visual_type="tableEx", x=x, y=y, width=width, height=height, title=title, projections={"Values": [{"queryRef": item} for item in spec["columns"]]}, references=list(spec["columns"]))
    if visual_type == "waterfall":
        return _create_chart_container(section, visual_type="waterfallChart", x=x, y=y, width=width, height=height, title=title, projections={"Category": [{"queryRef": spec["category"]}], "Y": [{"queryRef": spec["measure"]}]}, references=[spec["category"], spec["measure"]])
    if visual_type == "slicer":
        return _make_visual_container(section=section, visual_type="slicer", x=x, y=y, width=width, height=height, projections={"Values": [{"queryRef": spec["column"]}]}, references=[spec["column"]], extra_single_visual={"slicerType": str(spec.get("slicer_type", "dropdown")).casefold()})
    if visual_type in {"text", "text_box"}:
        return _make_visual_container(section=section, visual_type="textbox", x=x, y=y, width=width, height=height, extra_single_visual={"textContent": spec["text"], "textStyle": {"fontSize": int(spec.get("font_size", 16)), "bold": bool(spec.get("bold", False)), "color": str(spec.get("color", "#222222"))}, "prototypeQuery": {"Version": 2, "From": [], "Select": []}})
    if visual_type == "gauge":
        return _create_chart_container(section, visual_type="gauge", x=x, y=y, width=width, height=height, title=title, projections={"Value": [{"queryRef": spec["measure"]}]}, references=[spec["measure"]])
    if visual_type == "kpi":
        measures = [spec["measure"]]
        if spec.get("target_measure"):
            measures.append(spec["target_measure"])
        return _create_chart_container(section, visual_type="kpi", x=x, y=y, width=width, height=height, title=title, projections={"Value": [{"queryRef": spec["measure"]}], "Goal": [{"queryRef": spec["target_measure"]}]} if spec.get("target_measure") else {"Value": [{"queryRef": spec["measure"]}]}, references=measures)
    if visual_type == "map":
        refs = [spec["location"]]
        projections = {"Category": [{"queryRef": spec["location"]}]}
        if spec.get("measure"):
            refs.append(spec["measure"])
            projections["Y"] = [{"queryRef": spec["measure"]}]
        return _create_chart_container(section, visual_type="map", x=x, y=y, width=width, height=height, title=title, projections=projections, references=refs)
    raise PowerBIValidationError("Unsupported dashboard visual type.", details={"type": visual_type})


def pbi_build_dashboard_tool(extract_folder: str, page: str, layout: list[dict[str, Any]]) -> dict[str, Any]:
    def _impl() -> dict[str, Any]:
        if not isinstance(layout, list):
            raise PowerBIValidationError("layout must be a list of visual specifications.")
        folder, report_layout = _load_layout(extract_folder)
        section = _find_page(report_layout, page)
        section.setdefault("visualContainers", [])
        created = []
        for item in layout:
            if not isinstance(item, dict):
                raise PowerBIValidationError("Each layout item must be an object.", details={"item": item})
            container = _create_visual_from_spec(section, item)
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


__all__ = [
    "pbi_add_bar_chart_tool",
    "pbi_add_card_tool",
    "pbi_add_donut_chart_tool",
    "pbi_add_gauge_tool",
    "pbi_add_line_chart_tool",
    "pbi_add_slicer_tool",
    "pbi_add_table_visual_tool",
    "pbi_add_text_box_tool",
    "pbi_add_waterfall_tool",
    "pbi_apply_theme_tool",
    "pbi_build_dashboard_tool",
    "pbi_compile_report_tool",
    "pbi_create_page_tool",
    "pbi_delete_page_tool",
    "pbi_extract_report_tool",
    "pbi_get_page_tool",
    "pbi_list_pages_tool",
    "pbi_move_visual_tool",
    "pbi_remove_visual_tool",
    "pbi_set_page_size_tool",
]

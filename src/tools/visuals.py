"""Report page and visual automation tools using pbi-tools and Layout JSON."""

from __future__ import annotations

import json
import logging
import os
import re
import shutil
import subprocess
import tempfile
import time
import uuid
import zipfile
from pathlib import Path
from typing import Any, Callable

from pbi_connection import PowerBIError, PowerBINotFoundError, PowerBIValidationError, error_payload, ok
from security import SECURITY, resolve_local_path
from .model import pbi_model_info_tool

DEFAULT_PAGE_WIDTH = 1280
DEFAULT_PAGE_HEIGHT = 720
LAYOUT_RELATIVE_PATH = Path("Report") / "Layout"
THEMES_RELATIVE_DIR = Path("Report") / "StaticResources" / "Themes"
DESIGN_THEME_RELATIVE_PATH = Path("Report") / "StaticResources" / "SharedResources" / "BaseThemes" / "CY26SU02.json"
MODEL_TABLES_RELATIVE_DIR = Path("Model") / "tables"
HEX_COLOR_RE = re.compile(r"^#[0-9A-Fa-f]{6}$")
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
VISUAL_FIELD_ROLES = {
    "card": {"Values"},
    "clusteredBarChart": {"Category", "Y", "Series"},
    "lineChart": {"Category", "Y"},
    "donutChart": {"Category", "Y"},
    "tableEx": {"Values"},
    "waterfallChart": {"Category", "Y"},
    "slicer": {"Values"},
    "gauge": {"Y", "Goal"},
    "kpi": {"Value", "Goal"},
    "map": {"Category", "Y"},
}

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
                    "title": [{"show": True, "fontColor": {"solid": {"color": "#1E40AF"}}, "background": {"solid": {"color": "#FFFFFF"}}, "fontSize": 12, "fontFamily": "Segoe UI Semibold"}],
                    "lineStyles": [{"strokeWidth": 3}],
                    "categoryAxis": [{"showAxisTitle": False, "gridlineStyle": "dotted", "gridlineColor": {"solid": {"color": "#E2E8F0"}}}],
                    "valueAxis": [{"showAxisTitle": False, "gridlineStyle": "dotted", "gridlineColor": {"solid": {"color": "#E2E8F0"}}}],
                }
            },
            "card": {
                "*": {
                    "labels": [{"color": {"solid": {"color": "#1E293B"}}, "fontSize": 22, "fontBold": True, "fontFamily": "Segoe UI Semibold"}],
                    "categoryLabels": [{"color": {"solid": {"color": "#475569"}}, "fontSize": 11, "fontFamily": "Segoe UI"}],
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
                    "columnHeaders": [{"fontColor": {"solid": {"color": "#1E40AF"}}, "backColor": {"solid": {"color": "#EFF6FF"}}, "fontSize": 11, "fontBold": True}],
                    "values": [{"fontColor": {"solid": {"color": "#1E293B"}}, "backColor": {"solid": {"color": "#FFFFFF"}}, "altBackColor": {"solid": {"color": "#F8FAFC"}}, "fontSize": 10}],
                }
            },
        },
    }
}

logger = logging.getLogger(__name__)


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
    discovered = shutil.which("pbi-tools") or shutil.which("pbi-tools.exe") or shutil.which("pbi-tools.core.exe")
    if discovered:
        return discovered
    # Fallback: check common install locations
    fallback_paths = [
        Path.home() / "AppData" / "Local" / "pbi-tools" / "full" / "pbi-tools.exe",
        Path.home() / "AppData" / "Local" / "pbi-tools" / "pbi-tools.core.exe",
    ]
    for fallback in fallback_paths:
        if fallback.exists():
            return str(fallback)
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


def _query_ref(reference: str) -> str:
    """Return the short queryRef name (column part only, without table prefix)."""
    return reference.split(".", 1)[1] if "." in reference else reference


def _scan_measure_home_tables(extract_folder: Path) -> dict[str, str]:
    """Map measure name -> home table from extract metadata folders."""
    table_root = extract_folder / MODEL_TABLES_RELATIVE_DIR
    if not table_root.is_dir():
        return {}

    measure_home_map: dict[str, str] = {}
    for table_dir in table_root.iterdir():
        if not table_dir.is_dir():
            continue
        measures_dir = table_dir / "measures"
        if not measures_dir.is_dir():
            continue
        for dax_file in measures_dir.glob("*.dax"):
            measure_name = dax_file.stem.strip()
            if not measure_name:
                continue
            existing = measure_home_map.get(measure_name)
            if existing and existing != table_dir.name:
                logger.warning(
                    "Measure '%s' found in multiple tables ('%s', '%s'); keeping first.",
                    measure_name,
                    existing,
                    table_dir.name,
                )
                continue
            measure_home_map[measure_name] = table_dir.name
    return measure_home_map


def _build_select_entry(
    reference: str,
    aliases: dict[str, str],
    measure_home_map: dict[str, str] | None = None,
) -> dict[str, Any]:
    if "." in reference:
        table, column = _split_column_ref(reference)
        alias = aliases.setdefault(table, f"s{len(aliases)}")
        return {
            "Column": {"Expression": {"SourceRef": {"Source": alias}}, "Property": column},
            "Name": column,  # PBI expects short name without table prefix
            "NativeReferenceName": column,
        }
    measure_entity = (measure_home_map or {}).get(reference) or "$Measures"
    if measure_entity == "$Measures":
        logger.warning(
            "Measure '%s' home table not found in extract metadata; using '$Measures' fallback.",
            reference,
        )
    alias = aliases.setdefault(measure_entity, f"s{len(aliases)}")
    return {
        "Measure": {"Expression": {"SourceRef": {"Source": alias}}, "Property": reference},
        "Name": reference,
        "NativeReferenceName": reference,
    }


def _build_prototype_query(
    references: list[str],
    measure_home_map: dict[str, str] | None = None,
) -> dict[str, Any]:
    aliases: dict[str, str] = {}
    select = [_build_select_entry(reference, aliases, measure_home_map) for reference in references]
    from_entries = [{"Name": alias, "Entity": entity} for entity, alias in aliases.items()]
    return {"Version": 2, "From": from_entries, "Select": select}


def _select_name_map(prototype_query: dict[str, Any]) -> dict[str, str]:
    names: dict[str, str] = {}
    for entry in prototype_query.get("Select", []) or []:
        if not isinstance(entry, dict):
            continue
        name = str(entry.get("Name", ""))
        if not name:
            continue
        if "Column" in entry:
            column = entry.get("Column", {})
            if isinstance(column, dict):
                prop = str(column.get("Property", ""))
                if prop:
                    names[prop.casefold()] = name
        if "Measure" in entry:
            measure = entry.get("Measure", {})
            if isinstance(measure, dict):
                prop = str(measure.get("Property", ""))
                if prop:
                    names[prop.casefold()] = name
        names[name.casefold()] = name
    return names


def _from_entity_by_alias(prototype_query: dict[str, Any]) -> dict[str, str]:
    entities: dict[str, str] = {}
    for entry in prototype_query.get("From", []) or []:
        if isinstance(entry, dict):
            entities[str(entry.get("Name", ""))] = str(entry.get("Entity", ""))
    return entities


def _next_alias(existing: set[str]) -> str:
    index = 0
    while f"s{index}" in existing:
        index += 1
    alias = f"s{index}"
    existing.add(alias)
    return alias


def _sync_container_query(container: dict[str, Any], prototype_query: dict[str, Any]) -> None:
    query_payload = _parse_embedded_json(container.get("query"), {})
    try:
        commands = query_payload.setdefault("Commands", [])
        if not commands:
            commands.append({"SemanticQueryDataShapeCommand": {}})
        commands[0].setdefault("SemanticQueryDataShapeCommand", {})["Query"] = prototype_query
        container["query"] = _dump_embedded_json(query_payload)
    except Exception:
        container["query"] = _dump_embedded_json(
            {"Commands": [{"SemanticQueryDataShapeCommand": {"Query": prototype_query}}]}
        )


def _persistence_risks(issues: list[dict[str, Any]]) -> list[dict[str, Any]]:
    return [
        item for item in issues
        if item.get("source") == "live_model"
        and item.get("extract_metadata") == "missing"
    ]


def _live_model_field_index(manager: Any | None, *, include_hidden: bool) -> tuple[dict[str, Any] | None, dict[str, Any]]:
    if manager is None:
        return None, {"status": "unavailable", "reason": "manager_not_provided"}
    try:
        model = pbi_model_info_tool(manager, include_hidden=include_hidden, include_row_counts=False)
    except Exception as exc:
        return None, {"status": "unavailable", "error": error_payload(exc)["error"]}
    if not model.get("ok"):
        return None, {"status": "unavailable", "error": model.get("error")}

    columns: set[tuple[str, str]] = set()
    measures: dict[str, set[str]] = {}
    measure_tables: dict[str, set[str]] = {}
    for table in model.get("tables", []) or []:
        table_name = str(table.get("name", ""))
        for column in table.get("columns", []) or []:
            columns.add((table_name.casefold(), str(column.get("name", "")).casefold()))
    for measure in model.get("measures", []) or []:
        name = str(measure.get("name", ""))
        table_name = str(measure.get("table", ""))
        measures.setdefault(name.casefold(), set()).add(table_name.casefold())
        measure_tables.setdefault(name.casefold(), set()).add(table_name)
    return {"columns": columns, "measures": measures, "measure_tables": measure_tables}, {"status": "available"}


def _visual_binding_issues(
    container: dict[str, Any],
    page_name: str,
    measure_home_map: dict[str, str],
    model_fields: dict[str, Any] | None = None,
    *,
    repair: bool = False,
) -> tuple[list[dict[str, Any]], int]:
    config = _parse_embedded_json(container.get("config"), {})
    if not isinstance(config, dict):
        return ([{"page": page_name, "visual_id": "", "issue": "invalid_config"}], 0)
    single_visual = config.get("singleVisual", {})
    if not isinstance(single_visual, dict):
        return ([], 0)
    visual_id = str(config.get("name", ""))
    visual_type = str(single_visual.get("visualType", ""))
    prototype_query = single_visual.get("prototypeQuery", {})
    if not isinstance(prototype_query, dict):
        return ([], 0)

    issues: list[dict[str, Any]] = []
    repairs = 0
    select_names = _select_name_map(prototype_query)
    from_entities = _from_entity_by_alias(prototype_query)

    allowed_roles = VISUAL_FIELD_ROLES.get(visual_type)
    projections = single_visual.get("projections", {})
    if isinstance(projections, dict):
        if repair and visual_type == "gauge" and "Value" in projections and "Y" not in projections:
            projections["Y"] = projections.pop("Value")
            issues.append({"page": page_name, "visual_id": visual_id, "visual_type": visual_type, "issue": "projection_role_repaired", "from": "Value", "to": "Y"})
            repairs += 1
        for role, items in list(projections.items()):
            if allowed_roles is not None and role not in allowed_roles:
                issues.append({"page": page_name, "visual_id": visual_id, "visual_type": visual_type, "issue": "unexpected_projection_role", "role": role, "allowed_roles": sorted(allowed_roles)})
            if not isinstance(items, list):
                continue
            for item in items:
                if not isinstance(item, dict):
                    continue
                query_ref = str(item.get("queryRef", ""))
                expected = select_names.get(query_ref.casefold())
                if expected is None:
                    short = _query_ref(query_ref)
                    expected = select_names.get(short.casefold())
                if expected is None:
                    issues.append({"page": page_name, "visual_id": visual_id, "visual_type": visual_type, "issue": "query_ref_not_found", "queryRef": query_ref})
                    continue
                if query_ref != expected:
                    issues.append({"page": page_name, "visual_id": visual_id, "visual_type": visual_type, "issue": "query_ref_mismatch", "queryRef": query_ref, "expected": expected})
                    if repair:
                        item["queryRef"] = expected
                        repairs += 1

    from_entries = prototype_query.get("From", []) or []
    aliases = {str(entry.get("Name", "")) for entry in from_entries if isinstance(entry, dict)}
    for entry in prototype_query.get("Select", []) or []:
        if not isinstance(entry, dict):
            continue
        if "Column" in entry:
            column = entry.get("Column", {})
            if isinstance(column, dict):
                column_name = str(column.get("Property", ""))
                source_ref = column.get("Expression", {}).get("SourceRef", {}) if isinstance(column.get("Expression"), dict) else {}
                alias = str(source_ref.get("Source", "")) if isinstance(source_ref, dict) else ""
                table_name = from_entities.get(alias, "")
                if model_fields is not None and (table_name.casefold(), column_name.casefold()) not in model_fields["columns"]:
                    issues.append({"page": page_name, "visual_id": visual_id, "visual_type": visual_type, "issue": "column_not_found", "table": table_name, "column": column_name})
            continue
        if "Measure" not in entry:
            continue
        measure = entry.get("Measure", {})
        if not isinstance(measure, dict):
            continue
        measure_name = str(measure.get("Property", ""))
        source_ref = measure.get("Expression", {}).get("SourceRef", {}) if isinstance(measure.get("Expression"), dict) else {}
        alias = str(source_ref.get("Source", "")) if isinstance(source_ref, dict) else ""
        entity = from_entities.get(alias, "")
        home_table = measure_home_map.get(measure_name)
        home_table_source = "extract_metadata" if home_table is not None else ""
        if home_table is None and model_fields is not None:
            live_tables = sorted(model_fields.get("measure_tables", {}).get(measure_name.casefold(), set()))
            if len(live_tables) == 1:
                home_table = live_tables[0]
                home_table_source = "live_model"
        if model_fields is not None:
            measure_tables = model_fields["measures"].get(measure_name.casefold(), set())
            if not measure_tables:
                issues.append({"page": page_name, "visual_id": visual_id, "visual_type": visual_type, "issue": "measure_not_found", "measure": measure_name})
            elif entity and entity != "$Measures" and entity.casefold() not in measure_tables:
                issues.append({"page": page_name, "visual_id": visual_id, "visual_type": visual_type, "issue": "measure_table_mismatch", "measure": measure_name, "table": entity, "expected_tables": sorted(measure_tables)})
        if entity == "$Measures":
            if not home_table:
                issues.append({"page": page_name, "visual_id": visual_id, "visual_type": visual_type, "issue": "measure_home_table_unknown", "measure": measure_name})
                continue
            if not repair:
                item = {"page": page_name, "visual_id": visual_id, "visual_type": visual_type, "issue": "measure_home_table_needs_repair", "measure": measure_name, "home_table": home_table}
                if home_table_source == "live_model":
                    item.update({"source": "live_model", "extract_metadata": "missing"})
                issues.append(item)
                continue
            same_alias_measures = [
                str(item.get("Measure", {}).get("Property", ""))
                for item in prototype_query.get("Select", []) or []
                if isinstance(item, dict)
                and isinstance(item.get("Measure"), dict)
                and item.get("Measure", {}).get("Expression", {}).get("SourceRef", {}).get("Source") == alias
            ]
            def _resolved_measure_home(item: str) -> str | None:
                if item in measure_home_map:
                    return measure_home_map[item]
                if model_fields is not None:
                    live = sorted(model_fields.get("measure_tables", {}).get(item.casefold(), set()))
                    if len(live) == 1:
                        return live[0]
                return None

            if all(_resolved_measure_home(item) == home_table for item in same_alias_measures):
                for from_entry in from_entries:
                    if isinstance(from_entry, dict) and str(from_entry.get("Name", "")) == alias:
                        from_entry["Entity"] = home_table
                        break
            else:
                new_alias = _next_alias(aliases)
                from_entries.append({"Name": new_alias, "Entity": home_table})
                measure.setdefault("Expression", {}).setdefault("SourceRef", {})["Source"] = new_alias
            item = {"page": page_name, "visual_id": visual_id, "visual_type": visual_type, "issue": "measure_home_table_repaired", "measure": measure_name, "home_table": home_table}
            if home_table_source == "live_model":
                item.update({"source": "live_model", "extract_metadata": "missing"})
            issues.append(item)
            repairs += 1

    if repair and repairs:
        single_visual["prototypeQuery"] = prototype_query
        container["config"] = _dump_embedded_json(config)
        _sync_container_query(container, prototype_query)
    return issues, repairs


def _scan_visual_bindings(
    layout: dict[str, Any],
    measure_home_map: dict[str, str],
    model_fields: dict[str, Any] | None = None,
    *,
    page: str | None = None,
    repair: bool = False,
) -> tuple[list[dict[str, Any]], int]:
    issues: list[dict[str, Any]] = []
    repairs = 0
    sections = layout.get("sections", []) or []
    for section in sections:
        if not isinstance(section, dict):
            continue
        section_name = str(section.get("displayName") or section.get("name") or "")
        if page and page.casefold() not in {str(section.get("name", "")).casefold(), str(section.get("displayName", "")).casefold()}:
            continue
        for container in section.get("visualContainers", []) or []:
            if not isinstance(container, dict):
                continue
            found, fixed = _visual_binding_issues(container, section_name, measure_home_map, model_fields, repair=repair)
            issues.extend(found)
            repairs += fixed
    return issues, repairs


def _assert_container_bindings(container: dict[str, Any], measure_home_map: dict[str, str]) -> None:
    issues, _ = _visual_binding_issues(container, "", measure_home_map, repair=False)
    blocking = [item for item in issues if item.get("issue") in {"unexpected_projection_role", "query_ref_not_found", "query_ref_mismatch"}]
    if blocking:
        raise PowerBIValidationError("Visual field bindings are invalid.", details={"issues": blocking})


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
    measure_home_map: dict[str, str] | None = None,
    projections: dict[str, list[dict[str, str]]] | None = None,
    title: str | None = None,
    extra_single_visual: dict[str, Any] | None = None,
) -> tuple[dict[str, Any], dict[str, Any]]:
    position = {"x": x, "y": y, "width": width, "height": height}
    single_visual = {
        "visualType": visual_type,
        "projections": projections or {},
        "prototypeQuery": _build_prototype_query(references or [], measure_home_map),
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
    measure_home_map: dict[str, str] | None = None,
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
        measure_home_map=measure_home_map,
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


def _append_visual(
    extract_folder: str,
    page: str,
    factory: Callable[[dict[str, Any], dict[str, str]], dict[str, Any]],
    measure_home_map: dict[str, str],
) -> dict[str, Any]:
    folder, layout = _load_layout(extract_folder)
    section = _find_page(layout, page)
    section.setdefault("visualContainers", [])
    container = factory(section, measure_home_map)
    _assert_container_bindings(container, measure_home_map)
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
    measure_home_map: dict[str, str] | None = None,
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
        measure_home_map=measure_home_map,
        extra_single_visual=extra_single_visual,
    )


def pbi_extract_report_tool(pbix_path: str, extract_folder: str | None = None) -> dict[str, Any]:
    def _impl() -> dict[str, Any]:
        pbix = _resolve_pbix_path(pbix_path, must_exist=True)
        target = _resolve_extract_folder(str(extract_folder or pbix.with_name(f"{pbix.stem}_extracted")), must_exist=False)
        target.mkdir(parents=True, exist_ok=True)
        _run_pbi_tools(["extract", str(pbix), "-extractFolder", str(target), "-modelSerialization", "Legacy"])
        layout_path = target / LAYOUT_RELATIVE_PATH
        if not layout_path.exists():
            layout_path.parent.mkdir(parents=True, exist_ok=True)
            with zipfile.ZipFile(pbix, "r") as z:
                if "Report/Layout" in z.namelist():
                    layout_path.write_bytes(z.read("Report/Layout"))
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


def _run_powershell(script: str, *, timeout: float = 20.0) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        ["powershell", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-Command", script],
        capture_output=True,
        text=True,
        check=False,
        shell=False,
        timeout=timeout,
    )


def _save_and_close_powerbi_gracefully(pbix_path: Path | None = None) -> bool:
    target_path = str(pbix_path) if pbix_path is not None else ""
    script = "$TargetPath = " + json.dumps(target_path) + r"""
$ErrorActionPreference = 'SilentlyContinue'
$ws = New-Object -ComObject WScript.Shell
$names = @('PBIDesktop', 'pbidesktoprs')
$initialWrite = $null
if ($TargetPath -and (Test-Path -LiteralPath $TargetPath)) {
    $initialWrite = (Get-Item -LiteralPath $TargetPath).LastWriteTimeUtc
}
$procs = Get-Process -Name $names | Where-Object { $_.MainWindowHandle -ne 0 }
foreach ($proc in $procs) {
    [void]$ws.AppActivate($proc.Id)
    Start-Sleep -Milliseconds 500
    $ws.SendKeys('^s')
}
if ($initialWrite -ne $null) {
    $deadline = (Get-Date).AddSeconds(30)
    do {
        Start-Sleep -Seconds 1
        $currentWrite = (Get-Item -LiteralPath $TargetPath).LastWriteTimeUtc
    } while ($currentWrite -le $initialWrite -and (Get-Date) -lt $deadline)
} else {
    Start-Sleep -Seconds 8
}
foreach ($proc in @($procs)) {
    $proc.Refresh()
    if (-not $proc.HasExited) {
        [void]$proc.CloseMainWindow()
    }
}
$deadline = (Get-Date).AddSeconds(12)
do {
    Start-Sleep -Milliseconds 500
    $open = @(Get-Process -Name $names | Where-Object { $_.MainWindowHandle -ne 0 }).Count
} while ($open -gt 0 -and (Get-Date) -lt $deadline)
if ($open -gt 0) { exit 1 }
exit 0
"""
    try:
        return _run_powershell(script, timeout=45.0).returncode == 0
    except Exception:
        return False


def _force_kill_powerbi() -> None:
    for image in ("PBIDesktop.exe", "pbidesktoprs.exe"):
        try:
            subprocess.run(
                ["taskkill", "/F", "/IM", image],
                capture_output=True,
                text=True,
                check=False,
                shell=False,
            )
        except Exception:
            pass


def _maybe_force_close_powerbi(force: bool, pbix_path: Path | None = None) -> None:
    if not force:
        return
    if os.name != "nt":
        logger.debug("force=True ignored on non-Windows platform for PBIDesktop termination.")
        return
    if not _save_and_close_powerbi_gracefully(pbix_path):
        _force_kill_powerbi()
    time.sleep(1.5)


def _page_names_from_layout_bytes(layout_bytes: bytes) -> list[str]:
    try:
        layout = json.loads(layout_bytes.decode("utf-16-le"))
    except (UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise ReportLayoutError("Report/Layout content is invalid UTF-16-LE JSON.") from exc
    if not isinstance(layout, dict):
        raise ReportLayoutError("Report/Layout root must be a JSON object.")
    names: list[str] = []
    for section in layout.get("sections", []):
        if not isinstance(section, dict):
            continue
        names.append(str(section.get("displayName") or section.get("name") or ""))
    return names


def pbi_patch_layout_tool(
    extract_folder: str,
    pbix_path: str,
    force: bool = False,
    fail_on_persistence_risk: bool = True,
    manager: Any | None = None,
    include_hidden: bool = False,
) -> dict[str, Any]:
    def _impl() -> dict[str, Any]:
        folder = _resolve_extract_folder(extract_folder, must_exist=True)
        pbix = _resolve_pbix_path(pbix_path, must_exist=True)
        layout_path = _layout_path(folder)
        if not layout_path.exists():
            raise ReportLayoutError("Report/Layout file was not found in the extract folder.", details={"path": str(layout_path)})

        if fail_on_persistence_risk:
            _, layout = _load_layout(folder)
            measure_home_map = _scan_measure_home_tables(folder)
            model_fields, model_validation = _live_model_field_index(manager, include_hidden=include_hidden)
            issues, _ = _scan_visual_bindings(layout, measure_home_map, model_fields, repair=False)
            persistence_risks = _persistence_risks(issues)
            if persistence_risks:
                raise PowerBIValidationError(
                    "Layout patch blocked because field bindings rely on live-model metadata missing from the extract.",
                    details={
                        "persistence_risk_count": len(persistence_risks),
                        "persistence_risks": persistence_risks,
                        "model_validation": model_validation,
                    },
                )

        _maybe_force_close_powerbi(force, pbix)

        layout_bytes = layout_path.read_bytes()
        pages = _page_names_from_layout_bytes(layout_bytes)

        temp_path: Path | None = None
        try:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pbix", dir=str(pbix.parent)) as tmp_file:
                temp_path = Path(tmp_file.name)
            with zipfile.ZipFile(pbix, "r") as source_zip, zipfile.ZipFile(temp_path, "w") as target_zip:
                layout_written = False
                for info in source_zip.infolist():
                    name = info.filename
                    if name == "SecurityBindings":
                        continue
                    payload = layout_bytes if name == "Report/Layout" else source_zip.read(name)
                    if name == "Report/Layout":
                        layout_written = True
                    target_info = zipfile.ZipInfo(name, date_time=info.date_time)
                    target_info.compress_type = info.compress_type
                    target_info.comment = info.comment
                    target_info.extra = info.extra
                    target_info.internal_attr = info.internal_attr
                    target_info.external_attr = info.external_attr
                    target_info.create_system = info.create_system
                    target_info.create_version = info.create_version
                    target_info.extract_version = info.extract_version
                    target_info.volume = info.volume
                    target_info.flag_bits = info.flag_bits
                    target_zip.writestr(target_info, payload)
                if not layout_written:
                    target_info = zipfile.ZipInfo("Report/Layout")
                    target_info.compress_type = zipfile.ZIP_DEFLATED
                    target_zip.writestr(target_info, layout_bytes)

            temp_size = temp_path.stat().st_size
            try:
                temp_path.replace(pbix)
            except PermissionError as exc:
                raise ReportLayoutError(
                    "PBIX file is locked by Power BI Desktop. Close it or retry with force=True.",
                    details={"pbix_path": str(pbix), "force": force},
                ) from exc
        finally:
            if temp_path and temp_path.exists():
                temp_path.unlink(missing_ok=True)

        return ok(
            "Layout patched into PBIX successfully.",
            extract_folder=str(folder),
            pbix_path=str(pbix),
            bytes_written=temp_size,
            layout_size=len(layout_bytes),
            pages=pages,
            persistence_risk_checked=fail_on_persistence_risk,
        )

    return _run(_impl)


def pbi_compile_report_tool(extract_folder: str, output_path: str, force: bool = False) -> dict[str, Any]:
    def _impl() -> dict[str, Any]:
        folder = _resolve_extract_folder(extract_folder, must_exist=True)
        output = _resolve_pbix_path(output_path, must_exist=False)
        output.parent.mkdir(parents=True, exist_ok=True)
        _maybe_force_close_powerbi(force, output if output.exists() else None)
        _run_pbi_tools(["compile", str(folder), "-outPath", str(output), "-overwrite"])
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
    measure_home_map = _scan_measure_home_tables(_resolve_extract_folder(extract_folder, must_exist=True))
    return _append_visual(
        extract_folder,
        page,
        lambda section, home_map: _create_chart_container(
            section,
            visual_type="card",
            x=x,
            y=y,
            width=width,
            height=height,
            title=title or None,
            projections={"Values": [{"queryRef": _query_ref(measure)}]},
            references=[measure],
            measure_home_map=home_map,
        ),
        measure_home_map,
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
    measure_home_map = _scan_measure_home_tables(_resolve_extract_folder(extract_folder, must_exist=True))
    projections = {"Category": [{"queryRef": _query_ref(category_column)}], "Y": [{"queryRef": _query_ref(value_measure)}]}
    references = [category_column, value_measure]
    if legend_column:
        projections["Series"] = [{"queryRef": _query_ref(legend_column)}]
        references.append(legend_column)
    return _append_visual(
        extract_folder,
        page,
        lambda section, home_map: _create_chart_container(
            section,
            visual_type="clusteredBarChart",
            x=x,
            y=y,
            width=width,
            height=height,
            title=title or None,
            projections=projections,
            references=references,
            measure_home_map=home_map,
        ),
        measure_home_map,
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
    measure_home_map = _scan_measure_home_tables(_resolve_extract_folder(extract_folder, must_exist=True))
    return _append_visual(
        extract_folder,
        page,
        lambda section, home_map: _create_chart_container(
            section,
            visual_type="lineChart",
            x=x,
            y=y,
            width=width,
            height=height,
            title=title or None,
            projections={
                "Category": [{"queryRef": _query_ref(axis_column)}],
                "Y": [{"queryRef": _query_ref(measure)} for measure in value_measures],
            },
            references=[axis_column, *value_measures],
            measure_home_map=home_map,
        ),
        measure_home_map,
    )


def pbi_add_donut_chart_tool(extract_folder: str, page: str, category_column: str, value_measure: str, x: int, y: int, width: int = 320, height: int = 280, title: str = "") -> dict[str, Any]:
    measure_home_map = _scan_measure_home_tables(_resolve_extract_folder(extract_folder, must_exist=True))
    return _append_visual(
        extract_folder,
        page,
        lambda section, home_map: _create_chart_container(
            section,
            visual_type="donutChart",
            x=x,
            y=y,
            width=width,
            height=height,
            title=title or None,
            projections={"Category": [{"queryRef": _query_ref(category_column)}], "Y": [{"queryRef": _query_ref(value_measure)}]},
            references=[category_column, value_measure],
            measure_home_map=home_map,
        ),
        measure_home_map,
    )


def pbi_add_table_visual_tool(extract_folder: str, page: str, columns: list[str], x: int, y: int, width: int = 520, height: int = 320, title: str = "") -> dict[str, Any]:
    if not columns:
        raise PowerBIValidationError("columns must contain at least one field or measure.")
    measure_home_map = _scan_measure_home_tables(_resolve_extract_folder(extract_folder, must_exist=True))
    return _append_visual(
        extract_folder,
        page,
        lambda section, home_map: _create_chart_container(
            section,
            visual_type="tableEx",
            x=x,
            y=y,
            width=width,
            height=height,
            title=title or None,
            projections={"Values": [{"queryRef": _query_ref(item)} for item in columns]},
            references=list(columns),
            measure_home_map=home_map,
        ),
        measure_home_map,
    )


def pbi_add_waterfall_tool(extract_folder: str, page: str, category_column: str, value_measure: str, x: int, y: int, width: int = 420, height: int = 300, title: str = "") -> dict[str, Any]:
    measure_home_map = _scan_measure_home_tables(_resolve_extract_folder(extract_folder, must_exist=True))
    return _append_visual(
        extract_folder,
        page,
        lambda section, home_map: _create_chart_container(
            section,
            visual_type="waterfallChart",
            x=x,
            y=y,
            width=width,
            height=height,
            title=title or None,
            projections={"Category": [{"queryRef": _query_ref(category_column)}], "Y": [{"queryRef": _query_ref(value_measure)}]},
            references=[category_column, value_measure],
            measure_home_map=home_map,
        ),
        measure_home_map,
    )


def pbi_add_slicer_tool(extract_folder: str, page: str, column: str, x: int, y: int, width: int = 220, height: int = 120, slicer_type: str = "dropdown") -> dict[str, Any]:
    slicer_kind = slicer_type.strip().casefold()
    if slicer_kind not in {"dropdown", "list", "range"}:
        raise PowerBIValidationError("slicer_type must be one of: dropdown, list, range.", details={"slicer_type": slicer_type})
    measure_home_map = _scan_measure_home_tables(_resolve_extract_folder(extract_folder, must_exist=True))
    return _append_visual(
        extract_folder,
        page,
        lambda section, home_map: _create_chart_container(
            section,
            visual_type="slicer",
            x=x,
            y=y,
            width=width,
            height=height,
            title=None,
            projections={"Values": [{"queryRef": _query_ref(column)}]},
            references=[column],
            measure_home_map=home_map,
            extra_single_visual={"slicerType": slicer_kind},
        ),
        measure_home_map,
    )


def pbi_add_gauge_tool(extract_folder: str, page: str, measure: str, x: int, y: int, width: int = 280, height: int = 220, title: str = "", target_measure: str | None = None) -> dict[str, Any]:
    measure_home_map = _scan_measure_home_tables(_resolve_extract_folder(extract_folder, must_exist=True))
    projections = {"Y": [{"queryRef": _query_ref(measure)}]}
    references = [measure]
    if target_measure:
        projections["Goal"] = [{"queryRef": _query_ref(target_measure)}]
        references.append(target_measure)
    return _append_visual(
        extract_folder,
        page,
        lambda section, home_map: _create_chart_container(
            section,
            visual_type="gauge",
            x=x,
            y=y,
            width=width,
            height=height,
            title=title or None,
            projections=projections,
            references=references,
            measure_home_map=home_map,
        ),
        measure_home_map,
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
    measure_home_map = _scan_measure_home_tables(_resolve_extract_folder(extract_folder, must_exist=True))
    return _append_visual(
        extract_folder,
        page,
        lambda section, home_map: _make_visual_container(
            section=section,
            visual_type="textbox",
            x=x,
            y=y,
            width=width,
            height=height,
            references=[],
            measure_home_map=home_map,
            projections={},
            extra_single_visual={
                "textContent": text,
                "textStyle": {"fontSize": font_size, "bold": bold, "color": color},
                "prototypeQuery": {"Version": 2, "From": [], "Select": []},
                "objects": {"paragraphs": [{"text": text, "fontSize": font_size, "bold": bold, "color": color}]},
            },
        ),
        measure_home_map,
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
        return _create_chart_container(section, visual_type="card", x=x, y=y, width=width, height=height, title=title, projections={"Values": [{"queryRef": _query_ref(spec["measure"])}]}, references=[spec["measure"]], measure_home_map=measure_home_map)
    if visual_type in {"bar_chart", "bar"}:
        projections = {"Category": [{"queryRef": _query_ref(spec["category"])}], "Y": [{"queryRef": _query_ref(spec["measure"])}]}
        references = [spec["category"], spec["measure"]]
        if spec.get("legend"):
            projections["Series"] = [{"queryRef": _query_ref(spec["legend"])}]
            references.append(spec["legend"])
        return _create_chart_container(section, visual_type="clusteredBarChart", x=x, y=y, width=width, height=height, title=title, projections=projections, references=references, measure_home_map=measure_home_map)
    if visual_type in {"line_chart", "line"}:
        measures = list(spec.get("measures") or [spec.get("measure")])
        return _create_chart_container(section, visual_type="lineChart", x=x, y=y, width=width, height=height, title=title, projections={"Category": [{"queryRef": _query_ref(spec["axis"])}], "Y": [{"queryRef": _query_ref(item)} for item in measures]}, references=[spec["axis"], *measures], measure_home_map=measure_home_map)
    if visual_type in {"donut", "donut_chart", "pie", "pie_chart"}:
        return _create_chart_container(section, visual_type="donutChart", x=x, y=y, width=width, height=height, title=title, projections={"Category": [{"queryRef": _query_ref(spec["category"])}], "Y": [{"queryRef": _query_ref(spec["measure"])}]}, references=[spec["category"], spec["measure"]], measure_home_map=measure_home_map)
    if visual_type in {"table", "table_visual"}:
        return _create_chart_container(section, visual_type="tableEx", x=x, y=y, width=width, height=height, title=title, projections={"Values": [{"queryRef": _query_ref(item)} for item in spec["columns"]]}, references=list(spec["columns"]), measure_home_map=measure_home_map)
    if visual_type == "waterfall":
        return _create_chart_container(section, visual_type="waterfallChart", x=x, y=y, width=width, height=height, title=title, projections={"Category": [{"queryRef": _query_ref(spec["category"])}], "Y": [{"queryRef": _query_ref(spec["measure"])}]}, references=[spec["category"], spec["measure"]], measure_home_map=measure_home_map)
    if visual_type == "slicer":
        return _make_visual_container(section=section, visual_type="slicer", x=x, y=y, width=width, height=height, projections={"Values": [{"queryRef": _query_ref(spec["column"])}]}, references=[spec["column"]], measure_home_map=measure_home_map, extra_single_visual={"slicerType": str(spec.get("slicer_type", "dropdown")).casefold()})
    if visual_type in {"text", "text_box"}:
        return _make_visual_container(section=section, visual_type="textbox", x=x, y=y, width=width, height=height, measure_home_map=measure_home_map, extra_single_visual={"textContent": spec["text"], "textStyle": {"fontSize": int(spec.get("font_size", 16)), "bold": bool(spec.get("bold", False)), "color": str(spec.get("color", "#222222"))}, "prototypeQuery": {"Version": 2, "From": [], "Select": []}})
    if visual_type == "gauge":
        return _create_chart_container(section, visual_type="gauge", x=x, y=y, width=width, height=height, title=title, projections={"Y": [{"queryRef": _query_ref(spec["measure"])}]}, references=[spec["measure"]], measure_home_map=measure_home_map)
    if visual_type == "kpi":
        measures = [spec["measure"]]
        if spec.get("target_measure"):
            measures.append(spec["target_measure"])
        return _create_chart_container(section, visual_type="kpi", x=x, y=y, width=width, height=height, title=title, projections={"Value": [{"queryRef": _query_ref(spec["measure"])}], "Goal": [{"queryRef": _query_ref(spec["target_measure"])}]} if spec.get("target_measure") else {"Value": [{"queryRef": _query_ref(spec["measure"])}]}, references=measures, measure_home_map=measure_home_map)
    if visual_type == "map":
        refs = [spec["location"]]
        projections = {"Category": [{"queryRef": _query_ref(spec["location"])}]}
        if spec.get("measure"):
            refs.append(spec["measure"])
            projections["Y"] = [{"queryRef": _query_ref(spec["measure"])}]
        return _create_chart_container(section, visual_type="map", x=x, y=y, width=width, height=height, title=title, projections=projections, references=refs, measure_home_map=measure_home_map)
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


def pbi_validate_report_fields_tool(
    extract_folder: str,
    page: str | None = None,
    include_hidden: bool = False,
    manager: Any | None = None,
) -> dict[str, Any]:
    def _impl() -> dict[str, Any]:
        folder, layout = _load_layout(extract_folder)
        measure_home_map = _scan_measure_home_tables(folder)
        model_fields, model_validation = _live_model_field_index(manager, include_hidden=include_hidden)
        issues, _ = _scan_visual_bindings(layout, measure_home_map, model_fields, page=page, repair=False)
        blocking = [item for item in issues if item.get("issue") not in {"measure_home_table_repaired"}]
        persistence_risks = _persistence_risks(issues)
        return ok(
            f"Report field validation found {len(blocking)} issue(s).",
            extract_folder=str(folder),
            page=page,
            include_hidden=include_hidden,
            model_validation=model_validation,
            valid=not blocking,
            issue_count=len(blocking),
            issues=blocking,
            persistence_risk_count=len(persistence_risks),
            persistence_risks=persistence_risks,
        )

    return _run(_impl)


def pbi_repair_report_fields_tool(
    extract_folder: str,
    page: str | None = None,
    apply: bool = False,
    manager: Any | None = None,
    include_hidden: bool = False,
) -> dict[str, Any]:
    def _impl() -> dict[str, Any]:
        folder, layout = _load_layout(extract_folder)
        measure_home_map = _scan_measure_home_tables(folder)
        model_fields, model_validation = _live_model_field_index(manager, include_hidden=include_hidden)
        issues, repairs = _scan_visual_bindings(layout, measure_home_map, model_fields, page=page, repair=apply)
        planned_repairs = repairs if apply else sum(
            1
            for item in issues
            if item.get("issue") in {"query_ref_mismatch", "measure_home_table_needs_repair"}
            or (item.get("issue") == "unexpected_projection_role" and item.get("visual_type") == "gauge" and item.get("role") == "Value")
        )
        unresolved = [
            item for item in issues
            if item.get("issue") in {"query_ref_not_found", "unexpected_projection_role", "measure_home_table_unknown", "column_not_found", "measure_not_found", "measure_table_mismatch"}
            and not (item.get("visual_type") == "gauge" and item.get("role") == "Value")
        ]
        persistence_risks = _persistence_risks(issues)
        if apply and repairs:
            _save_layout(folder, layout)
        return ok(
            f"Report field repair {'applied' if apply else 'planned'}: {planned_repairs} deterministic fix(es), {len(unresolved)} unresolved issue(s).",
            extract_folder=str(folder),
            page=page,
            apply=apply,
            model_validation=model_validation,
            repairs=planned_repairs,
            unresolved=unresolved,
            persistence_risk_count=len(persistence_risks),
            persistence_risks=persistence_risks,
            issues=issues,
            needs_apply=not apply and planned_repairs > 0,
        )

    return _run(_impl)


_VISUAL_TYPE_DISPATCH: dict[str, Callable[..., dict[str, Any]]] = {}


def pbi_add_visual_tool(
    extract_folder: str,
    page: str,
    visual_type: str,
    x: int,
    y: int,
    width: int | None = None,
    height: int | None = None,
    title: str = "",
    config: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """Generic visual dispatcher. Keeps the 9 per-type tools as stable API surface.

    visual_type: one of card, bar_chart, line_chart, donut, table, waterfall,
                 slicer, gauge, text_box.
    config: per-type keyword arguments (e.g. {"measure": "Total Sales"} for card,
            {"category_column": "...", "value_measure": "..."} for bar_chart).
    """
    cfg = dict(config or {})
    visual_key = visual_type.strip().casefold()
    size = DEFAULT_VISUAL_SIZES.get(visual_key)
    effective_width = width if width is not None else (size[0] if size else 320)
    effective_height = height if height is not None else (size[1] if size else 240)

    handler = _VISUAL_TYPE_DISPATCH.get(visual_key)
    if handler is None:
        raise PowerBIValidationError(
            f"Unknown visual_type '{visual_type}'. Allowed: {sorted(_VISUAL_TYPE_DISPATCH)}",
            details={"visual_type": visual_type},
        )
    return handler(extract_folder, page, x, y, effective_width, effective_height, title, cfg)


def _dispatch_card(extract, page, x, y, w, h, title, cfg):
    measure = cfg.get("measure")
    if not measure:
        raise PowerBIValidationError("card visual requires config.measure", details={"visual_type": "card"})
    return pbi_add_card_tool(extract, page, measure, x, y, w, h, title)


def _dispatch_bar(extract, page, x, y, w, h, title, cfg):
    cat = cfg.get("category_column")
    value = cfg.get("value_measure")
    if not cat or not value:
        raise PowerBIValidationError(
            "bar_chart requires config.category_column and config.value_measure",
            details={"visual_type": "bar_chart"},
        )
    return pbi_add_bar_chart_tool(extract, page, cat, value, x, y, w, h, title, cfg.get("legend_column"))


def _dispatch_line(extract, page, x, y, w, h, title, cfg):
    axis = cfg.get("axis_column")
    measures = cfg.get("value_measures") or []
    if not axis or not measures:
        raise PowerBIValidationError(
            "line_chart requires config.axis_column and config.value_measures (list)",
            details={"visual_type": "line_chart"},
        )
    return pbi_add_line_chart_tool(extract, page, axis, measures, x, y, w, h, title)


def _dispatch_donut(extract, page, x, y, w, h, title, cfg):
    cat = cfg.get("category_column")
    value = cfg.get("value_measure")
    if not cat or not value:
        raise PowerBIValidationError(
            "donut requires config.category_column and config.value_measure",
            details={"visual_type": "donut"},
        )
    return pbi_add_donut_chart_tool(extract, page, cat, value, x, y, w, h, title)


def _dispatch_table(extract, page, x, y, w, h, title, cfg):
    columns = cfg.get("columns") or []
    if not columns:
        raise PowerBIValidationError("table requires config.columns (list)", details={"visual_type": "table"})
    return pbi_add_table_visual_tool(extract, page, columns, x, y, w, h, title)


def _dispatch_waterfall(extract, page, x, y, w, h, title, cfg):
    cat = cfg.get("category_column")
    value = cfg.get("value_measure")
    if not cat or not value:
        raise PowerBIValidationError(
            "waterfall requires config.category_column and config.value_measure",
            details={"visual_type": "waterfall"},
        )
    return pbi_add_waterfall_tool(extract, page, cat, value, x, y, w, h, title)


def _dispatch_slicer(extract, page, x, y, w, h, title, cfg):
    column = cfg.get("column")
    if not column:
        raise PowerBIValidationError("slicer requires config.column", details={"visual_type": "slicer"})
    return pbi_add_slicer_tool(extract, page, column, x, y, w, h, cfg.get("slicer_type", "dropdown"))


def _dispatch_gauge(extract, page, x, y, w, h, title, cfg):
    measure = cfg.get("measure")
    if not measure:
        raise PowerBIValidationError("gauge requires config.measure", details={"visual_type": "gauge"})
    return pbi_add_gauge_tool(extract, page, measure, x, y, w, h, title, cfg.get("target_measure"))


def _dispatch_text_box(extract, page, x, y, w, h, title, cfg):
    text = cfg.get("text")
    if text is None:
        raise PowerBIValidationError("text_box requires config.text", details={"visual_type": "text_box"})
    return pbi_add_text_box_tool(
        extract,
        page,
        str(text),
        x,
        y,
        w,
        h,
        int(cfg.get("font_size", 16)),
        bool(cfg.get("bold", False)),
        str(cfg.get("color", "#222222")),
    )


_VISUAL_TYPE_DISPATCH.update({
    "card": _dispatch_card,
    "bar_chart": _dispatch_bar,
    "line_chart": _dispatch_line,
    "donut": _dispatch_donut,
    "table": _dispatch_table,
    "waterfall": _dispatch_waterfall,
    "slicer": _dispatch_slicer,
    "gauge": _dispatch_gauge,
    "text_box": _dispatch_text_box,
    "textbox": _dispatch_text_box,
})


__all__ = [
    "pbi_add_visual_tool",
    "pbi_add_bar_chart_tool",
    "pbi_add_card_tool",
    "pbi_add_donut_chart_tool",
    "pbi_add_gauge_tool",
    "pbi_add_line_chart_tool",
    "pbi_add_slicer_tool",
    "pbi_add_table_visual_tool",
    "pbi_add_text_box_tool",
    "pbi_add_waterfall_tool",
    "pbi_apply_design_tool",
    "pbi_apply_theme_tool",
    "pbi_build_dashboard_tool",
    "pbi_compile_report_tool",
    "pbi_create_page_tool",
    "pbi_delete_page_tool",
    "pbi_extract_report_tool",
    "pbi_get_page_tool",
    "pbi_list_pages_tool",
    "pbi_move_visual_tool",
    "pbi_patch_layout_tool",
    "pbi_repair_report_fields_tool",
    "pbi_remove_visual_tool",
    "pbi_set_page_size_tool",
    "pbi_validate_report_fields_tool",
]

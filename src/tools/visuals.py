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
    "clusteredColumnChart": {"Category", "Y", "Series"},
    "lineChart": {"Category", "Y"},
    "donutChart": {"Category", "Y"},
    "tableEx": {"Values"},
    "waterfallChart": {"Category", "Y"},
    "slicer": {"Values"},
    "gauge": {"Y", "Goal"},
    "kpi": {"Indicator", "TrendLine", "Goal"},
    "map": {"Category", "Y"},
    "scatterChart": {"Category", "X", "Y", "Size", "Series"},
    "lineClusteredColumnComboChart": {"Category", "Y", "Y2", "Series"},
    "pivotTable": {"Rows", "Columns", "Values"},
}

# Per-role expected reference kind ("column", "measure", or "any"). Used by the
# pre-flight projection role validator to catch "wrong field kind in role"
# mistakes (e.g. an LLM puts a measure into Category) at tool-call time, before
# we ever write the layout.
VISUAL_ROLE_KINDS: dict[str, dict[str, str]] = {
    "card": {"Values": "measure"},
    "clusteredBarChart": {"Category": "column", "Y": "measure", "Series": "column"},
    "clusteredColumnChart": {"Category": "column", "Y": "measure", "Series": "column"},
    "lineChart": {"Category": "column", "Y": "measure"},
    "donutChart": {"Category": "column", "Y": "measure"},
    "tableEx": {"Values": "any"},
    "waterfallChart": {"Category": "column", "Y": "measure"},
    "slicer": {"Values": "column"},
    "gauge": {"Y": "measure", "Goal": "measure"},
    "kpi": {"Indicator": "measure", "TrendLine": "column", "Goal": "measure"},
    "map": {"Category": "column", "Y": "measure"},
    "scatterChart": {"Category": "column", "X": "measure", "Y": "measure", "Size": "measure", "Series": "column"},
    "lineClusteredColumnComboChart": {"Category": "column", "Y": "measure", "Y2": "measure", "Series": "column"},
    "pivotTable": {"Rows": "column", "Columns": "column", "Values": "measure"},
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
    bundled = Path(__file__).resolve().parents[2] / "tools-bin" / "pbi-tools.core.exe"
    if bundled.exists():
        return str(bundled)
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


# Match the standard Power BI bracket forms: Table[Column] or 'Table With Spaces'[Column].
_BRACKET_REF_RE = re.compile(r"^\s*'?(?P<table>[^'\[\]]+?)'?\s*\[\s*(?P<column>[^\[\]]+?)\s*\]\s*$")


def _normalize_reference(reference: str) -> str:
    """Normalise a user-supplied field reference into ``"Table.Column"`` form.

    Accepts (case-insensitive on whitespace; surrounding quotes optional):
    - ``"Table.Column"`` (existing canonical form, returned as-is)
    - ``"Table[Column]"``
    - ``"'Table With Spaces'[Column]"``
    - ``"BareMeasureName"`` (measure references stay unchanged)

    The downstream tooling (``_split_column_ref``, ``_query_ref``) treats the
    returned string as ``Table.Column``.
    """
    if not isinstance(reference, str):
        return reference  # type: ignore[return-value]
    raw = reference.strip()
    if not raw:
        return raw
    match = _BRACKET_REF_RE.match(raw)
    if match:
        table = match.group("table").strip()
        column = match.group("column").strip()
        return f"{table}.{column}"
    return raw


def _split_column_ref(reference: str) -> tuple[str, str]:
    normalized = _normalize_reference(reference)
    if "." not in normalized:
        raise PowerBIValidationError(
            "Column references must use 'TableName.ColumnName', 'TableName[ColumnName]', "
            "or '\\'Table Name\\'[Column Name]' format.",
            details={"reference": reference},
        )
    table, column = normalized.rsplit(".", 1)
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
    """Return the short queryRef name (column part only, without table prefix).

    Accepts the same flexible reference forms as :func:`_normalize_reference`
    (``Table.Column``, ``Table[Column]``, ``'Table'[Column]``, bare measure).
    """
    normalized = _normalize_reference(reference)
    return normalized.split(".", 1)[1] if "." in normalized else normalized


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


def _resolve_measure_home_map(
    extract_folder: str,
    manager: Any | None = None,
) -> dict[str, str]:
    """Build a measure → home table map combining on-disk PBIP metadata and the
    live model (when ``manager`` is supplied).

    Use at the top of every ``pbi_add_*_tool`` so the visual write carries the
    correct ``Entity`` reference and callers don't get the post-hoc
    ``measure_home_table_needs_repair`` validation issue.
    """
    home_map = _scan_measure_home_tables(_resolve_extract_folder(extract_folder, must_exist=True))
    return _augment_measure_home_map_with_live(home_map, manager)


def _augment_measure_home_map_with_live(
    measure_home_map: dict[str, str],
    manager: Any | None,
    *,
    include_hidden: bool = False,
) -> dict[str, str]:
    """Fill in missing measure → table mappings from the live model.

    The on-disk PBIP extract metadata is the canonical source, but it isn't
    always present (e.g. a layout-only extract from a closed PBIX). When a
    connection manager is supplied, we pull the same information from the
    live TOM so visual writes carry the correct ``Entity`` reference and
    callers don't see ``measure_home_table_needs_repair`` after the write.

    Returns the (possibly augmented) map. Existing entries take priority so
    the on-disk metadata always wins on conflict.
    """
    if manager is None:
        return measure_home_map
    try:
        model = pbi_model_info_tool(manager, include_hidden=include_hidden, include_row_counts=False)
    except Exception:
        return measure_home_map
    if not model.get("ok"):
        return measure_home_map
    existing_lower = {key.casefold() for key in measure_home_map}
    for measure in model.get("measures", []) or []:
        name = str(measure.get("name", ""))
        table_name = str(measure.get("table", ""))
        if not name or not table_name:
            continue
        if name.casefold() in existing_lower:
            continue
        measure_home_map[name] = table_name
        existing_lower.add(name.casefold())
    return measure_home_map


def _build_select_entry(
    reference: str,
    aliases: dict[str, str],
    measure_home_map: dict[str, str] | None = None,
) -> dict[str, Any]:
    # Normalise so Date[Année] / 'Date'[Année] / Date.Année all enter the same path.
    reference = _normalize_reference(reference)
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


def _validate_projection_roles(
    visual_type: str,
    projections: dict[str, list[dict[str, str]]] | None,
    *,
    manager: Any | None = None,
    include_hidden: bool = False,
) -> dict[str, Any]:
    """Pre-flight check that every role used in ``projections`` is allowed for
    ``visual_type`` and (when a live model is reachable) carries a reference of
    the expected kind (column vs measure).

    Use at tool-call time, before writing the layout. ``visual_type`` matches
    the singleVisual.visualType value (e.g. ``clusteredBarChart``,
    ``scatterChart``). Raises ``PowerBIValidationError`` with structured
    ``details`` listing every offending role + reference.
    """
    if not isinstance(projections, dict) or not projections:
        return {"status": "skipped", "reason": "no_projections"}
    allowed = VISUAL_FIELD_ROLES.get(visual_type)
    role_kinds = VISUAL_ROLE_KINDS.get(visual_type, {})

    unknown_roles: list[str] = []
    if allowed is not None:
        for role in projections:
            if role not in allowed:
                unknown_roles.append(role)
    if unknown_roles:
        raise PowerBIValidationError(
            f"Visual type '{visual_type}' does not accept role(s): {', '.join(sorted(unknown_roles))}.",
            details={
                "visual_type": visual_type,
                "unknown_roles": sorted(unknown_roles),
                "allowed_roles": sorted(allowed) if allowed else [],
            },
        )

    if manager is None or not role_kinds:
        return {"status": "roles_only_checked"}
    index, status = _live_model_field_index(manager, include_hidden=include_hidden)
    if index is None:
        return status

    role_kind_mismatches: list[dict[str, str]] = []
    for role, items in projections.items():
        expected_kind = role_kinds.get(role, "any")
        if expected_kind == "any":
            continue
        for item in items or []:
            ref = item.get("queryRef") if isinstance(item, dict) else None
            if not isinstance(ref, str) or not ref.strip():
                continue
            actual_kind: str | None = None
            if ref.casefold() in index["measures"]:
                actual_kind = "measure"
            else:
                # Look up by short column name across all tables.
                for table_lc, col_lc in index["columns"]:
                    if col_lc == ref.casefold():
                        actual_kind = "column"
                        break
            if actual_kind is not None and actual_kind != expected_kind:
                role_kind_mismatches.append({
                    "role": role,
                    "reference": ref,
                    "expected_kind": expected_kind,
                    "actual_kind": actual_kind,
                })
    if role_kind_mismatches:
        raise PowerBIValidationError(
            "Projection role/kind mismatch — at least one reference is the wrong kind for its role.",
            details={
                "visual_type": visual_type,
                "mismatches": role_kind_mismatches,
            },
        )
    return {"status": "roles_and_kinds_checked"}


def _validate_field_references_live(
    manager: Any | None,
    references: list[str],
    *,
    expected_kinds: dict[str, str] | None = None,
    include_hidden: bool = False,
) -> dict[str, Any]:
    """If a connection manager is available, verify each reference exists in
    the live model. Raises ``PowerBIValidationError`` on missing fields so
    callers fail fast before writing the layout.

    A reference is either ``"TableName.ColumnName"`` (column) or a bare
    ``"MeasureName"`` (measure). Bracket forms (``Table[Column]``,
    ``'Table'[Column]``) are accepted and normalised. Skips silently when the
    live model is unavailable so offline/test usage still works.

    ``expected_kinds`` (optional): map ``reference -> "column" | "measure"``
    so the error message can call out role expectation when a user passes the
    wrong format (e.g. ``"Year"`` for a Category role that wants a column).
    The reported ``kind`` for each missing entry then becomes the role
    expectation rather than only the inferred-from-format guess, so
    diagnostic messages match the visual's role contract.
    """
    if manager is None or not references:
        return {"status": "skipped"}
    index, status = _live_model_field_index(manager, include_hidden=include_hidden)
    if index is None:
        return status
    expected_kinds = expected_kinds or {}
    missing: list[dict[str, str]] = []
    for ref in references:
        if not isinstance(ref, str) or not ref.strip():
            continue
        normalized = _normalize_reference(ref)
        expected_kind = expected_kinds.get(ref)
        if "." in normalized:
            table, column = normalized.split(".", 1)
            if (table.casefold(), column.casefold()) not in index["columns"]:
                missing.append({
                    "reference": ref,
                    "kind": expected_kind or "column",
                    "inferred_kind": "column",
                    "hint": "use 'Table.Column', 'Table[Column]', or \"'Table With Spaces'[Column]\"",
                })
        else:
            # Bare form: it could be a measure OR an unqualified column name.
            measure_hit = normalized.casefold() in index["measures"]
            column_short_hit = any(col_lc == normalized.casefold() for _table_lc, col_lc in index["columns"])
            if expected_kind == "column":
                if not column_short_hit:
                    missing.append({
                        "reference": ref,
                        "kind": "column",
                        "inferred_kind": "measure" if measure_hit else "unknown",
                        "hint": "axis/category/rows expect a column — qualify with the table (e.g. 'Date.Year' or 'Date[Year]').",
                    })
            elif expected_kind == "measure":
                if not measure_hit:
                    missing.append({
                        "reference": ref,
                        "kind": "measure",
                        "inferred_kind": "column" if column_short_hit else "unknown",
                        "hint": "values/Y expect a measure — check spelling against the live model's measure list.",
                    })
            else:
                if not measure_hit:
                    missing.append({
                        "reference": ref,
                        "kind": "measure",
                        "inferred_kind": "column" if column_short_hit else "unknown",
                    })
    if missing:
        raise PowerBIValidationError(
            f"Field reference(s) not found in the live model: "
            f"{', '.join(item['reference'] for item in missing)}",
            details={"missing": missing, "checked": list(references)},
        )
    return {"status": "validated", "checked": len(references)}


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


def _decimal_literal(value: float) -> dict[str, Any]:
    """Power BI numeric literal (Decimal). Uses 'D' suffix expected by the report engine."""
    return {"expr": {"Literal": {"Value": f"{float(value)}D"}}}


def _int_literal(value: int) -> dict[str, Any]:
    return {"expr": {"Literal": {"Value": f"{int(value)}L"}}}


def _text_literal(value: str) -> dict[str, Any]:
    """Power BI canonical text literal: 'value' with embedded quotes doubled.

    PBI's Literal.Value field for text uses single-quoted form (the older
    json.dumps-derived '"…"' style is silently ignored by some visual
    serializers, which is why titles set that way never render).
    """
    escaped = str(value).replace("'", "''")
    return {"expr": {"Literal": {"Value": f"'{escaped}'"}}}


def _solid_color(color: str) -> dict[str, Any]:
    if not HEX_COLOR_RE.match(color):
        raise PowerBIValidationError(
            "color must match '#RRGGBB'.",
            details={"value": color},
        )
    return {"solid": {"color": {"expr": {"Literal": {"Value": f"'{color}'"}}}}}


def _gauge_axis_objects(min_value: float | None, max_value: float | None, target_value: float | None) -> list[dict[str, Any]]:
    properties: dict[str, Any] = {}
    if min_value is not None:
        properties["min"] = _decimal_literal(min_value)
    if max_value is not None:
        properties["max"] = _decimal_literal(max_value)
    if target_value is not None:
        properties["target"] = _decimal_literal(target_value)
    if not properties:
        return []
    return [{"properties": properties}]


_VISUAL_FORMAT_TYPES = frozenset({"auto", "text", "bool", "int", "decimal", "color", "raw"})


def _encode_visual_format_value(value: Any, hint: str | None = None) -> Any:
    """Encode a Python value as a Power BI visual format property.

    ``hint`` (optional, one of ``text``, ``bool``, ``int``, ``decimal``,
    ``color``, ``raw``) selects the literal form. ``auto`` (default) infers
    from the Python type. ``raw`` returns the value untouched so callers can
    pass an already-shaped dict (e.g. a measure binding).
    """
    if hint is not None and hint not in _VISUAL_FORMAT_TYPES:
        raise PowerBIValidationError(
            f"unknown property type hint '{hint}'.",
            details={"hint": hint, "allowed": sorted(_VISUAL_FORMAT_TYPES)},
        )
    if hint == "raw":
        return value
    if hint == "color":
        if not isinstance(value, str):
            raise PowerBIValidationError("color values must be strings.", details={"value": repr(value)})
        return _solid_color(value)
    if hint == "text":
        return _text_literal(str(value))
    if hint == "bool":
        return _literal_value(bool(value))
    if hint == "int":
        return _int_literal(int(value))
    if hint == "decimal":
        return _decimal_literal(float(value))
    # auto
    if isinstance(value, bool):
        return _literal_value(value)
    if isinstance(value, int):
        return _int_literal(value)
    if isinstance(value, float):
        return _decimal_literal(value)
    if isinstance(value, str):
        if HEX_COLOR_RE.match(value):
            return _solid_color(value)
        return _text_literal(value)
    if isinstance(value, dict):
        # Allow callers to pass already-shaped expr/Measure/etc. payloads
        return value
    raise PowerBIValidationError(
        f"cannot encode value of type {type(value).__name__} for visual format property.",
        details={"value": repr(value)},
    )


def _datapoint_fill_objects(fill_color: str | None, target_color: str | None) -> list[dict[str, Any]]:
    properties: dict[str, Any] = {}
    if fill_color is not None:
        properties["fill"] = _solid_color(fill_color)
    if target_color is not None:
        properties["targetFill"] = _solid_color(target_color)
    if not properties:
        return []
    return [{"properties": properties}]


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
        extra_objects = extra_single_visual.get("objects")
        if isinstance(extra_objects, dict):
            single_visual["objects"].update(extra_objects)
        for key, val in extra_single_visual.items():
            if key == "objects":
                continue
            single_visual[key] = val
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
    manager: Any | None = None,
) -> dict[str, Any]:
    # Pre-flight role validation: catches "role not allowed for this visual
    # type" and (with manager) "wrong reference kind in role" before we ever
    # write a layout that PBI Desktop would refuse to render.
    _validate_projection_roles(visual_type, projections, manager=manager)
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


def _extract_pbix_zip_natively(pbix: Path, target: Path) -> dict[str, Any]:
    """Fallback PBIX extraction using the standard ZIP. Used when the bundled
    pbi-tools.core does not support 'extract' (it only ships 'compile').

    Copies the Report payload (Layout, StaticResources/Themes) so downstream
    layout-touching tools work. The data model stays inside the PBIX —
    consumers needing model definitions should rely on the live TOM
    connection via pbi_connect.
    """
    target.mkdir(parents=True, exist_ok=True)
    extracted: list[str] = []
    layout_path = target / LAYOUT_RELATIVE_PATH
    layout_path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(pbix, "r") as zf:
        names = set(zf.namelist())
        if "Report/Layout" in names:
            layout_path.write_bytes(zf.read("Report/Layout"))
            extracted.append("Report/Layout")
        # Copy theme JSONs so apply_theme / build_dashboard with theme references resolve.
        for name in names:
            if name.startswith("Report/StaticResources/") and not name.endswith("/"):
                dest = target / name
                dest.parent.mkdir(parents=True, exist_ok=True)
                dest.write_bytes(zf.read(name))
                extracted.append(name)
    return {"method": "zip_native", "extracted_entries": extracted}


def pbi_extract_report_tool(pbix_path: str, extract_folder: str | None = None) -> dict[str, Any]:
    def _impl() -> dict[str, Any]:
        pbix = _resolve_pbix_path(pbix_path, must_exist=True)
        target = _resolve_extract_folder(str(extract_folder or pbix.with_name(f"{pbix.stem}_extracted")), must_exist=False)
        target.mkdir(parents=True, exist_ok=True)
        method = "pbi_tools_extract"
        try:
            _run_pbi_tools(["extract", str(pbix), "-extractFolder", str(target), "-modelSerialization", "Legacy"])
        except (VisualToolError, PBIToolsNotInstalledError) as exc:
            # pbi-tools.core (bundled) only ships the 'compile' action — the
            # CLI returns "Unknown action: 'extract'" or "No action was
            # specified". Fall back to a native ZIP-based extraction so the
            # tool stays usable for layout-touching workflows.
            details = getattr(exc, "details", {}) or {}
            stdout = str(details.get("stdout", "")) + str(details.get("stderr", ""))
            cli_lacks_extract = (
                "Unknown action" in stdout
                or "No action was specified" in stdout
                or isinstance(exc, PBIToolsNotInstalledError)
            )
            if not cli_lacks_extract:
                raise
            logger.info(
                "pbi-tools CLI cannot extract (likely the .core build); falling back to ZIP-native extraction."
            )
            fallback = _extract_pbix_zip_natively(pbix, target)
            method = fallback["method"]
        layout_path = target / LAYOUT_RELATIVE_PATH
        if not layout_path.exists():
            # Defensive last-mile fallback: even if the CLI reported success,
            # verify the layout landed and rebuild from the ZIP if not.
            _extract_pbix_zip_natively(pbix, target)
            method = method + "+zip_native_fallback"
        _, layout = _load_layout(target)
        pages = [_page_summary(section) for section in layout.get("sections", [])]
        return ok(
            "Report extracted successfully.",
            pbix_path=str(pbix),
            extract_folder=str(target),
            extraction_method=method,
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


def pbi_convert_visual_type_tool(
    extract_folder: str,
    page: str,
    visual_id: str,
    new_type: str,
) -> dict[str, Any]:
    """Migrate an existing visual to a different type while preserving
    compatible field bindings.

    Compatibility groups (bindings preserved as-is within a group):
    - ``card`` ↔ ``kpi`` (Values ↔ Indicator)
    - ``clusteredBarChart`` ↔ ``clusteredColumnChart`` ↔ ``lineChart`` ↔
      ``lineClusteredColumnComboChart`` (Category/Y/Series identical)
    - ``donutChart`` ↔ ``treemap``

    Raises ``PowerBIValidationError`` with a clear ``reason`` when the source
    and target are incompatible. Use this so an LLM can recover from "I picked
    the wrong visual type" without losing bindings.
    """
    new_type_clean = str(new_type).strip()
    if not new_type_clean:
        raise PowerBIValidationError("new_type must be non-empty.")
    if new_type_clean not in VISUAL_FIELD_ROLES:
        raise PowerBIValidationError(
            f"Unknown target visual type '{new_type_clean}'.",
            details={"new_type": new_type_clean, "known_types": sorted(VISUAL_FIELD_ROLES)},
        )

    # Source-target role rewrites. Each entry maps source role → target role.
    # If a source role isn't covered, the conversion is rejected (no silent loss).
    COMPATIBILITY: dict[tuple[str, str], dict[str, str]] = {
        ("card", "kpi"): {"Values": "Indicator"},
        ("kpi", "card"): {"Indicator": "Values"},
        ("donutChart", "treemap"): {"Category": "Category", "Y": "Y"},
        ("treemap", "donutChart"): {"Category": "Category", "Y": "Y"},
    }
    # Charts that share Category/Y/Series — identity rewrites are filled
    # automatically.
    chart_family = {"clusteredBarChart", "clusteredColumnChart", "lineChart", "lineClusteredColumnComboChart"}
    for src in chart_family:
        for tgt in chart_family:
            if src != tgt:
                COMPATIBILITY.setdefault((src, tgt), {"Category": "Category", "Y": "Y", "Series": "Series"})

    def _impl() -> dict[str, Any]:
        folder, layout = _load_layout(extract_folder)
        section = _find_page(layout, page)
        _, container, config = _find_visual(section, visual_id)
        sv = config.setdefault("singleVisual", {})
        old_type = str(sv.get("visualType") or "")
        if old_type == new_type_clean:
            return ok(
                f"Visual '{visual_id}' already has type '{new_type_clean}'; nothing to do.",
                visual_id=visual_id,
                old_type=old_type,
                new_type=new_type_clean,
                changed=False,
            )
        rewrite = COMPATIBILITY.get((old_type, new_type_clean))
        if rewrite is None:
            raise PowerBIValidationError(
                f"No compatible role rewrite from '{old_type}' to '{new_type_clean}'.",
                details={
                    "old_type": old_type,
                    "new_type": new_type_clean,
                    "supported_targets_from_old": sorted({tgt for (src, tgt) in COMPATIBILITY if src == old_type}),
                    "reason": "incompatible",
                },
            )

        old_projections = sv.get("projections", {}) or {}
        new_projections: dict[str, list[dict[str, str]]] = {}
        unmapped_roles: list[str] = []
        for role, items in old_projections.items():
            target_role = rewrite.get(role)
            if target_role is None:
                unmapped_roles.append(role)
                continue
            new_projections[target_role] = items

        if unmapped_roles:
            raise PowerBIValidationError(
                f"Source role(s) {unmapped_roles} not mapped by '{old_type}' → '{new_type_clean}'. Aborting to avoid silent data loss.",
                details={"unmapped_roles": unmapped_roles, "rewrite": rewrite},
            )

        # Validate the new projections against the new visual type's role schema.
        _validate_projection_roles(new_type_clean, new_projections)

        sv["visualType"] = new_type_clean
        sv["projections"] = new_projections
        # Refresh the embedded query so the visualType change is reflected in the query payload too.
        prototype_query = sv.get("prototypeQuery") or {"Version": 2, "From": [], "Select": []}
        sv["prototypeQuery"] = prototype_query
        container["config"] = _dump_embedded_json(config)
        query_payload = {"Commands": [{"SemanticQueryDataShapeCommand": {"Query": prototype_query}}]}
        container["query"] = _dump_embedded_json(query_payload)
        _save_layout(folder, layout)
        return ok(
            f"Visual '{visual_id}' converted: {old_type} → {new_type_clean}.",
            visual_id=visual_id,
            old_type=old_type,
            new_type=new_type_clean,
            role_rewrites=rewrite,
            changed=True,
        )

    return _run(_impl)


def pbi_auto_grid_layout_tool(
    specs: list[dict[str, Any]],
    *,
    cols: int = 4,
    gap: int = 16,
    start_x: int = 20,
    start_y: int = 60,
    cell_width: int | None = None,
    cell_height: int | None = None,
    page_width: int = DEFAULT_PAGE_WIDTH,
    page_height: int = DEFAULT_PAGE_HEIGHT,
) -> dict[str, Any]:
    """Compute non-overlapping (x, y, width, height) for a list of visual
    specs on a column-based grid.

    No layout writes happen here — the function is pure and offline. Each
    input ``spec`` is returned annotated with ``x``, ``y``, ``width``,
    ``height``. ``cell_width`` defaults to ``(page_width - 2*start_x - gap*(cols-1)) / cols``
    so the grid fits the page width. ``cell_height`` defaults to 200.
    Specs may set ``col_span`` / ``row_span`` to grow over neighbours; the
    walker advances per cell so spans never overlap.
    """
    if not isinstance(specs, list) or not specs:
        raise PowerBIValidationError("specs must be a non-empty list of visual configs.")
    if cols < 1:
        raise PowerBIValidationError("cols must be >= 1.", details={"cols": cols})
    if gap < 0:
        raise PowerBIValidationError("gap must be >= 0.", details={"gap": gap})

    # Derive cell sizes.
    usable_width = max(0, page_width - 2 * start_x - gap * max(0, cols - 1))
    cw = int(cell_width if cell_width is not None else (usable_width // cols if cols else usable_width))
    if cw <= 0:
        raise PowerBIValidationError(
            "Computed cell_width is non-positive; reduce cols or start_x or pass cell_width explicitly.",
            details={"page_width": page_width, "start_x": start_x, "gap": gap, "cols": cols},
        )
    ch = int(cell_height) if cell_height is not None else 200

    # Track which cells are occupied so spans don't overlap.
    occupied: set[tuple[int, int]] = set()
    placed: list[dict[str, Any]] = []
    cursor_row = 0
    cursor_col = 0

    def _next_free_cell(row: int, col: int) -> tuple[int, int]:
        while (row, col) in occupied:
            col += 1
            if col >= cols:
                col = 0
                row += 1
        return row, col

    for index, spec in enumerate(specs):
        if not isinstance(spec, dict):
            raise PowerBIValidationError(
                f"specs[{index}] must be a dict, got {type(spec).__name__}.",
                details={"index": index},
            )
        col_span = max(1, int(spec.get("col_span", 1)))
        if col_span > cols:
            col_span = cols
        row_span = max(1, int(spec.get("row_span", 1)))

        # Find a row/col where the whole span fits.
        cursor_row, cursor_col = _next_free_cell(cursor_row, cursor_col)
        while cursor_col + col_span > cols or any(
            (cursor_row + r, cursor_col + c) in occupied
            for r in range(row_span)
            for c in range(col_span)
        ):
            cursor_col += 1
            if cursor_col + col_span > cols:
                cursor_col = 0
                cursor_row += 1
            cursor_row, cursor_col = _next_free_cell(cursor_row, cursor_col)

        for r in range(row_span):
            for c in range(col_span):
                occupied.add((cursor_row + r, cursor_col + c))

        x = start_x + cursor_col * (cw + gap)
        y = start_y + cursor_row * (ch + gap)
        width = cw * col_span + gap * (col_span - 1)
        height = ch * row_span + gap * (row_span - 1)
        placed_spec = dict(spec)
        placed_spec.update({"x": x, "y": y, "width": width, "height": height})
        placed.append(placed_spec)

        cursor_col += col_span
        if cursor_col >= cols:
            cursor_col = 0
            cursor_row += 1

    total_height_used = start_y + (cursor_row + 1) * (ch + gap)
    return ok(
        f"Auto-grid: positioned {len(placed)} spec(s) on a {cols}-column grid.",
        cols=cols,
        gap=gap,
        cell_width=cw,
        cell_height=ch,
        start_x=start_x,
        start_y=start_y,
        page_width=page_width,
        page_height=page_height,
        used_height_estimate=total_height_used,
        specs=placed,
    )


def pbi_describe_page_tool(
    extract_folder: str,
    page: str,
    *,
    manager: Any | None = None,
) -> dict[str, Any]:
    """Return a structured, LLM-friendly snapshot of a report page.

    Unlike ``pbi_get_page_tool`` (which exposes the raw projections), this
    surface returns one entry per visual with:

    - ``id``, ``type``, ``position`` (x, y, width, height)
    - ``bindings`` mapping each role (Values, Category, Y, …) to a flat list
      of query refs
    - ``formatting``: extracted title / X axis title / Y axis title /
      ``label_display_units`` when present
    - ``binding_health``: ``ok`` | ``missing_field`` | ``wrong_role`` based on
      live-model validation when ``manager`` is supplied; ``unchecked``
      otherwise

    Use this so an LLM can introspect the current page without having to
    parse ``Report/Layout`` JSON itself.
    """
    def _impl() -> dict[str, Any]:
        folder, layout = _load_layout(extract_folder)
        section = _find_page(layout, page)
        measure_home_map = _scan_measure_home_tables(folder)
        model_fields, _ = _live_model_field_index(manager, include_hidden=False) if manager else (None, {"status": "unavailable"})

        visuals: list[dict[str, Any]] = []
        for container in section.get("visualContainers", []) or []:
            cfg = _parse_embedded_json(container.get("config"), {})
            sv = cfg.get("singleVisual", {}) if isinstance(cfg, dict) else {}
            visual_type = str(sv.get("visualType") or "unknown")
            projections = sv.get("projections", {}) or {}
            bindings: dict[str, list[str]] = {}
            for role, items in projections.items():
                refs: list[str] = []
                if isinstance(items, list):
                    for item in items:
                        ref = item.get("queryRef") if isinstance(item, dict) else None
                        if isinstance(ref, str) and ref:
                            refs.append(ref)
                bindings[role] = refs

            objects = sv.get("objects", {}) if isinstance(sv, dict) else {}
            def _extract_literal_text(obj_name: str, prop: str) -> str | None:
                entries = objects.get(obj_name)
                if not isinstance(entries, list) or not entries:
                    return None
                props = entries[0].get("properties", {}) if isinstance(entries[0], dict) else {}
                value = props.get(prop, {})
                literal = value.get("expr", {}).get("Literal", {}).get("Value") if isinstance(value, dict) else None
                if not isinstance(literal, str):
                    return None
                # Power BI text literals are wrapped in single quotes ('Value').
                if len(literal) >= 2 and literal[0] == "'" and literal[-1] == "'":
                    return literal[1:-1].replace("''", "'")
                return literal

            formatting: dict[str, Any] = {}
            title_text = _extract_literal_text("title", "text")
            if title_text is not None:
                formatting["title"] = title_text
            x_axis_title = _extract_literal_text("categoryAxis", "titleText") or _extract_literal_text("categoryAxis", "axisTitle")
            if x_axis_title is not None:
                formatting["x_axis_title"] = x_axis_title
            y_axis_title = _extract_literal_text("valueAxis", "titleText") or _extract_literal_text("valueAxis", "axisTitle")
            if y_axis_title is not None:
                formatting["y_axis_title"] = y_axis_title
            labels = objects.get("labels")
            if isinstance(labels, list) and labels:
                lu_value = labels[0].get("properties", {}).get("labelDisplayUnits", {}) if isinstance(labels[0], dict) else {}
                lu_literal = lu_value.get("expr", {}).get("Literal", {}).get("Value") if isinstance(lu_value, dict) else None
                if lu_literal is not None:
                    formatting["label_display_units"] = lu_literal

            issues, _ = _visual_binding_issues(container, str(section.get("displayName") or section.get("name", "")), measure_home_map, model_fields)
            if not issues:
                health = "ok"
            else:
                # Roll up the most actionable issue type into a single label.
                kinds = {item.get("issue") for item in issues}
                if "live_model_missing" in kinds or "live_model_unknown_field" in kinds:
                    health = "missing_field"
                elif any(k and "role" in k for k in kinds):
                    health = "wrong_role"
                else:
                    health = "issues"

            visuals.append({
                "id": str(cfg.get("name", "")),
                "type": visual_type,
                "position": {
                    "x": int(container.get("x", 0)),
                    "y": int(container.get("y", 0)),
                    "width": int(container.get("width", 0)),
                    "height": int(container.get("height", 0)),
                },
                "bindings": bindings,
                "formatting": formatting,
                "binding_health": health,
                "issues": issues,
            })

        return ok(
            f"Page '{section.get('displayName')}' described — {len(visuals)} visual(s).",
            extract_folder=str(folder),
            page={
                "name": str(section.get("name", "")),
                "display_name": str(section.get("displayName", "")),
                "width": int(section.get("width", DEFAULT_PAGE_WIDTH)),
                "height": int(section.get("height", DEFAULT_PAGE_HEIGHT)),
                "visual_count": len(visuals),
            },
            visuals=visuals,
        )

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


def pbi_add_card_tool(
    extract_folder: str,
    page: str,
    measure: str,
    x: int,
    y: int,
    width: int = 200,
    height: int = 120,
    title: str = "",
    *,
    manager: Any | None = None,
) -> dict[str, Any]:
    """Add a card visual.

    If a connection ``manager`` is supplied, the measure name is checked
    against the live model first and the call fails fast on a typo. Without
    a manager the tool still works (offline mode), preserving prior behavior.
    """
    _validate_field_references_live(manager, [measure])
    measure_home_map = _resolve_measure_home_map(extract_folder, manager=manager)
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
    *,
    manager: Any | None = None,
) -> dict[str, Any]:
    measure_home_map = _resolve_measure_home_map(extract_folder, manager=manager)
    projections = {"Category": [{"queryRef": _query_ref(category_column)}], "Y": [{"queryRef": _query_ref(value_measure)}]}
    references = [category_column, value_measure]
    if legend_column:
        projections["Series"] = [{"queryRef": _query_ref(legend_column)}]
        references.append(legend_column)
    _validate_field_references_live(manager, references)
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
    *,
    manager: Any | None = None,
) -> dict[str, Any]:
    if not value_measures:
        raise PowerBIValidationError("value_measures must contain at least one measure.")
    _validate_field_references_live(manager, [axis_column, *value_measures])
    measure_home_map = _resolve_measure_home_map(extract_folder, manager=manager)
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


def pbi_add_donut_chart_tool(extract_folder: str, page: str, category_column: str, value_measure: str, x: int, y: int, width: int = 320, height: int = 280, title: str = "", *, manager: Any | None = None) -> dict[str, Any]:
    _validate_field_references_live(manager, [category_column, value_measure])
    measure_home_map = _resolve_measure_home_map(extract_folder, manager=manager)
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


def pbi_add_table_visual_tool(extract_folder: str, page: str, columns: list[str], x: int, y: int, width: int = 520, height: int = 320, title: str = "", *, manager: Any | None = None) -> dict[str, Any]:
    if not columns:
        raise PowerBIValidationError("columns must contain at least one field or measure.")
    _validate_field_references_live(manager, list(columns))
    measure_home_map = _resolve_measure_home_map(extract_folder, manager=manager)
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


def pbi_add_waterfall_tool(extract_folder: str, page: str, category_column: str, value_measure: str, x: int, y: int, width: int = 420, height: int = 300, title: str = "", *, manager: Any | None = None) -> dict[str, Any]:
    _validate_field_references_live(manager, [category_column, value_measure])
    measure_home_map = _resolve_measure_home_map(extract_folder, manager=manager)
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


def pbi_add_slicer_tool(extract_folder: str, page: str, column: str, x: int, y: int, width: int = 220, height: int = 120, slicer_type: str = "dropdown", *, manager: Any | None = None) -> dict[str, Any]:
    slicer_kind = slicer_type.strip().casefold()
    if slicer_kind not in {"dropdown", "list", "range", "tile"}:
        raise PowerBIValidationError("slicer_type must be one of: dropdown, list, range, tile.", details={"slicer_type": slicer_type})
    _validate_field_references_live(manager, [column])
    measure_home_map = _resolve_measure_home_map(extract_folder, manager=manager)
    if slicer_kind == "tile":
        # Horizontal tile slicer: native list type with horizontal orientation flag.
        emitted_kind = "list"
        extra: dict[str, Any] = {
            "slicerType": emitted_kind,
            "objects": {
                "general": [
                    {"properties": {"orientation": _int_literal(1)}}
                ]
            },
        }
    else:
        extra = {"slicerType": slicer_kind}
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
            extra_single_visual=extra,
        ),
        measure_home_map,
    )


def pbi_add_gauge_tool(
    extract_folder: str,
    page: str,
    measure: str,
    x: int,
    y: int,
    width: int = 280,
    height: int = 220,
    title: str = "",
    target_measure: str | None = None,
    *,
    min_value: float | None = None,
    max_value: float | None = None,
    target_value: float | None = None,
    fill_color: str | None = None,
    target_color: str | None = None,
    fill_color_measure: str | None = None,
    manager: Any | None = None,
) -> dict[str, Any]:
    """Add a gauge visual.

    ``fill_color_measure`` (optional): name of a DAX measure that returns a
    ``"#RRGGBB"`` string. The gauge arc fill becomes a measure-binding to it
    (conditional formatting), so the colour reacts to slicer / filter context.
    Mutually exclusive with ``fill_color``; if both are provided, the measure
    binding wins.

    If a connection ``manager`` is supplied, every measure / column reference
    (Y, target, fill_color_measure) is verified against the live model so a
    typo fails fast instead of producing a broken visual.
    """
    if fill_color and fill_color_measure:
        # Measure binding takes precedence — drop the static fill silently to keep callers tidy.
        fill_color = None
    refs_to_validate: list[str] = [measure]
    if target_measure:
        refs_to_validate.append(target_measure)
    if fill_color_measure:
        refs_to_validate.append(fill_color_measure)
    _validate_field_references_live(manager, refs_to_validate)
    measure_home_map = _resolve_measure_home_map(extract_folder, manager=manager)
    projections = {"Y": [{"queryRef": _query_ref(measure)}]}
    references = [measure]
    if target_measure:
        projections["Goal"] = [{"queryRef": _query_ref(target_measure)}]
        references.append(target_measure)
    if fill_color_measure:
        # Pull the color measure into the visual's prototypeQuery so PBI can resolve the fill binding.
        references.append(fill_color_measure)
    extra_objects: dict[str, Any] = {}
    axis_obj = _gauge_axis_objects(min_value, max_value, target_value)
    if axis_obj:
        extra_objects["axis"] = axis_obj
    fill_obj = _datapoint_fill_objects(fill_color, target_color)
    if fill_color_measure:
        # Conditional fill via measure binding — overrides any static fill_color above.
        host_table = measure_home_map.get(fill_color_measure) or "$Measures"
        properties: dict[str, Any] = {}
        if fill_obj and fill_obj[0].get("properties"):
            properties = dict(fill_obj[0]["properties"])
        properties["fill"] = {
            "solid": {
                "color": {
                    "expr": {
                        "Measure": {
                            "Expression": {"SourceRef": {"Entity": host_table}},
                            "Property": fill_color_measure,
                        }
                    }
                }
            }
        }
        fill_obj = [{"properties": properties}]
    if fill_obj:
        extra_objects["dataPoint"] = fill_obj
    extra_single_visual = {"objects": extra_objects} if extra_objects else None
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
            extra_single_visual=extra_single_visual,
        ),
        measure_home_map,
    )


def pbi_add_labelled_card_tool(
    extract_folder: str,
    page: str,
    measure: str,
    label: str,
    x: int,
    y: int,
    width: int = 220,
    height: int = 110,
    *,
    label_height: int = 28,
    label_font_size: int = 11,
    label_bold: bool = True,
    label_color: str = "#1F2937",
    manager: Any | None = None,
) -> dict[str, Any]:
    """Place a text label above a card value, matching docx-style 'label-on-top' card layout.

    Returns both created visuals so callers can move/style them together later.
    If a connection ``manager`` is supplied, the measure name is verified
    against the live model before either visual is created.
    """
    if label_height <= 0 or label_height >= height:
        raise PowerBIValidationError(
            "label_height must be > 0 and smaller than height.",
            details={"label_height": label_height, "height": height},
        )
    _validate_field_references_live(manager, [measure])
    label_response = pbi_add_text_box_tool(
        extract_folder,
        page,
        label,
        x,
        y,
        width,
        label_height,
        font_size=label_font_size,
        bold=label_bold,
        color=label_color,
    )
    if not label_response.get("ok"):
        return label_response
    card_response = pbi_add_card_tool(
        extract_folder,
        page,
        measure,
        x,
        y + label_height,
        width,
        height - label_height,
        title="",
    )
    if not card_response.get("ok"):
        return card_response
    return ok(
        f"Labelled card '{label}' added.",
        page=card_response.get("page"),
        visuals={
            "label": label_response.get("visual"),
            "value": card_response.get("visual"),
        },
    )


def pbi_add_scatter_chart_tool(
    extract_folder: str,
    page: str,
    category_column: str,
    x_measure: str,
    y_measure: str,
    x: int,
    y: int,
    width: int = 420,
    height: int = 320,
    title: str = "",
    size_measure: str | None = None,
    legend_column: str | None = None,
    *,
    manager: Any | None = None,
) -> dict[str, Any]:
    """Add a scatter chart visual (``scatterChart``).

    Roles: ``Category`` (column — dot identity), ``X`` (measure), ``Y``
    (measure), ``Size`` (measure, optional), ``Series`` (column, optional).
    Use for correlation analysis between two measures grouped by a dimension.
    """
    projections = {
        "Category": [{"queryRef": _query_ref(category_column)}],
        "X": [{"queryRef": _query_ref(x_measure)}],
        "Y": [{"queryRef": _query_ref(y_measure)}],
    }
    references = [category_column, x_measure, y_measure]
    if size_measure:
        projections["Size"] = [{"queryRef": _query_ref(size_measure)}]
        references.append(size_measure)
    if legend_column:
        projections["Series"] = [{"queryRef": _query_ref(legend_column)}]
        references.append(legend_column)
    _validate_field_references_live(manager, references)
    measure_home_map = _resolve_measure_home_map(extract_folder, manager=manager)
    return _append_visual(
        extract_folder,
        page,
        lambda section, home_map: _create_chart_container(
            section,
            visual_type="scatterChart",
            x=x,
            y=y,
            width=width,
            height=height,
            title=title or None,
            projections=projections,
            references=references,
            measure_home_map=home_map,
            manager=manager,
        ),
        measure_home_map,
    )


def pbi_add_combo_chart_tool(
    extract_folder: str,
    page: str,
    category_column: str,
    bar_measures: list[str],
    line_measures: list[str],
    x: int,
    y: int,
    width: int = 480,
    height: int = 320,
    title: str = "",
    legend_column: str | None = None,
    *,
    manager: Any | None = None,
) -> dict[str, Any]:
    """Add a combo chart (column bars + line overlay).

    Roles: ``Category`` (column), ``Y`` (measure list — bars), ``Y2``
    (measure list — line). Useful for "actual vs target" with a target line
    over actual bars.
    """
    if not bar_measures:
        raise PowerBIValidationError("bar_measures must contain at least one measure.")
    if not line_measures:
        raise PowerBIValidationError("line_measures must contain at least one measure.")
    projections: dict[str, list[dict[str, str]]] = {
        "Category": [{"queryRef": _query_ref(category_column)}],
        "Y": [{"queryRef": _query_ref(item)} for item in bar_measures],
        "Y2": [{"queryRef": _query_ref(item)} for item in line_measures],
    }
    references = [category_column, *bar_measures, *line_measures]
    if legend_column:
        projections["Series"] = [{"queryRef": _query_ref(legend_column)}]
        references.append(legend_column)
    _validate_field_references_live(manager, references)
    measure_home_map = _resolve_measure_home_map(extract_folder, manager=manager)
    return _append_visual(
        extract_folder,
        page,
        lambda section, home_map: _create_chart_container(
            section,
            visual_type="lineClusteredColumnComboChart",
            x=x,
            y=y,
            width=width,
            height=height,
            title=title or None,
            projections=projections,
            references=references,
            measure_home_map=home_map,
            manager=manager,
        ),
        measure_home_map,
    )


def pbi_add_kpi_tool(
    extract_folder: str,
    page: str,
    indicator_measure: str,
    trend_axis_column: str,
    x: int,
    y: int,
    width: int = 240,
    height: int = 160,
    title: str = "",
    goal_measure: str | None = None,
    direction: str = "high_is_good",
    *,
    manager: Any | None = None,
) -> dict[str, Any]:
    """Add a native KPI visual (``kpi``).

    Roles: ``Indicator`` (measure — current value), ``TrendLine`` (column,
    typically a Date), ``Goal`` (measure, optional). ``direction`` controls
    the status colour interpretation: ``"high_is_good"`` means the green
    threshold sits above ``Goal``; ``"low_is_good"`` flips it.
    """
    if direction not in {"high_is_good", "low_is_good"}:
        raise PowerBIValidationError(
            "direction must be 'high_is_good' or 'low_is_good'.",
            details={"direction": direction},
        )
    projections: dict[str, list[dict[str, str]]] = {
        "Indicator": [{"queryRef": _query_ref(indicator_measure)}],
        "TrendLine": [{"queryRef": _query_ref(trend_axis_column)}],
    }
    references = [indicator_measure, trend_axis_column]
    if goal_measure:
        projections["Goal"] = [{"queryRef": _query_ref(goal_measure)}]
        references.append(goal_measure)
    _validate_field_references_live(manager, references)
    # Encode direction in the visual's objects so PBI's KPI rendering picks the right colour rule.
    extra_single_visual = {
        "objects": {
            "indicator": [
                {
                    "properties": {
                        "directionType": _text_literal(
                            "Increasing" if direction == "high_is_good" else "Decreasing"
                        )
                    }
                }
            ]
        }
    }
    measure_home_map = _resolve_measure_home_map(extract_folder, manager=manager)
    return _append_visual(
        extract_folder,
        page,
        lambda section, home_map: _create_chart_container(
            section,
            visual_type="kpi",
            x=x,
            y=y,
            width=width,
            height=height,
            title=title or None,
            projections=projections,
            references=references,
            measure_home_map=home_map,
            extra_single_visual=extra_single_visual,
            manager=manager,
        ),
        measure_home_map,
    )


def pbi_add_matrix_tool(
    extract_folder: str,
    page: str,
    rows: list[str],
    values: list[str],
    x: int,
    y: int,
    columns: list[str] | None = None,
    width: int = 540,
    height: int = 360,
    title: str = "",
    subtotals: bool = True,
    column_layout: str = "stepped",
    *,
    manager: Any | None = None,
) -> dict[str, Any]:
    """Add a matrix / pivot-table visual (``pivotTable``).

    Roles: ``Rows`` (column list — required), ``Columns`` (column list —
    optional), ``Values`` (measure list — required). ``column_layout``
    accepts ``"stepped"`` (compact, single column with indents) or
    ``"tabular"`` (one column per row level).
    """
    if not rows:
        raise PowerBIValidationError("rows must contain at least one column.")
    if not values:
        raise PowerBIValidationError("values must contain at least one measure.")
    layout_token = column_layout.strip().casefold()
    if layout_token not in {"stepped", "tabular"}:
        raise PowerBIValidationError(
            "column_layout must be 'stepped' or 'tabular'.",
            details={"column_layout": column_layout},
        )
    projections: dict[str, list[dict[str, str]]] = {
        "Rows": [{"queryRef": _query_ref(item)} for item in rows],
        "Values": [{"queryRef": _query_ref(item)} for item in values],
    }
    references = [*rows, *values]
    if columns:
        projections["Columns"] = [{"queryRef": _query_ref(item)} for item in columns]
        references.extend(columns)
    _validate_field_references_live(manager, references)
    extra_single_visual = {
        "objects": {
            "subTotals": [
                {"properties": {"rowSubtotals": _literal_value(bool(subtotals))}}
            ],
            "general": [
                {"properties": {"layout": _text_literal("Stepped" if layout_token == "stepped" else "Tabular")}}
            ],
        }
    }
    measure_home_map = _resolve_measure_home_map(extract_folder, manager=manager)
    return _append_visual(
        extract_folder,
        page,
        lambda section, home_map: _create_chart_container(
            section,
            visual_type="pivotTable",
            x=x,
            y=y,
            width=width,
            height=height,
            title=title or None,
            projections=projections,
            references=references,
            measure_home_map=home_map,
            extra_single_visual=extra_single_visual,
            manager=manager,
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
    # Text boxes have no field references so manager-augmented home table is unused.
    measure_home_map = _resolve_measure_home_map(extract_folder)
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


def pbi_set_visual_format_property_tool(
    extract_folder: str,
    page: str,
    visual_id: str,
    object_name: str,
    properties: dict[str, Any],
    property_types: dict[str, str] | None = None,
) -> dict[str, Any]:
    """Set formatting properties on an existing visual's ``singleVisual.objects[<object_name>][0].properties``.

    Use to override titles, axis labels, data labels, etc. on a visual that's
    already on the report — without rebuilding it. Encodes Python values as
    proper Power BI literals (text in single quotes, ints with ``L`` suffix,
    decimals with ``D`` suffix, ``#RRGGBB`` as solid color, bool as
    ``true``/``false``).

    ``property_types`` lets you force the encoding of a specific property:
    ``"text"``, ``"bool"``, ``"int"``, ``"decimal"``, ``"color"``, or ``"raw"``
    (pass-through).

    Existing properties under the same object are preserved and merged with
    the new values (write-through semantics).
    """
    def _impl() -> dict[str, Any]:
        if not object_name or not str(object_name).strip():
            raise PowerBIValidationError("object_name must be non-empty.")
        if not isinstance(properties, dict) or not properties:
            raise PowerBIValidationError(
                "properties must be a non-empty dict of {property_name: value}.",
                details={"properties": repr(properties)},
            )
        types_map = {k: str(v) for k, v in (property_types or {}).items()}

        folder, layout = _load_layout(extract_folder)
        section = _find_page(layout, page)
        _, container, config = _find_visual(section, visual_id)
        single_visual = config.setdefault("singleVisual", {})
        objects = single_visual.setdefault("objects", {})
        existing = objects.get(object_name) or [{"properties": {}}]
        if not isinstance(existing, list) or not existing:
            existing = [{"properties": {}}]
        merged_props = dict(existing[0].get("properties", {}))

        encoded: dict[str, Any] = {}
        for prop_name, raw_value in properties.items():
            if not prop_name or not str(prop_name).strip():
                raise PowerBIValidationError(
                    "property names must be non-empty strings.",
                    details={"name": repr(prop_name)},
                )
            hint = types_map.get(prop_name)
            encoded[prop_name] = _encode_visual_format_value(raw_value, hint=hint)
        merged_props.update(encoded)
        existing[0]["properties"] = merged_props
        objects[object_name] = existing
        container["config"] = _dump_embedded_json(config)
        _save_layout(folder, layout)
        return ok(
            f"Format properties applied to visual '{visual_id}' (object '{object_name}').",
            extract_folder=str(folder),
            page=str(section.get("displayName") or section.get("name")),
            visual_id=visual_id,
            object=object_name,
            applied=sorted(encoded.keys()),
        )

    return _run(_impl)


def pbi_disable_card_autoscale_tool(
    extract_folder: str,
    page: str | None = None,
    visual_ids: list[str] | None = None,
    label_precision: int = 0,
) -> dict[str, Any]:
    """Disable the auto K/M/B unit-scaling on card visuals.

    By default Power BI cards display large numeric values rescaled (e.g.
    ``119,229`` becomes ``119K``). When the underlying measure already uses a
    custom format string with ``K`` / ``€`` suffixes the result is the
    classic "119K K €" double-suffix bug. This tool sets ``labelDisplayUnits=1``
    (None) plus an explicit ``labelPrecision`` on every card on the report
    (or restricted to ``page`` / ``visual_ids``).

    ``label_precision``: 0 keeps integer display, raise to 2 if you need the
    decimal portion of the underlying measure to render.
    """
    def _impl() -> dict[str, Any]:
        if visual_ids is not None and not isinstance(visual_ids, list):
            raise PowerBIValidationError("visual_ids must be a list of visual ids or None.")
        ids_filter = {str(v).strip() for v in (visual_ids or []) if str(v).strip()}
        folder, layout = _load_layout(extract_folder)

        target_sections: list[dict[str, Any]]
        if page:
            target_sections = [_find_page(layout, page)]
        else:
            target_sections = list(layout.get("sections", []) or [])

        patched: list[dict[str, Any]] = []
        for section in target_sections:
            for container in section.get("visualContainers", []) or []:
                cfg_raw = container.get("config")
                if not isinstance(cfg_raw, str):
                    continue
                cfg = json.loads(cfg_raw)
                sv = cfg.get("singleVisual", {})
                if sv.get("visualType") != "card":
                    continue
                visual_id = str(cfg.get("name", ""))
                if ids_filter and visual_id not in ids_filter:
                    continue
                objects = sv.setdefault("objects", {})
                labels = objects.get("labels")
                if not isinstance(labels, list) or not labels:
                    labels = [{"properties": {}}]
                props = dict(labels[0].get("properties", {}))
                props["labelDisplayUnits"] = _decimal_literal(1)  # 1 = None
                props["labelPrecision"] = _decimal_literal(int(label_precision))
                labels[0]["properties"] = props
                objects["labels"] = labels
                container["config"] = _dump_embedded_json(cfg)
                patched.append({
                    "visual_id": visual_id,
                    "page": str(section.get("displayName") or section.get("name", "")),
                })
        _save_layout(folder, layout)
        return ok(
            f"Disabled autoscale on {len(patched)} card visual(s).",
            extract_folder=str(folder),
            patched=patched,
            patched_count=len(patched),
            label_precision=int(label_precision),
        )

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
    return pbi_add_gauge_tool(
        extract,
        page,
        measure,
        x,
        y,
        w,
        h,
        title,
        cfg.get("target_measure"),
        min_value=cfg.get("min_value"),
        max_value=cfg.get("max_value"),
        target_value=cfg.get("target_value"),
        fill_color=cfg.get("fill_color"),
        target_color=cfg.get("target_color"),
        fill_color_measure=cfg.get("fill_color_measure"),
    )


def _dispatch_labelled_card(extract, page, x, y, w, h, title, cfg):
    measure = cfg.get("measure")
    label = cfg.get("label") or title
    if not measure or not label:
        raise PowerBIValidationError(
            "labelled_card requires config.measure and config.label (or title)",
            details={"visual_type": "labelled_card"},
        )
    return pbi_add_labelled_card_tool(
        extract,
        page,
        measure,
        str(label),
        x,
        y,
        w,
        h,
        label_height=int(cfg.get("label_height", 28)),
        label_font_size=int(cfg.get("label_font_size", 11)),
        label_bold=bool(cfg.get("label_bold", True)),
        label_color=str(cfg.get("label_color", "#1F2937")),
    )


def _dispatch_map(extract, page, x, y, w, h, title, cfg):
    location = cfg.get("location") or cfg.get("category_column") or cfg.get("category")
    measure = cfg.get("measure") or cfg.get("value_measure")
    if not location:
        raise PowerBIValidationError(
            "map requires config.location (Table.Column with the geographic field)",
            details={"visual_type": "map"},
        )
    return pbi_add_map_tool(extract, page, location, measure, x, y, w, h, title)


def pbi_add_map_tool(
    extract_folder: str,
    page: str,
    location_column: str,
    value_measure: str | None,
    x: int,
    y: int,
    width: int = 420,
    height: int = 320,
    title: str = "",
    *,
    manager: Any | None = None,
) -> dict[str, Any]:
    """Add a bubble/map visual (``map``).

    Roles: ``Category`` (column with the geographic field — country, city,
    Lat/Long…) and optional ``Y`` (measure that drives bubble size).

    Provides feature parity with ``pbi_build_dashboard``'s ``map`` spec so
    callers can use the simpler ``pbi_add_visual_tool(visual_type="map", …)``
    surface without dropping into the dashboard builder.
    """
    projections: dict[str, list[dict[str, str]]] = {
        "Category": [{"queryRef": _query_ref(location_column)}]
    }
    references = [location_column]
    if value_measure:
        projections["Y"] = [{"queryRef": _query_ref(value_measure)}]
        references.append(value_measure)
    _validate_field_references_live(manager, references)
    measure_home_map = _resolve_measure_home_map(extract_folder, manager=manager)
    return _append_visual(
        extract_folder,
        page,
        lambda section, home_map: _create_chart_container(
            section,
            visual_type="map",
            x=x,
            y=y,
            width=width,
            height=height,
            title=title or None,
            projections=projections,
            references=references,
            measure_home_map=home_map,
            manager=manager,
        ),
        measure_home_map,
    )


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
    "labelled_card": _dispatch_labelled_card,
    "labeled_card": _dispatch_labelled_card,
    "bar_chart": _dispatch_bar,
    "line_chart": _dispatch_line,
    "donut": _dispatch_donut,
    "table": _dispatch_table,
    "waterfall": _dispatch_waterfall,
    "slicer": _dispatch_slicer,
    "gauge": _dispatch_gauge,
    "map": _dispatch_map,
    "text_box": _dispatch_text_box,
    "textbox": _dispatch_text_box,
})


__all__ = [
    "pbi_add_visual_tool",
    "pbi_add_bar_chart_tool",
    "pbi_add_card_tool",
    "pbi_add_combo_chart_tool",
    "pbi_add_donut_chart_tool",
    "pbi_add_gauge_tool",
    "pbi_add_kpi_tool",
    "pbi_add_labelled_card_tool",
    "pbi_add_line_chart_tool",
    "pbi_add_map_tool",
    "pbi_add_matrix_tool",
    "pbi_add_scatter_chart_tool",
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
    "pbi_auto_grid_layout_tool",
    "pbi_convert_visual_type_tool",
    "pbi_describe_page_tool",
    "pbi_disable_card_autoscale_tool",
    "pbi_remove_visual_tool",
    "pbi_set_page_size_tool",
    "pbi_set_visual_format_property_tool",
    "pbi_validate_report_fields_tool",
]
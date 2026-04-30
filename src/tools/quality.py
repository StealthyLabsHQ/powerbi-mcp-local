"""Quality gates for Power BI models, DAX, report layout, and scenario runs."""

from __future__ import annotations

import json
import zipfile
from datetime import datetime
from pathlib import Path
from typing import Any

from pbi_connection import PowerBIValidationError, dax_quote_table_name, ok
from security import resolve_local_path


LAYOUT_RELATIVE_PATH = Path("Report") / "Layout"
MIN_VISUAL_WIDTH = 120
MIN_VISUAL_HEIGHT = 80
MAX_VISUALS_PER_PAGE = 12
DATE_PARSE_FORMATS = (
    "%Y-%m-%d",
    "%Y/%m/%d",
    "%d/%m/%Y",
    "%m/%d/%Y",
    "%d-%m-%Y",
    "%Y%m%d",
)


def _load_layout(extract_folder: str) -> tuple[Path, dict[str, Any]]:
    folder = resolve_local_path(extract_folder, must_exist=True)
    layout_path = folder / LAYOUT_RELATIVE_PATH
    if not layout_path.exists():
        raise PowerBIValidationError("Report/Layout was not found.", details={"extract_folder": str(folder)})
    return folder, json.loads(layout_path.read_bytes().decode("utf-16-le"))


def _visual_config(container: dict[str, Any]) -> dict[str, Any]:
    raw = container.get("config", "{}")
    if isinstance(raw, str):
        try:
            return json.loads(raw)
        except json.JSONDecodeError:
            return {}
    return raw if isinstance(raw, dict) else {}


def _visual_name(container: dict[str, Any]) -> str:
    return str(_visual_config(container).get("name", ""))


def _visual_type(container: dict[str, Any]) -> str:
    cfg = _visual_config(container)
    return str((cfg.get("singleVisual") or {}).get("visualType", ""))


def _visual_has_title(container: dict[str, Any]) -> bool:
    single = (_visual_config(container).get("singleVisual") or {})
    objects = single.get("objects") or {}
    title = objects.get("title") or []
    return bool(title)


def _bounds(container: dict[str, Any]) -> tuple[float, float, float, float]:
    return (
        float(container.get("x", 0) or 0),
        float(container.get("y", 0) or 0),
        float(container.get("width", 0) or 0),
        float(container.get("height", 0) or 0),
    )


def _overlap_area(a: dict[str, Any], b: dict[str, Any]) -> float:
    ax, ay, aw, ah = _bounds(a)
    bx, by, bw, bh = _bounds(b)
    x_overlap = max(0.0, min(ax + aw, bx + bw) - max(ax, bx))
    y_overlap = max(0.0, min(ay + ah, by + bh) - max(ay, by))
    return x_overlap * y_overlap


def _model_snapshot(manager: Any, *, include_hidden: bool = False) -> dict[str, Any]:
    from .model import pbi_model_info_tool

    return pbi_model_info_tool(manager, include_hidden=include_hidden, include_row_counts=False)


def _model_audit_from_snapshot(snapshot: dict[str, Any]) -> dict[str, Any]:
    tables = snapshot.get("tables", [])
    relationships = snapshot.get("relationships", [])
    measures = snapshot.get("measures", [])
    issues: list[dict[str, Any]] = []
    warnings: list[dict[str, Any]] = []

    visible_tables = {item["name"]: item for item in tables if not item.get("is_hidden")}
    related_tables: set[str] = set()
    graph: dict[str, set[str]] = {name: set() for name in visible_tables}
    pair_count: dict[tuple[str, str], int] = {}

    for rel in relationships:
        from_table = str(rel.get("from_table", ""))
        to_table = str(rel.get("to_table", ""))
        related_tables.update({from_table, to_table})
        if from_table in graph and to_table in graph:
            graph[from_table].add(to_table)
            graph[to_table].add(from_table)
        pair = tuple(sorted((from_table, to_table)))
        pair_count[pair] = pair_count.get(pair, 0) + 1
        direction = str(rel.get("direction", ""))
        if direction and direction.casefold() not in {"onedirection", "single"}:
            issues.append({"type": "bidirectional_relationship", "relationship": rel})
        if pair_count[pair] > 1:
            warnings.append({"type": "parallel_relationships", "tables": list(pair), "count": pair_count[pair]})

    if len(visible_tables) > 1:
        for table_name in sorted(set(visible_tables) - related_tables):
            warnings.append({"type": "unrelated_table", "table": table_name})

    for table_name, table in visible_tables.items():
        date_columns = [
            col["name"]
            for col in table.get("columns", [])
            if "date" in str(col.get("name", "")).casefold() or str(col.get("data_type", "")).casefold() == "datetime"
        ]
        if date_columns and table_name not in related_tables and len(visible_tables) > 1:
            warnings.append({"type": "unrelated_date_columns", "table": table_name, "columns": date_columns})

    for a, neighbors in graph.items():
        for b in neighbors:
            for c in neighbors.intersection(graph.get(b, set())):
                if a < b < c:
                    warnings.append({"type": "ambiguous_relationship_triangle", "tables": [a, b, c]})

    measure_tables = {str(item.get("table", "")): [] for item in measures}
    for measure in measures:
        measure_tables.setdefault(str(measure.get("table", "")), []).append(measure)
    for table in visible_tables:
        if table not in related_tables and not measure_tables.get(table) and len(visible_tables) > 1:
            warnings.append({"type": "orphan_table", "table": table})

    return {
        "valid": not issues,
        "issue_count": len(issues),
        "warning_count": len(warnings),
        "issues": issues,
        "warnings": warnings,
    }


def _dax_column(table: str, column: str) -> str:
    return f"{dax_quote_table_name(table)}[{column.replace(']', ']]')}]"


def _row_value(row: dict[str, Any], alias: str) -> Any:
    return row.get(alias, row.get(f"[{alias}]"))


def _table_map(snapshot: dict[str, Any]) -> dict[str, dict[str, Any]]:
    return {str(item.get("name", "")).casefold(): item for item in snapshot.get("tables", [])}


def _find_table(snapshot: dict[str, Any], table: str) -> dict[str, Any] | None:
    return _table_map(snapshot).get(table.casefold())


def _find_column(snapshot: dict[str, Any], table: str, column: str) -> dict[str, Any] | None:
    found = _find_table(snapshot, table)
    if not found:
        return None
    for item in found.get("columns", []):
        if str(item.get("name", "")).casefold() == column.casefold():
            return item
    return None


def _column_profile(manager: Any, table: str, column: str) -> dict[str, Any]:
    query = (
        "EVALUATE ROW("
        '"__Rows", COUNTROWS(' + dax_quote_table_name(table) + "), "
        '"__Distinct", DISTINCTCOUNT(' + _dax_column(table, column) + "), "
        '"__Blank", COUNTBLANK(' + _dax_column(table, column) + ")"
        ")"
    )
    result = manager.run_adomd_query(query, max_rows=1)
    row = result.get("rows", [{}])[0] if result.get("rows") else {}
    return {
        "row_count": _row_value(row, "__Rows"),
        "distinct_count": _row_value(row, "__Distinct"),
        "blank_count": _row_value(row, "__Blank"),
    }


def _graph_paths(graph: dict[str, set[str]], start: str, end: str, *, limit: int = 2) -> list[list[str]]:
    paths: list[list[str]] = []

    def _walk(node: str, target: str, seen: list[str]) -> None:
        if len(paths) >= limit:
            return
        if node == target:
            paths.append(seen[:])
            return
        for nxt in sorted(graph.get(node, set())):
            if nxt not in seen:
                _walk(nxt, target, [*seen, nxt])

    _walk(start, end, [start])
    return paths


def _duplicate_relationship_key_issues(manager: Any, relationships: list[dict[str, Any]]) -> list[dict[str, Any]]:
    issues: list[dict[str, Any]] = []
    checked: set[tuple[str, str]] = set()
    candidates: list[tuple[str, str, str]] = []
    for rel in relationships:
        candidates.append((str(rel.get("to_table", "")), str(rel.get("to_column", "")), "one_side"))
        from_table = str(rel.get("from_table", ""))
        to_table = str(rel.get("to_table", ""))
        if to_table.casefold().startswith("fact") and not from_table.casefold().startswith("fact"):
            candidates.append((from_table, str(rel.get("from_column", "")), "non_fact_many_side"))
    for table, column, role in candidates:
        key = (table.casefold(), column.casefold())
        if key in checked:
            continue
        checked.add(key)
        query = (
            "EVALUATE ROW("
            '"__Rows", COUNTROWS(' + dax_quote_table_name(table) + "), "
            '"__Distinct", DISTINCTCOUNT(' + _dax_column(table, column) + ")"
            ")"
        )
        try:
            result = manager.run_adomd_query(query, max_rows=1)
        except Exception as exc:
            issues.append({"type": "relationship_key_check_failed", "table": table, "column": column, "error": str(exc)})
            continue
        rows = result.get("rows", [])
        if not rows:
            continue
        row_count = rows[0].get("__Rows", rows[0].get("[__Rows]"))
        distinct_count = rows[0].get("__Distinct", rows[0].get("[__Distinct]"))
        if row_count is not None and distinct_count is not None and row_count != distinct_count:
            issues.append(
                {
                    "type": "duplicate_relationship_key",
                    "table": table,
                    "column": column,
                    "relationship_role": role,
                    "row_count": row_count,
                    "distinct_count": distinct_count,
                }
            )
    return issues


def pbi_audit_model_tool(manager: Any, *, include_hidden: bool = False) -> dict[str, Any]:
    """Detect missing, ambiguous, bidirectional, and orphaned model structures."""
    snapshot = _model_snapshot(manager, include_hidden=include_hidden)
    audit = _model_audit_from_snapshot(snapshot)
    duplicate_key_issues = _duplicate_relationship_key_issues(manager, snapshot.get("relationships", []))
    if duplicate_key_issues:
        audit["issues"].extend(duplicate_key_issues)
        audit["issue_count"] = len(audit["issues"])
        audit["valid"] = False
    return ok(
        f"Model audit found {audit['issue_count']} issue(s), {audit['warning_count']} warning(s).",
        include_hidden=include_hidden,
        table_count=len(snapshot.get("tables", [])),
        measure_count=len(snapshot.get("measures", [])),
        relationship_count=len(snapshot.get("relationships", [])),
        **audit,
    )


def pbi_lint_dax_tool(manager: Any, *, include_hidden: bool = False, validate_expressions: bool = True) -> dict[str, Any]:
    """Validate measures, formats, and measure/column name collisions."""
    from .query import pbi_validate_dax_tool

    snapshot = _model_snapshot(manager, include_hidden=include_hidden)
    column_names_by_table = {
        table["name"]: {str(col.get("name", "")).casefold() for col in table.get("columns", [])}
        for table in snapshot.get("tables", [])
    }
    issues: list[dict[str, Any]] = []
    warnings: list[dict[str, Any]] = []

    for measure in snapshot.get("measures", []):
        table = str(measure.get("table", ""))
        name = str(measure.get("name", ""))
        expression = str(measure.get("expression", "") or "").strip()
        if name.casefold() in column_names_by_table.get(table, set()):
            issues.append({"type": "measure_column_name_collision", "table": table, "measure": name})
        if not expression:
            issues.append({"type": "empty_measure_expression", "table": table, "measure": name})
        if not str(measure.get("format_string", "") or "").strip() and not measure.get("is_hidden"):
            warnings.append({"type": "missing_format_string", "table": table, "measure": name})
        if validate_expressions and expression:
            result = pbi_validate_dax_tool(manager, expression=f"[{name}]", kind="scalar")
            if not result.get("valid"):
                issues.append({"type": "invalid_measure_dax", "table": table, "measure": name, "error": result.get("error")})

    return ok(
        f"DAX lint found {len(issues)} issue(s), {len(warnings)} warning(s).",
        include_hidden=include_hidden,
        validate_expressions=validate_expressions,
        valid=not issues,
        issue_count=len(issues),
        warning_count=len(warnings),
        issues=issues,
        warnings=warnings,
    )


def pbi_detect_name_collisions_tool(manager: Any, *, include_hidden: bool = False) -> dict[str, Any]:
    """Detect table, column, and measure name collisions before writes."""
    snapshot = _model_snapshot(manager, include_hidden=include_hidden)
    issues: list[dict[str, Any]] = []
    warnings: list[dict[str, Any]] = []
    table_names: dict[str, list[str]] = {}
    global_columns: dict[str, list[dict[str, str]]] = {}
    global_measures: dict[str, list[dict[str, str]]] = {}

    for table in snapshot.get("tables", []):
        table_name = str(table.get("name", ""))
        table_names.setdefault(table_name.casefold(), []).append(table_name)
        local_columns: dict[str, list[str]] = {}
        for column in table.get("columns", []):
            column_name = str(column.get("name", ""))
            local_columns.setdefault(column_name.casefold(), []).append(column_name)
            global_columns.setdefault(column_name.casefold(), []).append({"table": table_name, "column": column_name})
        for key, names in local_columns.items():
            if len(names) > 1:
                issues.append({"type": "duplicate_column_name", "table": table_name, "name": names[0], "count": len(names)})

    for measure in snapshot.get("measures", []):
        table = str(measure.get("table", ""))
        name = str(measure.get("name", ""))
        global_measures.setdefault(name.casefold(), []).append({"table": table, "measure": name})

    for names in table_names.values():
        if len(names) > 1:
            issues.append({"type": "duplicate_table_name", "name": names[0], "count": len(names)})
    for name, measures in global_measures.items():
        if len(measures) > 1:
            warnings.append({"type": "duplicate_measure_name", "name": measures[0]["measure"], "measures": measures})
        for measure in measures:
            same_table_columns = [
                item for item in global_columns.get(name, [])
                if item["table"].casefold() == measure["table"].casefold()
            ]
            if same_table_columns:
                issues.append({"type": "measure_column_name_collision", "table": measure["table"], "measure": measure["measure"], "columns": same_table_columns})
    for name, columns in global_columns.items():
        tables = {item["table"].casefold() for item in columns}
        if len(tables) > 1:
            warnings.append({"type": "same_column_name_across_tables", "name": columns[0]["column"], "columns": columns})

    return ok(
        f"Name collision scan found {len(issues)} issue(s), {len(warnings)} warning(s).",
        include_hidden=include_hidden,
        valid=not issues,
        issue_count=len(issues),
        warning_count=len(warnings),
        issues=issues,
        warnings=warnings,
    )


def pbi_detect_dirty_dates_tool(
    manager: Any,
    *,
    table: str | None = None,
    max_samples: int = 200,
    min_parse_success_rate: float = 0.8,
    scan_all_text_columns: bool = False,
) -> dict[str, Any]:
    """Detect text columns that look like dirty dates."""
    snapshot = _model_snapshot(manager, include_hidden=False)
    if max_samples < 1 or max_samples > 1000:
        raise PowerBIValidationError("max_samples must be between 1 and 1000.", details={"max_samples": max_samples})
    if not 0 <= min_parse_success_rate <= 1:
        raise PowerBIValidationError("min_parse_success_rate must be between 0 and 1.", details={"min_parse_success_rate": min_parse_success_rate})

    tables = snapshot.get("tables", [])
    if table:
        found = _find_table(snapshot, table)
        if found is None:
            raise PowerBIValidationError("Table was not found.", details={"table": table})
        tables = [found]

    issues: list[dict[str, Any]] = []
    warnings: list[dict[str, Any]] = []
    scanned: list[dict[str, Any]] = []
    for item in tables:
        table_name = str(item.get("name", ""))
        for column in item.get("columns", []):
            column_name = str(column.get("name", ""))
            data_type = str(column.get("data_type", ""))
            is_text = data_type.casefold() in {"string", "text"}
            if not is_text:
                continue
            name_suggests_date = "date" in column_name.casefold()
            if not (scan_all_text_columns or name_suggests_date):
                continue
            query = (
                "EVALUATE TOPN("
                + str(max_samples)
                + ", SELECTCOLUMNS("
                + dax_quote_table_name(table_name)
                + ', "__Value", '
                + _dax_column(table_name, column_name)
                + "))"
            )
            try:
                result = manager.run_adomd_query(query, max_rows=max_samples)
            except Exception as exc:
                issues.append({"type": "dirty_date_scan_failed", "table": table_name, "column": column_name, "error": str(exc)})
                continue
            values = [str(_row_value(row, "__Value") or "").strip() for row in result.get("rows", [])]
            non_blank = [value for value in values if value]
            parsed = 0
            formats: set[str] = set()
            invalid_examples: list[str] = []
            for value in non_blank:
                matched = False
                for fmt in DATE_PARSE_FORMATS:
                    try:
                        datetime.strptime(value, fmt)
                        parsed += 1
                        formats.add(fmt)
                        matched = True
                        break
                    except ValueError:
                        pass
                if not matched and len(invalid_examples) < 5:
                    invalid_examples.append(value)
            parse_rate = parsed / len(non_blank) if non_blank else 0.0
            profile = {
                "table": table_name,
                "column": column_name,
                "sample_count": len(values),
                "non_blank_count": len(non_blank),
                "blank_count": len(values) - len(non_blank),
                "parse_success_rate": round(parse_rate, 4),
                "formats": sorted(formats),
                "invalid_examples": invalid_examples,
            }
            scanned.append(profile)
            if name_suggests_date and (not non_blank or parse_rate < min_parse_success_rate):
                issues.append({"type": "dirty_text_date", **profile})
            elif len(formats) > 1:
                warnings.append({"type": "mixed_text_date_formats", **profile})

    return ok(
        f"Dirty date scan found {len(issues)} issue(s), {len(warnings)} warning(s).",
        table=table,
        max_samples=max_samples,
        min_parse_success_rate=min_parse_success_rate,
        valid=not issues,
        issue_count=len(issues),
        warning_count=len(warnings),
        issues=issues,
        warnings=warnings,
        scanned_columns=scanned,
    )


def pbi_validate_relationship_plan_tool(
    manager: Any,
    *,
    from_table: str,
    from_column: str,
    to_table: str,
    to_column: str,
    cardinality: str = "oneToMany",
    direction: str = "oneDirection",
    is_active: bool = True,
) -> dict[str, Any]:
    """Validate relationship cardinality, direction, duplicates, and ambiguity before creation."""
    snapshot = _model_snapshot(manager, include_hidden=False)
    issues: list[dict[str, Any]] = []
    warnings: list[dict[str, Any]] = []
    for table_name, column_name, side in (
        (from_table, from_column, "from"),
        (to_table, to_column, "to"),
    ):
        if _find_table(snapshot, table_name) is None:
            issues.append({"type": "table_not_found", "side": side, "table": table_name})
        elif _find_column(snapshot, table_name, column_name) is None:
            issues.append({"type": "column_not_found", "side": side, "table": table_name, "column": column_name})
    if issues:
        return ok(
            f"Relationship plan found {len(issues)} issue(s), {len(warnings)} warning(s).",
            valid=False,
            safe_to_create=False,
            issue_count=len(issues),
            warning_count=len(warnings),
            issues=issues,
            warnings=warnings,
        )

    normalized_direction = direction.casefold()
    normalized_cardinality = cardinality.casefold()
    if normalized_direction not in {"onedirection", "single"}:
        issues.append({"type": "unsafe_filter_direction", "direction": direction})
    if normalized_cardinality not in {"onetomany", "manytoone", "manytomany", "onetoone"}:
        issues.append({"type": "unknown_cardinality", "cardinality": cardinality})
    if normalized_cardinality == "manytomany":
        issues.append({"type": "many_to_many_relationship", "cardinality": cardinality})

    existing = snapshot.get("relationships", [])
    for rel in existing:
        endpoints_match = (
            str(rel.get("from_table", "")).casefold() == from_table.casefold()
            and str(rel.get("from_column", "")).casefold() == from_column.casefold()
            and str(rel.get("to_table", "")).casefold() == to_table.casefold()
            and str(rel.get("to_column", "")).casefold() == to_column.casefold()
        )
        if endpoints_match:
            issues.append({"type": "duplicate_relationship", "relationship": rel})

    from_profile = _column_profile(manager, from_table, from_column)
    to_profile = _column_profile(manager, to_table, to_column)
    if (from_profile.get("blank_count") or 0) > 0:
        warnings.append({"type": "from_column_has_blanks", "table": from_table, "column": from_column, **from_profile})
    if (to_profile.get("blank_count") or 0) > 0:
        warnings.append({"type": "to_column_has_blanks", "table": to_table, "column": to_column, **to_profile})

    from_unique = from_profile.get("row_count") == from_profile.get("distinct_count")
    to_unique = to_profile.get("row_count") == to_profile.get("distinct_count")
    if normalized_cardinality in {"onetomany", "manytoone"} and not (from_unique or to_unique):
        issues.append({"type": "no_unique_relationship_side", "from_profile": from_profile, "to_profile": to_profile})
    if normalized_cardinality == "onetoone" and not (from_unique and to_unique):
        issues.append({"type": "one_to_one_requires_both_unique", "from_profile": from_profile, "to_profile": to_profile})

    if is_active:
        graph: dict[str, set[str]] = {}
        for rel in existing:
            if not bool(rel.get("is_active", rel.get("active", True))):
                continue
            a = str(rel.get("from_table", ""))
            b = str(rel.get("to_table", ""))
            graph.setdefault(a, set()).add(b)
            graph.setdefault(b, set()).add(a)
        paths = _graph_paths(graph, from_table, to_table, limit=2)
        if paths:
            warnings.append({"type": "relationship_creates_parallel_path", "existing_paths": paths})
        graph.setdefault(from_table, set()).add(to_table)
        graph.setdefault(to_table, set()).add(from_table)
        if len(_graph_paths(graph, from_table, to_table, limit=3)) > 2:
            warnings.append({"type": "relationship_ambiguity_risk", "from_table": from_table, "to_table": to_table})

    return ok(
        f"Relationship plan found {len(issues)} issue(s), {len(warnings)} warning(s).",
        valid=not issues,
        safe_to_create=not issues,
        issue_count=len(issues),
        warning_count=len(warnings),
        issues=issues,
        warnings=warnings,
        profiles={"from": from_profile, "to": to_profile},
        proposed_relationship={
            "from_table": from_table,
            "from_column": from_column,
            "to_table": to_table,
            "to_column": to_column,
            "cardinality": cardinality,
            "direction": direction,
            "is_active": is_active,
        },
    )


def pbi_lint_report_layout_tool(extract_folder: str, page: str | None = None) -> dict[str, Any]:
    """Detect overlaps, excessive whitespace, tiny visuals, and missing titles."""
    folder, layout = _load_layout(extract_folder)
    issues: list[dict[str, Any]] = []
    warnings: list[dict[str, Any]] = []
    sections = layout.get("sections", [])
    for section in sections:
        if page and page not in {str(section.get("name")), str(section.get("displayName"))}:
            continue
        page_name = str(section.get("displayName") or section.get("name") or "")
        width = float(section.get("width", 1280) or 1280)
        height = float(section.get("height", 720) or 720)
        containers = [item for item in section.get("visualContainers", []) if isinstance(item, dict)]
        if len(containers) > MAX_VISUALS_PER_PAGE:
            warnings.append({"type": "too_many_visuals", "page": page_name, "count": len(containers), "limit": MAX_VISUALS_PER_PAGE})
        used_area = 0.0
        for index, container in enumerate(containers):
            x, y, visual_width, visual_height = _bounds(container)
            name = _visual_name(container) or f"visual_{index}"
            visual_type = _visual_type(container)
            used_area += visual_width * visual_height
            if visual_type not in {"textbox", "slicer"} and (visual_width < MIN_VISUAL_WIDTH or visual_height < MIN_VISUAL_HEIGHT):
                warnings.append({"type": "visual_too_small", "page": page_name, "visual": name, "width": visual_width, "height": visual_height})
            if visual_type not in {"textbox", "slicer"} and not _visual_has_title(container):
                warnings.append({"type": "missing_title", "page": page_name, "visual": name, "visual_type": visual_type})
            if x < 0 or y < 0 or x + visual_width > width or y + visual_height > height:
                issues.append({"type": "visual_outside_canvas", "page": page_name, "visual": name})
            for other in containers[index + 1:]:
                area = _overlap_area(container, other)
                if area > 1:
                    issues.append({"type": "visual_overlap", "page": page_name, "visual_a": name, "visual_b": _visual_name(other), "area": round(area, 2)})
        density = used_area / max(width * height, 1)
        if density < 0.35 and containers:
            warnings.append({"type": "excessive_whitespace", "page": page_name, "density": round(density, 3)})
        if density > 0.9:
            warnings.append({"type": "layout_overloaded", "page": page_name, "density": round(density, 3)})

    return ok(
        f"Layout lint found {len(issues)} issue(s), {len(warnings)} warning(s).",
        extract_folder=str(folder),
        page=page,
        valid=not issues,
        issue_count=len(issues),
        warning_count=len(warnings),
        issues=issues,
        warnings=warnings,
    )


def pbi_validate_visual_bindings_tool(
    extract_folder: str,
    page: str | None = None,
    include_hidden: bool = False,
    manager: Any | None = None,
) -> dict[str, Any]:
    """Alias-level visual binding validation with clearer tool naming."""
    from .visuals import pbi_validate_report_fields_tool

    return pbi_validate_report_fields_tool(
        extract_folder,
        page=page,
        include_hidden=include_hidden,
        manager=manager,
    )


def _measure_aliases(measures: list[str]) -> str:
    return ", ".join(f'"__M{idx}", {measure}' for idx, measure in enumerate(measures))


def _filtered_table_query(table_expression: str, filter_expression: str | None, max_rows: int) -> str:
    if filter_expression:
        return f"EVALUATE TOPN({max_rows}, CALCULATETABLE({table_expression}, {filter_expression}))"
    return f"EVALUATE TOPN({max_rows}, {table_expression})"


def pbi_validate_filter_expression_tool(manager: Any, *, filter_expression: str) -> dict[str, Any]:
    """Validate a DAX boolean filter expression before visual probes."""
    expression = str(filter_expression or "").strip()
    if not expression:
        raise PowerBIValidationError("filter_expression is required.")
    query = f"EVALUATE CALCULATETABLE(ROW(\"__probe\", 1), {expression})"
    try:
        manager.run_adomd_query(query, max_rows=1)
    except Exception as exc:
        return ok(
            "Filter expression is invalid.",
            valid=False,
            filter_expression=expression,
            error=str(exc),
        )
    return ok(
        "Filter expression is valid.",
        valid=True,
        filter_expression=expression,
    )


def _visual_query_parts(container: dict[str, Any]) -> tuple[list[str], list[str]]:
    cfg = _visual_config(container)
    single = cfg.get("singleVisual") or {}
    prototype = single.get("prototypeQuery") or {}
    aliases = {
        str(item.get("Name", "")): str(item.get("Entity", ""))
        for item in prototype.get("From", []) or []
        if isinstance(item, dict)
    }
    columns: list[str] = []
    measures: list[str] = []
    for item in prototype.get("Select", []) or []:
        if not isinstance(item, dict):
            continue
        column = item.get("Column")
        measure = item.get("Measure")
        if isinstance(column, dict):
            alias = str((column.get("Expression") or {}).get("SourceRef", {}).get("Source", ""))
            table = aliases.get(alias, alias)
            name = str(column.get("Property", ""))
            if table and name:
                columns.append(_dax_column(table, name))
        elif isinstance(measure, dict):
            name = str(measure.get("Property", ""))
            if name:
                measures.append(f"[{name.replace(']', ']]')}]")
    return columns, measures


def pbi_detect_empty_visuals_tool(
    manager: Any,
    *,
    extract_folder: str,
    page: str | None = None,
    include_slicers: bool = False,
    max_rows: int = 1,
    filter_expression: str | None = None,
) -> dict[str, Any]:
    """Execute lightweight DAX probes to detect visuals with no data."""
    if max_rows < 1 or max_rows > 10:
        raise PowerBIValidationError("max_rows must be between 1 and 10.", details={"max_rows": max_rows})
    if filter_expression is not None and not filter_expression.strip():
        raise PowerBIValidationError("filter_expression cannot be blank.")
    filter_validation = None
    if filter_expression:
        filter_validation = pbi_validate_filter_expression_tool(manager, filter_expression=filter_expression)
        if not filter_validation.get("valid"):
            return ok(
                "Empty visual scan skipped because filter_expression is invalid.",
                extract_folder=str(resolve_local_path(extract_folder, must_exist=True)),
                page=page,
                include_slicers=include_slicers,
                filter_expression=filter_expression,
                filter_validation=filter_validation,
                valid=False,
                issue_count=1,
                warning_count=0,
                issues=[{"type": "invalid_filter_expression", "error": filter_validation.get("error")}],
                warnings=[],
                checked_visuals=[],
            )
    folder, layout = _load_layout(extract_folder)
    issues: list[dict[str, Any]] = []
    warnings: list[dict[str, Any]] = []
    checked: list[dict[str, Any]] = []

    for section in layout.get("sections", []):
        page_name = str(section.get("displayName") or section.get("name") or "")
        if page and page not in {str(section.get("name")), page_name}:
            continue
        for index, container in enumerate(section.get("visualContainers", []) or []):
            if not isinstance(container, dict):
                continue
            visual_type = _visual_type(container)
            if visual_type == "textbox" or (visual_type == "slicer" and not include_slicers):
                continue
            visual_name = _visual_name(container) or f"visual_{index}"
            columns, measures = _visual_query_parts(container)
            if not columns and not measures:
                warnings.append({"type": "visual_has_no_bindings", "page": page_name, "visual": visual_name, "visual_type": visual_type})
                continue
            if measures:
                aliases = _measure_aliases(measures)
                if columns:
                    table_expression = f"SUMMARIZECOLUMNS({', '.join(columns)}, {aliases})"
                    query = _filtered_table_query(table_expression, filter_expression, max_rows)
                else:
                    if filter_expression:
                        query = f"EVALUATE CALCULATETABLE(ROW({aliases}), {filter_expression})"
                    else:
                        query = f"EVALUATE ROW({aliases})"
            else:
                table_expression = f"SUMMARIZECOLUMNS({', '.join(columns)})"
                query = _filtered_table_query(table_expression, filter_expression, max_rows)
            try:
                result = manager.run_adomd_query(query, max_rows=max_rows)
            except Exception as exc:
                issues.append({"type": "visual_query_failed", "page": page_name, "visual": visual_name, "visual_type": visual_type, "error": str(exc)})
                continue
            rows = result.get("rows", [])
            checked.append({"page": page_name, "visual": visual_name, "visual_type": visual_type, "row_count": len(rows)})
            if not rows:
                issues.append({"type": "empty_visual", "page": page_name, "visual": visual_name, "visual_type": visual_type})
                continue
            if measures:
                measure_values = [
                    value
                    for row in rows
                    for key, value in row.items()
                    if str(key).strip("[]").startswith("__M")
                ]
                if measure_values:
                    non_blank = [value for value in measure_values if value is not None]
                    if not non_blank:
                        warnings.append({"type": "visual_measures_all_blank", "page": page_name, "visual": visual_name, "visual_type": visual_type})
                    elif all(float(value or 0) == 0 for value in non_blank if isinstance(value, (int, float))):
                        warnings.append({"type": "visual_numeric_measures_all_zero", "page": page_name, "visual": visual_name, "visual_type": visual_type})

    return ok(
        f"Empty visual scan found {len(issues)} issue(s), {len(warnings)} warning(s).",
        extract_folder=str(folder),
        page=page,
        include_slicers=include_slicers,
        filter_expression=filter_expression,
        filter_validation=filter_validation,
        valid=not issues,
        issue_count=len(issues),
        warning_count=len(warnings),
        issues=issues,
        warnings=warnings,
        checked_visuals=checked,
    )


def _selected_measures(snapshot: dict[str, Any], measures: list[str] | None) -> list[dict[str, Any]]:
    all_measures = list(snapshot.get("measures", []))
    if not measures:
        return all_measures
    wanted = {item.casefold() for item in measures}
    return [item for item in all_measures if str(item.get("name", "")).casefold() in wanted]


def _measure_ref(name: str) -> str:
    return f"[{name.replace(']', ']]')}]"


def _measure_expected_format(name: str) -> str | None:
    lowered = name.casefold()
    if "coverage" in lowered or "/" in name:
        return "number"
    if "%" in name or "rate" in lowered or "retention" in lowered or "win rate" in lowered:
        return "percent"
    if any(token in lowered for token in ("revenue", "arr", "mrr", "margin", "ltv", "cac", "pipeline", "deal", "spend", "target", "forecast")):
        return "currency"
    return None


def _format_matches(format_string: str, expected: str | None) -> bool:
    if expected is None:
        return True
    fmt = (format_string or "").casefold()
    if expected == "percent":
        return "%" in fmt
    if expected == "currency":
        return "$" in fmt or "€" in fmt or "£" in fmt or "currency" in fmt
    if expected == "number":
        return bool(fmt) and "%" not in fmt and "$" not in fmt and "€" not in fmt and "£" not in fmt
    return True


def pbi_generate_measure_tests_tool(
    manager: Any,
    *,
    measures: list[str] | None = None,
    include_hidden: bool = False,
    max_measures: int = 200,
) -> dict[str, Any]:
    """Generate and execute smoke tests for DAX measures."""
    if max_measures < 1 or max_measures > 500:
        raise PowerBIValidationError("max_measures must be between 1 and 500.", details={"max_measures": max_measures})
    snapshot = _model_snapshot(manager, include_hidden=include_hidden)
    selected = _selected_measures(snapshot, measures)[:max_measures]
    issues: list[dict[str, Any]] = []
    warnings: list[dict[str, Any]] = []
    tests: list[dict[str, Any]] = []
    found = {str(item.get("name", "")).casefold() for item in selected}
    for requested in measures or []:
        if requested.casefold() not in found:
            issues.append({"type": "measure_not_found", "measure": requested})

    for measure in selected:
        name = str(measure.get("name", ""))
        expression = str(measure.get("expression", "") or "")
        format_string = str(measure.get("format_string", "") or "")
        expected_format = _measure_expected_format(name)
        ref = _measure_ref(name)
        query = f'EVALUATE ROW("__Value", {ref})'
        test: dict[str, Any] = {
            "table": measure.get("table"),
            "measure": name,
            "query": query,
            "format_string": format_string,
            "expected_format": expected_format,
        }
        if "/" in expression and "DIVIDE(" not in expression.upper():
            warnings.append({"type": "unsafe_division_operator", "measure": name})
        if not _format_matches(format_string, expected_format):
            warnings.append({"type": "unexpected_measure_format", "measure": name, "format_string": format_string, "expected": expected_format})
        try:
            result = manager.run_adomd_query(query, max_rows=1)
        except Exception as exc:
            issues.append({"type": "measure_execution_failed", "measure": name, "error": str(exc)})
            test["valid"] = False
            test["error"] = str(exc)
            tests.append(test)
            continue
        rows = result.get("rows", [])
        value = _row_value(rows[0], "__Value") if rows else None
        test["valid"] = True
        test["value"] = value
        test["blank"] = value is None
        test["zero"] = isinstance(value, (int, float)) and float(value) == 0.0
        if value is None:
            warnings.append({"type": "measure_returns_blank", "measure": name})
        tests.append(test)

    return ok(
        f"Measure test generation found {len(issues)} issue(s), {len(warnings)} warning(s).",
        include_hidden=include_hidden,
        requested_count=len(measures or []),
        tested_count=len(tests),
        valid=not issues,
        issue_count=len(issues),
        warning_count=len(warnings),
        issues=issues,
        warnings=warnings,
        tests=tests,
    )


def pbi_export_validation_report_tool(
    manager: Any,
    *,
    output_path: str,
    extract_folder: str | None = None,
    include_hidden: bool = False,
    include_empty_visual_scan: bool = False,
    empty_visual_filter_expression: str | None = None,
    include_measure_tests: bool = False,
) -> dict[str, Any]:
    """Export model, DAX, layout, binding, and score validation as JSON."""
    output = resolve_local_path(output_path, must_exist=False, allowed_extensions={".json"})
    output.parent.mkdir(parents=True, exist_ok=True)
    report: dict[str, Any] = {
        "model": pbi_audit_model_tool(manager, include_hidden=include_hidden),
        "dax": pbi_lint_dax_tool(manager, include_hidden=include_hidden),
        "name_collisions": pbi_detect_name_collisions_tool(manager, include_hidden=include_hidden),
        "dirty_dates": pbi_detect_dirty_dates_tool(manager, scan_all_text_columns=False),
    }
    if extract_folder:
        report["layout"] = pbi_lint_report_layout_tool(extract_folder)
        report["visual_bindings"] = pbi_validate_visual_bindings_tool(extract_folder, include_hidden=include_hidden, manager=manager)
        if include_empty_visual_scan:
            report["empty_visuals"] = pbi_detect_empty_visuals_tool(manager, extract_folder=extract_folder, filter_expression=empty_visual_filter_expression)
    if include_measure_tests:
        report["measure_tests"] = pbi_generate_measure_tests_tool(manager, include_hidden=include_hidden)
    report["score"] = pbi_score_dashboard_tool(manager, extract_folder=extract_folder, include_hidden=include_hidden)
    validation_sections = [
        item
        for item in report.values()
        if isinstance(item, dict) and "valid" in item
    ]
    report["summary"] = {
        "overall_valid": all(item.get("valid") for item in validation_sections),
        "score_total": report["score"].get("score_total"),
        "issue_count": sum(int(item.get("issue_count", 0) or 0) for item in validation_sections),
        "warning_count": sum(int(item.get("warning_count", 0) or 0) for item in validation_sections),
        "sections": sorted(report),
    }
    output.write_text(json.dumps(report, indent=2, default=str), encoding="utf-8")
    return ok(
        "Validation report exported successfully.",
        output_path=str(output),
        score_total=report["score"].get("score_total"),
        overall_valid=report["summary"]["overall_valid"],
        issue_count=report["summary"]["issue_count"],
        warning_count=report["summary"]["warning_count"],
        sections=sorted(report),
    )


def _score_parts(model: dict[str, Any], dax: dict[str, Any], layout: dict[str, Any], bindings: dict[str, Any] | None) -> dict[str, int]:
    model_score = max(0, 25 - model.get("issue_count", 0) * 10 - model.get("warning_count", 0) * 2)
    dax_score = max(0, 25 - dax.get("issue_count", 0) * 10 - dax.get("warning_count", 0) * 2)
    layout_score = max(0, 20 - layout.get("issue_count", 0) * 8 - layout.get("warning_count", 0) * 2)
    binding_penalty = 0 if not bindings else bindings.get("issue_count", 0) * 8
    readability_score = max(0, 20 - binding_penalty - model.get("warning_count", 0))
    robustness_score = 10 if model.get("valid") and dax.get("valid") and layout.get("valid") and (not bindings or bindings.get("valid")) else 5
    return {
        "model": int(model_score),
        "dax": int(dax_score),
        "layout": int(layout_score),
        "business_readability": int(readability_score),
        "error_robustness": int(robustness_score),
    }


def pbi_score_dashboard_tool(
    manager: Any,
    *,
    extract_folder: str | None = None,
    include_hidden: bool = False,
) -> dict[str, Any]:
    """Score dashboard quality across model, DAX, layout, and readability."""
    model = pbi_audit_model_tool(manager, include_hidden=include_hidden)
    dax = pbi_lint_dax_tool(manager, include_hidden=include_hidden)
    layout = (
        pbi_lint_report_layout_tool(extract_folder)
        if extract_folder
        else {"valid": True, "issue_count": 0, "warning_count": 0, "issues": [], "warnings": []}
    )
    bindings = (
        pbi_validate_visual_bindings_tool(extract_folder, include_hidden=include_hidden, manager=manager)
        if extract_folder
        else None
    )
    parts = _score_parts(model, dax, layout, bindings)
    total = sum(parts.values())
    return ok(
        "Dashboard scored successfully.",
        score_total=total,
        breakdown=parts,
        model=model,
        dax=dax,
        layout=layout,
        visual_bindings=bindings,
    )


def pbi_run_scenario_tool(
    manager: Any,
    *,
    scenario: str,
    extract_folder: str | None = None,
    include_hidden: bool = False,
) -> dict[str, Any]:
    """Run a complete QA scenario against the active model and optional extracted layout."""
    result = pbi_score_dashboard_tool(manager, extract_folder=extract_folder, include_hidden=include_hidden)
    return ok(
        "Scenario run completed.",
        scenario=scenario,
        score_total=result["score_total"],
        breakdown=result["breakdown"],
        model=result["model"],
        dax=result["dax"],
        layout=result["layout"],
        visual_bindings=result["visual_bindings"],
        patch_required=result["score_total"] < 85,
    )


def pbi_compare_report_versions_tool(
    *,
    extract_folder_a: str,
    extract_folder_b: str,
    label_a: str = "A",
    label_b: str = "B",
) -> dict[str, Any]:
    """Compare two extracted report versions by pages, visuals, and layout score."""
    _, layout_a = _load_layout(extract_folder_a)
    _, layout_b = _load_layout(extract_folder_b)

    def _summary(layout: dict[str, Any]) -> dict[str, Any]:
        pages = layout.get("sections", [])
        visuals = []
        for section in pages:
            for container in section.get("visualContainers", []) or []:
                if isinstance(container, dict):
                    visuals.append({"page": section.get("displayName") or section.get("name"), "name": _visual_name(container), "type": _visual_type(container)})
        return {"page_count": len(pages), "visual_count": len(visuals), "visuals": visuals}

    summary_a = _summary(layout_a)
    summary_b = _summary(layout_b)
    score_a = pbi_lint_report_layout_tool(extract_folder_a)
    score_b = pbi_lint_report_layout_tool(extract_folder_b)
    return ok(
        "Report versions compared successfully.",
        labels={"a": label_a, "b": label_b},
        a=summary_a,
        b=summary_b,
        delta={
            "page_count": summary_b["page_count"] - summary_a["page_count"],
            "visual_count": summary_b["visual_count"] - summary_a["visual_count"],
        },
        layout_lint={label_a: score_a, label_b: score_b},
    )


def _layout_summary(layout: dict[str, Any]) -> dict[str, Any]:
    pages: list[dict[str, Any]] = []
    visual_count = 0
    for section in layout.get("sections", []) or []:
        if not isinstance(section, dict):
            continue
        visuals = [item for item in section.get("visualContainers", []) or [] if isinstance(item, dict)]
        visual_count += len(visuals)
        pages.append(
            {
                "name": section.get("name"),
                "display_name": section.get("displayName") or section.get("name"),
                "visual_count": len(visuals),
            }
        )
    return {"page_count": len(pages), "visual_count": visual_count, "pages": pages}


def pbi_validate_pbix_persistence_tool(
    *,
    pbix_path: str,
    extract_folder: str | None = None,
    require_security_bindings_removed: bool = True,
) -> dict[str, Any]:
    """Validate that a patched PBIX still contains a readable, persistent report layout."""
    pbix = resolve_local_path(pbix_path, must_exist=True, allowed_extensions={".pbix"})
    issues: list[dict[str, Any]] = []
    warnings: list[dict[str, Any]] = []
    pbix_summary: dict[str, Any] | None = None
    extract_summary: dict[str, Any] | None = None

    if not zipfile.is_zipfile(pbix):
        issues.append({"type": "pbix_not_zip", "pbix_path": str(pbix)})
    else:
        with zipfile.ZipFile(pbix, "r") as archive:
            names = set(archive.namelist())
            if "Report/Layout" not in names:
                issues.append({"type": "missing_report_layout", "pbix_path": str(pbix)})
            else:
                try:
                    layout = json.loads(archive.read("Report/Layout").decode("utf-16-le"))
                    pbix_summary = _layout_summary(layout)
                    if pbix_summary["page_count"] == 0:
                        warnings.append({"type": "no_report_pages", "pbix_path": str(pbix)})
                    if pbix_summary["visual_count"] == 0:
                        warnings.append({"type": "no_report_visuals", "pbix_path": str(pbix)})
                except (UnicodeDecodeError, json.JSONDecodeError) as exc:
                    issues.append({"type": "invalid_report_layout", "pbix_path": str(pbix), "error": str(exc)})
            if require_security_bindings_removed and "SecurityBindings" in names:
                issues.append({"type": "security_bindings_present", "pbix_path": str(pbix)})

    if extract_folder:
        _, extract_layout = _load_layout(extract_folder)
        extract_summary = _layout_summary(extract_layout)
        if pbix_summary and extract_summary:
            if pbix_summary["page_count"] != extract_summary["page_count"]:
                issues.append(
                    {
                        "type": "page_count_mismatch",
                        "pbix_count": pbix_summary["page_count"],
                        "extract_count": extract_summary["page_count"],
                    }
                )
            if pbix_summary["visual_count"] != extract_summary["visual_count"]:
                issues.append(
                    {
                        "type": "visual_count_mismatch",
                        "pbix_count": pbix_summary["visual_count"],
                        "extract_count": extract_summary["visual_count"],
                    }
                )

    return ok(
        f"PBIX persistence validation found {len(issues)} issue(s), {len(warnings)} warning(s).",
        pbix_path=str(pbix),
        extract_folder=extract_folder,
        valid=not issues,
        issue_count=len(issues),
        warning_count=len(warnings),
        issues=issues,
        warnings=warnings,
        pbix_summary=pbix_summary,
        extract_summary=extract_summary,
    )


__all__ = [
    "pbi_audit_model_tool",
    "pbi_compare_report_versions_tool",
    "pbi_detect_dirty_dates_tool",
    "pbi_detect_empty_visuals_tool",
    "pbi_detect_name_collisions_tool",
    "pbi_export_validation_report_tool",
    "pbi_generate_measure_tests_tool",
    "pbi_lint_dax_tool",
    "pbi_lint_report_layout_tool",
    "pbi_run_scenario_tool",
    "pbi_score_dashboard_tool",
    "pbi_validate_filter_expression_tool",
    "pbi_validate_pbix_persistence_tool",
    "pbi_validate_relationship_plan_tool",
    "pbi_validate_visual_bindings_tool",
]

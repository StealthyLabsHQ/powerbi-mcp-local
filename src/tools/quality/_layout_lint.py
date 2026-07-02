"""Report layout lint tools: overlaps, bindings, empty/missing visuals, diffs."""

from __future__ import annotations

from typing import Any

from pbi_connection import PowerBIValidationError, ok

from ._shared import (
    MAX_VISUALS_PER_PAGE,
    MIN_VISUAL_HEIGHT,
    MIN_VISUAL_WIDTH,
    _bounds,
    _dax_column,
    _load_layout,
    _overlap_area,
    _visual_config,
    _visual_has_title,
    _visual_name,
    _visual_type,
)


def pbi_lint_report_layout_tool(
    extract_folder: str,
    page: str | None = None,
    *,
    ignore_warnings: list[str] | None = None,
    only_pages: list[str] | None = None,
    max_visuals_per_page: int | None = None,
) -> dict[str, Any]:
    """Detect overlaps, excessive whitespace, tiny visuals, and missing titles.

    Optional knobs to silence noise on intentionally dense pages:

    - ``ignore_warnings``: list of warning types to drop (e.g.
      ``["too_many_visuals", "visual_too_small", "missing_title",
      "excessive_whitespace", "layout_overloaded"]``). Issues are never
      ignored — only warnings.
    - ``only_pages``: restrict the scan to a list of page display names /
      internal names. Use to lint just the new pages an LLM produced.
    - ``max_visuals_per_page``: override the default ``too_many_visuals``
      threshold (defaults to ``MAX_VISUALS_PER_PAGE``). Set high to
      effectively disable the warning without ignoring it.
    """
    folder, layout = _load_layout(extract_folder)
    issues: list[dict[str, Any]] = []
    warnings: list[dict[str, Any]] = []
    sections = layout.get("sections", [])
    ignore_set = {str(item).strip() for item in (ignore_warnings or []) if str(item).strip()}
    pages_filter: set[str] | None = None
    if only_pages:
        pages_filter = {str(item).strip() for item in only_pages if str(item).strip()}
    visual_threshold = int(max_visuals_per_page) if max_visuals_per_page is not None else MAX_VISUALS_PER_PAGE

    def _add_warning(item: dict[str, Any]) -> None:
        if item.get("type") not in ignore_set:
            warnings.append(item)

    for section in sections:
        section_name = str(section.get("name", ""))
        section_display = str(section.get("displayName") or section.get("name") or "")
        if page and page not in {section_name, section_display}:
            continue
        if pages_filter is not None and section_name not in pages_filter and section_display not in pages_filter:
            continue
        page_name = section_display
        width = float(section.get("width", 1280) or 1280)
        height = float(section.get("height", 720) or 720)
        containers = [item for item in section.get("visualContainers", []) if isinstance(item, dict)]
        if len(containers) > visual_threshold:
            _add_warning(
                {"type": "too_many_visuals", "page": page_name, "count": len(containers), "limit": visual_threshold}
            )
        used_area = 0.0
        for index, container in enumerate(containers):
            x, y, visual_width, visual_height = _bounds(container)
            name = _visual_name(container) or f"visual_{index}"
            visual_type = _visual_type(container)
            used_area += visual_width * visual_height
            if visual_type not in {"textbox", "slicer"} and (
                visual_width < MIN_VISUAL_WIDTH or visual_height < MIN_VISUAL_HEIGHT
            ):
                _add_warning(
                    {
                        "type": "visual_too_small",
                        "page": page_name,
                        "visual": name,
                        "width": visual_width,
                        "height": visual_height,
                    }
                )
            if visual_type not in {"textbox", "slicer"} and not _visual_has_title(container):
                _add_warning({"type": "missing_title", "page": page_name, "visual": name, "visual_type": visual_type})
            if x < 0 or y < 0 or x + visual_width > width or y + visual_height > height:
                issues.append({"type": "visual_outside_canvas", "page": page_name, "visual": name})
            for other in containers[index + 1 :]:
                area = _overlap_area(container, other)
                if area > 1:
                    issues.append(
                        {
                            "type": "visual_overlap",
                            "page": page_name,
                            "visual_a": name,
                            "visual_b": _visual_name(other),
                            "area": round(area, 2),
                        }
                    )
        density = used_area / max(width * height, 1)
        if density < 0.35 and containers:
            _add_warning({"type": "excessive_whitespace", "page": page_name, "density": round(density, 3)})
        if density > 0.9:
            _add_warning({"type": "layout_overloaded", "page": page_name, "density": round(density, 3)})

    return ok(
        f"Layout lint found {len(issues)} issue(s), {len(warnings)} warning(s).",
        extract_folder=str(folder),
        page=page,
        valid=not issues,
        issue_count=len(issues),
        warning_count=len(warnings),
        ignored_warnings=sorted(ignore_set) or [],
        max_visuals_per_page=visual_threshold,
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
    from ..visuals import pbi_validate_report_fields_tool

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
    from . import pbi_validate_filter_expression_tool, resolve_local_path

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
                warnings.append(
                    {
                        "type": "visual_has_no_bindings",
                        "page": page_name,
                        "visual": visual_name,
                        "visual_type": visual_type,
                    }
                )
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
                issues.append(
                    {
                        "type": "visual_query_failed",
                        "page": page_name,
                        "visual": visual_name,
                        "visual_type": visual_type,
                        "error": str(exc),
                    }
                )
                continue
            rows = result.get("rows", [])
            checked.append(
                {"page": page_name, "visual": visual_name, "visual_type": visual_type, "row_count": len(rows)}
            )
            if not rows:
                issues.append(
                    {"type": "empty_visual", "page": page_name, "visual": visual_name, "visual_type": visual_type}
                )
                continue
            if measures:
                measure_values = [
                    value for row in rows for key, value in row.items() if str(key).strip("[]").startswith("__M")
                ]
                if measure_values:
                    non_blank = [value for value in measure_values if value is not None]
                    if not non_blank:
                        warnings.append(
                            {
                                "type": "visual_measures_all_blank",
                                "page": page_name,
                                "visual": visual_name,
                                "visual_type": visual_type,
                            }
                        )
                    else:
                        # FIX: don't flag text-returning measures (e.g. FORMAT()) as
                        # "numeric zero". Only emit the warning when ALL non-blank
                        # values are numeric AND all evaluate to 0. Otherwise the
                        # measure returns a non-numeric value (text, date, …) which
                        # is a legitimate non-zero result.
                        all_numeric = all(
                            isinstance(value, (int, float)) and not isinstance(value, bool) for value in non_blank
                        )
                        if all_numeric and all(float(value) == 0 for value in non_blank):
                            warnings.append(
                                {
                                    "type": "visual_numeric_measures_all_zero",
                                    "page": page_name,
                                    "visual": visual_name,
                                    "visual_type": visual_type,
                                }
                            )

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


def pbi_detect_missing_visuals_tool(
    extract_folder: str,
    *,
    page: str,
    requirements: list[dict[str, Any]],
) -> dict[str, Any]:
    """Detect required visuals that are absent from a page.

    Each requirement entry is a dict with at least ``visual_type`` and
    optionally ``count`` (default 1), ``contains_field`` (a reference like
    ``Date.Year`` that must appear in the visual's prototypeQuery), and
    ``label`` (free-form name surfaced in the report).
    """
    folder, layout = _load_layout(extract_folder)
    section = None
    for sec in layout.get("sections", []) or []:
        if isinstance(sec, dict) and (sec.get("displayName") == page or sec.get("name") == page):
            section = sec
            break
    if section is None:
        raise PowerBIValidationError(
            f"Page '{page}' was not found in the layout.",
            details={"extract_folder": str(folder), "page": page},
        )

    containers = section.get("visualContainers", []) or []
    parsed: list[dict[str, Any]] = []
    for container in containers:
        cfg = _visual_config(container)
        single = cfg.get("singleVisual") or {}
        vt = str(single.get("visualType", "")).casefold()
        proto = single.get("prototypeQuery") or {}
        fields: set[str] = set()
        for entry in proto.get("Select", []) or []:
            if not isinstance(entry, dict):
                continue
            name = str(entry.get("Name", "")).casefold()
            if name:
                fields.add(name)
        parsed.append({"visual_type": vt, "fields": fields})

    found: list[dict[str, Any]] = []
    missing: list[dict[str, Any]] = []
    for req in requirements:
        if not isinstance(req, dict):
            missing.append({"requirement": req, "reason": "invalid_entry"})
            continue
        wanted_type = str(req.get("visual_type", "")).casefold()
        if not wanted_type:
            missing.append({"requirement": req, "reason": "visual_type_missing"})
            continue
        wanted_count = int(req.get("count", 1))
        contains = str(req.get("contains_field", "") or "").casefold()
        contains_short = contains.rsplit(".", 1)[-1] if contains else ""
        matches = [
            v
            for v in parsed
            if v["visual_type"] == wanted_type and (not contains_short or contains_short in v["fields"])
        ]
        entry = {
            "requirement": req,
            "matched_count": len(matches),
            "expected_count": wanted_count,
        }
        if len(matches) >= wanted_count:
            found.append(entry)
        else:
            missing.append({**entry, "reason": "insufficient_matches"})

    return ok(
        f"Page '{page}': {len(found)}/{len(requirements)} requirements satisfied.",
        valid=not missing,
        page=page,
        visual_count=len(parsed),
        found_count=len(found),
        missing_count=len(missing),
        found=found,
        missing=missing,
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
                    visuals.append(
                        {
                            "page": section.get("displayName") or section.get("name"),
                            "name": _visual_name(container),
                            "type": _visual_type(container),
                        }
                    )
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

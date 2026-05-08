"""Report-field validation + repair tools."""

from __future__ import annotations

from typing import Any

from pbi_connection import ok

from ._base import _run
from ._bindings import _live_model_field_index, _scan_visual_bindings
from ._home_tables import _inspect_value_measures, _persistence_risks, _scan_measure_home_tables
from ._layout import _find_page, _load_layout, _parse_embedded_json, _save_layout

# Visual types whose Y role is an axis-driven series — these are the ones
# that fail with bug 0.92 (line chart + constant measure) and similar
# render errors.
_CARTESIAN_VISUAL_TYPES = {
    "lineChart",
    "lineClusteredColumnComboChart",
    "waterfallChart",
    "areaChart",
    "stackedAreaChart",
}


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
        planned_repairs = (
            repairs
            if apply
            else sum(
                1
                for item in issues
                if item.get("issue") in {"query_ref_mismatch", "measure_home_table_needs_repair"}
                or (
                    item.get("issue") == "unexpected_projection_role"
                    and item.get("visual_type") == "gauge"
                    and item.get("role") == "Value"
                )
            )
        )
        unresolved = [
            item
            for item in issues
            if item.get("issue")
            in {
                "query_ref_not_found",
                "unexpected_projection_role",
                "measure_home_table_unknown",
                "column_not_found",
                "measure_not_found",
                "measure_table_mismatch",
            }
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


def _collect_cartesian_y_measures(layout: dict[str, Any], page: str | None) -> list[dict[str, Any]]:
    """Walk visualContainers and yield ``{page, visual_id, visual_type,
    measures}`` for every cartesian chart whose Y/Y2 role binds measures.
    """
    findings: list[dict[str, Any]] = []
    sections = layout.get("sections", []) or []
    for section in sections:
        if not isinstance(section, dict):
            continue
        section_name = str(section.get("displayName") or section.get("name") or "")
        if page and page.casefold() not in {
            str(section.get("name", "")).casefold(),
            str(section.get("displayName", "")).casefold(),
        }:
            continue
        for container in section.get("visualContainers", []) or []:
            if not isinstance(container, dict):
                continue
            config = _parse_embedded_json(container.get("config"), {})
            if not isinstance(config, dict):
                continue
            sv = config.get("singleVisual", {})
            if not isinstance(sv, dict):
                continue
            visual_type = str(sv.get("visualType", ""))
            if visual_type not in _CARTESIAN_VISUAL_TYPES:
                continue
            projections = sv.get("projections", {}) or {}
            measures: list[str] = []
            for role in ("Y", "Y2"):
                role_items = projections.get(role) or []
                if not isinstance(role_items, list):
                    continue
                for item in role_items:
                    if isinstance(item, dict):
                        ref = str(item.get("queryRef", "")).strip()
                        if ref:
                            measures.append(ref)
            if not measures:
                continue
            findings.append(
                {
                    "page": section_name,
                    "visual_id": str(config.get("name", "")),
                    "visual_type": visual_type,
                    "measures": measures,
                }
            )
    return findings


def pbi_diagnose_render_risks_tool(
    extract_folder: str,
    page: str | None = None,
    visual_id: str | None = None,
    *,
    manager: Any | None = None,
    include_hidden: bool = False,
) -> dict[str, Any]:
    """Aggregate render-risk diagnostic for visuals on disk.

    Read-only. Walks the extracted layout and reports conditions known to
    make Power BI Desktop refuse to render a visual or surface an opaque
    internal error:

    - **Constant Y measure on a cartesian chart** (line / combo / area /
      waterfall). The bug-0.92 family: a measure whose DAX expression has
      no column or measure reference (or returns BLANK() unconditionally)
      collapses every Y point to the same scalar with no axis dependency,
      which some PBI builds reject.
    - **Unresolved measure home table** — the binding falls back to the
      synthetic ``$Measures`` entity which PBI refuses to plot.
    - **Missing column / measure / wrong reference kind** in the live
      model (only when ``manager`` is supplied).
    - **Query-ref mismatch** between projections and prototypeQuery
      Select entries.

    A live ``manager`` enables DAX-expression heuristics for the
    constant-measure check; without it the call still reports unresolved
    home tables and binding-shape issues that don't need the live model.

    Pass ``visual_id`` to narrow to a single visual; pass ``page`` to
    narrow to a specific page; both omitted scans the whole report.
    """

    def _impl() -> dict[str, Any]:
        folder, layout = _load_layout(extract_folder)
        if visual_id and page:
            section = _find_page(layout, page)
            sections_scoped = {"sections": [section]}
        else:
            sections_scoped = layout

        measure_home_map = _scan_measure_home_tables(folder)
        model_fields, model_validation = _live_model_field_index(manager, include_hidden=include_hidden)
        binding_issues, _ = _scan_visual_bindings(
            sections_scoped, measure_home_map, model_fields, page=page if not visual_id else None, repair=False
        )
        if visual_id:
            binding_issues = [item for item in binding_issues if item.get("visual_id") == visual_id]

        cartesian_findings = _collect_cartesian_y_measures(sections_scoped, page if not visual_id else None)
        if visual_id:
            cartesian_findings = [item for item in cartesian_findings if item["visual_id"] == visual_id]

        constant_findings: list[dict[str, Any]] = []
        for finding in cartesian_findings:
            warnings = _inspect_value_measures(finding["measures"], measure_home_map, manager)
            for warning in warnings:
                if warning.get("issue") != "constant_measure":
                    continue
                constant_findings.append(
                    {
                        "page": finding["page"],
                        "visual_id": finding["visual_id"],
                        "visual_type": finding["visual_type"],
                        "measure": warning.get("measure"),
                        "issue": warning.get("issue"),
                        "hint": warning.get("hint"),
                        "expression_preview": warning.get("expression_preview"),
                    }
                )

        risks_count = len(binding_issues) + len(constant_findings)
        return ok(
            f"Render-risk diagnostic found {risks_count} risk(s) "
            f"({len(binding_issues)} binding, {len(constant_findings)} constant-measure).",
            extract_folder=str(folder),
            page=page,
            visual_id=visual_id,
            include_hidden=include_hidden,
            model_validation=model_validation,
            risk_count=risks_count,
            binding_issues=binding_issues,
            constant_measure_risks=constant_findings,
            cartesian_visual_count=len(cartesian_findings),
            healthy=risks_count == 0,
        )

    return _run(_impl)

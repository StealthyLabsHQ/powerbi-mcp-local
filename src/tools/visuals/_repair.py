"""Report-field validation + repair tools."""

from __future__ import annotations

from typing import Any

from pbi_connection import ok

from ._bindings import _live_model_field_index, _scan_visual_bindings
from ._home_tables import _persistence_risks, _scan_measure_home_tables
from ._layout import _load_layout, _save_layout


def _run(callback):
    from pbi_connection import error_payload

    try:
        return callback()
    except Exception as exc:
        return error_payload(exc)


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

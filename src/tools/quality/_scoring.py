"""Scoring, rubric, scenario, and report-export tools for quality gates."""

from __future__ import annotations

import json
from datetime import UTC, datetime
from typing import Any

from pbi_connection import PowerBIValidationError, ok


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
    from . import (
        pbi_audit_model_tool,
        pbi_detect_dirty_dates_tool,
        pbi_detect_empty_visuals_tool,
        pbi_detect_name_collisions_tool,
        pbi_generate_measure_tests_tool,
        pbi_lint_dax_tool,
        pbi_lint_report_layout_tool,
        pbi_score_dashboard_tool,
        pbi_validate_visual_bindings_tool,
        resolve_local_path,
    )

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
        report["visual_bindings"] = pbi_validate_visual_bindings_tool(
            extract_folder, include_hidden=include_hidden, manager=manager
        )
        if include_empty_visual_scan:
            report["empty_visuals"] = pbi_detect_empty_visuals_tool(
                manager, extract_folder=extract_folder, filter_expression=empty_visual_filter_expression
            )
    if include_measure_tests:
        report["measure_tests"] = pbi_generate_measure_tests_tool(manager, include_hidden=include_hidden)
    report["score"] = pbi_score_dashboard_tool(manager, extract_folder=extract_folder, include_hidden=include_hidden)
    validation_sections = [item for item in report.values() if isinstance(item, dict) and "valid" in item]
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


def _score_parts(
    model: dict[str, Any], dax: dict[str, Any], layout: dict[str, Any], bindings: dict[str, Any] | None
) -> dict[str, int]:
    model_score = max(0, 25 - model.get("issue_count", 0) * 10 - model.get("warning_count", 0) * 2)
    dax_score = max(0, 25 - dax.get("issue_count", 0) * 10 - dax.get("warning_count", 0) * 2)
    layout_score = max(0, 20 - layout.get("issue_count", 0) * 8 - layout.get("warning_count", 0) * 2)
    binding_penalty = 0 if not bindings else bindings.get("issue_count", 0) * 8
    readability_score = max(0, 20 - binding_penalty - model.get("warning_count", 0))
    robustness_score = (
        10
        if model.get("valid") and dax.get("valid") and layout.get("valid") and (not bindings or bindings.get("valid"))
        else 5
    )
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
    from . import (
        pbi_audit_model_tool,
        pbi_lint_dax_tool,
        pbi_lint_report_layout_tool,
        pbi_validate_visual_bindings_tool,
    )

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
    from . import pbi_score_dashboard_tool

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


def pbi_score_rubric_tool(
    manager: Any,
    *,
    extract_folder: str | None = None,
    criteria: list[dict[str, Any]],
) -> dict[str, Any]:
    """Aggregate scoring across multiple validators.

    Each criterion is a dict with:
    - ``id`` (str): unique identifier
    - ``label`` (str): human description
    - ``check`` (str): one of ``star_schema``, ``no_circular_deps``,
      ``power_query_steps``, ``missing_visuals``, ``measure_exists``
    - ``weight`` (float, default 1.0)
    - ``params`` (dict): check-specific parameters

    Returns per-criterion verdicts plus a weighted total score in [0, 1].
    """
    from . import (
        _model_snapshot,
        pbi_detect_circular_dependencies_tool,
        pbi_detect_missing_visuals_tool,
        pbi_validate_power_query_steps_tool,
        pbi_validate_star_schema_tool,
    )

    if not criteria:
        raise PowerBIValidationError("criteria must contain at least one entry.")

    measure_names: set[str] | None = None

    def _ensure_measure_names() -> set[str]:
        nonlocal measure_names
        if measure_names is None:
            snapshot = _model_snapshot(manager, include_hidden=True)
            measure_names = {str(m.get("name", "")).casefold() for m in snapshot.get("measures", []) or []}
        return measure_names

    results: list[dict[str, Any]] = []
    total_weight = 0.0
    earned = 0.0

    for criterion in criteria:
        if not isinstance(criterion, dict):
            results.append({"id": None, "passed": False, "reason": "invalid_criterion"})
            continue
        cid = str(criterion.get("id", "") or f"criterion_{len(results)}")
        label = str(criterion.get("label", "") or cid)
        check = str(criterion.get("check", ""))
        weight = float(criterion.get("weight", 1.0))
        params = criterion.get("params") or {}
        passed = False
        details: dict[str, Any] = {}
        try:
            if check == "star_schema":
                payload = pbi_validate_star_schema_tool(manager, **params)
                passed = bool(payload.get("is_star_schema"))
                details = {"issue_count": payload.get("issue_count"), "fact_tables": payload.get("fact_tables")}
            elif check == "no_circular_deps":
                payload = pbi_detect_circular_dependencies_tool(manager, **params)
                passed = bool(payload.get("valid"))
                details = {"cycle_count": payload.get("cycle_count")}
            elif check == "power_query_steps":
                payload = pbi_validate_power_query_steps_tool(manager, **params)
                passed = bool(payload.get("valid"))
                details = {"found_count": payload.get("found_count"), "missing_count": payload.get("missing_count")}
            elif check == "missing_visuals":
                if not extract_folder:
                    raise PowerBIValidationError("extract_folder required for missing_visuals check.")
                payload = pbi_detect_missing_visuals_tool(extract_folder, **params)
                passed = bool(payload.get("valid"))
                details = {"found_count": payload.get("found_count"), "missing_count": payload.get("missing_count")}
            elif check == "measure_exists":
                target = str(params.get("name", "")).casefold()
                if not target:
                    raise PowerBIValidationError("measure_exists requires params.name")
                passed = target in _ensure_measure_names()
                details = {"name": params.get("name")}
            else:
                results.append(
                    {
                        "id": cid,
                        "label": label,
                        "passed": False,
                        "reason": "unknown_check",
                        "check": check,
                        "weight": weight,
                    }
                )
                total_weight += weight
                continue
        except Exception as exc:
            results.append(
                {
                    "id": cid,
                    "label": label,
                    "passed": False,
                    "reason": "check_failed",
                    "error": str(exc),
                    "weight": weight,
                }
            )
            total_weight += weight
            continue

        if passed:
            earned += weight
        total_weight += weight
        results.append(
            {
                "id": cid,
                "label": label,
                "check": check,
                "weight": weight,
                "passed": passed,
                "details": details,
            }
        )

    score = (earned / total_weight) if total_weight else 0.0
    passed_count = sum(1 for r in results if r.get("passed"))

    return ok(
        f"Rubric: {passed_count}/{len(results)} passed, score={score:.2%}.",
        score=round(score, 4),
        earned_weight=round(earned, 4),
        total_weight=round(total_weight, 4),
        passed_count=passed_count,
        criterion_count=len(results),
        results=results,
    )


def pbi_export_correction_report_tool(
    manager: Any,
    *,
    output_path: str,
    extract_folder: str | None = None,
    rubric_criteria: list[dict[str, Any]] | None = None,
    fact_table_hints: list[str] | None = None,
) -> dict[str, Any]:
    """Generate a Markdown correction report aggregating all analysis tools.

    Output sections: model overview, star-schema verdict, circular
    dependency scan, optional rubric scoring, audit issues. Writes to
    ``output_path`` and returns the path plus an inline summary.
    """
    from . import (
        _model_snapshot,
        pbi_audit_model_tool,
        pbi_detect_circular_dependencies_tool,
        pbi_score_rubric_tool,
        pbi_validate_star_schema_tool,
        resolve_local_path,
    )

    out = resolve_local_path(output_path, must_exist=False)
    if out.is_dir():
        raise PowerBIValidationError(
            "output_path must be a file path, not a directory.",
            details={"output_path": str(out)},
        )

    snapshot = _model_snapshot(manager, include_hidden=False)
    star = pbi_validate_star_schema_tool(manager, fact_table_hints=fact_table_hints)
    cycles = pbi_detect_circular_dependencies_tool(manager)
    audit = pbi_audit_model_tool(manager)
    rubric: dict[str, Any] | None = None
    if rubric_criteria:
        rubric = pbi_score_rubric_tool(manager, extract_folder=extract_folder, criteria=rubric_criteria)

    lines: list[str] = []
    now = datetime.now(UTC).strftime("%Y-%m-%d %H:%M:%S UTC")
    lines.append("# Power BI correction report")
    lines.append("")
    lines.append(f"_Generated {now}_")
    lines.append("")
    lines.append("## Model overview")
    lines.append("")
    lines.append(f"- Tables: {len(snapshot.get('tables', []) or [])}")
    lines.append(f"- Measures: {len(snapshot.get('measures', []) or [])}")
    lines.append(f"- Relationships: {len(snapshot.get('relationships', []) or [])}")
    lines.append("")
    lines.append("## Star schema")
    lines.append("")
    lines.append(f"- Verdict: **{'PASS' if star.get('is_star_schema') else 'FAIL'}**")
    lines.append(f"- Fact tables: {', '.join(star.get('fact_tables', []) or []) or '_none_'}")
    lines.append(f"- Dimension tables: {', '.join(star.get('dim_tables', []) or []) or '_none_'}")
    if star.get("issues"):
        lines.append("- Issues:")
        for issue in star["issues"]:
            lines.append(f"  - `{issue.get('type')}`: {issue}")
    lines.append("")
    lines.append("## Circular dependencies")
    lines.append("")
    lines.append(f"- Verdict: **{'PASS' if cycles.get('valid') else 'FAIL'}**")
    lines.append(f"- Cycles: {cycles.get('cycle_count', 0)}, self-refs: {cycles.get('self_reference_count', 0)}")
    for cycle in cycles.get("cycles", []) or []:
        lines.append(f"  - cycle: {' → '.join(cycle)}")
    for sr in cycles.get("self_references", []) or []:
        lines.append(f"  - self-ref: `{sr}`")
    lines.append("")
    lines.append("## Model audit")
    lines.append("")
    lines.append(f"- Issues: {audit.get('issue_count', 0)}")
    lines.append(f"- Warnings: {audit.get('warning_count', 0)}")
    for issue in audit.get("issues", []) or []:
        lines.append(f"  - issue: `{issue.get('type')}`")
    if rubric:
        lines.append("")
        lines.append("## Rubric scoring")
        lines.append("")
        lines.append(f"- Score: **{rubric.get('score', 0):.2%}**")
        lines.append(f"- Passed: {rubric.get('passed_count')}/{rubric.get('criterion_count')}")
        for entry in rubric.get("results", []) or []:
            mark = "✓" if entry.get("passed") else "✗"
            lines.append(f"  - {mark} `{entry.get('id')}` — {entry.get('label')} (weight {entry.get('weight')})")

    content = "\n".join(lines) + "\n"
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(content, encoding="utf-8")

    return ok(
        f"Correction report written to {out}.",
        output_path=str(out),
        bytes_written=len(content.encode("utf-8")),
        is_star_schema=star.get("is_star_schema"),
        cycle_count=cycles.get("cycle_count"),
        audit_issue_count=audit.get("issue_count"),
        rubric_score=(rubric.get("score") if rubric else None),
    )

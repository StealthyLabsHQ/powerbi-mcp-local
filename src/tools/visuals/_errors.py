"""Repairable-error registry + detect→classify→repair→verify loop.

Raw binding issues are useful for humans but hard for an LLM to act on.
This module classifies every known failure mode into a stable error code
with an ``llm_action`` — a concrete instruction telling the calling model
how to fix its *spec* instead of guessing — and runs deterministic repairs
in a loop until the layout converges.

Tools
-----
- ``pbi_list_repairable_errors_tool`` — the error vocabulary (codes,
  severity, whether auto-repairable, LLM repair instruction).
- ``pbi_repair_loop_tool`` — detect → auto-repair → re-verify loop;
  returns the residual errors in classified, LLM-actionable form.
"""

from __future__ import annotations

import re
from typing import Any

from atomic_io import snapshot_once
from pbi_connection import ok

from ._base import _run
from ._bindings import _live_model_field_index, _scan_visual_bindings
from ._home_tables import _persistence_risks, _scan_measure_home_tables
from ._layout import _load_layout, _parse_embedded_json, _save_layout
from ._paths import _layout_path

#: Stable error vocabulary. ``auto_repair`` errors are fixed by the loop
#: itself; the rest carry an ``llm_action`` the calling model can apply to
#: its spec before retrying.
REPAIRABLE_ERRORS: dict[str, dict[str, Any]] = {
    "measure_not_found": {
        "severity": "error",
        "auto_repair": False,
        "llm_action": "The referenced measure does not exist in the model. Create it first "
        "(pbi_create_measure) or change the spec to an existing 'Table.Measure' "
        "(pbi_list_measures shows the valid set).",
    },
    "column_not_found": {
        "severity": "error",
        "auto_repair": False,
        "llm_action": "The referenced column does not exist. Axis/legend/slicer roles expect a model "
        "column, not a measure — fix the 'Table.Column' reference in the spec.",
    },
    "measure_table_mismatch": {
        "severity": "error",
        "auto_repair": False,
        "llm_action": "The measure exists but lives in a different table. Use the home table reported "
        "in details as the prefix in the spec.",
    },
    "measure_home_table_unknown": {
        "severity": "error",
        "auto_repair": False,
        "llm_action": "The measure's home table could not be resolved, so the binding falls back to "
        "'$Measures' which Power BI refuses to plot. Reference the measure as "
        "'HomeTable.Measure' explicitly.",
    },
    "measure_home_table_needs_repair": {
        "severity": "warning",
        "auto_repair": True,
        "llm_action": "Deterministic — rerun pbi_repair_loop with apply=true.",
    },
    "query_ref_mismatch": {
        "severity": "warning",
        "auto_repair": True,
        "llm_action": "Deterministic — rerun pbi_repair_loop with apply=true.",
    },
    "query_ref_not_found": {
        "severity": "error",
        "auto_repair": False,
        "llm_action": "A projection points at a queryRef missing from the prototypeQuery. Rebuild the "
        "visual through pbi_add_visual_from_intent instead of patching projections by hand.",
    },
    "unexpected_projection_role": {
        "severity": "error",
        "auto_repair": False,
        "llm_action": "This visual type does not accept the projected role (e.g. an axis role bound to "
        "a measure). Rebuild via the intent layer so roles are assigned deterministically.",
    },
    "projection_role_repaired": {
        "severity": "info",
        "auto_repair": True,
        "llm_action": "Already repaired.",
    },
    "invalid_config": {
        "severity": "error",
        "auto_repair": False,
        "llm_action": "The visual's embedded config is not valid JSON. Remove the visual "
        "(pbi_remove_visual) and recreate it.",
    },
    "gauge_target_invalid": {
        "severity": "error",
        "auto_repair": False,
        "llm_action": "The gauge target/min/max literals are inconsistent (target outside [min, max] or "
        "min >= max). Fix target_value/min_value/max_value in the spec.",
    },
    "double_display_units": {
        "severity": "warning",
        "auto_repair": False,
        "llm_action": "The measure's format string already scales to K/M and the visual also sets "
        "labelDisplayUnits, so values render double-scaled (e.g. '1K' shown for 1M). Either "
        "reset labelDisplayUnits to Auto (pbi_set_visual_format_property) or remove the "
        "scaling from the measure format string.",
    },
    "empty_visual": {
        "severity": "warning",
        "auto_repair": False,
        "llm_action": "The visual's data query returns no rows or only BLANK. Check filters, the "
        "measure's DAX, and that the bound columns contain data.",
    },
}


def _classify(issue: dict[str, Any]) -> dict[str, Any]:
    code = str(issue.get("issue") or issue.get("code") or "unknown")
    meta = REPAIRABLE_ERRORS.get(code, {})
    details = {k: v for k, v in issue.items() if k not in {"issue", "page", "visual_id", "visual_type"}}
    return {
        "code": code,
        "severity": meta.get("severity", "error"),
        "auto_repair": bool(meta.get("auto_repair", False)),
        "page": issue.get("page", ""),
        "visual_id": issue.get("visual_id", ""),
        "visual_type": issue.get("visual_type", ""),
        "details": details,
        "llm_action": meta.get("llm_action", "Unknown error code; inspect details manually."),
    }


def _parse_pbi_literal_number(prop: Any) -> float | None:
    """Decode ``{"expr": {"Literal": {"Value": "100.0D"}}}`` → 100.0."""
    if not isinstance(prop, dict):
        return None
    value = prop.get("expr", {}).get("Literal", {}).get("Value")
    if not isinstance(value, str):
        return None
    raw = value.strip().rstrip("DLdl")
    try:
        return float(raw)
    except ValueError:
        return None


def _iter_single_visuals(layout: dict[str, Any], page: str | None):
    for section in layout.get("sections", []) or []:
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
            sv = config.get("singleVisual")
            if isinstance(sv, dict):
                yield section_name, str(config.get("name", "")), sv


def _detect_gauge_target_issues(layout: dict[str, Any], page: str | None) -> list[dict[str, Any]]:
    issues: list[dict[str, Any]] = []
    for section_name, visual_id, sv in _iter_single_visuals(layout, page):
        if str(sv.get("visualType", "")) != "gauge":
            continue
        axis_entries = (sv.get("objects", {}) or {}).get("axis") or []
        props: dict[str, Any] = {}
        for entry in axis_entries:
            if isinstance(entry, dict) and isinstance(entry.get("properties"), dict):
                props.update(entry["properties"])
        minimum = _parse_pbi_literal_number(props.get("min"))
        maximum = _parse_pbi_literal_number(props.get("max"))
        target = _parse_pbi_literal_number(props.get("target"))
        problem: str | None = None
        if minimum is not None and maximum is not None and minimum >= maximum:
            problem = f"min ({minimum}) >= max ({maximum})"
        elif target is not None and minimum is not None and target < minimum:
            problem = f"target ({target}) < min ({minimum})"
        elif target is not None and maximum is not None and target > maximum:
            problem = f"target ({target}) > max ({maximum})"
        if problem:
            issues.append(
                {
                    "issue": "gauge_target_invalid",
                    "page": section_name,
                    "visual_id": visual_id,
                    "visual_type": "gauge",
                    "problem": problem,
                    "min": minimum,
                    "max": maximum,
                    "target": target,
                }
            )
    return issues


# Format strings that already scale values: a thousands comma right before
# the decimal/end ("#,0,." → K, "#,0,," → M) or an explicit K/M suffix.
_SCALED_FORMAT_RE = re.compile(r'(?i)[#0]\s*,+\s*(?=[.;%"\s]|$)|"\s*[km]\s*"|[0#.,][km]\b')


def _measure_format_index(manager: Any) -> dict[str, str]:
    """``table.measure`` (casefold) → format string, from the live model."""
    if manager is None:
        return {}
    from ..quality import _model_snapshot

    index: dict[str, str] = {}
    for measure in _model_snapshot(manager, include_hidden=True).get("measures", []):
        key = f"{measure.get('table', '')}.{measure.get('name', '')}".casefold()
        index[key] = str(measure.get("format_string", "") or "")
    return index


def _detect_double_display_units(
    layout: dict[str, Any], page: str | None, format_index: dict[str, str]
) -> list[dict[str, Any]]:
    if not format_index:
        return []
    issues: list[dict[str, Any]] = []
    for section_name, visual_id, sv in _iter_single_visuals(layout, page):
        labels = (sv.get("objects", {}) or {}).get("labels") or []
        display_units: float | None = None
        for entry in labels:
            if isinstance(entry, dict) and isinstance(entry.get("properties"), dict):
                parsed = _parse_pbi_literal_number(entry["properties"].get("labelDisplayUnits"))
                if parsed is not None:
                    display_units = parsed
        if not display_units or display_units < 1000:
            continue
        projections = sv.get("projections", {}) or {}
        for role_items in projections.values():
            if not isinstance(role_items, list):
                continue
            for item in role_items:
                if not isinstance(item, dict):
                    continue
                query_ref = str(item.get("queryRef", "")).strip()
                fmt = format_index.get(query_ref.casefold())
                if fmt and _SCALED_FORMAT_RE.search(fmt):
                    issues.append(
                        {
                            "issue": "double_display_units",
                            "page": section_name,
                            "visual_id": visual_id,
                            "visual_type": str(sv.get("visualType", "")),
                            "measure": query_ref,
                            "format_string": fmt,
                            "label_display_units": display_units,
                        }
                    )
    return issues


def pbi_list_repairable_errors_tool() -> dict[str, Any]:
    """Return the repairable-error vocabulary used by ``pbi_repair_loop``.

    Each entry maps a stable error code to its severity, whether the loop
    can fix it automatically, and the spec-level repair instruction an LLM
    should apply when it cannot.
    """

    def _impl() -> dict[str, Any]:
        return ok(
            f"{len(REPAIRABLE_ERRORS)} repairable error code(s) registered.",
            errors=REPAIRABLE_ERRORS,
        )

    return _run(_impl)


def pbi_repair_loop_tool(
    extract_folder: str,
    page: str | None = None,
    apply: bool = True,
    max_rounds: int = 3,
    *,
    manager: Any | None = None,
    include_hidden: bool = False,
    check_empty_visuals: bool = False,
) -> dict[str, Any]:
    """Detect → classify → auto-repair → re-verify loop for a report extract.

    Each round scans visual bindings, applies the deterministic repairs
    (query-ref mismatches, measure home tables) when ``apply=True``, saves,
    and rescans — until no repair is applied or ``max_rounds`` is reached.
    Gauge target consistency and double display-unit scaling are then
    checked on the converged layout; ``check_empty_visuals=True`` (requires
    a live connection) additionally probes every visual's data query.

    The response's ``repairable_errors`` lists the residual issues in
    classified form, each with an ``llm_action`` telling the calling model
    how to fix its spec.
    """

    def _impl() -> dict[str, Any]:
        rounds_run = 0
        auto_repairs = 0
        issues: list[dict[str, Any]] = []
        model_fields, model_validation = _live_model_field_index(manager, include_hidden=include_hidden)

        folder, layout = _load_layout(extract_folder)
        # Each repair round rolls Layout.bak forward, so after round 2 the
        # backup is a machine-edited intermediate. Snapshot the pristine
        # layout once so a clean rollback stays possible.
        pristine_snapshot = snapshot_once(_layout_path(folder)) if apply else None
        while rounds_run < max(1, max_rounds):
            rounds_run += 1
            measure_home_map = _scan_measure_home_tables(folder)
            issues, repairs = _scan_visual_bindings(layout, measure_home_map, model_fields, page=page, repair=apply)
            if apply and repairs:
                auto_repairs += repairs
                _save_layout(folder, layout)
                folder, layout = _load_layout(extract_folder)
                continue
            break

        issues = [item for item in issues if item.get("issue") != "measure_home_table_repaired"]
        issues.extend(_detect_gauge_target_issues(layout, page))
        issues.extend(_detect_double_display_units(layout, page, _measure_format_index(manager)))

        empty_visuals_checked = False
        if check_empty_visuals and manager is not None:
            from ..quality import pbi_detect_empty_visuals_tool

            empty_result = pbi_detect_empty_visuals_tool(manager, extract_folder=str(folder), page=page)
            empty_visuals_checked = bool(empty_result.get("ok"))
            empty_findings = list(empty_result.get("issues", []) or []) + [
                item
                for item in empty_result.get("warnings", []) or []
                if item.get("type") == "visual_measures_all_blank"
            ]
            for finding in empty_findings:
                issues.append(
                    {
                        "issue": "empty_visual",
                        "page": finding.get("page", ""),
                        "visual_id": finding.get("visual", ""),
                        "visual_type": finding.get("visual_type", ""),
                        "probe": finding,
                    }
                )

        classified = [_classify(item) for item in issues]
        residual = [item for item in classified if not item["auto_repair"]]
        persistence_risks = _persistence_risks(issues)
        return ok(
            f"Repair loop ran {rounds_run} round(s): {auto_repairs} auto-repair(s) applied, "
            f"{len(residual)} repairable error(s) left for the caller.",
            extract_folder=str(folder),
            page=page,
            apply=apply,
            rounds_run=rounds_run,
            auto_repairs=auto_repairs,
            pristine_snapshot=str(pristine_snapshot) if pristine_snapshot else None,
            model_validation=model_validation,
            empty_visuals_checked=empty_visuals_checked,
            repairable_errors=residual,
            persistence_risk_count=len(persistence_risks),
            persistence_risks=persistence_risks,
            healthy=not residual,
        )

    return _run(_impl)

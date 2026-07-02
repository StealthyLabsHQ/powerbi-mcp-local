"""DAX lint tools: measure lint, filter/Power Query validation, measure tests."""

from __future__ import annotations

from typing import Any

from pbi_connection import PowerBIValidationError, ok


def pbi_lint_dax_tool(
    manager: Any, *, include_hidden: bool = False, validate_expressions: bool = True
) -> dict[str, Any]:
    """Validate measures, formats, and measure/column name collisions."""
    from ..query import pbi_validate_dax_tool
    from . import _model_snapshot

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
                issues.append(
                    {"type": "invalid_measure_dax", "table": table, "measure": name, "error": result.get("error")}
                )

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


def pbi_validate_filter_expression_tool(manager: Any, *, filter_expression: str) -> dict[str, Any]:
    """Validate a DAX boolean filter expression before visual probes."""
    expression = str(filter_expression or "").strip()
    if not expression:
        raise PowerBIValidationError("filter_expression is required.")
    query = f'EVALUATE CALCULATETABLE(ROW("__probe", 1), {expression})'
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


def pbi_validate_power_query_steps_tool(
    manager: Any,
    *,
    table: str,
    expected_steps: list[str],
    partition_name: str | None = None,
    case_sensitive: bool = False,
) -> dict[str, Any]:
    """Verify that a Power Query (M) expression contains expected step patterns.

    Each entry in ``expected_steps`` is treated as a substring (or regex if it
    starts with ``re:``) that must appear at least once in the M expression.
    Useful for grading exercises: e.g. checking that a postal-code column has
    been left-padded to 5 chars, or that rows with null customer ids are
    filtered out.
    """
    import re

    from ..power_query import pbi_get_power_query_tool

    if not expected_steps:
        raise PowerBIValidationError("expected_steps must contain at least one entry.")

    payload = pbi_get_power_query_tool(manager, table=table, partition_name=partition_name)
    expression = str(payload.get("expression", "") or "")
    haystack = expression if case_sensitive else expression.casefold()

    found: list[dict[str, Any]] = []
    missing: list[dict[str, Any]] = []
    for step in expected_steps:
        is_regex = step.startswith("re:")
        needle = step[3:] if is_regex else step
        if not is_regex:
            target = needle if case_sensitive else needle.casefold()
            ok_match = target in haystack
        else:
            flags = 0 if case_sensitive else re.IGNORECASE
            ok_match = re.search(needle, expression, flags) is not None
        entry = {"step": step, "is_regex": is_regex}
        (found if ok_match else missing).append(entry)

    return ok(
        f"Power Query steps: {len(found)}/{len(expected_steps)} found.",
        valid=not missing,
        table=table,
        partition_name=partition_name,
        found_count=len(found),
        missing_count=len(missing),
        found=found,
        missing=missing,
        expression_length=len(expression),
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
    if any(
        token in lowered
        for token in (
            "revenue",
            "arr",
            "mrr",
            "margin",
            "ltv",
            "cac",
            "pipeline",
            "deal",
            "spend",
            "target",
            "forecast",
        )
    ):
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
    from . import _model_snapshot, _row_value

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
            warnings.append(
                {
                    "type": "unexpected_measure_format",
                    "measure": name,
                    "format_string": format_string,
                    "expected": expected_format,
                }
            )
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

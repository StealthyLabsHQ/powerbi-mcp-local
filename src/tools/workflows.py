"""High-level workflow tools for common Power BI agent tasks."""

from __future__ import annotations

from typing import Any

from pbi_connection import PowerBIValidationError, error_payload, ok
from security import validate_measure_name, validate_model_expression, validate_model_object_name

from .excel import excel_workbook_info_tool
from .measures import pbi_create_measures_tool
from .model import pbi_list_tables_tool, pbi_model_info_tool, pbi_validate_model_tool
from .power_query import pbi_import_excel_workbook_tool, pbi_list_power_queries_tool
from .query import pbi_measure_dependencies_tool, pbi_refresh_tool, pbi_validate_dax_tool


def _soft_call(callback: Any, *args: Any, **kwargs: Any) -> dict[str, Any]:
    try:
        return callback(*args, **kwargs)
    except Exception as exc:
        return error_payload(exc)


def _table_names(tables: list[dict[str, Any]]) -> set[str]:
    return {str(item.get("name", "")).casefold() for item in tables}


def _build_sheet_table_map(sheets: list[dict[str, Any]], tables: list[dict[str, Any]]) -> dict[str, str]:
    table_lookup = {str(item.get("name", "")).casefold(): str(item.get("name", "")) for item in tables}
    mapping: dict[str, str] = {}
    for sheet in sheets:
        name = str(sheet.get("name", ""))
        table_name = table_lookup.get(name.casefold())
        if table_name:
            mapping[name] = table_name
    return mapping


def pbi_model_audit_workflow_tool(
    manager: Any,
    *,
    include_hidden: bool = False,
    include_row_counts: bool = True,
) -> dict[str, Any]:
    """Run a compact model audit with recommendations for an agent."""
    model = pbi_model_info_tool(manager, include_hidden=include_hidden, include_row_counts=include_row_counts)
    validation = pbi_validate_model_tool(manager)
    dependencies = _soft_call(pbi_measure_dependencies_tool, manager)
    power_queries = _soft_call(pbi_list_power_queries_tool, manager, include_hidden=include_hidden)

    tables = model.get("tables", [])
    measures = model.get("measures", [])
    relationships = model.get("relationships", [])
    findings: list[dict[str, Any]] = []
    recommendations: list[dict[str, Any]] = []

    findings.extend(validation.get("issues", []))
    findings.extend(validation.get("warnings", []))
    if not relationships and tables:
        recommendations.append({"priority": "high", "message": "Create relationships before report automation."})
    if validation.get("warning_count", 0):
        recommendations.append({"priority": "medium", "message": "Fix model warnings before adding new measures."})
    if dependencies.get("ok") and dependencies.get("truncated"):
        recommendations.append({"priority": "medium", "message": "Dependency graph was truncated; narrow follow-up scans by table."})
    if not recommendations:
        recommendations.append({"priority": "low", "message": "Model audit found no immediate workflow blockers."})

    return ok(
        "Model audit workflow completed.",
        plan=["Collect model snapshot", "Validate model", "Inspect measure dependencies", "Inspect Power Query partitions"],
        summary={
            "table_count": len(tables),
            "measure_count": len(measures),
            "relationship_count": len(relationships),
            "issue_count": validation.get("issue_count", 0),
            "warning_count": validation.get("warning_count", 0),
        },
        findings=findings,
        recommendations=recommendations[:3],
        validation=validation,
        diagnostics={
            "dependencies": dependencies,
            "power_queries": power_queries,
        },
        needs_apply=False,
    )


def pbi_excel_import_workflow_tool(
    manager: Any,
    *,
    excel_path: str,
    sheet_table_map: dict[str, str] | None = None,
    promote_headers: bool = True,
    refresh_after: bool = True,
    apply: bool = False,
) -> dict[str, Any]:
    """Plan or run an Excel workbook import into Power BI."""
    workbook = excel_workbook_info_tool(excel_path)
    if not workbook.get("ok"):
        return workbook
    tables_payload = pbi_list_tables_tool(manager, include_hidden=False, include_row_counts=False)
    sheets = [item for item in workbook.get("sheets", []) if item.get("has_data")]
    tables = tables_payload.get("tables", [])
    mapping = dict(sheet_table_map or _build_sheet_table_map(sheets, tables))
    table_set = _table_names(tables)
    sheet_set = {str(item.get("name", "")).casefold() for item in sheets}

    findings = []
    for sheet, table in mapping.items():
        sheet_name = str(sheet)
        table_name = str(table)
        if sheet_name.casefold() not in sheet_set:
            findings.append({"type": "sheet_not_found", "sheet": sheet_name})
        if table_name.casefold() not in table_set:
            findings.append({"type": "table_not_found", "table": table_name, "sheet": sheet_name})
    for sheet in sheets:
        name = str(sheet.get("name", ""))
        if name not in mapping:
            findings.append({"type": "unmapped_sheet", "sheet": name})

    actions = [
        {
            "tool": "pbi_import_excel_workbook",
            "excel_path": workbook.get("file_path", excel_path),
            "sheet_table_map": mapping,
            "promote_headers": promote_headers,
        }
    ]
    if refresh_after:
        actions.append({"tool": "pbi_refresh", "target": "model", "refresh_type": "full"})

    if not apply:
        return ok(
            "Excel import workflow planned.",
            plan=["Inspect workbook", "Map sheets to model tables", "Import workbook", "Refresh model", "Validate row counts"],
            workbook=workbook,
            sheet_table_map=mapping,
            findings=findings,
            actions=actions,
            validation={"ready": not any(item["type"].endswith("_not_found") for item in findings)},
            needs_apply=True,
        )

    if any(item["type"].endswith("_not_found") for item in findings):
        raise PowerBIValidationError("Excel import workflow has blocking mapping issues.", details={"findings": findings})

    import_result = pbi_import_excel_workbook_tool(
        manager,
        excel_path=excel_path,
        sheet_table_map=mapping,
        promote_headers=promote_headers,
        refresh_after=False,
    )
    refresh_result = None
    if refresh_after:
        refresh_result = pbi_refresh_tool(manager, target="model", refresh_type="full")
    validation = pbi_list_tables_tool(manager, include_hidden=False, include_row_counts=True)
    return ok(
        "Excel import workflow applied.",
        plan=["Inspect workbook", "Map sheets to model tables", "Import workbook", "Refresh model", "Validate row counts"],
        workbook=workbook,
        sheet_table_map=mapping,
        findings=findings,
        actions=[import_result, refresh_result] if refresh_result else [import_result],
        validation=validation,
        needs_apply=False,
    )


def pbi_measure_workflow_tool(
    manager: Any,
    *,
    table: str,
    measures: list[dict[str, Any]],
    overwrite: bool = True,
    apply: bool = False,
) -> dict[str, Any]:
    """Plan or run validated batch measure creation."""
    validate_model_object_name(table)
    if not measures:
        raise PowerBIValidationError("At least one measure is required.")

    tables_payload = pbi_list_tables_tool(manager, include_hidden=True, include_row_counts=False)
    model = pbi_model_info_tool(manager, include_hidden=True, include_row_counts=False)
    target_exists = table.casefold() in _table_names(tables_payload.get("tables", []))
    existing = {
        str(item.get("name", "")).casefold()
        for item in model.get("measures", [])
        if str(item.get("table", "")).casefold() == table.casefold()
    }

    findings: list[dict[str, Any]] = []
    validations: list[dict[str, Any]] = []
    if not target_exists:
        findings.append({"type": "table_not_found", "table": table})

    for item in measures:
        name = str(item.get("name", ""))
        expression = str(item.get("expression", ""))
        try:
            validate_measure_name(name)
            validate_model_expression(expression, kind="measure expression")
        except Exception as exc:
            findings.append({"type": "invalid_measure", "measure": name, "error": error_payload(exc)["error"]})
            continue
        if name.casefold() in existing and not overwrite:
            findings.append({"type": "measure_exists", "measure": name, "table": table})
        validations.append(pbi_validate_dax_tool(manager, expression=expression, kind="scalar"))

    for result in validations:
        if result.get("valid") is False:
            findings.append({"type": "invalid_dax", "error": result.get("error"), "error_code": result.get("error_code")})

    blocking = [item for item in findings if item["type"] in {"table_not_found", "invalid_measure", "measure_exists", "invalid_dax"}]
    actions = [{"tool": "pbi_create_measures", "table": table, "count": len(measures), "overwrite": overwrite}]
    if not apply:
        return ok(
            "Measure workflow planned.",
            plan=["Inspect target table", "Check measure conflicts", "Validate DAX", "Create measures in one batch"],
            findings=findings,
            actions=actions,
            validation={"dax": validations, "ready": not blocking},
            needs_apply=True,
        )

    if blocking:
        raise PowerBIValidationError("Measure workflow has blocking validation issues.", details={"findings": findings})
    create_result = pbi_create_measures_tool(manager, table=table, measures=measures, overwrite=overwrite, stop_on_error=True)
    return ok(
        "Measure workflow applied.",
        plan=["Inspect target table", "Check measure conflicts", "Validate DAX", "Create measures in one batch"],
        findings=findings,
        actions=[create_result],
        validation={"dax": validations, "ready": True},
        needs_apply=False,
    )


__all__ = [
    "pbi_excel_import_workflow_tool",
    "pbi_measure_workflow_tool",
    "pbi_model_audit_workflow_tool",
]

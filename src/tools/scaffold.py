"""High-level PBIX scaffolding (v0.13).

``pbi_scaffold_pbix_tool`` wraps :func:`pbi_create_persistent_report_tool`
with named templates (``blank``, ``finance``, ``sales``, ``analytics``).
Templates ship a date table, a baseline measure pack, a "Summary" page,
and optionally apply a user-supplied theme JSON after the PBIX exists on
disk. The goal is a one-call path from "give me a starter PBIX" to a
fully-bootstrapped file the LLM can iterate on, without forcing the
caller to assemble the spec by hand.
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from pbi_connection import PowerBIValidationError, ok
from security import resolve_local_path

from .persistent_report import pbi_create_persistent_report_tool
from .visuals._themes import (
    MAX_THEME_BYTES,
    ThemeValidationError,
    assert_theme_within_size_limit,
    validate_theme_payload,
)


def _date_table_spec() -> dict[str, Any]:
    return {
        "name": "DateTable",
        "columns": [
            {"name": "Date", "data_type": "DateTime"},
            {"name": "Year", "data_type": "Int64"},
            {"name": "Month", "data_type": "Int64"},
            {"name": "MonthName", "data_type": "String"},
            {"name": "Quarter", "data_type": "String"},
        ],
        "rows": [],
    }


def _baseline_measures(fact_table: str, amount_column: str) -> list[dict[str, Any]]:
    fact = fact_table
    amount = amount_column
    return [
        {
            "table": fact,
            "name": "Total",
            "expression": f"SUM('{fact}'[{amount}])",
            "format_string": "#,##0",
        },
        {
            "table": fact,
            "name": "Total YTD",
            "expression": f"TOTALYTD([Total], 'DateTable'[Date])",
            "format_string": "#,##0",
        },
        {
            "table": fact,
            "name": "Total MTD",
            "expression": f"TOTALMTD([Total], 'DateTable'[Date])",
            "format_string": "#,##0",
        },
        {
            "table": fact,
            "name": "Total YoY %",
            "expression": (
                "DIVIDE([Total] - CALCULATE([Total], SAMEPERIODLASTYEAR('DateTable'[Date])),"
                " CALCULATE([Total], SAMEPERIODLASTYEAR('DateTable'[Date])))"
            ),
            "format_string": "0.0%",
        },
    ]


SCAFFOLD_TEMPLATES: dict[str, dict[str, Any]] = {
    "blank": {
        "description": "Empty model with a date table only.",
        "tables": [_date_table_spec()],
        "measures": [],
        "relationships": [],
        "pages": [{"name": "Page 1", "visuals": []}],
    },
    "finance": {
        "description": "Finance starter: date table + GL fact + baseline KPIs.",
        "tables": [
            _date_table_spec(),
            {
                "name": "GL",
                "columns": [
                    {"name": "Date", "data_type": "DateTime"},
                    {"name": "Account", "data_type": "String"},
                    {"name": "Amount", "data_type": "Decimal"},
                ],
                "rows": [],
            },
        ],
        "measures": _baseline_measures("GL", "Amount"),
        "relationships": [
            {"from_table": "GL", "from_column": "Date", "to_table": "DateTable", "to_column": "Date"},
        ],
        "pages": [{"name": "Summary", "visuals": []}, {"name": "Detail", "visuals": []}],
    },
    "sales": {
        "description": "Sales starter: date + sales fact + product dim.",
        "tables": [
            _date_table_spec(),
            {
                "name": "Product",
                "columns": [
                    {"name": "ProductId", "data_type": "Int64"},
                    {"name": "Name", "data_type": "String"},
                    {"name": "Category", "data_type": "String"},
                ],
                "rows": [],
            },
            {
                "name": "Sales",
                "columns": [
                    {"name": "Date", "data_type": "DateTime"},
                    {"name": "ProductId", "data_type": "Int64"},
                    {"name": "Quantity", "data_type": "Int64"},
                    {"name": "Amount", "data_type": "Decimal"},
                ],
                "rows": [],
            },
        ],
        "measures": _baseline_measures("Sales", "Amount"),
        "relationships": [
            {"from_table": "Sales", "from_column": "Date", "to_table": "DateTable", "to_column": "Date"},
            {"from_table": "Sales", "from_column": "ProductId", "to_table": "Product", "to_column": "ProductId"},
        ],
        "pages": [
            {"name": "Overview", "visuals": []},
            {"name": "By Product", "visuals": []},
            {"name": "By Period", "visuals": []},
        ],
    },
    "analytics": {
        "description": "Analytics starter: date + events fact + user dim.",
        "tables": [
            _date_table_spec(),
            {
                "name": "User",
                "columns": [
                    {"name": "UserId", "data_type": "Int64"},
                    {"name": "Segment", "data_type": "String"},
                ],
                "rows": [],
            },
            {
                "name": "Events",
                "columns": [
                    {"name": "Date", "data_type": "DateTime"},
                    {"name": "UserId", "data_type": "Int64"},
                    {"name": "EventType", "data_type": "String"},
                    {"name": "Value", "data_type": "Decimal"},
                ],
                "rows": [],
            },
        ],
        "measures": _baseline_measures("Events", "Value"),
        "relationships": [
            {"from_table": "Events", "from_column": "Date", "to_table": "DateTable", "to_column": "Date"},
            {"from_table": "Events", "from_column": "UserId", "to_table": "User", "to_column": "UserId"},
        ],
        "pages": [{"name": "Funnel", "visuals": []}, {"name": "Cohorts", "visuals": []}],
    },
}


def list_scaffold_templates() -> list[dict[str, Any]]:
    return [
        {
            "key": key,
            "description": spec["description"],
            "table_count": len(spec.get("tables", [])),
            "measure_count": len(spec.get("measures", [])),
        }
        for key, spec in SCAFFOLD_TEMPLATES.items()
    ]


def _load_theme(theme_json_path: str) -> dict[str, Any]:
    theme_path = resolve_local_path(theme_json_path, must_exist=True, allowed_extensions={".json"})
    raw = theme_path.read_bytes()
    assert_theme_within_size_limit(len(raw))
    try:
        payload = json.loads(raw.decode("utf-8"))
    except json.JSONDecodeError as exc:
        raise PowerBIValidationError(
            "Theme JSON is invalid.", details={"path": str(theme_path), "line": exc.lineno}
        ) from exc
    issues = validate_theme_payload(payload)
    errors = [issue for issue in issues if issue.get("level") == "error"]
    if errors:
        raise ThemeValidationError(
            "Theme JSON failed schema validation.",
            details={"path": str(theme_path), "errors": errors[:20]},
        )
    return payload


def pbi_scaffold_pbix_tool(
    output_path: str,
    template: str = "blank",
    *,
    theme_json_path: str | None = None,
    extra_measures: list[dict[str, Any]] | None = None,
    open_after_create: bool = False,
) -> dict[str, Any]:
    """Create a starter PBIX from a named template.

    Templates: ``blank`` (date table only), ``finance``, ``sales``,
    ``analytics``. Optional ``theme_json_path`` is validated against the
    Power BI theme schema (size, key allowlist, colour format,
    URL-bearing values) and embedded in the generated file. Additional
    measures supplied via ``extra_measures`` are appended after the
    template's baseline pack.
    """
    if template not in SCAFFOLD_TEMPLATES:
        raise PowerBIValidationError(
            "Unknown scaffold template.",
            details={"template": template, "available_templates": sorted(SCAFFOLD_TEMPLATES)},
        )
    spec = SCAFFOLD_TEMPLATES[template]
    tables = [dict(table) for table in spec.get("tables", [])]
    measures = [dict(measure) for measure in spec.get("measures", [])]
    if extra_measures:
        if not isinstance(extra_measures, list):
            raise PowerBIValidationError("extra_measures must be a list of measure dicts.")
        for entry in extra_measures:
            if not isinstance(entry, dict):
                raise PowerBIValidationError("extra_measures entries must be objects.")
            measures.append(dict(entry))
    relationships = [dict(rel) for rel in spec.get("relationships", [])]
    pages = [dict(page) for page in spec.get("pages", [])]

    theme_payload: dict[str, Any] | None = None
    if theme_json_path:
        theme_payload = _load_theme(theme_json_path)

    creation = pbi_create_persistent_report_tool(
        output_path=output_path,
        tables=tables,
        measures=measures,
        relationships=relationships,
        pages=pages,
        open_after_create=open_after_create,
    )

    return ok(
        f"PBIX scaffold '{template}' created.",
        template=template,
        output_path=creation.get("output_path"),
        size_bytes=creation.get("size_bytes"),
        table_count=creation.get("table_count"),
        measure_count=creation.get("measure_count"),
        relationship_count=creation.get("relationship_count"),
        page_count=creation.get("page_count"),
        theme_applied=bool(theme_payload),
        theme_size_limit_bytes=MAX_THEME_BYTES if theme_payload else None,
        opened=creation.get("opened", False),
        pre_build_issues=creation.get("pre_build_issues", []),
        validation_issues=creation.get("validation_issues", []),
    )


def pbi_list_scaffold_templates_tool() -> dict[str, Any]:
    """List the available scaffold templates with summary metrics."""
    return ok("Scaffold templates listed.", templates=list_scaffold_templates())


__all__ = [
    "SCAFFOLD_TEMPLATES",
    "list_scaffold_templates",
    "pbi_list_scaffold_templates_tool",
    "pbi_scaffold_pbix_tool",
]

"""Time-intelligence, variance, contribution, Top-N, and rolling-average generators."""

from __future__ import annotations

from typing import Any

from pbi_connection import PowerBIValidationError, find_named, ok
from security import (
    validate_measure_name,
    validate_model_object_name,
)

# ---------------------------------------------------------------------------
# Time intelligence templates
# ---------------------------------------------------------------------------


def _dax_table_ref(table: str) -> str:
    """Quote a DAX table reference with single quotes.

    Power BI allows ``Sales`` and ``'Sales'`` interchangeably, but when the
    table name collides with a reserved word (``Date``, ``Time``, ``Year``…)
    only the quoted form parses correctly. Always quoting is safe and avoids
    surprising syntax errors when callers pass perfectly normal names like
    ``Date``. Embedded single quotes are doubled per DAX grammar.
    """
    return "'" + str(table).replace("'", "''") + "'"


def _dax_column_ref(table: str, column: str) -> str:
    return f"{_dax_table_ref(table)}[{column}]"


_TIME_INTELLIGENCE_TEMPLATES: dict[str, dict[str, str]] = {
    # Each entry: name suffix + DAX template parameterised on {base} and {date_ref}.
    # ``{date_ref}`` is the already-quoted ``'Date'[Date]`` form so reserved-word
    # collisions (Date, Time, …) stay safe.
    "YTD": {
        "suffix": "YTD",
        "template": "CALCULATE([{base}], DATESYTD({date_ref}))",
        "description": "Year-to-date of [{base}].",
    },
    "MTD": {
        "suffix": "MTD",
        "template": "CALCULATE([{base}], DATESMTD({date_ref}))",
        "description": "Month-to-date of [{base}].",
    },
    "QTD": {
        "suffix": "QTD",
        "template": "CALCULATE([{base}], DATESQTD({date_ref}))",
        "description": "Quarter-to-date of [{base}].",
    },
    "SPY": {
        "suffix": "SPY",
        "template": "CALCULATE([{base}], SAMEPERIODLASTYEAR({date_ref}))",
        "description": "Same period last year of [{base}].",
    },
    "YOY": {
        "suffix": "YOY",
        "template": "[{base}] - [{base} SPY]",
        "description": "Year-over-year delta of [{base}] (requires SPY companion).",
        "depends_on": ["SPY"],
    },
    "YOY%": {
        "suffix": "YOY %",
        "template": "DIVIDE([{base} YOY], [{base} SPY])",
        "description": "Year-over-year % growth of [{base}] (requires YOY + SPY companions).",
        "format_hint": "0.00%",
        "depends_on": ["YOY", "SPY"],
    },
    "MA3": {
        "suffix": "MA3",
        "template": ("AVERAGEX(DATESINPERIOD({date_ref}, LASTDATE({date_ref}), -3, MONTH), [{base}])"),
        "description": "Trailing 3-month moving average of [{base}].",
    },
}

_DEFAULT_TIME_INTELLIGENCE_PATTERNS = ["YTD", "MTD", "QTD", "SPY", "YOY", "YOY%", "MA3"]


def _resolve_ti_patterns(patterns: list[str] | None) -> list[str]:
    if patterns is None:
        return list(_DEFAULT_TIME_INTELLIGENCE_PATTERNS)
    if not patterns:
        raise PowerBIValidationError("patterns must be a non-empty list.")
    resolved: list[str] = []
    seen: set[str] = set()
    for raw in patterns:
        token = str(raw).strip().upper()
        if token not in _TIME_INTELLIGENCE_TEMPLATES:
            raise PowerBIValidationError(
                f"Unknown time-intelligence pattern '{raw}'.",
                details={"pattern": raw, "supported": sorted(_TIME_INTELLIGENCE_TEMPLATES)},
            )
        if token in seen:
            continue
        seen.add(token)
        resolved.append(token)
    # Topo-sort dependencies: YOY needs SPY, YOY% needs YOY+SPY. Auto-add silently
    # when missing so a single "create YOY%" call still works.
    expanded: list[str] = []
    expanded_set: set[str] = set()

    def _visit(token: str) -> None:
        if token in expanded_set:
            return
        for dep in _TIME_INTELLIGENCE_TEMPLATES[token].get("depends_on", []) or []:
            _visit(dep)
        expanded.append(token)
        expanded_set.add(token)

    for token in resolved:
        _visit(token)
    return expanded


def pbi_create_time_intelligence_pack_tool(
    manager: Any,
    *,
    table: str,
    base_measure: str,
    date_table: str,
    date_column: str,
    patterns: list[str] | None = None,
    display_folder: str = "Time intelligence",
    format_inherit: bool = True,
    format_string: str = "",
    overwrite: bool = False,
    dry_run: bool = False,
) -> dict[str, Any]:
    """Create a family of time-intelligence measures from a base measure.

    Default patterns: YTD, MTD, QTD, SPY, YOY, YOY %, MA3. Each generated
    measure is named ``"{base_measure} {suffix}"`` (e.g. ``"Sales YTD"``)
    and lives on ``table`` (typically the same fact table as the base measure).
    Dependency-aware: requesting ``YOY%`` automatically adds ``YOY`` and ``SPY``
    so the whole family resolves.

    With ``format_inherit=True``, each measure inherits ``base_measure``'s
    format string (best-effort lookup on the live model). ``format_string``
    overrides that when supplied. Patterns that prescribe their own format
    (e.g. ``YOY %`` → ``"0.00%"``) win unless ``format_string`` is explicit.

    With ``dry_run=True`` no model mutation occurs — the response carries a
    ``plan`` listing every measure that would be created/updated/skipped.
    """
    validate_model_object_name(table)
    validate_measure_name(base_measure)
    validate_model_object_name(date_table)
    validate_model_object_name(date_column)
    pattern_list = _resolve_ti_patterns(patterns)

    inherited_format: str | None = None
    if format_inherit and not format_string and not dry_run:
        # Best-effort: read the base measure's format from the live model.
        def _reader(state: Any) -> str | None:
            target_table = find_named(state.database.Model.Tables, table)
            if target_table is None:
                return None
            existing = find_named(target_table.Measures, base_measure)
            if existing is None:
                return None
            return str(getattr(existing, "FormatString", "") or "") or None

        try:
            inherited_format = manager.run_read("ti_pack_inherit_format", _reader)
        except Exception:
            inherited_format = None

    plan: list[dict[str, Any]] = []
    measure_specs: list[dict[str, Any]] = []
    for token in pattern_list:
        tmpl = _TIME_INTELLIGENCE_TEMPLATES[token]
        suffix = tmpl["suffix"]
        new_name = f"{base_measure} {suffix}"
        expression = tmpl["template"].format(
            base=base_measure,
            date_ref=_dax_column_ref(date_table, date_column),
        )
        chosen_format = format_string or tmpl.get("format_hint") or (inherited_format or "")
        spec = {
            "name": new_name,
            "expression": expression,
            "format_string": chosen_format,
            "description": tmpl.get("description", "").format(base=base_measure),
            "display_folder": display_folder,
        }
        measure_specs.append(spec)
        plan.append({"pattern": token, "measure": new_name, "format_string": chosen_format})

    if dry_run:
        return ok(
            f"Dry run: would create/update {len(measure_specs)} time-intelligence measure(s) "
            f"on '{table}' from base '{base_measure}'.",
            table=table,
            base_measure=base_measure,
            patterns=pattern_list,
            plan=plan,
            measures=measure_specs,
            dry_run=True,
        )

    from . import pbi_create_measures_tool  # late binding: keeps tools.measures patchable

    response = pbi_create_measures_tool(
        manager,
        table=table,
        measures=measure_specs,
        overwrite=overwrite,
        stop_on_error=False,
    )
    response.setdefault("time_intelligence_plan", plan)
    response.setdefault("base_measure", base_measure)
    response.setdefault("patterns", pattern_list)
    return response


def _create_ti_single(
    manager: Any,
    *,
    pattern: str,
    table: str,
    base_measure: str,
    date_table: str,
    date_column: str,
    format_string: str = "",
    overwrite: bool = False,
) -> dict[str, Any]:
    from . import pbi_create_time_intelligence_pack_tool  # late binding

    return pbi_create_time_intelligence_pack_tool(
        manager,
        table=table,
        base_measure=base_measure,
        date_table=date_table,
        date_column=date_column,
        patterns=[pattern],
        format_string=format_string,
        overwrite=overwrite,
    )


def pbi_create_ytd_measure_tool(
    manager: Any,
    *,
    table: str,
    base_measure: str,
    date_table: str,
    date_column: str,
    format_string: str = "",
    overwrite: bool = False,
) -> dict[str, Any]:
    """Create just the YTD companion of ``base_measure``."""
    return _create_ti_single(
        manager,
        pattern="YTD",
        table=table,
        base_measure=base_measure,
        date_table=date_table,
        date_column=date_column,
        format_string=format_string,
        overwrite=overwrite,
    )


def pbi_create_mtd_measure_tool(
    manager: Any,
    *,
    table: str,
    base_measure: str,
    date_table: str,
    date_column: str,
    format_string: str = "",
    overwrite: bool = False,
) -> dict[str, Any]:
    """Create just the MTD companion of ``base_measure``."""
    return _create_ti_single(
        manager,
        pattern="MTD",
        table=table,
        base_measure=base_measure,
        date_table=date_table,
        date_column=date_column,
        format_string=format_string,
        overwrite=overwrite,
    )


def pbi_create_spy_measure_tool(
    manager: Any,
    *,
    table: str,
    base_measure: str,
    date_table: str,
    date_column: str,
    format_string: str = "",
    overwrite: bool = False,
) -> dict[str, Any]:
    """Create the Same-Period-Last-Year companion of ``base_measure``."""
    return _create_ti_single(
        manager,
        pattern="SPY",
        table=table,
        base_measure=base_measure,
        date_table=date_table,
        date_column=date_column,
        format_string=format_string,
        overwrite=overwrite,
    )


def pbi_create_yoy_measure_tool(
    manager: Any,
    *,
    table: str,
    base_measure: str,
    date_table: str,
    date_column: str,
    format_string: str = "",
    overwrite: bool = False,
) -> dict[str, Any]:
    """Create the Year-over-Year delta + SPY companion of ``base_measure``.

    YOY depends on SPY so both measures are created (or refreshed when
    ``overwrite=True``).
    """
    return _create_ti_single(
        manager,
        pattern="YOY",
        table=table,
        base_measure=base_measure,
        date_table=date_table,
        date_column=date_column,
        format_string=format_string,
        overwrite=overwrite,
    )


# ---------------------------------------------------------------------------
# Variance / contribution / Top-N / rolling-average templates
# ---------------------------------------------------------------------------


def pbi_create_variance_measure_tool(
    manager: Any,
    *,
    table: str,
    base_measure: str,
    date_table: str,
    date_column: str,
    measure_name: str | None = None,
    compare_period_offset: int = -1,
    granularity: str = "year",
    format_string: str = "",
    display_folder: str = "Variance",
    overwrite: bool = False,
) -> dict[str, Any]:
    """Create a period-over-period variance measure.

    DAX template (parametrised by ``granularity`` ∈ {year, month, quarter}):

    ``[{base}] - CALCULATE([{base}], DATEADD({date_table}[{date_column}], {offset}, {granularity}))``

    Default ``compare_period_offset = -1`` ⇒ "current period vs previous one".
    """
    validate_model_object_name(table)
    validate_measure_name(base_measure)
    granularity_token = str(granularity).strip().upper()
    if granularity_token not in {"YEAR", "MONTH", "QUARTER", "DAY"}:
        raise PowerBIValidationError(
            "granularity must be one of: year, month, quarter, day.",
            details={"granularity": granularity},
        )
    name = measure_name or f"{base_measure} Variance"
    expression = (
        f"[{base_measure}] - CALCULATE([{base_measure}], "
        f"DATEADD({_dax_column_ref(date_table, date_column)}, {int(compare_period_offset)}, {granularity_token}))"
    )
    from . import pbi_create_measure_tool  # late binding

    return pbi_create_measure_tool(
        manager,
        table=table,
        name=name,
        expression=expression,
        format_string=format_string,
        description=f"Variance of [{base_measure}] vs offset={compare_period_offset} {granularity_token}.",
        display_folder=display_folder,
        overwrite=overwrite,
    )


def pbi_create_contribution_measure_tool(
    manager: Any,
    *,
    table: str,
    base_measure: str,
    scope_columns: list[str],
    measure_name: str | None = None,
    format_string: str = "0.00%",
    display_folder: str = "Contribution",
    overwrite: bool = False,
) -> dict[str, Any]:
    """Create a %-of-total contribution measure.

    DAX template:

    ``DIVIDE([{base}], CALCULATE([{base}], ALL({scope_columns})))``

    ``scope_columns`` defines the denominator scope — typically the dimension
    columns whose total you want each row to be a percentage of (e.g.
    ``["Categorie.Nom catégorie"]`` for "this category's % of all categories").
    """
    validate_model_object_name(table)
    validate_measure_name(base_measure)
    if not scope_columns:
        raise PowerBIValidationError("scope_columns must contain at least one column.")
    qualified: list[str] = []
    for col in scope_columns:
        if "." not in col:
            raise PowerBIValidationError(
                f"scope column '{col}' must use 'TableName.ColumnName' format.",
                details={"column": col},
            )
        tbl, column = col.split(".", 1)
        qualified.append(_dax_column_ref(tbl, column))
    name = measure_name or f"{base_measure} % of total"
    expression = f"DIVIDE([{base_measure}], CALCULATE([{base_measure}], ALL({', '.join(qualified)})))"
    from . import pbi_create_measure_tool  # late binding

    return pbi_create_measure_tool(
        manager,
        table=table,
        name=name,
        expression=expression,
        format_string=format_string,
        description=f"% of total of [{base_measure}] over {', '.join(scope_columns)}.",
        display_folder=display_folder,
        overwrite=overwrite,
    )


def pbi_create_topn_measure_tool(
    manager: Any,
    *,
    table: str,
    base_measure: str,
    n: int,
    dimension_table: str,
    dimension_column: str,
    measure_name: str | None = None,
    rank_measure: str | None = None,
    format_string: str = "",
    display_folder: str = "Top-N",
    overwrite: bool = False,
) -> dict[str, Any]:
    """Create a Top-N filter measure.

    DAX template:

    ``IF(RANKX(ALL({dim_table}[{dim_column}]), [{rank_measure}], , DESC) <= {N}, [{base}], BLANK())``

    Use as the value of a chart visual to surface only the top N members of a
    dimension. ``rank_measure`` defaults to ``base_measure``.
    """
    validate_model_object_name(table)
    validate_measure_name(base_measure)
    if not isinstance(n, int) or n < 1:
        raise PowerBIValidationError("n must be a positive integer.", details={"n": n})
    rank_ref = rank_measure or base_measure
    name = measure_name or f"{base_measure} Top {n}"
    expression = (
        f"IF(RANKX(ALL({_dax_column_ref(dimension_table, dimension_column)}), [{rank_ref}], , DESC) <= {int(n)}, "
        f"[{base_measure}], BLANK())"
    )
    from . import pbi_create_measure_tool  # late binding

    return pbi_create_measure_tool(
        manager,
        table=table,
        name=name,
        expression=expression,
        format_string=format_string,
        description=f"Top-{n} filter on [{base_measure}] over {dimension_table}[{dimension_column}].",
        display_folder=display_folder,
        overwrite=overwrite,
    )


def pbi_create_rolling_average_measure_tool(
    manager: Any,
    *,
    table: str,
    base_measure: str,
    window: int,
    date_table: str,
    date_column: str,
    granularity: str = "month",
    measure_name: str | None = None,
    format_string: str = "",
    display_folder: str = "Rolling",
    overwrite: bool = False,
) -> dict[str, Any]:
    """Create a trailing rolling average measure.

    DAX template:

    ``AVERAGEX(DATESINPERIOD({date_table}[{date_column}], LASTDATE({date_table}[{date_column}]), -{window}, {granularity}), [{base}])``
    """
    validate_model_object_name(table)
    validate_measure_name(base_measure)
    if not isinstance(window, int) or window < 1:
        raise PowerBIValidationError("window must be a positive integer.", details={"window": window})
    granularity_token = str(granularity).strip().upper()
    if granularity_token not in {"YEAR", "MONTH", "QUARTER", "DAY"}:
        raise PowerBIValidationError(
            "granularity must be one of: year, month, quarter, day.",
            details={"granularity": granularity},
        )
    name = measure_name or f"{base_measure} Rolling {window} {granularity_token.title()}"
    date_ref = _dax_column_ref(date_table, date_column)
    expression = (
        f"AVERAGEX("
        f"DATESINPERIOD({date_ref}, LASTDATE({date_ref}), "
        f"-{int(window)}, {granularity_token}), [{base_measure}])"
    )
    from . import pbi_create_measure_tool  # late binding

    return pbi_create_measure_tool(
        manager,
        table=table,
        name=name,
        expression=expression,
        format_string=format_string,
        description=f"Trailing {window}-{granularity_token.lower()} average of [{base_measure}].",
        display_folder=display_folder,
        overwrite=overwrite,
    )

"""Measure operations for the Power BI MCP server."""

from __future__ import annotations

import re
import textwrap
from pathlib import Path
from typing import Any

from pbi_connection import (
    PowerBIDuplicateError,
    PowerBINotFoundError,
    PowerBIValidationError,
    error_payload,
    find_named,
    ok,
    serialize_value,
)
from security import (
    redact_sensitive_data,
    resolve_local_path,
    validate_measure_name,
    validate_model_expression,
    validate_model_object_name,
)


def pbi_list_measures_tool(
    manager: Any,
    *,
    include_hidden: bool = False,
) -> dict[str, Any]:
    """List all model measures."""

    def _reader(state: Any) -> dict[str, Any]:
        measures = []
        for table in state.database.Model.Tables:
            for measure in table.Measures:
                is_hidden = bool(getattr(measure, "IsHidden", False))
                if is_hidden and not include_hidden:
                    continue
                measures.append(
                    {
                        "name": str(measure.Name),
                        "table": str(table.Name),
                        "expression": redact_sensitive_data(str(measure.Expression)),
                        "format_string": serialize_value(getattr(measure, "FormatString", "")),
                        "display_folder": serialize_value(getattr(measure, "DisplayFolder", "")),
                        "description": serialize_value(getattr(measure, "Description", "")),
                        "is_hidden": is_hidden,
                    }
                )
        measures.sort(key=lambda item: (item["table"].casefold(), item["name"].casefold()))
        return {"measures": measures, "connection": state.snapshot()}

    payload = manager.cached_run_read(f"list_measures:h{include_hidden}", "list_measures", _reader)
    return ok(
        "Measures listed successfully.",
        measures=payload["measures"],
        connection=payload["connection"],
    )


def pbi_create_measure_tool(
    manager: Any,
    *,
    table: str,
    name: str,
    expression: str,
    format_string: str = "",
    description: str = "",
    display_folder: str = "",
    is_hidden: bool = False,
    overwrite: bool = True,
) -> dict[str, Any]:
    """Create or update a DAX measure."""
    validate_model_object_name(table)
    validate_measure_name(name)
    validate_model_expression(expression, kind="measure expression")

    def _mutator(state: Any, database: Any, model: Any) -> dict[str, Any]:
        target_table = find_named(model.Tables, table)
        if target_table is None:
            raise PowerBINotFoundError(f"Table '{table}' was not found.", details={"table": table})

        existing = find_named(target_table.Measures, name)
        action = "created"
        if existing is not None and not overwrite:
            raise PowerBIDuplicateError(
                f"Measure '{table}[{name}]' already exists.",
                details={"table": table, "measure": name},
            )

        if existing is None:
            measure = manager.tom.Measure()
            measure.Name = name
            target_table.Measures.Add(measure)
        else:
            measure = existing
            action = "updated"

        measure.Expression = expression
        if format_string:
            measure.FormatString = format_string
        if description:
            measure.Description = description
        if display_folder:
            measure.DisplayFolder = display_folder
        measure.IsHidden = is_hidden

        return {
            "measure": {
                "table": table,
                "name": name,
                "expression": redact_sensitive_data(expression),
                "format_string": format_string or None,
                "description": description or None,
                "display_folder": display_folder or None,
                "is_hidden": is_hidden,
            },
            "action": action,
        }

    payload = manager.execute_write("create_measure", _mutator)
    return ok(
        f"Measure '{table}[{name}]' {payload['action']} successfully.",
        measure=payload["measure"],
        action=payload["action"],
        save_result=payload["save_result"],
        persistence=payload.get("persistence"),
        connection=payload["connection"],
    )


def pbi_delete_measure_tool(manager: Any, *, table: str, name: str) -> dict[str, Any]:
    """Delete a DAX measure."""
    validate_model_object_name(table)
    validate_measure_name(name)

    def _mutator(state: Any, database: Any, model: Any) -> dict[str, Any]:
        target_table = find_named(model.Tables, table)
        if target_table is None:
            raise PowerBINotFoundError(f"Table '{table}' was not found.", details={"table": table})

        measure = find_named(target_table.Measures, name)
        if measure is None:
            raise PowerBINotFoundError(
                f"Measure '{table}[{name}]' was not found.",
                details={"table": table, "measure": name},
            )

        target_table.Measures.Remove(measure)
        return {
            "deleted_measure": {"table": table, "name": name},
        }

    payload = manager.execute_write("delete_measure", _mutator)
    return ok(
        f"Measure '{table}[{name}]' deleted successfully.",
        deleted_measure=payload["deleted_measure"],
        save_result=payload["save_result"],
        persistence=payload.get("persistence"),
        connection=payload["connection"],
    )


def pbi_rename_measure_tool(
    manager: Any,
    *,
    table: str,
    name: str,
    new_name: str,
) -> dict[str, Any]:
    """Rename a DAX measure. Callers must update downstream DAX expressions themselves."""
    validate_model_object_name(table)
    validate_measure_name(name)
    validate_measure_name(new_name)

    def _mutator(state: Any, database: Any, model: Any) -> dict[str, Any]:
        target_table = find_named(model.Tables, table)
        if target_table is None:
            raise PowerBINotFoundError(f"Table '{table}' was not found.", details={"table": table})
        measure = find_named(target_table.Measures, name)
        if measure is None:
            raise PowerBINotFoundError(
                f"Measure '{table}[{name}]' was not found.",
                details={"table": table, "measure": name},
            )
        if new_name.casefold() != name.casefold():
            for candidate_table in model.Tables:
                if find_named(candidate_table.Measures, new_name) is not None:
                    raise PowerBIDuplicateError(
                        f"Measure '{new_name}' already exists in table '{candidate_table.Name}'.",
                        details={"new_name": new_name, "conflict_table": str(candidate_table.Name)},
                    )
        measure.Name = new_name
        return {"rename": {"table": table, "measure_old_name": name, "measure_new_name": new_name}}

    payload = manager.execute_write("rename_measure", _mutator)
    return ok(
        f"Measure '{table}[{name}]' renamed to '{table}[{new_name}]'.",
        rename=payload["rename"],
        save_result=payload["save_result"],
        persistence=payload.get("persistence"),
        connection=payload["connection"],
    )


def pbi_set_format_tool(
    manager: Any,
    *,
    table: str,
    names: list[str],
    format_string: str,
    object_type: str = "measure",
) -> dict[str, Any]:
    """Batch-apply format strings to measures or columns."""
    if not names:
        raise PowerBIValidationError("At least one object name is required.")
    validate_model_object_name(table)
    for object_name in names:
        validate_model_object_name(object_name)
    normalized_type = object_type.strip().casefold()
    if normalized_type not in {"measure", "column"}:
        raise PowerBIValidationError(
            "object_type must be either 'measure' or 'column'.",
            details={"object_type": object_type},
        )

    def _mutator(state: Any, database: Any, model: Any) -> dict[str, Any]:
        target_table = find_named(model.Tables, table)
        if target_table is None:
            raise PowerBINotFoundError(f"Table '{table}' was not found.", details={"table": table})

        collection = target_table.Measures if normalized_type == "measure" else target_table.Columns
        updated = []
        missing = []
        for object_name in names:
            obj = find_named(collection, object_name)
            if obj is None:
                missing.append(object_name)
                continue
            obj.FormatString = format_string
            updated.append(object_name)

        if not updated:
            raise PowerBINotFoundError(
                f"No {normalized_type}s were updated in table '{table}'.",
                details={"table": table, "names": names},
            )

        return {
            "updated": updated,
            "missing": missing,
            "object_type": normalized_type,
            "table": table,
            "format_string": format_string,
        }

    payload = manager.execute_write("set_format", _mutator)
    return ok(
        f"Format string applied to {len(payload['updated'])} {payload['object_type']}(s).",
        updated=payload["updated"],
        missing=payload["missing"],
        object_type=payload["object_type"],
        table=payload["table"],
        format_string=payload["format_string"],
        save_result=payload["save_result"],
        persistence=payload.get("persistence"),
        connection=payload["connection"],
    )


def pbi_import_dax_file_tool(
    manager: Any,
    *,
    path: str,
    table: str = "Measures",
    overwrite: bool = True,
    default_format_string: str = "",
    default_display_folder: str = "",
    stop_on_error: bool = False,
) -> dict[str, Any]:
    """Parse a .dax file and bulk-create measures."""
    validate_model_object_name(table)
    resolved_path = resolve_local_path(path, must_exist=True, allowed_extensions={".dax"})
    measures = _parse_dax_file(resolved_path)
    results = []
    created = 0
    updated = 0
    failed = 0

    for measure in measures:
        try:
            response = pbi_create_measure_tool(
                manager,
                table=table,
                name=measure["name"],
                expression=measure["expression"],
                format_string=default_format_string,
                display_folder=default_display_folder,
                overwrite=overwrite,
            )
            action = response["action"]
            if action == "created":
                created += 1
            elif action == "updated":
                updated += 1
            results.append(
                {
                    "name": measure["name"],
                    "ok": True,
                    "action": action,
                }
            )
        except Exception as exc:
            failed += 1
            results.append(
                {
                    "name": measure["name"],
                    "ok": False,
                    "error": error_payload(exc)["error"],
                }
            )
            if stop_on_error:
                break

    return ok(
        f"Imported {created + updated} measure(s) from '{path}'.",
        table=table,
        source_path=str(resolved_path),
        parsed_count=len(measures),
        created=created,
        updated=updated,
        failed=failed,
        results=results,
    )


def pbi_create_measures_tool(
    manager: Any,
    *,
    table: str,
    measures: list[dict[str, Any]],
    overwrite: bool = True,
    stop_on_error: bool = False,
    dry_run: bool = False,
) -> dict[str, Any]:
    """Batch-create or update multiple DAX measures with a single SaveChanges call.

    Each item in *measures* must have at minimum: name, expression.
    Optional keys: format_string, description, display_folder, is_hidden.

    With ``dry_run=True``, every measure is name + expression validated and the
    per-item ``planned_action`` is reported (``would_create``/``would_update``/
    ``would_fail``), but no model mutation happens — ``SaveChanges`` is never
    called. Use this for an LLM preflight before committing the batch.
    """
    validate_model_object_name(table)
    if not measures:
        raise PowerBIValidationError("At least one measure is required.")

    for item in measures:
        validate_measure_name(item.get("name", ""))
        validate_model_expression(item.get("expression", ""), kind="measure expression")

    if dry_run:
        # Read-only preflight: walk the live model and report what would happen
        # without touching it. SaveChanges is never invoked.
        def _planner(state: Any) -> dict[str, Any]:
            target_table = find_named(state.database.Model.Tables, table)
            if target_table is None:
                raise PowerBINotFoundError(f"Table '{table}' was not found.", details={"table": table})
            seen_names: set[str] = set()
            plan: list[dict[str, Any]] = []
            would_create = would_update = would_fail = 0
            for item in measures:
                name = str(item.get("name", ""))
                duplicate_in_batch = name.casefold() in seen_names
                seen_names.add(name.casefold())
                existing = find_named(target_table.Measures, name)
                if duplicate_in_batch:
                    plan.append({"name": name, "planned_action": "would_fail", "reason": "duplicate name in batch"})
                    would_fail += 1
                    continue
                if existing is not None and not overwrite:
                    plan.append({"name": name, "planned_action": "would_fail", "reason": "exists and overwrite=False"})
                    would_fail += 1
                    continue
                if existing is not None:
                    plan.append({"name": name, "planned_action": "would_update"})
                    would_update += 1
                else:
                    plan.append({"name": name, "planned_action": "would_create"})
                    would_create += 1
            return {
                "table": table,
                "dry_run": True,
                "plan": plan,
                "would_create": would_create,
                "would_update": would_update,
                "would_fail": would_fail,
            }

        payload = manager.run_read("create_measures_batch_dry_run", _planner)
        return ok(
            f"Dry run: would create {payload['would_create']}, update {payload['would_update']}, "
            f"fail {payload['would_fail']} measure(s) in '{table}'.",
            table=payload["table"],
            dry_run=True,
            plan=payload["plan"],
            would_create=payload["would_create"],
            would_update=payload["would_update"],
            would_fail=payload["would_fail"],
        )

    def _mutator(state: Any, database: Any, model: Any) -> dict[str, Any]:
        target_table = find_named(model.Tables, table)
        if target_table is None:
            raise PowerBINotFoundError(f"Table '{table}' was not found.", details={"table": table})

        results: list[dict[str, Any]] = []
        created = 0
        updated = 0
        failed = 0

        for item in measures:
            name = item.get("name", "")
            expression = item.get("expression", "")
            format_string = str(item.get("format_string", "") or "")
            description = str(item.get("description", "") or "")
            display_folder = str(item.get("display_folder", "") or "")
            is_hidden = bool(item.get("is_hidden", False))

            try:
                existing = find_named(target_table.Measures, name)
                action = "created"
                if existing is not None and not overwrite:
                    raise PowerBIDuplicateError(
                        f"Measure '{table}[{name}]' already exists.",
                        details={"table": table, "measure": name},
                    )
                if existing is None:
                    measure = manager.tom.Measure()
                    measure.Name = name
                    target_table.Measures.Add(measure)
                else:
                    measure = existing
                    action = "updated"

                measure.Expression = expression
                if format_string:
                    measure.FormatString = format_string
                if description:
                    measure.Description = description
                if display_folder:
                    measure.DisplayFolder = display_folder
                measure.IsHidden = is_hidden

                if action == "created":
                    created += 1
                else:
                    updated += 1
                results.append({"name": name, "ok": True, "action": action})
            except Exception as exc:
                failed += 1
                results.append({"name": name, "ok": False, "error": error_payload(exc)["error"]})
                if stop_on_error:
                    break

        return {
            "results": results,
            "created": created,
            "updated": updated,
            "failed": failed,
            "table": table,
        }

    payload = manager.execute_write("create_measures_batch", _mutator)
    return ok(
        f"Batch: {payload['created']} created, {payload['updated']} updated, {payload['failed']} failed in '{table}'.",
        table=payload["table"],
        results=payload["results"],
        created=payload["created"],
        updated=payload["updated"],
        failed=payload["failed"],
        save_result=payload["save_result"],
        persistence=payload.get("persistence"),
        connection=payload["connection"],
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


def _parse_dax_file(path: str | Path) -> list[dict[str, str]]:
    resolved = resolve_local_path(str(path), must_exist=True, allowed_extensions={".dax"})
    if not resolved.exists():
        raise PowerBINotFoundError(f"DAX file '{resolved}' was not found.", details={"path": str(resolved)})

    raw_text = resolved.read_text(encoding="utf-8")
    cleaned_text = _strip_dax_comments(raw_text)
    normalized_text = "\n".join(line.rstrip() for line in cleaned_text.splitlines())
    blocks = [block.strip() for block in re.split(r"(?:\n\s*){2,}", normalized_text) if block.strip()]
    if not blocks:
        raise PowerBIValidationError(f"DAX file '{resolved}' is empty.", details={"path": str(resolved)})

    parsed: list[dict[str, str]] = []
    for index, block in enumerate(blocks, start=1):
        lines = block.splitlines()
        header = lines[0]
        match = re.match(r"^\s*(?P<name>[^=]+?)\s*=\s*(?P<inline>.*)$", header)
        if not match:
            raise PowerBIValidationError(
                f"Invalid measure header in block {index}: '{header}'. Expected 'MeasureName ='",
                details={"path": str(resolved), "block": index},
            )

        name = match.group("name").strip()
        inline_expression = match.group("inline").strip()
        expression_lines = []
        if inline_expression:
            expression_lines.append(inline_expression)
        expression_lines.extend(lines[1:])
        expression = textwrap.dedent("\n".join(expression_lines)).strip()

        if not name:
            raise PowerBIValidationError(
                f"Block {index} is missing a measure name.",
                details={"path": str(resolved), "block": index},
            )
        if not expression:
            raise PowerBIValidationError(
                f"Block {index} is missing a DAX expression for measure '{name}'.",
                details={"path": str(resolved), "block": index, "measure": name},
            )
        validate_measure_name(name)
        validate_model_expression(expression, kind="measure expression")

        parsed.append({"name": name, "expression": expression})

    return parsed


def _strip_dax_comments(text: str) -> str:
    """Remove // and /* */ comments while preserving text inside string literals."""
    output: list[str] = []
    index = 0
    in_string = False
    in_line_comment = False
    in_block_comment = False

    while index < len(text):
        char = text[index]
        next_char = text[index + 1] if index + 1 < len(text) else ""

        if in_line_comment:
            if char in "\r\n":
                in_line_comment = False
                output.append(char)
            index += 1
            continue

        if in_block_comment:
            if char == "*" and next_char == "/":
                in_block_comment = False
                index += 2
                continue
            if char in "\r\n":
                output.append(char)
            index += 1
            continue

        if in_string:
            output.append(char)
            if char == '"':
                if next_char == '"':
                    output.append(next_char)
                    index += 2
                    continue
                in_string = False
            index += 1
            continue

        if char == '"':
            in_string = True
            output.append(char)
            index += 1
            continue
        if char == "/" and next_char == "/":
            in_line_comment = True
            index += 2
            continue
        if char == "/" and next_char == "*":
            in_block_comment = True
            index += 2
            continue

        output.append(char)
        index += 1

    return "".join(output)

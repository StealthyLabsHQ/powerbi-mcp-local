"""Visual binding builders + validators.

Builds the ``prototypeQuery`` Select / From entries that Power BI's
report engine needs, and runs pre-flight checks against the live model
so we never ship a layout PBI Desktop will refuse to render.
"""

from __future__ import annotations

import logging
from typing import Any

from pbi_connection import PowerBIValidationError, error_payload

from ._base import VISUAL_FIELD_ROLES, VISUAL_ROLE_KINDS
from ._layout import _dump_embedded_json, _parse_embedded_json
from ._refs import _normalize_reference, _query_ref, _split_column_ref


logger = logging.getLogger("tools.visuals._bindings")


def _build_select_entry(
    reference: str,
    aliases: dict[str, str],
    measure_home_map: dict[str, str] | None = None,
) -> dict[str, Any]:
    # Normalise so Date[Année] / 'Date'[Année] / Date.Année all enter the same path.
    reference = _normalize_reference(reference)
    if "." in reference:
        table, column = _split_column_ref(reference)
        alias = aliases.setdefault(table, f"s{len(aliases)}")
        return {
            "Column": {"Expression": {"SourceRef": {"Source": alias}}, "Property": column},
            "Name": column,
            "NativeReferenceName": column,
        }
    measure_entity = (measure_home_map or {}).get(reference) or "$Measures"
    if measure_entity == "$Measures":
        logger.warning(
            "Measure '%s' home table not found in extract metadata; using '$Measures' fallback.",
            reference,
        )
    alias = aliases.setdefault(measure_entity, f"s{len(aliases)}")
    return {
        "Measure": {"Expression": {"SourceRef": {"Source": alias}}, "Property": reference},
        "Name": reference,
        "NativeReferenceName": reference,
    }


def _build_prototype_query(
    references: list[str],
    measure_home_map: dict[str, str] | None = None,
) -> dict[str, Any]:
    aliases: dict[str, str] = {}
    select = [_build_select_entry(reference, aliases, measure_home_map) for reference in references]
    from_entries = [{"Name": alias, "Entity": entity} for entity, alias in aliases.items()]
    return {"Version": 2, "From": from_entries, "Select": select}


def _select_name_map(prototype_query: dict[str, Any]) -> dict[str, str]:
    names: dict[str, str] = {}
    for entry in prototype_query.get("Select", []) or []:
        if not isinstance(entry, dict):
            continue
        name = str(entry.get("Name", ""))
        if not name:
            continue
        if "Column" in entry:
            column = entry.get("Column", {})
            if isinstance(column, dict):
                prop = str(column.get("Property", ""))
                if prop:
                    names[prop.casefold()] = name
        if "Measure" in entry:
            measure = entry.get("Measure", {})
            if isinstance(measure, dict):
                prop = str(measure.get("Property", ""))
                if prop:
                    names[prop.casefold()] = name
        names[name.casefold()] = name
    return names


def _from_entity_by_alias(prototype_query: dict[str, Any]) -> dict[str, str]:
    entities: dict[str, str] = {}
    for entry in prototype_query.get("From", []) or []:
        if isinstance(entry, dict):
            entities[str(entry.get("Name", ""))] = str(entry.get("Entity", ""))
    return entities


def _next_alias(existing: set[str]) -> str:
    index = 0
    while f"s{index}" in existing:
        index += 1
    alias = f"s{index}"
    existing.add(alias)
    return alias


def _sync_container_query(container: dict[str, Any], prototype_query: dict[str, Any]) -> None:
    query_payload = _parse_embedded_json(container.get("query"), {})
    try:
        commands = query_payload.setdefault("Commands", [])
        if not commands:
            commands.append({"SemanticQueryDataShapeCommand": {}})
        commands[0].setdefault("SemanticQueryDataShapeCommand", {})["Query"] = prototype_query
        container["query"] = _dump_embedded_json(query_payload)
    except Exception:
        container["query"] = _dump_embedded_json(
            {"Commands": [{"SemanticQueryDataShapeCommand": {"Query": prototype_query}}]}
        )


def _live_model_field_index(manager: Any | None, *, include_hidden: bool) -> tuple[dict[str, Any] | None, dict[str, Any]]:
    if manager is None:
        return None, {"status": "unavailable", "reason": "manager_not_provided"}
    # Resolve through the visuals package re-export so test patches against
    # ``tools.visuals.pbi_model_info_tool`` keep working.
    from . import pbi_model_info_tool

    try:
        model = pbi_model_info_tool(manager, include_hidden=include_hidden, include_row_counts=False)
    except Exception as exc:
        return None, {"status": "unavailable", "error": error_payload(exc)["error"]}
    if not model.get("ok"):
        return None, {"status": "unavailable", "error": model.get("error")}

    columns: set[tuple[str, str]] = set()
    measures: dict[str, set[str]] = {}
    measure_tables: dict[str, set[str]] = {}
    for table in model.get("tables", []) or []:
        table_name = str(table.get("name", ""))
        for column in table.get("columns", []) or []:
            columns.add((table_name.casefold(), str(column.get("name", "")).casefold()))
    for measure in model.get("measures", []) or []:
        name = str(measure.get("name", ""))
        table_name = str(measure.get("table", ""))
        measures.setdefault(name.casefold(), set()).add(table_name.casefold())
        measure_tables.setdefault(name.casefold(), set()).add(table_name)
    return {"columns": columns, "measures": measures, "measure_tables": measure_tables}, {"status": "available"}


def _validate_projection_roles(
    visual_type: str,
    projections: dict[str, list[dict[str, str]]] | None,
    *,
    manager: Any | None = None,
    include_hidden: bool = False,
) -> dict[str, Any]:
    """Pre-flight check that every role used in ``projections`` is allowed for
    ``visual_type`` and (when a live model is reachable) carries a reference of
    the expected kind (column vs measure).
    """
    # Resolve through the package re-export so tests can patch
    # ``tools.visuals._live_model_field_index``.
    from . import _live_model_field_index

    if not isinstance(projections, dict) or not projections:
        return {"status": "skipped", "reason": "no_projections"}
    allowed = VISUAL_FIELD_ROLES.get(visual_type)
    role_kinds = VISUAL_ROLE_KINDS.get(visual_type, {})

    unknown_roles: list[str] = []
    if allowed is not None:
        for role in projections:
            if role not in allowed:
                unknown_roles.append(role)
    if unknown_roles:
        raise PowerBIValidationError(
            f"Visual type '{visual_type}' does not accept role(s): {', '.join(sorted(unknown_roles))}.",
            details={
                "visual_type": visual_type,
                "unknown_roles": sorted(unknown_roles),
                "allowed_roles": sorted(allowed) if allowed else [],
            },
        )

    if manager is None or not role_kinds:
        return {"status": "roles_only_checked"}
    index, status = _live_model_field_index(manager, include_hidden=include_hidden)
    if index is None:
        return status

    role_kind_mismatches: list[dict[str, str]] = []
    for role, items in projections.items():
        expected_kind = role_kinds.get(role, "any")
        if expected_kind == "any":
            continue
        for item in items or []:
            ref = item.get("queryRef") if isinstance(item, dict) else None
            if not isinstance(ref, str) or not ref.strip():
                continue
            actual_kind: str | None = None
            if ref.casefold() in index["measures"]:
                actual_kind = "measure"
            else:
                for table_lc, col_lc in index["columns"]:
                    if col_lc == ref.casefold():
                        actual_kind = "column"
                        break
            if actual_kind is not None and actual_kind != expected_kind:
                role_kind_mismatches.append({
                    "role": role,
                    "reference": ref,
                    "expected_kind": expected_kind,
                    "actual_kind": actual_kind,
                })
    if role_kind_mismatches:
        raise PowerBIValidationError(
            "Projection role/kind mismatch — at least one reference is the wrong kind for its role.",
            details={
                "visual_type": visual_type,
                "mismatches": role_kind_mismatches,
            },
        )
    return {"status": "roles_and_kinds_checked"}


def _validate_field_references_live(
    manager: Any | None,
    references: list[str],
    *,
    expected_kinds: dict[str, str] | None = None,
    include_hidden: bool = False,
) -> dict[str, Any]:
    """If a connection manager is available, verify each reference exists in
    the live model. Raises ``PowerBIValidationError`` on missing fields so
    callers fail fast before writing the layout.
    """
    # Resolve through the package re-export for patchability.
    from . import _live_model_field_index

    if manager is None or not references:
        return {"status": "skipped"}
    index, status = _live_model_field_index(manager, include_hidden=include_hidden)
    if index is None:
        return status
    expected_kinds = expected_kinds or {}

    import difflib
    measure_names = sorted(index.get("measures", {}).keys())
    column_short_names = sorted({col_lc for _t, col_lc in index.get("columns", set())})

    def _close_measure_names(needle: str, n: int = 5) -> list[str]:
        return difflib.get_close_matches(needle.casefold(), measure_names, n=n, cutoff=0.6)

    def _close_column_names(needle: str, n: int = 5) -> list[str]:
        return difflib.get_close_matches(needle.casefold(), column_short_names, n=n, cutoff=0.6)

    missing: list[dict[str, Any]] = []
    for ref in references:
        if not isinstance(ref, str) or not ref.strip():
            continue
        normalized = _normalize_reference(ref)
        expected_kind = expected_kinds.get(ref)
        if "." in normalized:
            table, column = normalized.split(".", 1)
            if (table.casefold(), column.casefold()) not in index["columns"]:
                suggestions = _close_column_names(column)
                missing.append({
                    "reference": ref,
                    "kind": expected_kind or "column",
                    "inferred_kind": "column",
                    "hint": "use 'Table.Column', 'Table[Column]', or \"'Table With Spaces'[Column]\"",
                    "did_you_mean": suggestions,
                })
        else:
            measure_hit = normalized.casefold() in index["measures"]
            column_short_hit = any(col_lc == normalized.casefold() for _table_lc, col_lc in index["columns"])
            if expected_kind == "column":
                qualified_examples = sorted(
                    {f"{tbl}.{col}" for tbl_lc, col_lc in index["columns"]
                     for tbl, col in [(_t, _c) for (_t, _c) in [(tbl_lc, col_lc)]]
                     if col_lc == normalized.casefold()}
                )
                hint = (
                    "axis/category/rows expect a column — qualify with the table "
                    "(e.g. 'Date.Year' or 'Date[Year]')."
                )
                if qualified_examples:
                    hint += f" Try one of: {', '.join(qualified_examples)}."
                missing.append({
                    "reference": ref,
                    "kind": "column",
                    "inferred_kind": "measure" if measure_hit else ("column" if column_short_hit else "unknown"),
                    "hint": hint,
                    "did_you_mean": _close_column_names(normalized),
                })
            elif expected_kind == "measure":
                if not measure_hit:
                    missing.append({
                        "reference": ref,
                        "kind": "measure",
                        "inferred_kind": "column" if column_short_hit else "unknown",
                        "hint": "values/Y/indicator expect a measure — check spelling against the live model's measure list.",
                        "did_you_mean": _close_measure_names(normalized),
                    })
            else:
                if not measure_hit:
                    suggestions = _close_measure_names(normalized) or _close_column_names(normalized)
                    missing.append({
                        "reference": ref,
                        "kind": "measure",
                        "inferred_kind": "column" if column_short_hit else "unknown",
                        "hint": (
                            "no measure with that exact name in the live model — "
                            "check spelling, or pass a 'Table.Column' / 'Table[Column]' "
                            "form if you meant a column."
                        ),
                        "did_you_mean": suggestions,
                    })
    if missing:
        raise PowerBIValidationError(
            f"Field reference(s) not found in the live model: "
            f"{', '.join(item['reference'] for item in missing)}",
            details={
                "missing": missing,
                "checked": list(references),
                "available_measure_count": len(measure_names),
                "available_column_count": len(column_short_names),
            },
        )
    return {"status": "validated", "checked": len(references)}


def _visual_binding_issues(
    container: dict[str, Any],
    page_name: str,
    measure_home_map: dict[str, str],
    model_fields: dict[str, Any] | None = None,
    *,
    repair: bool = False,
) -> tuple[list[dict[str, Any]], int]:
    config = _parse_embedded_json(container.get("config"), {})
    if not isinstance(config, dict):
        return ([{"page": page_name, "visual_id": "", "issue": "invalid_config"}], 0)
    single_visual = config.get("singleVisual", {})
    if not isinstance(single_visual, dict):
        return ([], 0)
    visual_id = str(config.get("name", ""))
    visual_type = str(single_visual.get("visualType", ""))
    prototype_query = single_visual.get("prototypeQuery", {})
    if not isinstance(prototype_query, dict):
        return ([], 0)

    issues: list[dict[str, Any]] = []
    repairs = 0
    select_names = _select_name_map(prototype_query)
    from_entities = _from_entity_by_alias(prototype_query)

    allowed_roles = VISUAL_FIELD_ROLES.get(visual_type)
    projections = single_visual.get("projections", {})
    if isinstance(projections, dict):
        if repair and visual_type == "gauge" and "Value" in projections and "Y" not in projections:
            projections["Y"] = projections.pop("Value")
            issues.append({"page": page_name, "visual_id": visual_id, "visual_type": visual_type, "issue": "projection_role_repaired", "from": "Value", "to": "Y"})
            repairs += 1
        for role, items in list(projections.items()):
            if allowed_roles is not None and role not in allowed_roles:
                issues.append({"page": page_name, "visual_id": visual_id, "visual_type": visual_type, "issue": "unexpected_projection_role", "role": role, "allowed_roles": sorted(allowed_roles)})
            if not isinstance(items, list):
                continue
            for item in items:
                if not isinstance(item, dict):
                    continue
                query_ref = str(item.get("queryRef", ""))
                expected = select_names.get(query_ref.casefold())
                if expected is None:
                    short = _query_ref(query_ref)
                    expected = select_names.get(short.casefold())
                if expected is None:
                    issues.append({"page": page_name, "visual_id": visual_id, "visual_type": visual_type, "issue": "query_ref_not_found", "queryRef": query_ref})
                    continue
                if query_ref != expected:
                    issues.append({"page": page_name, "visual_id": visual_id, "visual_type": visual_type, "issue": "query_ref_mismatch", "queryRef": query_ref, "expected": expected})
                    if repair:
                        item["queryRef"] = expected
                        repairs += 1

    from_entries = prototype_query.get("From", []) or []
    aliases = {str(entry.get("Name", "")) for entry in from_entries if isinstance(entry, dict)}
    for entry in prototype_query.get("Select", []) or []:
        if not isinstance(entry, dict):
            continue
        if "Column" in entry:
            column = entry.get("Column", {})
            if isinstance(column, dict):
                column_name = str(column.get("Property", ""))
                source_ref = column.get("Expression", {}).get("SourceRef", {}) if isinstance(column.get("Expression"), dict) else {}
                alias = str(source_ref.get("Source", "")) if isinstance(source_ref, dict) else ""
                table_name = from_entities.get(alias, "")
                if model_fields is not None and (table_name.casefold(), column_name.casefold()) not in model_fields["columns"]:
                    issues.append({"page": page_name, "visual_id": visual_id, "visual_type": visual_type, "issue": "column_not_found", "table": table_name, "column": column_name})
            continue
        if "Measure" not in entry:
            continue
        measure = entry.get("Measure", {})
        if not isinstance(measure, dict):
            continue
        measure_name = str(measure.get("Property", ""))
        source_ref = measure.get("Expression", {}).get("SourceRef", {}) if isinstance(measure.get("Expression"), dict) else {}
        alias = str(source_ref.get("Source", "")) if isinstance(source_ref, dict) else ""
        entity = from_entities.get(alias, "")
        home_table = measure_home_map.get(measure_name)
        home_table_source = "extract_metadata" if home_table is not None else ""
        if home_table is None and model_fields is not None:
            live_tables = sorted(model_fields.get("measure_tables", {}).get(measure_name.casefold(), set()))
            if len(live_tables) == 1:
                home_table = live_tables[0]
                home_table_source = "live_model"
        if model_fields is not None:
            measure_tables = model_fields["measures"].get(measure_name.casefold(), set())
            if not measure_tables:
                issues.append({"page": page_name, "visual_id": visual_id, "visual_type": visual_type, "issue": "measure_not_found", "measure": measure_name})
            elif entity and entity != "$Measures" and entity.casefold() not in measure_tables:
                issues.append({"page": page_name, "visual_id": visual_id, "visual_type": visual_type, "issue": "measure_table_mismatch", "measure": measure_name, "table": entity, "expected_tables": sorted(measure_tables)})
        if entity == "$Measures":
            if not home_table:
                issues.append({"page": page_name, "visual_id": visual_id, "visual_type": visual_type, "issue": "measure_home_table_unknown", "measure": measure_name})
                continue
            if not repair:
                item = {"page": page_name, "visual_id": visual_id, "visual_type": visual_type, "issue": "measure_home_table_needs_repair", "measure": measure_name, "home_table": home_table}
                if home_table_source == "live_model":
                    item.update({"source": "live_model", "extract_metadata": "missing"})
                issues.append(item)
                continue
            same_alias_measures = [
                str(item.get("Measure", {}).get("Property", ""))
                for item in prototype_query.get("Select", []) or []
                if isinstance(item, dict)
                and isinstance(item.get("Measure"), dict)
                and item.get("Measure", {}).get("Expression", {}).get("SourceRef", {}).get("Source") == alias
            ]

            def _resolved_measure_home(item: str) -> str | None:
                if item in measure_home_map:
                    return measure_home_map[item]
                if model_fields is not None:
                    live = sorted(model_fields.get("measure_tables", {}).get(item.casefold(), set()))
                    if len(live) == 1:
                        return live[0]
                return None

            if all(_resolved_measure_home(item) == home_table for item in same_alias_measures):
                for from_entry in from_entries:
                    if isinstance(from_entry, dict) and str(from_entry.get("Name", "")) == alias:
                        from_entry["Entity"] = home_table
                        break
            else:
                new_alias = _next_alias(aliases)
                from_entries.append({"Name": new_alias, "Entity": home_table})
                measure.setdefault("Expression", {}).setdefault("SourceRef", {})["Source"] = new_alias
            item = {"page": page_name, "visual_id": visual_id, "visual_type": visual_type, "issue": "measure_home_table_repaired", "measure": measure_name, "home_table": home_table}
            if home_table_source == "live_model":
                item.update({"source": "live_model", "extract_metadata": "missing"})
            issues.append(item)
            repairs += 1

    if repair and repairs:
        single_visual["prototypeQuery"] = prototype_query
        container["config"] = _dump_embedded_json(config)
        _sync_container_query(container, prototype_query)
    return issues, repairs


def _scan_visual_bindings(
    layout: dict[str, Any],
    measure_home_map: dict[str, str],
    model_fields: dict[str, Any] | None = None,
    *,
    page: str | None = None,
    repair: bool = False,
) -> tuple[list[dict[str, Any]], int]:
    issues: list[dict[str, Any]] = []
    repairs = 0
    sections = layout.get("sections", []) or []
    for section in sections:
        if not isinstance(section, dict):
            continue
        section_name = str(section.get("displayName") or section.get("name") or "")
        if page and page.casefold() not in {str(section.get("name", "")).casefold(), str(section.get("displayName", "")).casefold()}:
            continue
        for container in section.get("visualContainers", []) or []:
            if not isinstance(container, dict):
                continue
            found, fixed = _visual_binding_issues(container, section_name, measure_home_map, model_fields, repair=repair)
            issues.extend(found)
            repairs += fixed
    return issues, repairs


def _assert_container_bindings(container: dict[str, Any], measure_home_map: dict[str, str]) -> None:
    issues, _ = _visual_binding_issues(container, "", measure_home_map, repair=False)
    blocking = [item for item in issues if item.get("issue") in {"unexpected_projection_role", "query_ref_not_found", "query_ref_mismatch"}]
    if blocking:
        raise PowerBIValidationError("Visual field bindings are invalid.", details={"issues": blocking})

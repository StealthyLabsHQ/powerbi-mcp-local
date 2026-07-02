"""Model audit tools: structure, collisions, dirty dates, relationships, schema."""

from __future__ import annotations

from datetime import datetime
from typing import Any

from pbi_connection import PowerBIValidationError, dax_quote_table_name, ok

from ._shared import DATE_PARSE_FORMATS, _dax_column, _row_value


def _model_audit_from_snapshot(snapshot: dict[str, Any]) -> dict[str, Any]:
    tables = snapshot.get("tables", [])
    relationships = snapshot.get("relationships", [])
    measures = snapshot.get("measures", [])
    issues: list[dict[str, Any]] = []
    warnings: list[dict[str, Any]] = []

    visible_tables = {item["name"]: item for item in tables if not item.get("is_hidden")}
    related_tables: set[str] = set()
    graph: dict[str, set[str]] = {name: set() for name in visible_tables}
    pair_count: dict[tuple[str, str], int] = {}

    for rel in relationships:
        from_table = str(rel.get("from_table", ""))
        to_table = str(rel.get("to_table", ""))
        related_tables.update({from_table, to_table})
        if from_table in graph and to_table in graph:
            graph[from_table].add(to_table)
            graph[to_table].add(from_table)
        pair = tuple(sorted((from_table, to_table)))
        pair_count[pair] = pair_count.get(pair, 0) + 1
        direction = str(rel.get("direction", ""))
        if direction and direction.casefold() not in {"onedirection", "single"}:
            issues.append({"type": "bidirectional_relationship", "relationship": rel})
        if pair_count[pair] > 1:
            warnings.append({"type": "parallel_relationships", "tables": list(pair), "count": pair_count[pair]})

    if len(visible_tables) > 1:
        for table_name in sorted(set(visible_tables) - related_tables):
            warnings.append({"type": "unrelated_table", "table": table_name})

    for table_name, table in visible_tables.items():
        date_columns = [
            col["name"]
            for col in table.get("columns", [])
            if "date" in str(col.get("name", "")).casefold() or str(col.get("data_type", "")).casefold() == "datetime"
        ]
        if date_columns and table_name not in related_tables and len(visible_tables) > 1:
            warnings.append({"type": "unrelated_date_columns", "table": table_name, "columns": date_columns})

    for a, neighbors in graph.items():
        for b in neighbors:
            for c in neighbors.intersection(graph.get(b, set())):
                if a < b < c:
                    warnings.append({"type": "ambiguous_relationship_triangle", "tables": [a, b, c]})

    measure_tables = {str(item.get("table", "")): [] for item in measures}
    for measure in measures:
        measure_tables.setdefault(str(measure.get("table", "")), []).append(measure)
    for table in visible_tables:
        if table not in related_tables and not measure_tables.get(table) and len(visible_tables) > 1:
            warnings.append({"type": "orphan_table", "table": table})

    return {
        "valid": not issues,
        "issue_count": len(issues),
        "warning_count": len(warnings),
        "issues": issues,
        "warnings": warnings,
    }


def _table_map(snapshot: dict[str, Any]) -> dict[str, dict[str, Any]]:
    return {str(item.get("name", "")).casefold(): item for item in snapshot.get("tables", [])}


def _find_table(snapshot: dict[str, Any], table: str) -> dict[str, Any] | None:
    return _table_map(snapshot).get(table.casefold())


def _find_column(snapshot: dict[str, Any], table: str, column: str) -> dict[str, Any] | None:
    found = _find_table(snapshot, table)
    if not found:
        return None
    for item in found.get("columns", []):
        if str(item.get("name", "")).casefold() == column.casefold():
            return item
    return None


def _column_profile(manager: Any, table: str, column: str) -> dict[str, Any]:
    query = (
        "EVALUATE ROW("
        '"__Rows", COUNTROWS(' + dax_quote_table_name(table) + "), "
        '"__Distinct", DISTINCTCOUNT(' + _dax_column(table, column) + "), "
        '"__Blank", COUNTBLANK(' + _dax_column(table, column) + ")"
        ")"
    )
    result = manager.run_adomd_query(query, max_rows=1)
    row = result.get("rows", [{}])[0] if result.get("rows") else {}
    return {
        "row_count": _row_value(row, "__Rows"),
        "distinct_count": _row_value(row, "__Distinct"),
        "blank_count": _row_value(row, "__Blank"),
    }


def _graph_paths(graph: dict[str, set[str]], start: str, end: str, *, limit: int = 2) -> list[list[str]]:
    paths: list[list[str]] = []

    def _walk(node: str, target: str, seen: list[str]) -> None:
        if len(paths) >= limit:
            return
        if node == target:
            paths.append(seen[:])
            return
        for nxt in sorted(graph.get(node, set())):
            if nxt not in seen:
                _walk(nxt, target, [*seen, nxt])

    _walk(start, end, [start])
    return paths


def _duplicate_relationship_key_issues(manager: Any, relationships: list[dict[str, Any]]) -> list[dict[str, Any]]:
    issues: list[dict[str, Any]] = []
    checked: set[tuple[str, str]] = set()
    candidates: list[tuple[str, str, str]] = []
    for rel in relationships:
        candidates.append((str(rel.get("to_table", "")), str(rel.get("to_column", "")), "one_side"))
        from_table = str(rel.get("from_table", ""))
        to_table = str(rel.get("to_table", ""))
        if to_table.casefold().startswith("fact") and not from_table.casefold().startswith("fact"):
            candidates.append((from_table, str(rel.get("from_column", "")), "non_fact_many_side"))
    for table, column, role in candidates:
        key = (table.casefold(), column.casefold())
        if key in checked:
            continue
        checked.add(key)
        query = (
            "EVALUATE ROW("
            '"__Rows", COUNTROWS(' + dax_quote_table_name(table) + "), "
            '"__Distinct", DISTINCTCOUNT(' + _dax_column(table, column) + ")"
            ")"
        )
        try:
            result = manager.run_adomd_query(query, max_rows=1)
        except Exception as exc:
            issues.append(
                {"type": "relationship_key_check_failed", "table": table, "column": column, "error": str(exc)}
            )
            continue
        rows = result.get("rows", [])
        if not rows:
            continue
        row_count = rows[0].get("__Rows", rows[0].get("[__Rows]"))
        distinct_count = rows[0].get("__Distinct", rows[0].get("[__Distinct]"))
        if row_count is not None and distinct_count is not None and row_count != distinct_count:
            issues.append(
                {
                    "type": "duplicate_relationship_key",
                    "table": table,
                    "column": column,
                    "relationship_role": role,
                    "row_count": row_count,
                    "distinct_count": distinct_count,
                }
            )
    return issues


def pbi_audit_model_tool(manager: Any, *, include_hidden: bool = False) -> dict[str, Any]:
    """Detect missing, ambiguous, bidirectional, and orphaned model structures."""
    from . import _duplicate_relationship_key_issues, _model_snapshot

    snapshot = _model_snapshot(manager, include_hidden=include_hidden)
    audit = _model_audit_from_snapshot(snapshot)
    duplicate_key_issues = _duplicate_relationship_key_issues(manager, snapshot.get("relationships", []))
    if duplicate_key_issues:
        audit["issues"].extend(duplicate_key_issues)
        audit["issue_count"] = len(audit["issues"])
        audit["valid"] = False
    return ok(
        f"Model audit found {audit['issue_count']} issue(s), {audit['warning_count']} warning(s).",
        include_hidden=include_hidden,
        table_count=len(snapshot.get("tables", [])),
        measure_count=len(snapshot.get("measures", [])),
        relationship_count=len(snapshot.get("relationships", [])),
        **audit,
    )


def pbi_detect_name_collisions_tool(manager: Any, *, include_hidden: bool = False) -> dict[str, Any]:
    """Detect table, column, and measure name collisions before writes."""
    from . import _model_snapshot

    snapshot = _model_snapshot(manager, include_hidden=include_hidden)
    issues: list[dict[str, Any]] = []
    warnings: list[dict[str, Any]] = []
    table_names: dict[str, list[str]] = {}
    global_columns: dict[str, list[dict[str, str]]] = {}
    global_measures: dict[str, list[dict[str, str]]] = {}

    for table in snapshot.get("tables", []):
        table_name = str(table.get("name", ""))
        table_names.setdefault(table_name.casefold(), []).append(table_name)
        local_columns: dict[str, list[str]] = {}
        for column in table.get("columns", []):
            column_name = str(column.get("name", ""))
            local_columns.setdefault(column_name.casefold(), []).append(column_name)
            global_columns.setdefault(column_name.casefold(), []).append({"table": table_name, "column": column_name})
        for _key, names in local_columns.items():
            if len(names) > 1:
                issues.append(
                    {"type": "duplicate_column_name", "table": table_name, "name": names[0], "count": len(names)}
                )

    for measure in snapshot.get("measures", []):
        table = str(measure.get("table", ""))
        name = str(measure.get("name", ""))
        global_measures.setdefault(name.casefold(), []).append({"table": table, "measure": name})

    for names in table_names.values():
        if len(names) > 1:
            issues.append({"type": "duplicate_table_name", "name": names[0], "count": len(names)})
    for name, measures in global_measures.items():
        if len(measures) > 1:
            warnings.append({"type": "duplicate_measure_name", "name": measures[0]["measure"], "measures": measures})
        for measure in measures:
            same_table_columns = [
                item for item in global_columns.get(name, []) if item["table"].casefold() == measure["table"].casefold()
            ]
            if same_table_columns:
                issues.append(
                    {
                        "type": "measure_column_name_collision",
                        "table": measure["table"],
                        "measure": measure["measure"],
                        "columns": same_table_columns,
                    }
                )
    for name, columns in global_columns.items():
        tables = {item["table"].casefold() for item in columns}
        if len(tables) > 1:
            warnings.append(
                {"type": "same_column_name_across_tables", "name": columns[0]["column"], "columns": columns}
            )

    return ok(
        f"Name collision scan found {len(issues)} issue(s), {len(warnings)} warning(s).",
        include_hidden=include_hidden,
        valid=not issues,
        issue_count=len(issues),
        warning_count=len(warnings),
        issues=issues,
        warnings=warnings,
    )


def pbi_detect_dirty_dates_tool(
    manager: Any,
    *,
    table: str | None = None,
    max_samples: int = 200,
    min_parse_success_rate: float = 0.8,
    scan_all_text_columns: bool = False,
) -> dict[str, Any]:
    """Detect text columns that look like dirty dates."""
    from . import _model_snapshot

    snapshot = _model_snapshot(manager, include_hidden=False)
    if max_samples < 1 or max_samples > 1000:
        raise PowerBIValidationError("max_samples must be between 1 and 1000.", details={"max_samples": max_samples})
    if not 0 <= min_parse_success_rate <= 1:
        raise PowerBIValidationError(
            "min_parse_success_rate must be between 0 and 1.",
            details={"min_parse_success_rate": min_parse_success_rate},
        )

    tables = snapshot.get("tables", [])
    if table:
        found = _find_table(snapshot, table)
        if found is None:
            raise PowerBIValidationError("Table was not found.", details={"table": table})
        tables = [found]

    issues: list[dict[str, Any]] = []
    warnings: list[dict[str, Any]] = []
    scanned: list[dict[str, Any]] = []
    for item in tables:
        table_name = str(item.get("name", ""))
        for column in item.get("columns", []):
            column_name = str(column.get("name", ""))
            data_type = str(column.get("data_type", ""))
            is_text = data_type.casefold() in {"string", "text"}
            if not is_text:
                continue
            name_suggests_date = "date" in column_name.casefold()
            if not (scan_all_text_columns or name_suggests_date):
                continue
            query = (
                "EVALUATE TOPN("
                + str(max_samples)
                + ", SELECTCOLUMNS("
                + dax_quote_table_name(table_name)
                + ', "__Value", '
                + _dax_column(table_name, column_name)
                + "))"
            )
            try:
                result = manager.run_adomd_query(query, max_rows=max_samples)
            except Exception as exc:
                issues.append(
                    {"type": "dirty_date_scan_failed", "table": table_name, "column": column_name, "error": str(exc)}
                )
                continue
            values = [str(_row_value(row, "__Value") or "").strip() for row in result.get("rows", [])]
            non_blank = [value for value in values if value]
            parsed = 0
            formats: set[str] = set()
            invalid_examples: list[str] = []
            for value in non_blank:
                matched = False
                for fmt in DATE_PARSE_FORMATS:
                    try:
                        datetime.strptime(value, fmt)
                        parsed += 1
                        formats.add(fmt)
                        matched = True
                        break
                    except ValueError:
                        pass
                if not matched and len(invalid_examples) < 5:
                    invalid_examples.append(value)
            parse_rate = parsed / len(non_blank) if non_blank else 0.0
            profile = {
                "table": table_name,
                "column": column_name,
                "sample_count": len(values),
                "non_blank_count": len(non_blank),
                "blank_count": len(values) - len(non_blank),
                "parse_success_rate": round(parse_rate, 4),
                "formats": sorted(formats),
                "invalid_examples": invalid_examples,
            }
            scanned.append(profile)
            if name_suggests_date and (not non_blank or parse_rate < min_parse_success_rate):
                issues.append({"type": "dirty_text_date", **profile})
            elif len(formats) > 1:
                warnings.append({"type": "mixed_text_date_formats", **profile})

    return ok(
        f"Dirty date scan found {len(issues)} issue(s), {len(warnings)} warning(s).",
        table=table,
        max_samples=max_samples,
        min_parse_success_rate=min_parse_success_rate,
        valid=not issues,
        issue_count=len(issues),
        warning_count=len(warnings),
        issues=issues,
        warnings=warnings,
        scanned_columns=scanned,
    )


def pbi_validate_relationship_plan_tool(
    manager: Any,
    *,
    from_table: str,
    from_column: str,
    to_table: str,
    to_column: str,
    cardinality: str = "oneToMany",
    direction: str = "oneDirection",
    is_active: bool = True,
) -> dict[str, Any]:
    """Validate relationship cardinality, direction, duplicates, and ambiguity before creation."""
    from . import _model_snapshot

    snapshot = _model_snapshot(manager, include_hidden=False)
    issues: list[dict[str, Any]] = []
    warnings: list[dict[str, Any]] = []
    for table_name, column_name, side in (
        (from_table, from_column, "from"),
        (to_table, to_column, "to"),
    ):
        if _find_table(snapshot, table_name) is None:
            issues.append({"type": "table_not_found", "side": side, "table": table_name})
        elif _find_column(snapshot, table_name, column_name) is None:
            issues.append({"type": "column_not_found", "side": side, "table": table_name, "column": column_name})
    if issues:
        return ok(
            f"Relationship plan found {len(issues)} issue(s), {len(warnings)} warning(s).",
            valid=False,
            safe_to_create=False,
            issue_count=len(issues),
            warning_count=len(warnings),
            issues=issues,
            warnings=warnings,
        )

    normalized_direction = direction.casefold()
    normalized_cardinality = cardinality.casefold()
    if normalized_direction not in {"onedirection", "single"}:
        issues.append({"type": "unsafe_filter_direction", "direction": direction})
    if normalized_cardinality not in {"onetomany", "manytoone", "manytomany", "onetoone"}:
        issues.append({"type": "unknown_cardinality", "cardinality": cardinality})
    if normalized_cardinality == "manytomany":
        issues.append({"type": "many_to_many_relationship", "cardinality": cardinality})

    existing = snapshot.get("relationships", [])
    for rel in existing:
        endpoints_match = (
            str(rel.get("from_table", "")).casefold() == from_table.casefold()
            and str(rel.get("from_column", "")).casefold() == from_column.casefold()
            and str(rel.get("to_table", "")).casefold() == to_table.casefold()
            and str(rel.get("to_column", "")).casefold() == to_column.casefold()
        )
        if endpoints_match:
            issues.append({"type": "duplicate_relationship", "relationship": rel})

    from_profile = _column_profile(manager, from_table, from_column)
    to_profile = _column_profile(manager, to_table, to_column)
    if (from_profile.get("blank_count") or 0) > 0:
        warnings.append({"type": "from_column_has_blanks", "table": from_table, "column": from_column, **from_profile})
    if (to_profile.get("blank_count") or 0) > 0:
        warnings.append({"type": "to_column_has_blanks", "table": to_table, "column": to_column, **to_profile})

    from_unique = from_profile.get("row_count") == from_profile.get("distinct_count")
    to_unique = to_profile.get("row_count") == to_profile.get("distinct_count")
    if normalized_cardinality in {"onetomany", "manytoone"} and not (from_unique or to_unique):
        issues.append({"type": "no_unique_relationship_side", "from_profile": from_profile, "to_profile": to_profile})
    if normalized_cardinality == "onetoone" and not (from_unique and to_unique):
        issues.append(
            {"type": "one_to_one_requires_both_unique", "from_profile": from_profile, "to_profile": to_profile}
        )

    if is_active:
        graph: dict[str, set[str]] = {}
        for rel in existing:
            if not bool(rel.get("is_active", rel.get("active", True))):
                continue
            a = str(rel.get("from_table", ""))
            b = str(rel.get("to_table", ""))
            graph.setdefault(a, set()).add(b)
            graph.setdefault(b, set()).add(a)
        paths = _graph_paths(graph, from_table, to_table, limit=2)
        if paths:
            warnings.append({"type": "relationship_creates_parallel_path", "existing_paths": paths})
        graph.setdefault(from_table, set()).add(to_table)
        graph.setdefault(to_table, set()).add(from_table)
        if len(_graph_paths(graph, from_table, to_table, limit=3)) > 2:
            warnings.append({"type": "relationship_ambiguity_risk", "from_table": from_table, "to_table": to_table})

    return ok(
        f"Relationship plan found {len(issues)} issue(s), {len(warnings)} warning(s).",
        valid=not issues,
        safe_to_create=not issues,
        issue_count=len(issues),
        warning_count=len(warnings),
        issues=issues,
        warnings=warnings,
        profiles={"from": from_profile, "to": to_profile},
        proposed_relationship={
            "from_table": from_table,
            "from_column": from_column,
            "to_table": to_table,
            "to_column": to_column,
            "cardinality": cardinality,
            "direction": direction,
            "is_active": is_active,
        },
    )


def pbi_validate_star_schema_tool(
    manager: Any,
    *,
    include_hidden: bool = False,
    fact_table_hints: list[str] | None = None,
) -> dict[str, Any]:
    """Confirm a model follows star schema topology.

    Heuristic: a *fact* table participates as the many-side in ≥1 relationship,
    a *dimension* table participates as the one-side. Tables that are both
    are flagged as bridge tables. Direct dimension-to-dimension relationships
    are flagged as snowflake violations. Fact tables wired to other facts are
    flagged as fact-to-fact (constellation).

    Optional ``fact_table_hints`` lets callers force-tag tables as facts
    (matched case-insensitively); useful when a fact table has no incoming
    relationships yet because the model is being built incrementally.
    """
    from . import _model_snapshot

    snapshot = _model_snapshot(manager, include_hidden=include_hidden)
    relationships = snapshot.get("relationships", []) or []
    visible_tables = {
        str(item.get("name", "")): item
        for item in snapshot.get("tables", []) or []
        if include_hidden or not item.get("is_hidden")
    }

    one_side: set[str] = set()
    many_side: set[str] = set()
    for rel in relationships:
        from_table = str(rel.get("from_table", ""))
        to_table = str(rel.get("to_table", ""))
        if from_table:
            many_side.add(from_table)
        if to_table:
            one_side.add(to_table)

    hints = {h.casefold() for h in (fact_table_hints or [])}

    fact_tables: list[str] = []
    dim_tables: list[str] = []
    bridge_tables: list[str] = []
    isolated_tables: list[str] = []
    for name in visible_tables:
        is_many = name in many_side
        is_one = name in one_side
        forced_fact = name.casefold() in hints
        if forced_fact or (is_many and not is_one):
            fact_tables.append(name)
        elif is_one and not is_many:
            dim_tables.append(name)
        elif is_one and is_many:
            bridge_tables.append(name)
        else:
            isolated_tables.append(name)

    fact_set = {n.casefold() for n in fact_tables}
    non_fact = {n.casefold() for n in (dim_tables + bridge_tables)}

    issues: list[dict[str, Any]] = []
    warnings: list[dict[str, Any]] = []

    for rel in relationships:
        a = str(rel.get("from_table", "")).casefold()
        b = str(rel.get("to_table", "")).casefold()
        if a in non_fact and b in non_fact:
            issues.append({"type": "snowflake_dim_to_dim", "from": rel.get("from_table"), "to": rel.get("to_table")})
        elif a in fact_set and b in fact_set:
            issues.append({"type": "fact_to_fact", "from": rel.get("from_table"), "to": rel.get("to_table")})

    if not fact_tables:
        issues.append({"type": "no_fact_table_detected"})
    if len(fact_tables) > 1:
        warnings.append({"type": "multiple_fact_tables", "tables": fact_tables})
    if bridge_tables:
        warnings.append({"type": "bridge_tables", "tables": bridge_tables})
    if isolated_tables and len(visible_tables) > 1:
        warnings.append({"type": "isolated_tables", "tables": isolated_tables})

    return ok(
        f"Star-schema validation: {len(fact_tables)} fact, {len(dim_tables)} dim, "
        f"{len(issues)} issue(s), {len(warnings)} warning(s).",
        valid=not issues,
        is_star_schema=not issues and len(fact_tables) >= 1,
        fact_tables=fact_tables,
        dim_tables=dim_tables,
        bridge_tables=bridge_tables,
        isolated_tables=isolated_tables,
        relationship_count=len(relationships),
        issue_count=len(issues),
        warning_count=len(warnings),
        issues=issues,
        warnings=warnings,
    )


def pbi_detect_circular_dependencies_tool(
    manager: Any,
    *,
    include_hidden: bool = False,
) -> dict[str, Any]:
    """Detect cycles in the measure dependency graph.

    Builds a graph of measure → referenced measures (parsed from DAX
    ``[Name]`` tokens that match a known measure name) and runs a DFS to
    find strongly connected cycles. Self-references are reported separately.
    """
    import re

    from . import _model_snapshot

    snapshot = _model_snapshot(manager, include_hidden=include_hidden)
    measures = snapshot.get("measures", []) or []
    measure_names = {str(m.get("name", "")).casefold(): str(m.get("name", "")) for m in measures}
    expressions: dict[str, str] = {str(m.get("name", "")): str(m.get("expression", "") or "") for m in measures}

    ref_pattern = re.compile(r"(?<!')\[([^\[\]]+)\]")
    graph: dict[str, set[str]] = {name: set() for name in expressions}
    self_refs: list[str] = []
    for name, expr in expressions.items():
        for match in ref_pattern.findall(expr):
            target_key = match.casefold()
            if target_key in measure_names:
                target = measure_names[target_key]
                if target == name:
                    self_refs.append(name)
                else:
                    graph[name].add(target)

    cycles: list[list[str]] = []
    visited: set[str] = set()
    stack_set: set[str] = set()
    stack: list[str] = []

    def _dfs(node: str) -> None:
        if node in stack_set:
            idx = stack.index(node)
            cycle = stack[idx:] + [node]
            if cycle not in cycles:
                cycles.append(cycle)
            return
        if node in visited:
            return
        visited.add(node)
        stack.append(node)
        stack_set.add(node)
        for nxt in sorted(graph.get(node, set())):
            _dfs(nxt)
        stack.pop()
        stack_set.discard(node)

    for name in sorted(graph):
        _dfs(name)

    return ok(
        f"Circular dependency scan: {len(cycles)} cycle(s), {len(self_refs)} self-reference(s).",
        valid=not cycles and not self_refs,
        cycle_count=len(cycles),
        self_reference_count=len(self_refs),
        cycles=cycles,
        self_references=self_refs,
        measure_count=len(measures),
    )

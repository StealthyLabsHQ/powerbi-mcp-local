"""MCP wrappers — domain: quality."""

from __future__ import annotations

from typing import Any

from mcp_core import CONNECTION_MANAGER, _run, mcp
from tools import (
    pbi_audit_model_tool,
    pbi_compare_report_versions_tool,
    pbi_detect_circular_dependencies_tool,
    pbi_detect_dirty_dates_tool,
    pbi_detect_empty_visuals_tool,
    pbi_detect_missing_visuals_tool,
    pbi_detect_name_collisions_tool,
    pbi_export_correction_report_tool,
    pbi_export_validation_report_tool,
    pbi_generate_measure_tests_tool,
    pbi_lint_dax_tool,
    pbi_lint_report_layout_tool,
    pbi_run_scenario_tool,
    pbi_score_dashboard_tool,
    pbi_score_rubric_tool,
    pbi_validate_pbix_persistence_tool,
    pbi_validate_pbix_reopen_tool,
    pbi_validate_power_query_steps_tool,
    pbi_validate_relationship_plan_tool,
    pbi_validate_star_schema_tool,
    pbi_validate_visual_bindings_tool,
)


@mcp.tool()
def pbi_audit_model(include_hidden: bool = False) -> dict[str, Any]:
    """Audit relationships, ambiguous paths, bidirectional filters, and orphan structures."""
    return _run(
        "pbi_audit_model",
        pbi_audit_model_tool,
        CONNECTION_MANAGER,
        include_hidden=include_hidden,
    )


@mcp.tool()
def pbi_validate_star_schema(
    include_hidden: bool = False,
    fact_table_hints: list[str] | None = None,
) -> dict[str, Any]:
    """Verify the model is a star schema (1 fact table per group, dim tables only on the one-side)."""
    return _run(
        "pbi_validate_star_schema",
        pbi_validate_star_schema_tool,
        CONNECTION_MANAGER,
        include_hidden=include_hidden,
        fact_table_hints=fact_table_hints,
    )


@mcp.tool()
def pbi_detect_circular_dependencies(include_hidden: bool = False) -> dict[str, Any]:
    """Find cycles and self-references in the measure dependency graph."""
    return _run(
        "pbi_detect_circular_dependencies",
        pbi_detect_circular_dependencies_tool,
        CONNECTION_MANAGER,
        include_hidden=include_hidden,
    )


@mcp.tool()
def pbi_validate_power_query_steps(
    table: str,
    expected_steps: list[str],
    partition_name: str | None = None,
    case_sensitive: bool = False,
) -> dict[str, Any]:
    """Verify that a Power Query (M) expression contains expected step patterns (substrings or `re:` regex)."""
    return _run(
        "pbi_validate_power_query_steps",
        pbi_validate_power_query_steps_tool,
        CONNECTION_MANAGER,
        table=table,
        expected_steps=expected_steps,
        partition_name=partition_name,
        case_sensitive=case_sensitive,
    )


@mcp.tool()
def pbi_detect_missing_visuals(
    extract_folder: str,
    page: str,
    requirements: list[dict[str, Any]],
) -> dict[str, Any]:
    """Detect required visuals missing from a page (visual_type + optional contains_field/count)."""
    return _run(
        "pbi_detect_missing_visuals",
        pbi_detect_missing_visuals_tool,
        extract_folder,
        page=page,
        requirements=requirements,
    )


@mcp.tool()
def pbi_score_rubric(
    criteria: list[dict[str, Any]],
    extract_folder: str | None = None,
) -> dict[str, Any]:
    """Run a weighted rubric across multiple validators (star_schema, no_circular_deps, power_query_steps, missing_visuals, measure_exists)."""
    return _run(
        "pbi_score_rubric",
        pbi_score_rubric_tool,
        CONNECTION_MANAGER,
        criteria=criteria,
        extract_folder=extract_folder,
    )


@mcp.tool()
def pbi_export_correction_report(
    output_path: str,
    extract_folder: str | None = None,
    rubric_criteria: list[dict[str, Any]] | None = None,
    fact_table_hints: list[str] | None = None,
) -> dict[str, Any]:
    """Generate a Markdown correction report (model overview, star schema, cycles, audit, optional rubric)."""
    return _run(
        "pbi_export_correction_report",
        pbi_export_correction_report_tool,
        CONNECTION_MANAGER,
        output_path=output_path,
        extract_folder=extract_folder,
        rubric_criteria=rubric_criteria,
        fact_table_hints=fact_table_hints,
    )


@mcp.tool()
def pbi_lint_dax(include_hidden: bool = False, validate_expressions: bool = True) -> dict[str, Any]:
    """Lint measures for DAX validity, format strings, and column/measure name collisions."""
    return _run(
        "pbi_lint_dax",
        pbi_lint_dax_tool,
        CONNECTION_MANAGER,
        include_hidden=include_hidden,
        validate_expressions=validate_expressions,
    )


@mcp.tool()
def pbi_detect_name_collisions(include_hidden: bool = False) -> dict[str, Any]:
    """Detect table, column, and measure name collisions before writes."""
    return _run(
        "pbi_detect_name_collisions",
        pbi_detect_name_collisions_tool,
        CONNECTION_MANAGER,
        include_hidden=include_hidden,
    )


@mcp.tool()
def pbi_detect_dirty_dates(
    table: str | None = None,
    max_samples: int = 200,
    min_parse_success_rate: float = 0.8,
    scan_all_text_columns: bool = False,
) -> dict[str, Any]:
    """Detect text columns that look like dirty dates."""
    return _run(
        "pbi_detect_dirty_dates",
        pbi_detect_dirty_dates_tool,
        CONNECTION_MANAGER,
        table=table,
        max_samples=max_samples,
        min_parse_success_rate=min_parse_success_rate,
        scan_all_text_columns=scan_all_text_columns,
    )


@mcp.tool()
def pbi_validate_relationship_plan(
    from_table: str,
    from_column: str,
    to_table: str,
    to_column: str,
    cardinality: str = "oneToMany",
    direction: str = "oneDirection",
    is_active: bool = True,
) -> dict[str, Any]:
    """Validate relationship cardinality, direction, duplicates, and ambiguity before creation."""
    return _run(
        "pbi_validate_relationship_plan",
        pbi_validate_relationship_plan_tool,
        CONNECTION_MANAGER,
        from_table=from_table,
        from_column=from_column,
        to_table=to_table,
        to_column=to_column,
        cardinality=cardinality,
        direction=direction,
        is_active=is_active,
    )


@mcp.tool()
def pbi_detect_empty_visuals(
    extract_folder: str,
    page: str | None = None,
    include_slicers: bool = False,
    max_rows: int = 1,
    filter_expression: str | None = None,
) -> dict[str, Any]:
    """Execute lightweight DAX probes to detect visuals with no data."""
    return _run(
        "pbi_detect_empty_visuals",
        pbi_detect_empty_visuals_tool,
        CONNECTION_MANAGER,
        extract_folder=extract_folder,
        page=page,
        include_slicers=include_slicers,
        max_rows=max_rows,
        filter_expression=filter_expression,
    )


@mcp.tool()
def pbi_generate_measure_tests(
    measures: list[str] | None = None,
    include_hidden: bool = False,
    max_measures: int = 200,
) -> dict[str, Any]:
    """Generate and execute smoke tests for DAX measures."""
    return _run(
        "pbi_generate_measure_tests",
        pbi_generate_measure_tests_tool,
        CONNECTION_MANAGER,
        measures=measures,
        include_hidden=include_hidden,
        max_measures=max_measures,
    )


@mcp.tool()
def pbi_validate_pbix_persistence(
    pbix_path: str,
    extract_folder: str | None = None,
    require_security_bindings_removed: bool = True,
) -> dict[str, Any]:
    """Validate that a patched PBIX still has a readable, persistent report layout."""
    return _run(
        "pbi_validate_pbix_persistence",
        pbi_validate_pbix_persistence_tool,
        pbix_path=pbix_path,
        extract_folder=extract_folder,
        require_security_bindings_removed=require_security_bindings_removed,
    )


@mcp.tool()
def pbi_validate_pbix_reopen(
    pbix_path: str,
    timeout_seconds: int = 60,
    screenshot_path: str | None = None,
    close_after: bool = False,
    analyze_screenshot: bool = True,
    use_windows_ocr: bool = True,
) -> dict[str, Any]:
    """Open a PBIX in Power BI Desktop and scan for visible repair-error signals."""
    return _run(
        "pbi_validate_pbix_reopen",
        pbi_validate_pbix_reopen_tool,
        pbix_path=pbix_path,
        timeout_seconds=timeout_seconds,
        screenshot_path=screenshot_path,
        close_after=close_after,
        analyze_screenshot=analyze_screenshot,
        use_windows_ocr=use_windows_ocr,
    )


@mcp.tool()
def pbi_export_validation_report(
    output_path: str,
    extract_folder: str | None = None,
    include_hidden: bool = False,
    include_empty_visual_scan: bool = False,
    empty_visual_filter_expression: str | None = None,
    include_measure_tests: bool = False,
) -> dict[str, Any]:
    """Export model, DAX, layout, binding, and score validation as JSON."""
    return _run(
        "pbi_export_validation_report",
        pbi_export_validation_report_tool,
        CONNECTION_MANAGER,
        output_path=output_path,
        extract_folder=extract_folder,
        include_hidden=include_hidden,
        include_empty_visual_scan=include_empty_visual_scan,
        empty_visual_filter_expression=empty_visual_filter_expression,
        include_measure_tests=include_measure_tests,
    )


@mcp.tool()
def pbi_lint_report_layout(
    extract_folder: str,
    page: str | None = None,
    ignore_warnings: list[str] | None = None,
    only_pages: list[str] | None = None,
    max_visuals_per_page: int | None = None,
) -> dict[str, Any]:
    """Lint extracted Report/Layout for overlaps, whitespace, tiny visuals, and missing titles.

    ``ignore_warnings``: list of warning types to drop (e.g.
    ``["too_many_visuals", "visual_too_small", "missing_title",
    "excessive_whitespace"]``) — useful on intentionally dense pages.
    ``only_pages``: restrict the scan to specific page names. ``max_visuals_per_page``:
    override the default ``too_many_visuals`` threshold.
    """
    return _run(
        "pbi_lint_report_layout",
        pbi_lint_report_layout_tool,
        extract_folder=extract_folder,
        page=page,
        ignore_warnings=ignore_warnings,
        only_pages=only_pages,
        max_visuals_per_page=max_visuals_per_page,
    )


@mcp.tool()
def pbi_validate_visual_bindings(
    extract_folder: str, page: str | None = None, include_hidden: bool = False
) -> dict[str, Any]:
    """Validate every field referenced by extracted report visuals against the active model."""
    return _run(
        "pbi_validate_visual_bindings",
        pbi_validate_visual_bindings_tool,
        extract_folder=extract_folder,
        page=page,
        include_hidden=include_hidden,
        manager=CONNECTION_MANAGER,
    )


@mcp.tool()
def pbi_score_dashboard(extract_folder: str | None = None, include_hidden: bool = False) -> dict[str, Any]:
    """Score active dashboard quality: model, DAX, layout, and business readability."""
    return _run(
        "pbi_score_dashboard",
        pbi_score_dashboard_tool,
        CONNECTION_MANAGER,
        extract_folder=extract_folder,
        include_hidden=include_hidden,
    )


@mcp.tool()
def pbi_run_scenario(scenario: str, extract_folder: str | None = None, include_hidden: bool = False) -> dict[str, Any]:
    """Run a complete QA scenario against active model and optional extracted layout."""
    return _run(
        "pbi_run_scenario",
        pbi_run_scenario_tool,
        CONNECTION_MANAGER,
        scenario=scenario,
        extract_folder=extract_folder,
        include_hidden=include_hidden,
    )


@mcp.tool()
def pbi_compare_report_versions(
    extract_folder_a: str,
    extract_folder_b: str,
    label_a: str = "A",
    label_b: str = "B",
) -> dict[str, Any]:
    """Compare two extracted report versions by pages, visuals, and layout lint."""
    return _run(
        "pbi_compare_report_versions",
        pbi_compare_report_versions_tool,
        extract_folder_a=extract_folder_a,
        extract_folder_b=extract_folder_b,
        label_a=label_a,
        label_b=label_b,
    )

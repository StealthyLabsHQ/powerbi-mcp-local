"""Report page and visual automation tools using pbi-tools and Layout JSON."""

from __future__ import annotations

import json
import logging
import os
import shutil
import subprocess
import tempfile
import time
import uuid
import zipfile
from pathlib import Path
from typing import Any, Callable

from pbi_connection import PowerBIError, PowerBINotFoundError, PowerBIValidationError, error_payload, ok
from security import SECURITY, resolve_local_path
from ..model import pbi_model_info_tool

# Constants, errors, and design presets live in dedicated submodules so the
# rest of the package can import them without circular-import risk.
from ._base import (
    DEFAULT_PAGE_HEIGHT,
    DEFAULT_PAGE_WIDTH,
    DEFAULT_VISUAL_SIZES,
    DESIGN_THEME_RELATIVE_PATH,
    HEX_COLOR_RE,
    LAYOUT_RELATIVE_PATH,
    MODEL_TABLES_RELATIVE_DIR,
    PBIToolsNotInstalledError,
    PageNotFoundError,
    ReportLayoutError,
    THEMES_RELATIVE_DIR,
    VISUAL_FIELD_ROLES,
    VISUAL_ROLE_KINDS,
    VisualNotFoundError,
    VisualToolError,
)
from ._layout import (
    _LAYOUT_WRITE_TL,
    _dump_embedded_json,
    _find_page,
    _is_dry_run,
    _load_layout,
    _next_page_name,
    _normalize_page_name,
    _page_summary,
    _parse_embedded_json,
    _record_dry_run_write,
    _save_layout,
    dry_run_layout_writes,
)
from ._paths import (
    _layout_path,
    _resolve_extract_folder,
    _resolve_pbix_path,
    _resolve_theme_path,
)
from ._refs import (
    _BRACKET_REF_RE,
    _normalize_reference,
    _query_ref,
    _split_column_ref,
)

logger = logging.getLogger(__name__)


def _run(callback: Callable[..., dict[str, Any]], *args: Any, **kwargs: Any) -> dict[str, Any]:
    try:
        return callback(*args, **kwargs)
    except Exception as exc:
        return error_payload(exc)


from ._pages import (
    pbi_create_page_tool,
    pbi_delete_page_tool,
    pbi_describe_page_tool,
    pbi_get_page_tool,
    pbi_list_pages_tool,
    pbi_set_page_size_tool,
)


from ._design import (
    DESIGN_PRESETS,
    pbi_apply_design_tool,
    pbi_apply_theme_tool,
    pbi_build_dashboard_tool,
)
from ._repair import (
    pbi_repair_report_fields_tool,
    pbi_validate_report_fields_tool,
)


from ._io import (
    _extract_pbix_zip_natively,
    _find_pbi_tools,
    _force_kill_powerbi,
    _maybe_force_close_powerbi,
    _page_names_from_layout_bytes,
    _run_pbi_tools,
    _run_powershell,
    _save_and_close_powerbi_gracefully,
    pbi_compile_report_tool,
    pbi_extract_report_tool,
)


from ._containers import (
    _append_visual,
    _base_visual_config,
    _create_chart_container,
    _find_visual,
    _make_visual_container,
    _page_next_z,
    _unique_visual_id,
    _validate_dimensions,
    _visual_payload,
)


from ._home_tables import (
    _augment_measure_home_map_with_live,
    _inspect_value_measures,
    _persistence_risks,
    _resolve_measure_home_map,
    _scan_measure_home_tables,
)


from ._bindings import (
    _assert_container_bindings,
    _build_prototype_query,
    _build_select_entry,
    _from_entity_by_alias,
    _live_model_field_index,
    _next_alias,
    _scan_visual_bindings,
    _select_name_map,
    _sync_container_query,
    _validate_field_references_live,
    _validate_projection_roles,
    _visual_binding_issues,
)


from ._formatting import (
    _VISUAL_FORMAT_TYPES,
    _datapoint_fill_objects,
    _decimal_literal,
    _encode_visual_format_value,
    _gauge_axis_objects,
    _int_literal,
    _literal_value,
    _solid_color,
    _text_literal,
    _title_objects,
)


def pbi_patch_layout_tool(
    extract_folder: str,
    pbix_path: str,
    force: bool = False,
    fail_on_persistence_risk: bool = True,
    manager: Any | None = None,
    include_hidden: bool = False,
) -> dict[str, Any]:
    def _impl() -> dict[str, Any]:
        folder = _resolve_extract_folder(extract_folder, must_exist=True)
        pbix = _resolve_pbix_path(pbix_path, must_exist=True)
        layout_path = _layout_path(folder)
        if not layout_path.exists():
            raise ReportLayoutError("Report/Layout file was not found in the extract folder.", details={"path": str(layout_path)})

        if fail_on_persistence_risk:
            _, layout = _load_layout(folder)
            measure_home_map = _scan_measure_home_tables(folder)
            model_fields, model_validation = _live_model_field_index(manager, include_hidden=include_hidden)
            issues, _ = _scan_visual_bindings(layout, measure_home_map, model_fields, repair=False)
            persistence_risks = _persistence_risks(issues)
            if persistence_risks:
                raise PowerBIValidationError(
                    "Layout patch blocked because field bindings rely on live-model metadata missing from the extract.",
                    details={
                        "persistence_risk_count": len(persistence_risks),
                        "persistence_risks": persistence_risks,
                        "model_validation": model_validation,
                    },
                )

        _maybe_force_close_powerbi(force, pbix)

        layout_bytes = layout_path.read_bytes()
        pages = _page_names_from_layout_bytes(layout_bytes)

        temp_path: Path | None = None
        try:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pbix", dir=str(pbix.parent)) as tmp_file:
                temp_path = Path(tmp_file.name)
            with zipfile.ZipFile(pbix, "r") as source_zip, zipfile.ZipFile(temp_path, "w") as target_zip:
                layout_written = False
                for info in source_zip.infolist():
                    name = info.filename
                    if name == "SecurityBindings":
                        continue
                    payload = layout_bytes if name == "Report/Layout" else source_zip.read(name)
                    if name == "Report/Layout":
                        layout_written = True
                    target_info = zipfile.ZipInfo(name, date_time=info.date_time)
                    target_info.compress_type = info.compress_type
                    target_info.comment = info.comment
                    target_info.extra = info.extra
                    target_info.internal_attr = info.internal_attr
                    target_info.external_attr = info.external_attr
                    target_info.create_system = info.create_system
                    target_info.create_version = info.create_version
                    target_info.extract_version = info.extract_version
                    target_info.volume = info.volume
                    target_info.flag_bits = info.flag_bits
                    target_zip.writestr(target_info, payload)
                if not layout_written:
                    target_info = zipfile.ZipInfo("Report/Layout")
                    target_info.compress_type = zipfile.ZIP_DEFLATED
                    target_zip.writestr(target_info, layout_bytes)

            temp_size = temp_path.stat().st_size
            try:
                temp_path.replace(pbix)
            except PermissionError as exc:
                raise ReportLayoutError(
                    "PBIX file is locked by Power BI Desktop. Close it or retry with force=True.",
                    details={"pbix_path": str(pbix), "force": force},
                ) from exc
        finally:
            if temp_path and temp_path.exists():
                temp_path.unlink(missing_ok=True)

        return ok(
            "Layout patched into PBIX successfully.",
            extract_folder=str(folder),
            pbix_path=str(pbix),
            bytes_written=temp_size,
            layout_size=len(layout_bytes),
            pages=pages,
            persistence_risk_checked=fail_on_persistence_risk,
        )

    return _run(_impl)


def pbi_convert_visual_type_tool(
    extract_folder: str,
    page: str,
    visual_id: str,
    new_type: str,
) -> dict[str, Any]:
    """Migrate an existing visual to a different type while preserving
    compatible field bindings.

    Compatibility groups (bindings preserved as-is within a group):
    - ``card`` ↔ ``kpi`` (Values ↔ Indicator)
    - ``clusteredBarChart`` ↔ ``clusteredColumnChart`` ↔ ``lineChart`` ↔
      ``lineClusteredColumnComboChart`` (Category/Y/Series identical)
    - ``donutChart`` ↔ ``treemap``

    Raises ``PowerBIValidationError`` with a clear ``reason`` when the source
    and target are incompatible. Use this so an LLM can recover from "I picked
    the wrong visual type" without losing bindings.
    """
    new_type_clean = str(new_type).strip()
    if not new_type_clean:
        raise PowerBIValidationError("new_type must be non-empty.")
    if new_type_clean not in VISUAL_FIELD_ROLES:
        raise PowerBIValidationError(
            f"Unknown target visual type '{new_type_clean}'.",
            details={"new_type": new_type_clean, "known_types": sorted(VISUAL_FIELD_ROLES)},
        )

    # Source-target role rewrites. Each entry maps source role → target role.
    # If a source role isn't covered, the conversion is rejected (no silent loss).
    COMPATIBILITY: dict[tuple[str, str], dict[str, str]] = {
        ("card", "kpi"): {"Values": "Indicator"},
        ("kpi", "card"): {"Indicator": "Values"},
        ("donutChart", "treemap"): {"Category": "Category", "Y": "Y"},
        ("treemap", "donutChart"): {"Category": "Category", "Y": "Y"},
    }
    # Charts that share Category/Y/Series — identity rewrites are filled
    # automatically.
    chart_family = {"clusteredBarChart", "clusteredColumnChart", "lineChart", "lineClusteredColumnComboChart"}
    for src in chart_family:
        for tgt in chart_family:
            if src != tgt:
                COMPATIBILITY.setdefault((src, tgt), {"Category": "Category", "Y": "Y", "Series": "Series"})

    def _impl() -> dict[str, Any]:
        folder, layout = _load_layout(extract_folder)
        section = _find_page(layout, page)
        _, container, config = _find_visual(section, visual_id)
        sv = config.setdefault("singleVisual", {})
        old_type = str(sv.get("visualType") or "")
        if old_type == new_type_clean:
            return ok(
                f"Visual '{visual_id}' already has type '{new_type_clean}'; nothing to do.",
                visual_id=visual_id,
                old_type=old_type,
                new_type=new_type_clean,
                changed=False,
            )
        rewrite = COMPATIBILITY.get((old_type, new_type_clean))
        if rewrite is None:
            raise PowerBIValidationError(
                f"No compatible role rewrite from '{old_type}' to '{new_type_clean}'.",
                details={
                    "old_type": old_type,
                    "new_type": new_type_clean,
                    "supported_targets_from_old": sorted({tgt for (src, tgt) in COMPATIBILITY if src == old_type}),
                    "reason": "incompatible",
                },
            )

        old_projections = sv.get("projections", {}) or {}
        new_projections: dict[str, list[dict[str, str]]] = {}
        unmapped_roles: list[str] = []
        for role, items in old_projections.items():
            target_role = rewrite.get(role)
            if target_role is None:
                unmapped_roles.append(role)
                continue
            new_projections[target_role] = items

        if unmapped_roles:
            raise PowerBIValidationError(
                f"Source role(s) {unmapped_roles} not mapped by '{old_type}' → '{new_type_clean}'. Aborting to avoid silent data loss.",
                details={"unmapped_roles": unmapped_roles, "rewrite": rewrite},
            )

        # Validate the new projections against the new visual type's role schema.
        _validate_projection_roles(new_type_clean, new_projections)

        sv["visualType"] = new_type_clean
        sv["projections"] = new_projections
        # Refresh the embedded query so the visualType change is reflected in the query payload too.
        prototype_query = sv.get("prototypeQuery") or {"Version": 2, "From": [], "Select": []}
        sv["prototypeQuery"] = prototype_query
        container["config"] = _dump_embedded_json(config)
        query_payload = {"Commands": [{"SemanticQueryDataShapeCommand": {"Query": prototype_query}}]}
        container["query"] = _dump_embedded_json(query_payload)
        _save_layout(folder, layout)
        return ok(
            f"Visual '{visual_id}' converted: {old_type} → {new_type_clean}.",
            visual_id=visual_id,
            old_type=old_type,
            new_type=new_type_clean,
            role_rewrites=rewrite,
            changed=True,
        )

    return _run(_impl)


def pbi_auto_grid_layout_tool(
    specs: list[dict[str, Any]],
    *,
    cols: int = 4,
    gap: int = 16,
    start_x: int = 20,
    start_y: int = 60,
    cell_width: int | None = None,
    cell_height: int | None = None,
    page_width: int = DEFAULT_PAGE_WIDTH,
    page_height: int = DEFAULT_PAGE_HEIGHT,
) -> dict[str, Any]:
    """Compute non-overlapping (x, y, width, height) for a list of visual
    specs on a column-based grid.

    No layout writes happen here — the function is pure and offline. Each
    input ``spec`` is returned annotated with ``x``, ``y``, ``width``,
    ``height``. ``cell_width`` defaults to ``(page_width - 2*start_x - gap*(cols-1)) / cols``
    so the grid fits the page width. ``cell_height`` defaults to 200.
    Specs may set ``col_span`` / ``row_span`` to grow over neighbours; the
    walker advances per cell so spans never overlap.
    """
    if not isinstance(specs, list) or not specs:
        raise PowerBIValidationError("specs must be a non-empty list of visual configs.")
    if cols < 1:
        raise PowerBIValidationError("cols must be >= 1.", details={"cols": cols})
    if gap < 0:
        raise PowerBIValidationError("gap must be >= 0.", details={"gap": gap})

    # Derive cell sizes.
    usable_width = max(0, page_width - 2 * start_x - gap * max(0, cols - 1))
    cw = int(cell_width if cell_width is not None else (usable_width // cols if cols else usable_width))
    if cw <= 0:
        raise PowerBIValidationError(
            "Computed cell_width is non-positive; reduce cols or start_x or pass cell_width explicitly.",
            details={"page_width": page_width, "start_x": start_x, "gap": gap, "cols": cols},
        )
    ch = int(cell_height) if cell_height is not None else 200

    # Track which cells are occupied so spans don't overlap.
    occupied: set[tuple[int, int]] = set()
    placed: list[dict[str, Any]] = []
    cursor_row = 0
    cursor_col = 0

    def _next_free_cell(row: int, col: int) -> tuple[int, int]:
        while (row, col) in occupied:
            col += 1
            if col >= cols:
                col = 0
                row += 1
        return row, col

    for index, spec in enumerate(specs):
        if not isinstance(spec, dict):
            raise PowerBIValidationError(
                f"specs[{index}] must be a dict, got {type(spec).__name__}.",
                details={"index": index},
            )
        col_span = max(1, int(spec.get("col_span", 1)))
        if col_span > cols:
            col_span = cols
        row_span = max(1, int(spec.get("row_span", 1)))

        # Find a row/col where the whole span fits.
        cursor_row, cursor_col = _next_free_cell(cursor_row, cursor_col)
        while cursor_col + col_span > cols or any(
            (cursor_row + r, cursor_col + c) in occupied
            for r in range(row_span)
            for c in range(col_span)
        ):
            cursor_col += 1
            if cursor_col + col_span > cols:
                cursor_col = 0
                cursor_row += 1
            cursor_row, cursor_col = _next_free_cell(cursor_row, cursor_col)

        for r in range(row_span):
            for c in range(col_span):
                occupied.add((cursor_row + r, cursor_col + c))

        x = start_x + cursor_col * (cw + gap)
        y = start_y + cursor_row * (ch + gap)
        width = cw * col_span + gap * (col_span - 1)
        height = ch * row_span + gap * (row_span - 1)
        placed_spec = dict(spec)
        placed_spec.update({"x": x, "y": y, "width": width, "height": height})
        placed.append(placed_spec)

        cursor_col += col_span
        if cursor_col >= cols:
            cursor_col = 0
            cursor_row += 1

    total_height_used = start_y + (cursor_row + 1) * (ch + gap)
    return ok(
        f"Auto-grid: positioned {len(placed)} spec(s) on a {cols}-column grid.",
        cols=cols,
        gap=gap,
        cell_width=cw,
        cell_height=ch,
        start_x=start_x,
        start_y=start_y,
        page_width=page_width,
        page_height=page_height,
        used_height_estimate=total_height_used,
        specs=placed,
    )


def pbi_add_card_tool(
    extract_folder: str,
    page: str,
    measure: str,
    x: int,
    y: int,
    width: int = 200,
    height: int = 120,
    title: str = "",
    *,
    manager: Any | None = None,
) -> dict[str, Any]:
    """Add a card visual.

    If a connection ``manager`` is supplied, the measure name is checked
    against the live model first and the call fails fast on a typo. Without
    a manager the tool still works (offline mode), preserving prior behavior.
    """
    _validate_field_references_live(manager, [measure], expected_kinds={measure: "measure"})
    measure_home_map = _resolve_measure_home_map(extract_folder, manager=manager)
    return _append_visual(
        extract_folder,
        page,
        lambda section, home_map: _create_chart_container(
            section,
            visual_type="card",
            x=x,
            y=y,
            width=width,
            height=height,
            title=title or None,
            projections={"Values": [{"queryRef": _query_ref(measure)}]},
            references=[measure],
            measure_home_map=home_map,
        ),
        measure_home_map,
    )


def pbi_add_bar_chart_tool(
    extract_folder: str,
    page: str,
    category_column: str,
    value_measure: str,
    x: int,
    y: int,
    width: int = 400,
    height: int = 300,
    title: str = "",
    legend_column: str | None = None,
    *,
    manager: Any | None = None,
) -> dict[str, Any]:
    measure_home_map = _resolve_measure_home_map(extract_folder, manager=manager)
    projections = {"Category": [{"queryRef": _query_ref(category_column)}], "Y": [{"queryRef": _query_ref(value_measure)}]}
    references = [category_column, value_measure]
    expected_kinds = {category_column: "column", value_measure: "measure"}
    if legend_column:
        projections["Series"] = [{"queryRef": _query_ref(legend_column)}]
        references.append(legend_column)
        expected_kinds[legend_column] = "column"
    _validate_field_references_live(manager, references, expected_kinds=expected_kinds)
    return _append_visual(
        extract_folder,
        page,
        lambda section, home_map: _create_chart_container(
            section,
            visual_type="clusteredBarChart",
            x=x,
            y=y,
            width=width,
            height=height,
            title=title or None,
            projections=projections,
            references=references,
            measure_home_map=home_map,
        ),
        measure_home_map,
    )


def pbi_add_line_chart_tool(
    extract_folder: str,
    page: str,
    axis_column: str,
    value_measures: list[str],
    x: int,
    y: int,
    width: int = 420,
    height: int = 300,
    title: str = "",
    *,
    manager: Any | None = None,
) -> dict[str, Any]:
    if not value_measures:
        raise PowerBIValidationError("value_measures must contain at least one measure.")
    expected_kinds = {axis_column: "column"}
    for m in value_measures:
        expected_kinds[m] = "measure"
    _validate_field_references_live(manager, [axis_column, *value_measures], expected_kinds=expected_kinds)
    measure_home_map = _resolve_measure_home_map(extract_folder, manager=manager)
    diagnostics = _inspect_value_measures(value_measures, measure_home_map, manager)
    result = _append_visual(
        extract_folder,
        page,
        lambda section, home_map: _create_chart_container(
            section,
            visual_type="lineChart",
            x=x,
            y=y,
            width=width,
            height=height,
            title=title or None,
            projections={
                "Category": [{"queryRef": _query_ref(axis_column)}],
                "Y": [{"queryRef": _query_ref(measure)} for measure in value_measures],
            },
            references=[axis_column, *value_measures],
            measure_home_map=home_map,
        ),
        measure_home_map,
    )
    if diagnostics:
        result["warnings"] = diagnostics
    return result


def pbi_add_donut_chart_tool(extract_folder: str, page: str, category_column: str, value_measure: str, x: int, y: int, width: int = 320, height: int = 280, title: str = "", *, manager: Any | None = None) -> dict[str, Any]:
    _validate_field_references_live(manager, [category_column, value_measure], expected_kinds={category_column: "column", value_measure: "measure"})
    measure_home_map = _resolve_measure_home_map(extract_folder, manager=manager)
    return _append_visual(
        extract_folder,
        page,
        lambda section, home_map: _create_chart_container(
            section,
            visual_type="donutChart",
            x=x,
            y=y,
            width=width,
            height=height,
            title=title or None,
            projections={"Category": [{"queryRef": _query_ref(category_column)}], "Y": [{"queryRef": _query_ref(value_measure)}]},
            references=[category_column, value_measure],
            measure_home_map=home_map,
        ),
        measure_home_map,
    )


def pbi_add_table_visual_tool(extract_folder: str, page: str, columns: list[str], x: int, y: int, width: int = 520, height: int = 320, title: str = "", *, manager: Any | None = None) -> dict[str, Any]:
    if not columns:
        raise PowerBIValidationError("columns must contain at least one field or measure.")
    _validate_field_references_live(manager, list(columns))
    measure_home_map = _resolve_measure_home_map(extract_folder, manager=manager)
    return _append_visual(
        extract_folder,
        page,
        lambda section, home_map: _create_chart_container(
            section,
            visual_type="tableEx",
            x=x,
            y=y,
            width=width,
            height=height,
            title=title or None,
            projections={"Values": [{"queryRef": _query_ref(item)} for item in columns]},
            references=list(columns),
            measure_home_map=home_map,
        ),
        measure_home_map,
    )


def pbi_add_waterfall_tool(extract_folder: str, page: str, category_column: str, value_measure: str, x: int, y: int, width: int = 420, height: int = 300, title: str = "", *, manager: Any | None = None) -> dict[str, Any]:
    _validate_field_references_live(manager, [category_column, value_measure], expected_kinds={category_column: "column", value_measure: "measure"})
    measure_home_map = _resolve_measure_home_map(extract_folder, manager=manager)
    return _append_visual(
        extract_folder,
        page,
        lambda section, home_map: _create_chart_container(
            section,
            visual_type="waterfallChart",
            x=x,
            y=y,
            width=width,
            height=height,
            title=title or None,
            projections={"Category": [{"queryRef": _query_ref(category_column)}], "Y": [{"queryRef": _query_ref(value_measure)}]},
            references=[category_column, value_measure],
            measure_home_map=home_map,
        ),
        measure_home_map,
    )


def pbi_add_slicer_tool(extract_folder: str, page: str, column: str, x: int, y: int, width: int = 220, height: int = 120, slicer_type: str = "dropdown", *, manager: Any | None = None) -> dict[str, Any]:
    slicer_kind = slicer_type.strip().casefold()
    if slicer_kind not in {"dropdown", "list", "range", "tile"}:
        raise PowerBIValidationError("slicer_type must be one of: dropdown, list, range, tile.", details={"slicer_type": slicer_type})
    _validate_field_references_live(manager, [column], expected_kinds={column: "column"})
    measure_home_map = _resolve_measure_home_map(extract_folder, manager=manager)
    if slicer_kind == "tile":
        # Horizontal tile slicer: native list type with horizontal orientation flag.
        emitted_kind = "list"
        extra: dict[str, Any] = {
            "slicerType": emitted_kind,
            "objects": {
                "general": [
                    {"properties": {"orientation": _int_literal(1)}}
                ]
            },
        }
    else:
        extra = {"slicerType": slicer_kind}
    return _append_visual(
        extract_folder,
        page,
        lambda section, home_map: _create_chart_container(
            section,
            visual_type="slicer",
            x=x,
            y=y,
            width=width,
            height=height,
            title=None,
            projections={"Values": [{"queryRef": _query_ref(column)}]},
            references=[column],
            measure_home_map=home_map,
            extra_single_visual=extra,
        ),
        measure_home_map,
    )


def pbi_add_gauge_tool(
    extract_folder: str,
    page: str,
    measure: str,
    x: int,
    y: int,
    width: int = 280,
    height: int = 220,
    title: str = "",
    target_measure: str | None = None,
    *,
    min_value: float | None = None,
    max_value: float | None = None,
    target_value: float | None = None,
    fill_color: str | None = None,
    target_color: str | None = None,
    fill_color_measure: str | None = None,
    manager: Any | None = None,
) -> dict[str, Any]:
    """Add a gauge visual.

    ``fill_color_measure`` (optional): name of a DAX measure that returns a
    ``"#RRGGBB"`` string. The gauge arc fill becomes a measure-binding to it
    (conditional formatting), so the colour reacts to slicer / filter context.
    Mutually exclusive with ``fill_color``; if both are provided, the measure
    binding wins.

    If a connection ``manager`` is supplied, every measure / column reference
    (Y, target, fill_color_measure) is verified against the live model so a
    typo fails fast instead of producing a broken visual.
    """
    if fill_color and fill_color_measure:
        # Measure binding takes precedence — drop the static fill silently to keep callers tidy.
        fill_color = None
    refs_to_validate: list[str] = [measure]
    expected_kinds = {measure: "measure"}
    if target_measure:
        refs_to_validate.append(target_measure)
        expected_kinds[target_measure] = "measure"
    if fill_color_measure:
        refs_to_validate.append(fill_color_measure)
        expected_kinds[fill_color_measure] = "measure"
    _validate_field_references_live(manager, refs_to_validate, expected_kinds=expected_kinds)
    measure_home_map = _resolve_measure_home_map(extract_folder, manager=manager)
    projections = {"Y": [{"queryRef": _query_ref(measure)}]}
    references = [measure]
    if target_measure:
        projections["Goal"] = [{"queryRef": _query_ref(target_measure)}]
        references.append(target_measure)
    if fill_color_measure:
        # Pull the color measure into the visual's prototypeQuery so PBI can resolve the fill binding.
        references.append(fill_color_measure)
    extra_objects: dict[str, Any] = {}
    axis_obj = _gauge_axis_objects(min_value, max_value, target_value)
    if axis_obj:
        extra_objects["axis"] = axis_obj
    fill_obj = _datapoint_fill_objects(fill_color, target_color)
    if fill_color_measure:
        # Conditional fill via measure binding — overrides any static fill_color above.
        host_table = measure_home_map.get(fill_color_measure) or "$Measures"
        properties: dict[str, Any] = {}
        if fill_obj and fill_obj[0].get("properties"):
            properties = dict(fill_obj[0]["properties"])
        properties["fill"] = {
            "solid": {
                "color": {
                    "expr": {
                        "Measure": {
                            "Expression": {"SourceRef": {"Entity": host_table}},
                            "Property": fill_color_measure,
                        }
                    }
                }
            }
        }
        fill_obj = [{"properties": properties}]
    if fill_obj:
        extra_objects["dataPoint"] = fill_obj
    extra_single_visual = {"objects": extra_objects} if extra_objects else None
    return _append_visual(
        extract_folder,
        page,
        lambda section, home_map: _create_chart_container(
            section,
            visual_type="gauge",
            x=x,
            y=y,
            width=width,
            height=height,
            title=title or None,
            projections=projections,
            references=references,
            measure_home_map=home_map,
            extra_single_visual=extra_single_visual,
        ),
        measure_home_map,
    )


def pbi_add_labelled_card_tool(
    extract_folder: str,
    page: str,
    measure: str,
    label: str,
    x: int,
    y: int,
    width: int = 220,
    height: int = 110,
    *,
    label_height: int = 28,
    label_font_size: int = 11,
    label_bold: bool = True,
    label_color: str = "#1F2937",
    manager: Any | None = None,
) -> dict[str, Any]:
    """Place a text label above a card value, matching docx-style 'label-on-top' card layout.

    Returns both created visuals so callers can move/style them together later.
    If a connection ``manager`` is supplied, the measure name is verified
    against the live model before either visual is created.
    """
    if label_height <= 0 or label_height >= height:
        raise PowerBIValidationError(
            "label_height must be > 0 and smaller than height.",
            details={"label_height": label_height, "height": height},
        )
    _validate_field_references_live(manager, [measure], expected_kinds={measure: "measure"})
    label_response = pbi_add_text_box_tool(
        extract_folder,
        page,
        label,
        x,
        y,
        width,
        label_height,
        font_size=label_font_size,
        bold=label_bold,
        color=label_color,
    )
    if not label_response.get("ok"):
        return label_response
    card_response = pbi_add_card_tool(
        extract_folder,
        page,
        measure,
        x,
        y + label_height,
        width,
        height - label_height,
        title="",
    )
    if not card_response.get("ok"):
        return card_response
    return ok(
        f"Labelled card '{label}' added.",
        page=card_response.get("page"),
        visuals={
            "label": label_response.get("visual"),
            "value": card_response.get("visual"),
        },
    )


def pbi_add_scatter_chart_tool(
    extract_folder: str,
    page: str,
    category_column: str,
    x_measure: str,
    y_measure: str,
    x: int,
    y: int,
    width: int = 420,
    height: int = 320,
    title: str = "",
    size_measure: str | None = None,
    legend_column: str | None = None,
    *,
    manager: Any | None = None,
) -> dict[str, Any]:
    """Add a scatter chart visual (``scatterChart``).

    Roles: ``Category`` (column — dot identity), ``X`` (measure), ``Y``
    (measure), ``Size`` (measure, optional), ``Series`` (column, optional).
    Use for correlation analysis between two measures grouped by a dimension.
    """
    projections = {
        "Category": [{"queryRef": _query_ref(category_column)}],
        "X": [{"queryRef": _query_ref(x_measure)}],
        "Y": [{"queryRef": _query_ref(y_measure)}],
    }
    references = [category_column, x_measure, y_measure]
    expected_kinds = {category_column: "column", x_measure: "measure", y_measure: "measure"}
    if size_measure:
        projections["Size"] = [{"queryRef": _query_ref(size_measure)}]
        references.append(size_measure)
        expected_kinds[size_measure] = "measure"
    if legend_column:
        projections["Series"] = [{"queryRef": _query_ref(legend_column)}]
        references.append(legend_column)
        expected_kinds[legend_column] = "column"
    _validate_field_references_live(manager, references, expected_kinds=expected_kinds)
    measure_home_map = _resolve_measure_home_map(extract_folder, manager=manager)
    return _append_visual(
        extract_folder,
        page,
        lambda section, home_map: _create_chart_container(
            section,
            visual_type="scatterChart",
            x=x,
            y=y,
            width=width,
            height=height,
            title=title or None,
            projections=projections,
            references=references,
            measure_home_map=home_map,
            manager=manager,
        ),
        measure_home_map,
    )


def pbi_add_combo_chart_tool(
    extract_folder: str,
    page: str,
    category_column: str,
    bar_measures: list[str],
    line_measures: list[str],
    x: int,
    y: int,
    width: int = 480,
    height: int = 320,
    title: str = "",
    legend_column: str | None = None,
    *,
    manager: Any | None = None,
) -> dict[str, Any]:
    """Add a combo chart (column bars + line overlay).

    Roles: ``Category`` (column), ``Y`` (measure list — bars), ``Y2``
    (measure list — line). Useful for "actual vs target" with a target line
    over actual bars.
    """
    if not bar_measures:
        raise PowerBIValidationError("bar_measures must contain at least one measure.")
    if not line_measures:
        raise PowerBIValidationError("line_measures must contain at least one measure.")
    projections: dict[str, list[dict[str, str]]] = {
        "Category": [{"queryRef": _query_ref(category_column)}],
        "Y": [{"queryRef": _query_ref(item)} for item in bar_measures],
        "Y2": [{"queryRef": _query_ref(item)} for item in line_measures],
    }
    references = [category_column, *bar_measures, *line_measures]
    expected_kinds = {category_column: "column"}
    for m in bar_measures:
        expected_kinds[m] = "measure"
    for m in line_measures:
        expected_kinds[m] = "measure"
    if legend_column:
        projections["Series"] = [{"queryRef": _query_ref(legend_column)}]
        references.append(legend_column)
        expected_kinds[legend_column] = "column"
    _validate_field_references_live(manager, references, expected_kinds=expected_kinds)
    measure_home_map = _resolve_measure_home_map(extract_folder, manager=manager)
    return _append_visual(
        extract_folder,
        page,
        lambda section, home_map: _create_chart_container(
            section,
            visual_type="lineClusteredColumnComboChart",
            x=x,
            y=y,
            width=width,
            height=height,
            title=title or None,
            projections=projections,
            references=references,
            measure_home_map=home_map,
            manager=manager,
        ),
        measure_home_map,
    )


def pbi_add_kpi_tool(
    extract_folder: str,
    page: str,
    indicator_measure: str,
    trend_axis_column: str,
    x: int,
    y: int,
    width: int = 240,
    height: int = 160,
    title: str = "",
    goal_measure: str | None = None,
    direction: str = "high_is_good",
    *,
    manager: Any | None = None,
) -> dict[str, Any]:
    """Add a native KPI visual (``kpi``).

    Roles: ``Indicator`` (measure — current value), ``TrendLine`` (column,
    typically a Date), ``Goal`` (measure, optional). ``direction`` controls
    the status colour interpretation: ``"high_is_good"`` means the green
    threshold sits above ``Goal``; ``"low_is_good"`` flips it.
    """
    if direction not in {"high_is_good", "low_is_good"}:
        raise PowerBIValidationError(
            "direction must be 'high_is_good' or 'low_is_good'.",
            details={"direction": direction},
        )
    projections: dict[str, list[dict[str, str]]] = {
        "Indicator": [{"queryRef": _query_ref(indicator_measure)}],
        "TrendLine": [{"queryRef": _query_ref(trend_axis_column)}],
    }
    references = [indicator_measure, trend_axis_column]
    expected_kinds = {indicator_measure: "measure", trend_axis_column: "column"}
    if goal_measure:
        projections["Goal"] = [{"queryRef": _query_ref(goal_measure)}]
        references.append(goal_measure)
        expected_kinds[goal_measure] = "measure"
    _validate_field_references_live(manager, references, expected_kinds=expected_kinds)
    # Encode direction in the visual's objects so PBI's KPI rendering picks the right colour rule.
    extra_single_visual = {
        "objects": {
            "indicator": [
                {
                    "properties": {
                        "directionType": _text_literal(
                            "Increasing" if direction == "high_is_good" else "Decreasing"
                        )
                    }
                }
            ]
        }
    }
    measure_home_map = _resolve_measure_home_map(extract_folder, manager=manager)
    return _append_visual(
        extract_folder,
        page,
        lambda section, home_map: _create_chart_container(
            section,
            visual_type="kpi",
            x=x,
            y=y,
            width=width,
            height=height,
            title=title or None,
            projections=projections,
            references=references,
            measure_home_map=home_map,
            extra_single_visual=extra_single_visual,
            manager=manager,
        ),
        measure_home_map,
    )


def pbi_add_matrix_tool(
    extract_folder: str,
    page: str,
    rows: list[str],
    values: list[str],
    x: int,
    y: int,
    columns: list[str] | None = None,
    width: int = 540,
    height: int = 360,
    title: str = "",
    subtotals: bool = True,
    column_layout: str = "stepped",
    *,
    manager: Any | None = None,
) -> dict[str, Any]:
    """Add a matrix / pivot-table visual (``pivotTable``).

    Roles: ``Rows`` (column list — required), ``Columns`` (column list —
    optional), ``Values`` (measure list — required). ``column_layout``
    accepts ``"stepped"`` (compact, single column with indents) or
    ``"tabular"`` (one column per row level).
    """
    if not rows:
        raise PowerBIValidationError("rows must contain at least one column.")
    if not values:
        raise PowerBIValidationError("values must contain at least one measure.")
    layout_token = column_layout.strip().casefold()
    if layout_token not in {"stepped", "tabular"}:
        raise PowerBIValidationError(
            "column_layout must be 'stepped' or 'tabular'.",
            details={"column_layout": column_layout},
        )
    projections: dict[str, list[dict[str, str]]] = {
        "Rows": [{"queryRef": _query_ref(item)} for item in rows],
        "Values": [{"queryRef": _query_ref(item)} for item in values],
    }
    references = [*rows, *values]
    expected_kinds = {r: "column" for r in rows}
    expected_kinds.update({v: "measure" for v in values})
    if columns:
        projections["Columns"] = [{"queryRef": _query_ref(item)} for item in columns]
        references.extend(columns)
        for c in columns:
            expected_kinds[c] = "column"
    _validate_field_references_live(manager, references, expected_kinds=expected_kinds)
    extra_single_visual = {
        "objects": {
            "subTotals": [
                {"properties": {"rowSubtotals": _literal_value(bool(subtotals))}}
            ],
            "general": [
                {"properties": {"layout": _text_literal("Stepped" if layout_token == "stepped" else "Tabular")}}
            ],
        }
    }
    measure_home_map = _resolve_measure_home_map(extract_folder, manager=manager)
    return _append_visual(
        extract_folder,
        page,
        lambda section, home_map: _create_chart_container(
            section,
            visual_type="pivotTable",
            x=x,
            y=y,
            width=width,
            height=height,
            title=title or None,
            projections=projections,
            references=references,
            measure_home_map=home_map,
            extra_single_visual=extra_single_visual,
            manager=manager,
        ),
        measure_home_map,
    )


def pbi_add_text_box_tool(
    extract_folder: str,
    page: str,
    text: str,
    x: int,
    y: int,
    width: int = 280,
    height: int = 80,
    font_size: int = 16,
    bold: bool = False,
    color: str = "#222222",
) -> dict[str, Any]:
    # Text boxes have no field references so manager-augmented home table is unused.
    measure_home_map = _resolve_measure_home_map(extract_folder)
    return _append_visual(
        extract_folder,
        page,
        lambda section, home_map: _make_visual_container(
            section=section,
            visual_type="textbox",
            x=x,
            y=y,
            width=width,
            height=height,
            references=[],
            measure_home_map=home_map,
            projections={},
            extra_single_visual={
                "textContent": text,
                "textStyle": {"fontSize": font_size, "bold": bold, "color": color},
                "prototypeQuery": {"Version": 2, "From": [], "Select": []},
                "objects": {"paragraphs": [{"text": text, "fontSize": font_size, "bold": bold, "color": color}]},
            },
        ),
        measure_home_map,
    )


def pbi_remove_visual_tool(extract_folder: str, page: str, visual_id: str) -> dict[str, Any]:
    def _impl() -> dict[str, Any]:
        folder, layout = _load_layout(extract_folder)
        section = _find_page(layout, page)
        index, _, _ = _find_visual(section, visual_id)
        removed = section["visualContainers"].pop(index)
        _save_layout(folder, layout)
        return ok("Visual removed successfully.", extract_folder=str(folder), page=str(section.get("displayName") or section.get("name")), visual=_visual_payload(removed))

    return _run(_impl)


def pbi_move_visual_tool(extract_folder: str, page: str, visual_id: str, x: int, y: int, width: int | None = None, height: int | None = None) -> dict[str, Any]:
    def _impl() -> dict[str, Any]:
        folder, layout = _load_layout(extract_folder)
        section = _find_page(layout, page)
        _, container, config = _find_visual(section, visual_id)
        new_width = width if width is not None else int(container.get("width", 0))
        new_height = height if height is not None else int(container.get("height", 0))
        _validate_dimensions(x, y, new_width, new_height)
        container.update({"x": x, "y": y, "width": new_width, "height": new_height})
        layouts = config.get("layouts", [])
        if layouts:
            layouts[0].setdefault("position", {})
            layouts[0]["position"].update({"x": x, "y": y, "width": new_width, "height": new_height})
        container["config"] = _dump_embedded_json(config)
        _save_layout(folder, layout)
        return ok("Visual moved successfully.", extract_folder=str(folder), page=str(section.get("displayName") or section.get("name")), visual=_visual_payload(container))

    return _run(_impl)


def pbi_set_visual_format_property_tool(
    extract_folder: str,
    page: str,
    visual_id: str,
    object_name: str,
    properties: dict[str, Any],
    property_types: dict[str, str] | None = None,
) -> dict[str, Any]:
    """Set formatting properties on an existing visual's ``singleVisual.objects[<object_name>][0].properties``.

    Use to override titles, axis labels, data labels, etc. on a visual that's
    already on the report — without rebuilding it. Encodes Python values as
    proper Power BI literals (text in single quotes, ints with ``L`` suffix,
    decimals with ``D`` suffix, ``#RRGGBB`` as solid color, bool as
    ``true``/``false``).

    ``property_types`` lets you force the encoding of a specific property:
    ``"text"``, ``"bool"``, ``"int"``, ``"decimal"``, ``"color"``, or ``"raw"``
    (pass-through).

    Existing properties under the same object are preserved and merged with
    the new values (write-through semantics).
    """
    def _impl() -> dict[str, Any]:
        if not object_name or not str(object_name).strip():
            raise PowerBIValidationError("object_name must be non-empty.")
        if not isinstance(properties, dict) or not properties:
            raise PowerBIValidationError(
                "properties must be a non-empty dict of {property_name: value}.",
                details={"properties": repr(properties)},
            )
        types_map = {k: str(v) for k, v in (property_types or {}).items()}

        folder, layout = _load_layout(extract_folder)
        section = _find_page(layout, page)
        _, container, config = _find_visual(section, visual_id)
        single_visual = config.setdefault("singleVisual", {})
        objects = single_visual.setdefault("objects", {})
        existing = objects.get(object_name) or [{"properties": {}}]
        if not isinstance(existing, list) or not existing:
            existing = [{"properties": {}}]
        merged_props = dict(existing[0].get("properties", {}))

        encoded: dict[str, Any] = {}
        for prop_name, raw_value in properties.items():
            if not prop_name or not str(prop_name).strip():
                raise PowerBIValidationError(
                    "property names must be non-empty strings.",
                    details={"name": repr(prop_name)},
                )
            hint = types_map.get(prop_name)
            encoded[prop_name] = _encode_visual_format_value(raw_value, hint=hint)
        merged_props.update(encoded)
        existing[0]["properties"] = merged_props
        objects[object_name] = existing
        container["config"] = _dump_embedded_json(config)
        _save_layout(folder, layout)
        return ok(
            f"Format properties applied to visual '{visual_id}' (object '{object_name}').",
            extract_folder=str(folder),
            page=str(section.get("displayName") or section.get("name")),
            visual_id=visual_id,
            object=object_name,
            applied=sorted(encoded.keys()),
        )

    return _run(_impl)


def pbi_disable_card_autoscale_tool(
    extract_folder: str,
    page: str | None = None,
    visual_ids: list[str] | None = None,
    label_precision: int = 0,
) -> dict[str, Any]:
    """Disable the auto K/M/B unit-scaling on card visuals.

    By default Power BI cards display large numeric values rescaled (e.g.
    ``119,229`` becomes ``119K``). When the underlying measure already uses a
    custom format string with ``K`` / ``€`` suffixes the result is the
    classic "119K K €" double-suffix bug. This tool sets ``labelDisplayUnits=1``
    (None) plus an explicit ``labelPrecision`` on every card on the report
    (or restricted to ``page`` / ``visual_ids``).

    ``label_precision``: 0 keeps integer display, raise to 2 if you need the
    decimal portion of the underlying measure to render.
    """
    def _impl() -> dict[str, Any]:
        if visual_ids is not None and not isinstance(visual_ids, list):
            raise PowerBIValidationError("visual_ids must be a list of visual ids or None.")
        ids_filter = {str(v).strip() for v in (visual_ids or []) if str(v).strip()}
        folder, layout = _load_layout(extract_folder)

        target_sections: list[dict[str, Any]]
        if page:
            target_sections = [_find_page(layout, page)]
        else:
            target_sections = list(layout.get("sections", []) or [])

        patched: list[dict[str, Any]] = []
        for section in target_sections:
            for container in section.get("visualContainers", []) or []:
                cfg_raw = container.get("config")
                if not isinstance(cfg_raw, str):
                    continue
                cfg = json.loads(cfg_raw)
                sv = cfg.get("singleVisual", {})
                if sv.get("visualType") != "card":
                    continue
                visual_id = str(cfg.get("name", ""))
                if ids_filter and visual_id not in ids_filter:
                    continue
                objects = sv.setdefault("objects", {})
                labels = objects.get("labels")
                if not isinstance(labels, list) or not labels:
                    labels = [{"properties": {}}]
                props = dict(labels[0].get("properties", {}))
                props["labelDisplayUnits"] = _decimal_literal(1)  # 1 = None
                props["labelPrecision"] = _decimal_literal(int(label_precision))
                labels[0]["properties"] = props
                objects["labels"] = labels
                container["config"] = _dump_embedded_json(cfg)
                patched.append({
                    "visual_id": visual_id,
                    "page": str(section.get("displayName") or section.get("name", "")),
                })
        _save_layout(folder, layout)
        return ok(
            f"Disabled autoscale on {len(patched)} card visual(s).",
            extract_folder=str(folder),
            patched=patched,
            patched_count=len(patched),
            label_precision=int(label_precision),
        )

    return _run(_impl)


_VISUAL_TYPE_DISPATCH: dict[str, Callable[..., dict[str, Any]]] = {}


def pbi_add_visual_tool(
    extract_folder: str,
    page: str,
    visual_type: str,
    x: int,
    y: int,
    width: int | None = None,
    height: int | None = None,
    title: str = "",
    config: dict[str, Any] | None = None,
    *,
    manager: Any | None = None,
    dry_run: bool = False,
) -> dict[str, Any]:
    """Generic visual dispatcher. Keeps the per-type tools as stable API surface.

    visual_type: one of card, bar_chart, line_chart, donut, table, waterfall,
                 slicer, gauge, kpi, scatter_chart, combo_chart, matrix, map,
                 text_box, labelled_card.
    config: per-type keyword arguments (e.g. {"measure": "Total Sales"} for card,
            {"category_column": "...", "value_measure": "..."} for bar_chart).
    dry_run: when True, run all validation and binding logic but skip the
             layout disk write. Useful to preview what the call would produce
             before committing — the response carries ``dry_run=True`` and a
             ``write_log`` with one entry per intercepted save.

    When ``manager`` is supplied the dispatchers forward it to the underlying
    ``pbi_add_*_tool`` so live field validation and home-table resolution
    happen — preventing post-write ``measure_home_table_needs_repair`` issues
    on map / scatter / combo / kpi / matrix visuals.
    """
    cfg = dict(config or {})
    visual_key = visual_type.strip().casefold()
    size = DEFAULT_VISUAL_SIZES.get(visual_key)
    effective_width = width if width is not None else (size[0] if size else 320)
    effective_height = height if height is not None else (size[1] if size else 240)

    handler = _VISUAL_TYPE_DISPATCH.get(visual_key)
    if handler is None:
        raise PowerBIValidationError(
            f"Unknown visual_type '{visual_type}'. Allowed: {sorted(_VISUAL_TYPE_DISPATCH)}",
            details={"visual_type": visual_type},
        )
    if manager is not None and "__manager__" not in cfg:
        cfg["__manager__"] = manager

    if dry_run:
        with dry_run_layout_writes() as write_log:
            try:
                result = handler(extract_folder, page, x, y, effective_width, effective_height, title, cfg)
            except Exception:
                raise
        result = dict(result or {})
        result["dry_run"] = True
        result["write_log"] = list(write_log)
        result["message"] = "[dry-run] " + str(result.get("message", ""))
        return result
    return handler(extract_folder, page, x, y, effective_width, effective_height, title, cfg)


def _dispatch_card(extract, page, x, y, w, h, title, cfg):
    measure = cfg.get("measure")
    if not measure:
        raise PowerBIValidationError("card visual requires config.measure", details={"visual_type": "card"})
    return pbi_add_card_tool(extract, page, measure, x, y, w, h, title, manager=cfg.get("__manager__"))


def _dispatch_bar(extract, page, x, y, w, h, title, cfg):
    cat = cfg.get("category_column")
    value = cfg.get("value_measure")
    if not cat or not value:
        raise PowerBIValidationError(
            "bar_chart requires config.category_column and config.value_measure",
            details={"visual_type": "bar_chart"},
        )
    return pbi_add_bar_chart_tool(extract, page, cat, value, x, y, w, h, title, cfg.get("legend_column"), manager=cfg.get("__manager__"))


def _dispatch_line(extract, page, x, y, w, h, title, cfg):
    axis = cfg.get("axis_column")
    measures = cfg.get("value_measures") or []
    if not axis or not measures:
        raise PowerBIValidationError(
            "line_chart requires config.axis_column and config.value_measures (list)",
            details={"visual_type": "line_chart"},
        )
    return pbi_add_line_chart_tool(extract, page, axis, measures, x, y, w, h, title, manager=cfg.get("__manager__"))


def _dispatch_donut(extract, page, x, y, w, h, title, cfg):
    cat = cfg.get("category_column")
    value = cfg.get("value_measure")
    if not cat or not value:
        raise PowerBIValidationError(
            "donut requires config.category_column and config.value_measure",
            details={"visual_type": "donut"},
        )
    return pbi_add_donut_chart_tool(extract, page, cat, value, x, y, w, h, title, manager=cfg.get("__manager__"))


def _dispatch_table(extract, page, x, y, w, h, title, cfg):
    columns = cfg.get("columns") or []
    if not columns:
        raise PowerBIValidationError("table requires config.columns (list)", details={"visual_type": "table"})
    return pbi_add_table_visual_tool(extract, page, columns, x, y, w, h, title, manager=cfg.get("__manager__"))


def _dispatch_waterfall(extract, page, x, y, w, h, title, cfg):
    cat = cfg.get("category_column")
    value = cfg.get("value_measure")
    if not cat or not value:
        raise PowerBIValidationError(
            "waterfall requires config.category_column and config.value_measure",
            details={"visual_type": "waterfall"},
        )
    return pbi_add_waterfall_tool(extract, page, cat, value, x, y, w, h, title, manager=cfg.get("__manager__"))


def _dispatch_slicer(extract, page, x, y, w, h, title, cfg):
    column = cfg.get("column")
    if not column:
        raise PowerBIValidationError("slicer requires config.column", details={"visual_type": "slicer"})
    return pbi_add_slicer_tool(extract, page, column, x, y, w, h, cfg.get("slicer_type", "dropdown"), manager=cfg.get("__manager__"))


def _dispatch_gauge(extract, page, x, y, w, h, title, cfg):
    measure = cfg.get("measure")
    if not measure:
        raise PowerBIValidationError("gauge requires config.measure", details={"visual_type": "gauge"})
    return pbi_add_gauge_tool(
        extract,
        page,
        measure,
        x,
        y,
        w,
        h,
        title,
        cfg.get("target_measure"),
        min_value=cfg.get("min_value"),
        max_value=cfg.get("max_value"),
        target_value=cfg.get("target_value"),
        fill_color=cfg.get("fill_color"),
        target_color=cfg.get("target_color"),
        fill_color_measure=cfg.get("fill_color_measure"),
        manager=cfg.get("__manager__"),
    )


def _dispatch_labelled_card(extract, page, x, y, w, h, title, cfg):
    measure = cfg.get("measure")
    label = cfg.get("label") or title
    if not measure or not label:
        raise PowerBIValidationError(
            "labelled_card requires config.measure and config.label (or title)",
            details={"visual_type": "labelled_card"},
        )
    return pbi_add_labelled_card_tool(
        extract,
        page,
        measure,
        str(label),
        x,
        y,
        w,
        h,
        label_height=int(cfg.get("label_height", 28)),
        label_font_size=int(cfg.get("label_font_size", 11)),
        label_bold=bool(cfg.get("label_bold", True)),
        label_color=str(cfg.get("label_color", "#1F2937")),
        manager=cfg.get("__manager__"),
    )


def _dispatch_map(extract, page, x, y, w, h, title, cfg):
    location = cfg.get("location") or cfg.get("category_column") or cfg.get("category")
    measure = cfg.get("measure") or cfg.get("value_measure")
    if not location:
        raise PowerBIValidationError(
            "map requires config.location (Table.Column with the geographic field)",
            details={"visual_type": "map"},
        )
    return pbi_add_map_tool(extract, page, location, measure, x, y, w, h, title, manager=cfg.get("__manager__"))


def pbi_add_map_tool(
    extract_folder: str,
    page: str,
    location_column: str,
    value_measure: str | None,
    x: int,
    y: int,
    width: int = 420,
    height: int = 320,
    title: str = "",
    *,
    manager: Any | None = None,
) -> dict[str, Any]:
    """Add a bubble/map visual (``map``).

    Roles: ``Category`` (column with the geographic field — country, city,
    Lat/Long…) and optional ``Y`` (measure that drives bubble size).

    Provides feature parity with ``pbi_build_dashboard``'s ``map`` spec so
    callers can use the simpler ``pbi_add_visual_tool(visual_type="map", …)``
    surface without dropping into the dashboard builder.
    """
    projections: dict[str, list[dict[str, str]]] = {
        "Category": [{"queryRef": _query_ref(location_column)}]
    }
    references = [location_column]
    expected_kinds = {location_column: "column"}
    if value_measure:
        projections["Y"] = [{"queryRef": _query_ref(value_measure)}]
        references.append(value_measure)
        expected_kinds[value_measure] = "measure"
    _validate_field_references_live(manager, references, expected_kinds=expected_kinds)
    measure_home_map = _resolve_measure_home_map(extract_folder, manager=manager)
    return _append_visual(
        extract_folder,
        page,
        lambda section, home_map: _create_chart_container(
            section,
            visual_type="map",
            x=x,
            y=y,
            width=width,
            height=height,
            title=title or None,
            projections=projections,
            references=references,
            measure_home_map=home_map,
            manager=manager,
        ),
        measure_home_map,
    )


def _dispatch_text_box(extract, page, x, y, w, h, title, cfg):
    text = cfg.get("text")
    if text is None:
        raise PowerBIValidationError("text_box requires config.text", details={"visual_type": "text_box"})
    return pbi_add_text_box_tool(
        extract,
        page,
        str(text),
        x,
        y,
        w,
        h,
        int(cfg.get("font_size", 16)),
        bool(cfg.get("bold", False)),
        str(cfg.get("color", "#222222")),
    )


_VISUAL_TYPE_DISPATCH.update({
    "card": _dispatch_card,
    "labelled_card": _dispatch_labelled_card,
    "labeled_card": _dispatch_labelled_card,
    "bar_chart": _dispatch_bar,
    "line_chart": _dispatch_line,
    "donut": _dispatch_donut,
    "table": _dispatch_table,
    "waterfall": _dispatch_waterfall,
    "slicer": _dispatch_slicer,
    "gauge": _dispatch_gauge,
    "map": _dispatch_map,
    "text_box": _dispatch_text_box,
    "textbox": _dispatch_text_box,
})


__all__ = [
    "pbi_add_visual_tool",
    "pbi_add_bar_chart_tool",
    "pbi_add_card_tool",
    "pbi_add_combo_chart_tool",
    "pbi_add_donut_chart_tool",
    "pbi_add_gauge_tool",
    "pbi_add_kpi_tool",
    "pbi_add_labelled_card_tool",
    "pbi_add_line_chart_tool",
    "pbi_add_map_tool",
    "pbi_add_matrix_tool",
    "pbi_add_scatter_chart_tool",
    "pbi_add_slicer_tool",
    "pbi_add_table_visual_tool",
    "pbi_add_text_box_tool",
    "pbi_add_waterfall_tool",
    "pbi_apply_design_tool",
    "pbi_apply_theme_tool",
    "pbi_build_dashboard_tool",
    "pbi_compile_report_tool",
    "pbi_create_page_tool",
    "pbi_delete_page_tool",
    "pbi_extract_report_tool",
    "pbi_get_page_tool",
    "pbi_list_pages_tool",
    "pbi_move_visual_tool",
    "pbi_patch_layout_tool",
    "pbi_repair_report_fields_tool",
    "pbi_auto_grid_layout_tool",
    "pbi_convert_visual_type_tool",
    "pbi_describe_page_tool",
    "pbi_disable_card_autoscale_tool",
    "pbi_remove_visual_tool",
    "pbi_set_page_size_tool",
    "pbi_set_visual_format_property_tool",
    "pbi_validate_report_fields_tool",
]
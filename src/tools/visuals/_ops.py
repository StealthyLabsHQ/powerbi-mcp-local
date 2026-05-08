"""Visual + layout ops: remove, move, format, convert type, auto-grid,
patch layout, disable card autoscale.
"""

from __future__ import annotations

import json
import tempfile
import zipfile
from pathlib import Path
from typing import Any

from pbi_connection import PowerBIValidationError, ok

from ._base import (
    DEFAULT_PAGE_HEIGHT,
    DEFAULT_PAGE_WIDTH,
    VISUAL_FIELD_ROLES,
    ReportLayoutError,
    _run,
)
from ._bindings import (
    _live_model_field_index,
    _scan_visual_bindings,
    _validate_projection_roles,
)
from ._containers import _find_visual, _validate_dimensions, _visual_payload
from ._formatting import _decimal_literal, _encode_visual_format_value
from ._home_tables import _persistence_risks, _scan_measure_home_tables
from ._io import _maybe_force_close_powerbi, _page_names_from_layout_bytes
from ._layout import (
    _dump_embedded_json,
    _find_page,
    _load_layout,
    _save_layout,
)
from ._paths import _layout_path, _resolve_extract_folder, _resolve_pbix_path


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
            raise ReportLayoutError(
                "Report/Layout file was not found in the extract folder.", details={"path": str(layout_path)}
            )

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
    """
    new_type_clean = str(new_type).strip()
    if not new_type_clean:
        raise PowerBIValidationError("new_type must be non-empty.")
    if new_type_clean not in VISUAL_FIELD_ROLES:
        raise PowerBIValidationError(
            f"Unknown target visual type '{new_type_clean}'.",
            details={"new_type": new_type_clean, "known_types": sorted(VISUAL_FIELD_ROLES)},
        )

    COMPATIBILITY: dict[tuple[str, str], dict[str, str]] = {
        ("card", "kpi"): {"Values": "Indicator"},
        ("kpi", "card"): {"Indicator": "Values"},
        ("donutChart", "treemap"): {"Category": "Category", "Y": "Y"},
        ("treemap", "donutChart"): {"Category": "Category", "Y": "Y"},
    }
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

        _validate_projection_roles(new_type_clean, new_projections)

        sv["visualType"] = new_type_clean
        sv["projections"] = new_projections
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
    """Compute non-overlapping (x, y, width, height) for visual specs on a
    column-based grid. Pure offline calculation.
    """
    if not isinstance(specs, list) or not specs:
        raise PowerBIValidationError("specs must be a non-empty list of visual configs.")
    if cols < 1:
        raise PowerBIValidationError("cols must be >= 1.", details={"cols": cols})
    if gap < 0:
        raise PowerBIValidationError("gap must be >= 0.", details={"gap": gap})

    usable_width = max(0, page_width - 2 * start_x - gap * max(0, cols - 1))
    cw = int(cell_width if cell_width is not None else (usable_width // cols if cols else usable_width))
    if cw <= 0:
        raise PowerBIValidationError(
            "Computed cell_width is non-positive; reduce cols or start_x or pass cell_width explicitly.",
            details={"page_width": page_width, "start_x": start_x, "gap": gap, "cols": cols},
        )
    ch = int(cell_height) if cell_height is not None else 200

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

        cursor_row, cursor_col = _next_free_cell(cursor_row, cursor_col)
        while cursor_col + col_span > cols or any(
            (cursor_row + r, cursor_col + c) in occupied for r in range(row_span) for c in range(col_span)
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


def pbi_remove_visual_tool(extract_folder: str, page: str, visual_id: str) -> dict[str, Any]:
    def _impl() -> dict[str, Any]:
        folder, layout = _load_layout(extract_folder)
        section = _find_page(layout, page)
        index, _, _ = _find_visual(section, visual_id)
        removed = section["visualContainers"].pop(index)
        _save_layout(folder, layout)
        return ok(
            "Visual removed successfully.",
            extract_folder=str(folder),
            page=str(section.get("displayName") or section.get("name")),
            visual=_visual_payload(removed),
        )

    return _run(_impl)


def pbi_move_visual_tool(
    extract_folder: str, page: str, visual_id: str, x: int, y: int, width: int | None = None, height: int | None = None
) -> dict[str, Any]:
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
        return ok(
            "Visual moved successfully.",
            extract_folder=str(folder),
            page=str(section.get("displayName") or section.get("name")),
            visual=_visual_payload(container),
        )

    return _run(_impl)


def pbi_set_visual_format_property_tool(
    extract_folder: str,
    page: str,
    visual_id: str,
    object_name: str,
    properties: dict[str, Any],
    property_types: dict[str, str] | None = None,
) -> dict[str, Any]:
    """Set formatting properties on an existing visual's
    ``singleVisual.objects[<object_name>][0].properties``. Encodes Python
    values as proper Power BI literals via :func:`_encode_visual_format_value`.
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
    """Disable the auto K/M/B unit-scaling on card visuals (sets
    ``labelDisplayUnits=1`` plus an explicit ``labelPrecision``).
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
                props["labelDisplayUnits"] = _decimal_literal(1)
                props["labelPrecision"] = _decimal_literal(int(label_precision))
                labels[0]["properties"] = props
                objects["labels"] = labels
                container["config"] = _dump_embedded_json(cfg)
                patched.append(
                    {
                        "visual_id": visual_id,
                        "page": str(section.get("displayName") or section.get("name", "")),
                    }
                )
        _save_layout(folder, layout)
        return ok(
            f"Disabled autoscale on {len(patched)} card visual(s).",
            extract_folder=str(folder),
            patched=patched,
            patched_count=len(patched),
            label_precision=int(label_precision),
        )

    return _run(_impl)

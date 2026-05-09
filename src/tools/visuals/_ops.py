"""Visual + layout ops: remove, move, format, convert type, auto-grid,
patch layout, disable card autoscale, update bindings.
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
    VISUAL_ROLE_KINDS,
    ReportLayoutError,
    _run,
)
from ._bindings import (
    _assert_container_bindings,
    _build_prototype_query,
    _live_model_field_index,
    _scan_visual_bindings,
    _sync_container_query,
    _validate_field_references_live,
    _validate_projection_roles,
)
from ._containers import _find_visual, _validate_dimensions, _visual_payload
from ._formatting import _decimal_literal, _encode_visual_format_value
from ._home_tables import _persistence_risks, _resolve_measure_home_map, _scan_measure_home_tables
from ._io import _maybe_force_close_powerbi, _page_names_from_layout_bytes, attempt_pbi_save_before_close
from ._layout import (
    _dump_embedded_json,
    _find_page,
    _load_layout,
    _save_layout,
    dry_run_layout_writes,
)
from ._paths import _layout_path, _resolve_extract_folder, _resolve_pbix_path
from ._refs import _projection, _query_ref


def pbi_patch_layout_tool(
    extract_folder: str,
    pbix_path: str,
    force: bool = False,
    fail_on_persistence_risk: bool = True,
    manager: Any | None = None,
    include_hidden: bool = False,
    save_before_close: bool = True,
) -> dict[str, Any]:
    """Patch the modified Report/Layout back into the PBIX archive.

    When ``force=True`` the call closes (and if necessary kills) Power BI
    Desktop so the PBIX can be overwritten. ``save_before_close`` (default
    ``True``) sends Ctrl+S to every running PBI Desktop window via
    ``PostMessage`` *before* the kill, then waits up to 10 seconds for the
    PBIX mtime to change. This flushes any in-memory TOM mutations
    (measures, columns, role filters) that have not yet been persisted —
    without this, those changes are lost when the process is killed.

    The save attempt is best-effort: it never raises, and the layout patch
    proceeds regardless of whether the save succeeded. The response
    includes ``save_attempt`` with telemetry so the caller can detect
    silent loss.
    """

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

        save_attempt: dict[str, Any] | None = None
        if force and save_before_close:
            save_attempt = attempt_pbi_save_before_close(pbix, timeout_seconds=10.0)

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
            save_attempt=save_attempt,
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
    properties: dict[str, Any] | None = None,
    property_types: dict[str, str] | None = None,
    reset_properties: list[str] | None = None,
) -> dict[str, Any]:
    """Set / reset formatting properties on a visual's
    ``singleVisual.objects[<object_name>][0].properties``.

    Parameters
    ----------
    extract_folder, page, visual_id, object_name:
        Identify the target. ``object_name`` is the Power BI object key
        ('title', 'dataPoint', 'labels', 'background', 'general', etc.).
    properties:
        ``{property_name: value}`` to set. Values are encoded as Power BI
        literals via the type hint in ``property_types``.
    property_types:
        ``{property_name: type_hint}``. Valid hints:

        - ``auto`` (default) — infer from Python type
        - ``bool`` — booleans
        - ``int`` — integers (alias: ``integer``)
        - ``decimal`` — floats / numerics (alias: ``float``, ``number``)
        - ``text`` — strings (alias: ``string``)
        - ``color`` — hex strings ``"#RRGGBB"`` (alias: ``fill``, ``hex``,
          ``rgb``); a previously-encoded ``{"solid": {"color": ...}}``
          object is also accepted and unwrapped automatically
        - ``raw`` — pass the value untouched (advanced)

        Type names are case-insensitive. Unknown hints raise with the
        full list.

    reset_properties:
        Optional list of property names to *delete* from the visual's
        property bag. Use this to revert a property to Power BI's default
        — passing an empty string would leave the key set to a blank
        value instead.

    Examples
    --------
    Set a title (text):

        pbi_set_visual_format_property(
            extract_folder, page, visual_id,
            object_name="title",
            properties={"text": "Sales", "show": True},
            property_types={"text": "text", "show": "bool"},
        )

    Set a fill color (single-series default):

        pbi_set_visual_format_property(
            ...,
            object_name="dataPoint",
            properties={"defaultColor": "#4472C4"},
            property_types={"defaultColor": "color"},
        )

    Reset a property to Power BI's default:

        pbi_set_visual_format_property(
            ...,
            object_name="title",
            reset_properties=["text"],
        )
    """

    def _impl() -> dict[str, Any]:
        if not object_name or not str(object_name).strip():
            raise PowerBIValidationError("object_name must be non-empty.")
        properties_dict = properties or {}
        reset_list = list(reset_properties or [])
        if not properties_dict and not reset_list:
            raise PowerBIValidationError(
                "pass at least one of `properties` (to set) or `reset_properties` (to clear).",
                details={"properties": repr(properties_dict), "reset_properties": repr(reset_list)},
            )
        if properties_dict and not isinstance(properties_dict, dict):
            raise PowerBIValidationError(
                "properties must be a dict of {property_name: value}.",
                details={"properties": repr(properties_dict)},
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
        for prop_name, raw_value in properties_dict.items():
            if not prop_name or not str(prop_name).strip():
                raise PowerBIValidationError(
                    "property names must be non-empty strings.",
                    details={"name": repr(prop_name)},
                )
            # Sentinel: pass "__reset__" as the value to clear a property.
            # Sits alongside the explicit ``reset_properties`` list so
            # callers that build a single dict (e.g. JSON payload) have
            # an in-band reset path.
            if isinstance(raw_value, str) and raw_value == "__reset__":
                reset_list.append(prop_name)
                continue
            hint = types_map.get(prop_name)
            encoded[prop_name] = _encode_visual_format_value(raw_value, hint=hint)
        merged_props.update(encoded)
        cleared: list[str] = []
        for prop_name in reset_list:
            if not prop_name or not str(prop_name).strip():
                continue
            if prop_name in merged_props:
                merged_props.pop(prop_name)
                cleared.append(prop_name)
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
            reset=sorted(cleared),
        )

    return _run(_impl)


def _recover_full_refs_from_prototype(
    prototype_query: dict[str, Any],
    projections: dict[str, Any],
) -> dict[str, list[str]]:
    """Recover ``{role: [full_ref, ...]}`` from a visual's existing
    ``prototypeQuery`` + ``projections``. Columns return ``Table.Column``,
    measures return the bare measure name.
    """
    select_by_name: dict[str, dict[str, Any]] = {}
    for entry in prototype_query.get("Select", []) or []:
        if isinstance(entry, dict) and entry.get("Name"):
            select_by_name[str(entry["Name"])] = entry
    from_alias_to_entity: dict[str, str] = {}
    for entry in prototype_query.get("From", []) or []:
        if isinstance(entry, dict):
            from_alias_to_entity[str(entry.get("Name", ""))] = str(entry.get("Entity", ""))

    full_by_role: dict[str, list[str]] = {}
    if not isinstance(projections, dict):
        return full_by_role
    for role, items in projections.items():
        refs: list[str] = []
        if not isinstance(items, list):
            continue
        for item in items:
            if not isinstance(item, dict):
                continue
            qref = str(item.get("queryRef", ""))
            entry = select_by_name.get(qref)
            if isinstance(entry, dict) and isinstance(entry.get("Column"), dict):
                column = entry["Column"]
                source_ref = (
                    column.get("Expression", {}).get("SourceRef", {})
                    if isinstance(column.get("Expression"), dict)
                    else {}
                )
                alias = str(source_ref.get("Source", "")) if isinstance(source_ref, dict) else ""
                table = from_alias_to_entity.get(alias, "")
                prop = str(column.get("Property", qref))
                refs.append(f"{table}.{prop}" if table else prop)
            elif isinstance(entry, dict) and isinstance(entry.get("Measure"), dict):
                measure = entry["Measure"]
                refs.append(str(measure.get("Property", qref)))
            else:
                refs.append(qref)
        full_by_role[str(role)] = refs
    return full_by_role


def pbi_update_visual_bindings_tool(
    extract_folder: str,
    page: str,
    visual_id: str,
    projections: dict[str, list[str]] | None = None,
    add_to_role: dict[str, list[str]] | None = None,
    remove_from_role: dict[str, list[str]] | None = None,
    *,
    manager: Any | None = None,
    include_hidden: bool = False,
    dry_run: bool = False,
) -> dict[str, Any]:
    """Update an existing visual's field bindings without removing and
    recreating it. Modes:

    - ``projections``: full replacement of all roles. dict keyed by role
      name (e.g. ``"Category"``, ``"Y"``, ``"Values"``); each value is a
      list of field references (``"Table.Column"`` for columns or bare
      measure names).
    - ``add_to_role`` / ``remove_from_role``: incremental edits to specific
      roles. Same value shape. Roles that end up empty after a remove are
      dropped from ``projections``.

    ``projections`` is mutually exclusive with the incremental params.
    Roles are validated against ``VISUAL_FIELD_ROLES[visual_type]``; field
    references are validated against the live model when ``manager`` is
    supplied. The ``prototypeQuery`` is rebuilt from the new reference set
    so Power BI Desktop renders the visual correctly. ``dry_run=True`` runs
    every check but skips the layout disk write.
    """

    def _impl() -> dict[str, Any]:
        if projections is not None and (add_to_role is not None or remove_from_role is not None):
            raise PowerBIValidationError(
                "projections is mutually exclusive with add_to_role / remove_from_role.",
                details={"visual_id": visual_id},
            )
        if projections is None and add_to_role is None and remove_from_role is None:
            raise PowerBIValidationError(
                "pass projections OR at least one of add_to_role / remove_from_role.",
                details={"visual_id": visual_id},
            )
        for source in (projections, add_to_role, remove_from_role):
            if source is None:
                continue
            if not isinstance(source, dict):
                raise PowerBIValidationError(
                    "projections / add_to_role / remove_from_role must be dicts of {role: [refs]}.",
                    details={"visual_id": visual_id},
                )
            for role, refs in source.items():
                if not isinstance(role, str) or not role.strip():
                    raise PowerBIValidationError(
                        "role names must be non-empty strings.",
                        details={"role": repr(role)},
                    )
                if not isinstance(refs, list):
                    raise PowerBIValidationError(
                        f"role '{role}' value must be a list of field reference strings.",
                        details={"role": role},
                    )

        folder, layout = _load_layout(extract_folder)
        section = _find_page(layout, page)
        _, container, config = _find_visual(section, visual_id)
        single_visual = config.setdefault("singleVisual", {})
        visual_type = str(single_visual.get("visualType", "") or "")
        if not visual_type:
            raise PowerBIValidationError(
                f"Visual '{visual_id}' has no visualType set; cannot update bindings.",
                details={"visual_id": visual_id},
            )

        current_projections = single_visual.get("projections", {}) or {}
        current_prototype = single_visual.get("prototypeQuery", {}) or {}
        old_full_refs = _recover_full_refs_from_prototype(current_prototype, current_projections)

        if projections is not None:
            new_full_refs: dict[str, list[str]] = {}
            for role, refs in projections.items():
                cleaned = [str(ref).strip() for ref in refs if isinstance(ref, str) and str(ref).strip()]
                # dedupe while preserving order
                seen: set[str] = set()
                deduped: list[str] = []
                for ref in cleaned:
                    if ref not in seen:
                        seen.add(ref)
                        deduped.append(ref)
                if deduped:
                    new_full_refs[role] = deduped
        else:
            new_full_refs = {role: list(refs) for role, refs in old_full_refs.items()}
            if remove_from_role:
                for role, refs in remove_from_role.items():
                    if role not in new_full_refs:
                        continue
                    short_to_drop = {_query_ref(str(ref)).casefold() for ref in refs if isinstance(ref, str)}
                    new_full_refs[role] = [
                        item for item in new_full_refs[role] if _query_ref(item).casefold() not in short_to_drop
                    ]
                    if not new_full_refs[role]:
                        new_full_refs.pop(role, None)
            if add_to_role:
                for role, refs in add_to_role.items():
                    bucket = new_full_refs.setdefault(role, [])
                    existing_short = {_query_ref(item).casefold() for item in bucket}
                    for ref in refs:
                        if not isinstance(ref, str) or not ref.strip():
                            continue
                        cleaned = ref.strip()
                        short = _query_ref(cleaned).casefold()
                        if short in existing_short:
                            continue
                        bucket.append(cleaned)
                        existing_short.add(short)

        if not new_full_refs:
            raise PowerBIValidationError(
                "Update would leave the visual with no field bindings; remove the visual instead.",
                details={"visual_id": visual_id, "visual_type": visual_type},
            )

        # Convert {role: [full_ref]} → {role: [{"queryRef": short, "active": True}]}
        new_projections: dict[str, list[dict[str, Any]]] = {
            role: [_projection(ref) for ref in refs] for role, refs in new_full_refs.items()
        }

        # Role allowlist + (with manager) ref-kind validation per role.
        _validate_projection_roles(visual_type, new_projections, manager=manager, include_hidden=include_hidden)

        # Live-model existence check per ref, with expected_kind from VISUAL_ROLE_KINDS.
        all_refs: list[str] = []
        expected_kinds: dict[str, str] = {}
        role_kinds = VISUAL_ROLE_KINDS.get(visual_type, {})
        for role, refs in new_full_refs.items():
            kind = role_kinds.get(role)
            for ref in refs:
                if ref not in all_refs:
                    all_refs.append(ref)
                if kind and kind != "any":
                    expected_kinds.setdefault(ref, kind)
        _validate_field_references_live(manager, all_refs, expected_kinds=expected_kinds, include_hidden=include_hidden)

        # Rebuild prototypeQuery from full refs so PBI Desktop renders.
        measure_home_map = _resolve_measure_home_map(extract_folder, manager=manager)
        new_prototype = _build_prototype_query(all_refs, measure_home_map)

        single_visual["projections"] = new_projections
        single_visual["prototypeQuery"] = new_prototype
        container["config"] = _dump_embedded_json(config)
        _sync_container_query(container, new_prototype)
        _assert_container_bindings(container, measure_home_map)
        _save_layout(folder, layout)

        added = [
            {"role": role, "reference": ref}
            for role, refs in new_full_refs.items()
            for ref in refs
            if ref not in old_full_refs.get(role, [])
        ]
        removed = [
            {"role": role, "reference": ref}
            for role, refs in old_full_refs.items()
            for ref in refs
            if ref not in new_full_refs.get(role, [])
        ]

        return ok(
            f"Visual '{visual_id}' bindings updated.",
            extract_folder=str(folder),
            page=str(section.get("displayName") or section.get("name")),
            visual_id=visual_id,
            visual_type=visual_type,
            old_projections=old_full_refs,
            new_projections=new_full_refs,
            added=added,
            removed=removed,
            references=all_refs,
            changed=bool(added or removed),
        )

    if dry_run:
        with dry_run_layout_writes() as write_log:
            result = _run(_impl)
        result = dict(result or {})
        result["dry_run"] = True
        result["write_log"] = list(write_log)
        result["message"] = "[dry-run] " + str(result.get("message", ""))
        return result
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


# ─────────────────────────────────────────────────────────────────────
# Per-series colour + conditional formatting (v0.12.5)
# ─────────────────────────────────────────────────────────────────────


_SERIES_COLOR_DEFAULT_ROLE_ORDER = ("Y", "Values", "Series", "Category")


def _resolve_series_target(
    single_visual: dict[str, Any],
    *,
    series_index: int | None,
    series_name: str | None,
    role_hint: str | None = None,
) -> dict[str, Any]:
    """Resolve a series identifier (index or name) to the matching Select
    entry in the visual's prototypeQuery. Returns the metadata needed to
    build the ``dataPoint`` selector.

    The result keys: ``query_ref``, ``role``, ``kind`` (``measure`` |
    ``column``), ``entity``, ``property``.
    """
    projections = single_visual.get("projections", {}) or {}
    prototype = single_visual.get("prototypeQuery", {}) or {}
    select_by_name: dict[str, dict[str, Any]] = {}
    for entry in prototype.get("Select", []) or []:
        if isinstance(entry, dict) and entry.get("Name"):
            select_by_name[str(entry["Name"])] = entry
    from_alias_to_entity: dict[str, str] = {}
    for entry in prototype.get("From", []) or []:
        if isinstance(entry, dict):
            from_alias_to_entity[str(entry.get("Name", ""))] = str(entry.get("Entity", ""))

    candidate_roles: list[str]
    if role_hint and role_hint in projections:
        candidate_roles = [role_hint]
    else:
        # Walk roles in a predictable order so series_index is stable.
        ordered = [r for r in _SERIES_COLOR_DEFAULT_ROLE_ORDER if r in projections]
        rest = [r for r in projections if r not in ordered]
        candidate_roles = ordered + rest

    flat: list[tuple[str, str]] = []  # (role, queryRef)
    for role in candidate_roles:
        items = projections.get(role) or []
        if not isinstance(items, list):
            continue
        for item in items:
            if not isinstance(item, dict):
                continue
            qref = str(item.get("queryRef", ""))
            if qref:
                flat.append((role, qref))

    if not flat:
        raise PowerBIValidationError(
            "Visual has no projections; cannot target a series.",
            details={"available_roles": list(projections.keys())},
        )

    chosen: tuple[str, str] | None = None
    if series_index is not None:
        if series_index < 0 or series_index >= len(flat):
            raise PowerBIValidationError(
                f"series_index {series_index} out of range; visual has {len(flat)} series.",
                details={
                    "series_index": series_index,
                    "series_count": len(flat),
                    "available": [{"role": r, "queryRef": q} for r, q in flat],
                },
            )
        chosen = flat[series_index]
    elif series_name is not None:
        target = str(series_name).strip()
        for role, qref in flat:
            if qref == target or qref.casefold() == target.casefold():
                chosen = (role, qref)
                break
        if chosen is None:
            # Fall back to matching against the underlying measure / column property name.
            for role, qref in flat:
                entry = select_by_name.get(qref) or {}
                measure = entry.get("Measure") if isinstance(entry, dict) else None
                column = entry.get("Column") if isinstance(entry, dict) else None
                prop = ""
                if isinstance(measure, dict):
                    prop = str(measure.get("Property", ""))
                elif isinstance(column, dict):
                    prop = str(column.get("Property", ""))
                if prop and (prop == target or prop.casefold() == target.casefold()):
                    chosen = (role, qref)
                    break
        if chosen is None:
            raise PowerBIValidationError(
                f"series_name '{series_name}' did not match any series in the visual.",
                details={"available": [{"role": r, "queryRef": q} for r, q in flat]},
            )
    else:
        raise PowerBIValidationError("pass either series_index or series_name.")

    role, qref = chosen
    entry = select_by_name.get(qref) or {}
    measure = entry.get("Measure") if isinstance(entry, dict) else None
    column = entry.get("Column") if isinstance(entry, dict) else None
    if isinstance(measure, dict):
        source_ref = (
            measure.get("Expression", {}).get("SourceRef", {}) if isinstance(measure.get("Expression"), dict) else {}
        )
        alias = str(source_ref.get("Source", "")) if isinstance(source_ref, dict) else ""
        return {
            "kind": "measure",
            "role": role,
            "query_ref": qref,
            "entity": from_alias_to_entity.get(alias, ""),
            "property": str(measure.get("Property", qref)),
        }
    if isinstance(column, dict):
        source_ref = (
            column.get("Expression", {}).get("SourceRef", {}) if isinstance(column.get("Expression"), dict) else {}
        )
        alias = str(source_ref.get("Source", "")) if isinstance(source_ref, dict) else ""
        return {
            "kind": "column",
            "role": role,
            "query_ref": qref,
            "entity": from_alias_to_entity.get(alias, ""),
            "property": str(column.get("Property", qref)),
        }
    return {
        "kind": "queryref",
        "role": role,
        "query_ref": qref,
        "entity": "",
        "property": qref,
    }


def _build_series_selector(target: dict[str, Any]) -> dict[str, Any]:
    """Build the ``id`` selector that pins a dataPoint entry to one series.

    Power BI uses a measure / column reference inside the selector to
    differentiate per-series overrides from the default fill.
    """
    entity = target.get("entity") or ""
    prop = target.get("property") or target.get("query_ref") or ""
    expr = {"SourceRef": {"Entity": entity}} if entity else {"SourceRef": {}}
    if target.get("kind") == "measure":
        return {"measure": {"Expression": expr, "Property": prop}}
    return {"column": {"Expression": expr, "Property": prop}}


def pbi_set_series_color_tool(
    extract_folder: str,
    page: str,
    visual_id: str,
    color: str,
    series_index: int | None = None,
    series_name: str | None = None,
    role: str | None = None,
) -> dict[str, Any]:
    """Set the fill colour for a single series of a chart visual.

    Power BI stores per-series overrides inside ``singleVisual.objects.dataPoint``
    as additional list entries with an ``id`` selector pinned to the target
    series' measure or column. The bare ``defaultColor`` property only
    affects series that do not have an explicit override — using it on a
    multi-series chart paints every series the same colour.

    Specify the target series by ``series_index`` (0-based across the
    visual's projections, in role order) or ``series_name`` (matches the
    queryRef or the underlying measure / column ``Property`` name).
    Optional ``role`` narrows the lookup to a specific projection role
    (e.g. ``"Y"``, ``"Values"``).
    """

    def _impl() -> dict[str, Any]:
        from ._formatting import _coerce_color_value, _solid_color

        hex_color = _coerce_color_value(color)
        encoded = _solid_color(hex_color)

        if series_index is None and not series_name:
            raise PowerBIValidationError("pass either series_index or series_name.")

        folder, layout = _load_layout(extract_folder)
        section = _find_page(layout, page)
        _, container, config = _find_visual(section, visual_id)
        single_visual = config.setdefault("singleVisual", {})
        target = _resolve_series_target(
            single_visual,
            series_index=series_index,
            series_name=series_name,
            role_hint=role,
        )

        objects = single_visual.setdefault("objects", {})
        data_point = objects.get("dataPoint")
        if not isinstance(data_point, list):
            data_point = []

        selector = _build_series_selector(target)
        # Each entry has either a global selector (None / empty id → default
        # for the visual) or a per-series id matching this measure / column.
        match_index: int | None = None
        for idx, entry in enumerate(data_point):
            if not isinstance(entry, dict):
                continue
            entry_id = entry.get("selector") or entry.get("id") or {}
            entry_target = entry_id.get("metadata") if isinstance(entry_id, dict) else None
            measure_meta = entry_id.get("measure") if isinstance(entry_id, dict) else None
            column_meta = entry_id.get("column") if isinstance(entry_id, dict) else None
            prop = None
            if isinstance(measure_meta, dict):
                prop = str(measure_meta.get("Property", ""))
            elif isinstance(column_meta, dict):
                prop = str(column_meta.get("Property", ""))
            if prop and prop == target.get("property"):
                match_index = idx
                break
            if entry_target and str(entry_target) == target.get("query_ref"):
                match_index = idx
                break

        new_entry: dict[str, Any] = {
            "selector": selector,
            "properties": {"fill": encoded},
        }
        if match_index is None:
            data_point.append(new_entry)
        else:
            existing = data_point[match_index] or {}
            props = dict(existing.get("properties", {}))
            props["fill"] = encoded
            existing["properties"] = props
            existing["selector"] = selector
            data_point[match_index] = existing

        objects["dataPoint"] = data_point
        container["config"] = _dump_embedded_json(config)
        _save_layout(folder, layout)
        return ok(
            f"Series colour set on visual '{visual_id}'.",
            extract_folder=str(folder),
            page=str(section.get("displayName") or section.get("name")),
            visual_id=visual_id,
            target=target,
            color=hex_color,
            mode="updated" if match_index is not None else "appended",
        )

    return _run(_impl)


_CONDITIONAL_FORMATS = ("dataBar", "colorScale", "iconSet")
_ICON_SETS = ("threeArrows", "threeArrowsGray", "threeTrafficLights", "threeSymbols", "threeFlags", "fiveArrows")


def pbi_add_conditional_formatting_tool(
    extract_folder: str,
    page: str,
    visual_id: str,
    column_name: str,
    format_type: str,
    min_color: str = "#FF0000",
    mid_color: str | None = None,
    max_color: str = "#00FF00",
    bar_color: str = "#4472C4",
    icon_set: str = "threeArrows",
) -> dict[str, Any]:
    """Add table / matrix conditional formatting (data bar, colour scale, icon set).

    Power BI stores conditional formatting on table-style visuals under
    ``singleVisual.objects.values`` (or ``columnFormatting`` on older
    builds), keyed by a queryRef that matches the column / measure
    bound to the visual's ``Values`` role. We write the canonical shape
    Power BI Desktop emits in current builds.

    Parameters
    ----------
    column_name:
        The display name (Property) of the measure / column inside the
        Values projection. Matched case-insensitively against the
        prototypeQuery Select entries.
    format_type:
        One of ``dataBar``, ``colorScale``, ``iconSet``.
    min_color, mid_color, max_color:
        Colour scale endpoints (and optional midpoint). Hex strings.
    bar_color:
        Data bar fill colour.
    icon_set:
        One of ``threeArrows``, ``threeArrowsGray``, ``threeTrafficLights``,
        ``threeSymbols``, ``threeFlags``, ``fiveArrows``.
    """

    def _impl() -> dict[str, Any]:
        from ._formatting import _coerce_color_value, _literal_value, _solid_color

        if format_type not in _CONDITIONAL_FORMATS:
            raise PowerBIValidationError(
                f"format_type must be one of {_CONDITIONAL_FORMATS}.",
                details={"format_type": format_type},
            )
        if format_type == "iconSet" and icon_set not in _ICON_SETS:
            raise PowerBIValidationError(
                f"icon_set must be one of {_ICON_SETS}.",
                details={"icon_set": icon_set},
            )

        folder, layout = _load_layout(extract_folder)
        section = _find_page(layout, page)
        _, container, config = _find_visual(section, visual_id)
        single_visual = config.setdefault("singleVisual", {})
        prototype = single_visual.get("prototypeQuery", {}) or {}
        target_qref: str | None = None
        for entry in prototype.get("Select", []) or []:
            if not isinstance(entry, dict):
                continue
            measure = entry.get("Measure") if isinstance(entry, dict) else None
            column = entry.get("Column") if isinstance(entry, dict) else None
            prop = ""
            if isinstance(measure, dict):
                prop = str(measure.get("Property", ""))
            elif isinstance(column, dict):
                prop = str(column.get("Property", ""))
            if prop.casefold() == column_name.strip().casefold():
                target_qref = str(entry.get("Name") or "")
                break
        if not target_qref:
            available = [
                str(entry.get("Name", "")) for entry in prototype.get("Select", []) or [] if isinstance(entry, dict)
            ]
            raise PowerBIValidationError(
                f"column '{column_name}' is not bound to this visual.",
                details={"column_name": column_name, "available": available},
            )

        objects = single_visual.setdefault("objects", {})
        values_bag = objects.get("values")
        if not isinstance(values_bag, list) or not values_bag:
            values_bag = [{"properties": {}}]

        properties: dict[str, Any] = {}
        if format_type == "dataBar":
            properties["dataBar"] = {
                "solid": {
                    "show": _literal_value(True),
                    "fill": _solid_color(_coerce_color_value(bar_color)),
                }
            }
        elif format_type == "colorScale":
            scale: dict[str, Any] = {
                "minColor": _solid_color(_coerce_color_value(min_color)),
                "maxColor": _solid_color(_coerce_color_value(max_color)),
            }
            if mid_color is not None:
                scale["midColor"] = _solid_color(_coerce_color_value(mid_color))
            properties["backColor"] = {"gradient": scale}
        else:  # iconSet
            properties["icon"] = {
                "set": _literal_value(icon_set),
                "show": _literal_value(True),
            }

        existing_props = dict(values_bag[0].get("properties", {}) or {})
        # Conditional formatting is stored per-column under a selector; mirror
        # that shape so Power BI Desktop applies it to the right field.
        column_overrides = existing_props.setdefault(
            f"_columnFormatting_{target_qref}",
            {"selector": {"metadata": target_qref}, "properties": {}},
        )
        if isinstance(column_overrides, dict):
            override_props = dict(column_overrides.get("properties", {}) or {})
            override_props.update(properties)
            column_overrides["properties"] = override_props
            existing_props[f"_columnFormatting_{target_qref}"] = column_overrides

        values_bag[0]["properties"] = existing_props
        objects["values"] = values_bag
        container["config"] = _dump_embedded_json(config)
        _save_layout(folder, layout)
        return ok(
            f"Conditional formatting ({format_type}) applied to '{column_name}' on visual '{visual_id}'.",
            extract_folder=str(folder),
            page=str(section.get("displayName") or section.get("name")),
            visual_id=visual_id,
            column_name=column_name,
            queryRef=target_qref,
            format_type=format_type,
        )

    return _run(_impl)

"""Card-style visual tools: card, gauge, KPI, labelled card, text box."""

from __future__ import annotations

from typing import Any

from pbi_connection import PowerBIValidationError, ok

from ._bindings import _validate_field_references_live
from ._containers import _append_visual, _create_chart_container, _make_visual_container
from ._formatting import _datapoint_fill_objects, _gauge_axis_objects, _text_literal
from ._home_tables import _resolve_measure_home_map
from ._refs import _query_ref


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
    against the live model first and the call fails fast on a typo.
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
    (conditional formatting).
    """
    if fill_color and fill_color_measure:
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
        references.append(fill_color_measure)
    extra_objects: dict[str, Any] = {}
    axis_obj = _gauge_axis_objects(min_value, max_value, target_value)
    if axis_obj:
        extra_objects["axis"] = axis_obj
    fill_obj = _datapoint_fill_objects(fill_color, target_color)
    if fill_color_measure:
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
    """Add a native KPI visual.

    Roles: Indicator (measure), TrendLine (column, typically Date), Goal
    (measure, optional). ``direction`` controls the status colour
    interpretation.
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
    """Place a text label above a card value (docx-style label-on-top layout)."""
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

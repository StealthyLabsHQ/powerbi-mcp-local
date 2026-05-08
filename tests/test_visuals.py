"""Standalone tests for report layout and visual tools."""

from __future__ import annotations

import json
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

from security import SECURITY
from tools.formats import (
    PRESETS as FORMAT_PRESETS,
)
from tools.formats import (
    pbi_apply_format_preset_tool,
    pbi_list_format_presets_tool,
)
from tools.visuals import (
    LAYOUT_RELATIVE_PATH,
    _find_pbi_tools,
    _maybe_force_close_powerbi,
    _query_ref,
    pbi_add_bar_chart_tool,
    pbi_add_card_tool,
    pbi_add_donut_chart_tool,
    pbi_add_gauge_tool,
    pbi_add_labelled_card_tool,
    pbi_add_slicer_tool,
    pbi_apply_theme_tool,
    pbi_auto_grid_layout_tool,
    pbi_build_dashboard_tool,
    pbi_compile_report_tool,
    pbi_convert_visual_type_tool,
    pbi_create_page_tool,
    pbi_delete_page_tool,
    pbi_describe_page_tool,
    pbi_disable_card_autoscale_tool,
    pbi_extract_report_tool,
    pbi_get_page_tool,
    pbi_list_pages_tool,
    pbi_move_visual_tool,
    pbi_patch_layout_tool,
    pbi_remove_visual_tool,
    pbi_set_visual_format_property_tool,
)


def _base_layout() -> dict:
    return {
        "id": 0,
        "reportId": "unit-test-report",
        "sections": [
            {
                "name": "ReportSection1",
                "displayName": "Overview",
                "displayOption": 0,
                "width": 1280,
                "height": 720,
                "visualContainers": [],
                "filters": "[]",
            }
        ],
    }


def _write_layout(folder: Path, layout: dict) -> None:
    layout_path = folder / LAYOUT_RELATIVE_PATH
    layout_path.parent.mkdir(parents=True, exist_ok=True)
    layout_path.write_text(json.dumps(layout, ensure_ascii=False, indent=2), encoding="utf-16-le")


def _read_layout(folder: Path) -> dict:
    return json.loads((folder / LAYOUT_RELATIVE_PATH).read_text(encoding="utf-16-le"))


class VisualToolTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp_dir = tempfile.TemporaryDirectory()
        self.root = Path(self.temp_dir.name)
        self.extract_folder = self.root / "report_extracted"
        self.extract_folder.mkdir(parents=True, exist_ok=True)
        _write_layout(self.extract_folder, _base_layout())
        self.pbix_path = self.root / "report.pbix"
        self.pbix_path.write_bytes(b"fake-pbix")
        self.theme_path = self.root / "theme.json"
        self.theme_path.write_text(json.dumps({"name": "Contoso"}), encoding="utf-8")
        self.previous_allowed = [str(item) for item in SECURITY.allowed_base_dirs()]
        SECURITY.configure_allowed_dirs([str(self.root)])

    def tearDown(self) -> None:
        SECURITY.configure_allowed_dirs(self.previous_allowed)
        self.temp_dir.cleanup()

    def _single_visual_from_dashboard_spec(self, spec: dict[str, object]) -> dict[str, object]:
        _write_layout(self.extract_folder, _base_layout())
        response = pbi_build_dashboard_tool(str(self.extract_folder), "Overview", [spec])
        self.assertTrue(response["ok"], response)
        layout = _read_layout(self.extract_folder)
        containers = layout["sections"][0]["visualContainers"]
        self.assertEqual(len(containers), 1)
        config = json.loads(containers[0]["config"])
        return config["singleVisual"]

    def _assert_prototype_query_structure(self, single_visual: dict[str, object], references: list[str]) -> None:
        prototype_query = single_visual["prototypeQuery"]
        self.assertEqual(prototype_query["Version"], 2)

        expected_entities: list[str] = []
        for reference in references:
            entity = reference.rsplit(".", 1)[0] if "." in reference else "$Measures"
            if entity not in expected_entities:
                expected_entities.append(entity)

        from_entries = prototype_query["From"]
        self.assertEqual(len(from_entries), len(expected_entities))
        entity_to_alias = {}
        for entry in from_entries:
            self.assertIn("Name", entry)
            self.assertIn("Entity", entry)
            entity = entry["Entity"]
            alias = entry["Name"]
            self.assertNotIn(entity, entity_to_alias)
            entity_to_alias[entity] = alias
        self.assertEqual(set(entity_to_alias.keys()), set(expected_entities))

        select_entries = prototype_query["Select"]
        self.assertEqual(len(select_entries), len(references))
        for index, (entry, reference) in enumerate(zip(select_entries, references, strict=False)):
            short_name = _query_ref(reference)
            self.assertEqual(entry["Name"], short_name, f"Select[{index}] Name should be short name.")
            self.assertEqual(
                entry["NativeReferenceName"],
                short_name,
                f"Select[{index}] NativeReferenceName should be short name.",
            )

            has_column = "Column" in entry
            has_measure = "Measure" in entry
            self.assertNotEqual(has_column, has_measure, f"Select[{index}] must have either Column or Measure.")

            if "." in reference:
                table = reference.rsplit(".", 1)[0]
                self.assertTrue(has_column, f"Select[{index}] should contain Column for table reference.")
                self.assertEqual(entry["Column"]["Property"], short_name)
                self.assertEqual(entry["Column"]["Expression"]["SourceRef"]["Source"], entity_to_alias[table])
            else:
                self.assertTrue(has_measure, f"Select[{index}] should contain Measure for measure reference.")
                self.assertEqual(entry["Measure"]["Property"], short_name)
                self.assertEqual(entry["Measure"]["Expression"]["SourceRef"]["Source"], entity_to_alias["$Measures"])

    def _assert_projection_structure(
        self,
        single_visual: dict[str, object],
        expected_projections: dict[str, list[str]],
    ) -> None:
        projections = single_visual["projections"]
        self.assertEqual(set(projections.keys()), set(expected_projections.keys()))
        for role, expected_query_refs in expected_projections.items():
            actual_query_refs = [item["queryRef"] for item in projections[role]]
            self.assertEqual(actual_query_refs, expected_query_refs)
            for query_ref in actual_query_refs:
                self.assertNotIn(".", query_ref)

    def _assert_visual_query_projection(
        self,
        *,
        spec: dict[str, object],
        references: list[str],
        expected_projections: dict[str, list[str]],
    ) -> None:
        single_visual = self._single_visual_from_dashboard_spec(spec)
        self._assert_prototype_query_structure(single_visual, references)
        self._assert_projection_structure(single_visual, expected_projections)

    def test_query_ref_returns_short_name_for_columns(self) -> None:
        self.assertEqual(_query_ref("Sales.Category"), "Category")
        self.assertEqual(_query_ref("TotalAmount"), "TotalAmount")
        self.assertEqual(_query_ref("Period.Year"), "Year")

    def test_dashboard_card_prototype_query_and_projection_structure(self) -> None:
        self._assert_visual_query_projection(
            spec={"type": "card", "measure": "TotalAmount", "x": 20, "y": 20, "title": "CA"},
            references=["TotalAmount"],
            expected_projections={"Values": ["TotalAmount"]},
        )

    def test_dashboard_bar_chart_without_legend_structure(self) -> None:
        self._assert_visual_query_projection(
            spec={"type": "bar_chart", "category": "Sales.Category", "measure": "TotalAmount", "x": 20, "y": 20},
            references=["Sales.Category", "TotalAmount"],
            expected_projections={"Category": ["Category"], "Y": ["TotalAmount"]},
        )

    def test_dashboard_bar_chart_with_legend_structure(self) -> None:
        self._assert_visual_query_projection(
            spec={
                "type": "bar_chart",
                "category": "Sales.Category",
                "measure": "TotalAmount",
                "legend": "Products.Family",
                "x": 20,
                "y": 20,
            },
            references=["Sales.Category", "TotalAmount", "Products.Family"],
            expected_projections={"Category": ["Category"], "Y": ["TotalAmount"], "Series": ["Family"]},
        )

    def test_dashboard_line_chart_structure(self) -> None:
        self._assert_visual_query_projection(
            spec={
                "type": "line_chart",
                "axis": "Period.Year",
                "measures": ["TotalAmount", "GrowthRate"],
                "x": 20,
                "y": 20,
            },
            references=["Period.Year", "TotalAmount", "GrowthRate"],
            expected_projections={"Category": ["Year"], "Y": ["TotalAmount", "GrowthRate"]},
        )

    def test_dashboard_donut_structure(self) -> None:
        self._assert_visual_query_projection(
            spec={"type": "donut", "category": "Sales.Category", "measure": "GrossMargin", "x": 20, "y": 20},
            references=["Sales.Category", "GrossMargin"],
            expected_projections={"Category": ["Category"], "Y": ["GrossMargin"]},
        )

    def test_dashboard_table_structure(self) -> None:
        self._assert_visual_query_projection(
            spec={"type": "table", "columns": ["Period.Year", "Sales.Category", "TotalAmount"], "x": 20, "y": 20},
            references=["Period.Year", "Sales.Category", "TotalAmount"],
            expected_projections={"Values": ["Year", "Category", "TotalAmount"]},
        )

    def test_dashboard_waterfall_structure(self) -> None:
        self._assert_visual_query_projection(
            spec={"type": "waterfall", "category": "Period.Year", "measure": "TotalAmount", "x": 20, "y": 20},
            references=["Period.Year", "TotalAmount"],
            expected_projections={"Category": ["Year"], "Y": ["TotalAmount"]},
        )

    def test_dashboard_slicer_structure(self) -> None:
        self._assert_visual_query_projection(
            spec={"type": "slicer", "column": "Customers.Region", "x": 20, "y": 20},
            references=["Customers.Region"],
            expected_projections={"Values": ["Region"]},
        )

    def test_dashboard_gauge_structure(self) -> None:
        self._assert_visual_query_projection(
            spec={"type": "gauge", "measure": "SalesTarget", "x": 20, "y": 20},
            references=["SalesTarget"],
            expected_projections={"Y": ["SalesTarget"]},
        )

    def test_dashboard_text_box_structure(self) -> None:
        self._assert_visual_query_projection(
            spec={"type": "text_box", "text": "Monitoring", "x": 20, "y": 20},
            references=[],
            expected_projections={},
        )

    def test_dashboard_map_without_measure_structure(self) -> None:
        self._assert_visual_query_projection(
            spec={"type": "map", "location": "Customers.Region", "x": 20, "y": 20},
            references=["Customers.Region"],
            expected_projections={"Category": ["Region"]},
        )

    def test_dashboard_map_with_measure_structure(self) -> None:
        self._assert_visual_query_projection(
            spec={"type": "map", "location": "Customers.Region", "measure": "TotalAmount", "x": 20, "y": 20},
            references=["Customers.Region", "TotalAmount"],
            expected_projections={"Category": ["Region"], "Y": ["TotalAmount"]},
        )

    def test_page_and_visual_operations(self) -> None:
        created_page = pbi_create_page_tool(str(self.extract_folder), "KPI")
        self.assertTrue(created_page["ok"], created_page)

        added_card = pbi_add_card_tool(str(self.extract_folder), "Overview", "CA Total", 20, 30, title="Revenue")
        self.assertTrue(added_card["ok"], added_card)
        visual_id = added_card["visual"]["id"]
        self.assertEqual(len(visual_id), 20)

        added_chart = pbi_add_bar_chart_tool(
            str(self.extract_folder),
            "Overview",
            "Dim_Date.Year",
            "CA Total",
            250,
            80,
            title="Revenue by Year",
        )
        self.assertTrue(added_chart["ok"], added_chart)

        moved = pbi_move_visual_tool(str(self.extract_folder), "Overview", visual_id, 60, 70, width=240, height=140)
        self.assertTrue(moved["ok"], moved)
        self.assertEqual(moved["visual"]["x"], 60)
        self.assertEqual(moved["visual"]["width"], 240)

        page = pbi_get_page_tool(str(self.extract_folder), "Overview")
        self.assertTrue(page["ok"], page)
        self.assertEqual(len(page["page"]["visuals"]), 2)
        self.assertEqual(page["page"]["visuals"][0]["type"], "card")

        removed = pbi_remove_visual_tool(str(self.extract_folder), "Overview", visual_id)
        self.assertTrue(removed["ok"], removed)
        page_after = pbi_get_page_tool(str(self.extract_folder), "Overview")
        self.assertEqual(len(page_after["page"]["visuals"]), 1)

        deleted_page = pbi_delete_page_tool(str(self.extract_folder), "KPI")
        self.assertTrue(deleted_page["ok"], deleted_page)

    def test_build_dashboard_with_multiple_visuals(self) -> None:
        response = pbi_build_dashboard_tool(
            str(self.extract_folder),
            "Overview",
            [
                {"type": "card", "measure": "CA Total", "x": 20, "y": 20, "title": "CA"},
                {"type": "bar_chart", "category": "Dim_Date.Year", "measure": "CA Total", "x": 260, "y": 20},
                {"type": "text", "text": "Monitoring", "x": 20, "y": 180, "width": 300, "height": 60},
                {"type": "gauge", "measure": "Margin %", "x": 620, "y": 20},
            ],
        )
        self.assertTrue(response["ok"], response)
        self.assertEqual(len(response["created_visuals"]), 4)

        layout = _read_layout(self.extract_folder)
        page = layout["sections"][0]
        self.assertEqual(len(page["visualContainers"]), 4)
        config = json.loads(page["visualContainers"][1]["config"])
        self.assertEqual(config["singleVisual"]["visualType"], "clusteredBarChart")

    def test_extract_and_compile_reports_with_mocked_subprocess(self) -> None:
        output_pbix = self.root / "compiled.pbix"

        def _fake_run(command, **kwargs):
            self.assertIsInstance(command, list)
            self.assertFalse(kwargs.get("shell", False))
            if "extract" in command:
                target = Path(command[command.index("-extractFolder") + 1])
                target.mkdir(parents=True, exist_ok=True)
                _write_layout(target, _base_layout())
            if "compile" in command:
                compiled = Path(command[command.index("-outPath") + 1])
                compiled.write_bytes(b"compiled")
            return SimpleNamespace(returncode=0, stdout="ok", stderr="")

        with (
            patch("tools.visuals._find_pbi_tools", return_value="pbi-tools"),
            patch("tools.visuals.subprocess.run", side_effect=_fake_run),
        ):
            extracted = pbi_extract_report_tool(str(self.pbix_path))
            compiled = pbi_compile_report_tool(str(self.extract_folder), str(output_pbix))

        self.assertTrue(extracted["ok"], extracted)
        self.assertTrue(compiled["ok"], compiled)
        self.assertEqual(compiled["size_bytes"], len(b"compiled"))

    def test_find_pbi_tools_uses_bundled_binary(self) -> None:
        bundled = Path(__file__).resolve().parent.parent / "tools-bin" / "pbi-tools.core.exe"
        if not bundled.exists():
            self.skipTest("bundled pbi-tools binary is not present")

        with (
            patch.dict("tools.visuals.os.environ", {"PBI_TOOLS_PATH": ""}, clear=False),
            patch(
                "tools.visuals.shutil.which",
                return_value=None,
            ),
            patch("tools.visuals.Path.home", return_value=self.root),
        ):
            self.assertEqual(_find_pbi_tools(), str(bundled))

    def test_force_close_saves_before_kill_fallback(self) -> None:
        with (
            patch("tools.visuals.os.name", "nt"),
            patch(
                "tools.visuals._save_and_close_powerbi_gracefully",
                return_value=False,
            ) as graceful,
            patch("tools.visuals._force_kill_powerbi") as force_kill,
            patch("tools.visuals.time.sleep"),
        ):
            _maybe_force_close_powerbi(True)

        graceful.assert_called_once_with(None)
        force_kill.assert_called_once_with()

    def test_force_close_skips_kill_after_graceful_close(self) -> None:
        with (
            patch("tools.visuals.os.name", "nt"),
            patch(
                "tools.visuals._save_and_close_powerbi_gracefully",
                return_value=True,
            ) as graceful,
            patch("tools.visuals._force_kill_powerbi") as force_kill,
            patch("tools.visuals.time.sleep"),
        ):
            _maybe_force_close_powerbi(True)

        graceful.assert_called_once_with(None)
        force_kill.assert_not_called()

    def test_patch_layout_reports_locked_pbix(self) -> None:
        valid_pbix = self.root / "valid.pbix"
        with zipfile.ZipFile(valid_pbix, "w") as archive:
            archive.writestr("Report/Layout", (self.extract_folder / LAYOUT_RELATIVE_PATH).read_bytes())

        with patch("tools.visuals.Path.replace", side_effect=PermissionError("locked")):
            result = pbi_patch_layout_tool(str(self.extract_folder), str(valid_pbix), fail_on_persistence_risk=False)

        self.assertFalse(result["ok"], result)
        self.assertEqual(result["error"]["code"], "report_layout_error")
        self.assertIn("locked", result["error"]["message"])

    def test_apply_theme_updates_layout(self) -> None:
        applied = pbi_apply_theme_tool(str(self.extract_folder), str(self.theme_path))
        self.assertTrue(applied["ok"], applied)
        layout = _read_layout(self.extract_folder)
        self.assertEqual(layout["activeTheme"]["name"], "theme")
        target = self.extract_folder / "Report" / "StaticResources" / "Themes" / "theme.json"
        self.assertTrue(target.exists())

    def test_gauge_fill_color_measure_binds_conditional_formatting(self) -> None:
        response = pbi_add_gauge_tool(
            str(self.extract_folder),
            "Overview",
            "Margin %",
            20,
            20,
            fill_color_measure="Couleur objectif marge",
        )
        self.assertTrue(response["ok"], response)
        layout = _read_layout(self.extract_folder)
        single_visual = json.loads(layout["sections"][0]["visualContainers"][0]["config"])["singleVisual"]
        fill = single_visual["objects"]["dataPoint"][0]["properties"]["fill"]
        # The fill is a Measure binding (not a Literal hex). It points at the color measure.
        binding = fill["solid"]["color"]["expr"]["Measure"]
        self.assertEqual(binding["Property"], "Couleur objectif marge")

    def test_gauge_fill_color_measure_overrides_static_fill(self) -> None:
        response = pbi_add_gauge_tool(
            str(self.extract_folder),
            "Overview",
            "Margin %",
            20,
            20,
            fill_color="#DC2626",
            fill_color_measure="Couleur objectif marge",
        )
        self.assertTrue(response["ok"], response)
        layout = _read_layout(self.extract_folder)
        single_visual = json.loads(layout["sections"][0]["visualContainers"][0]["config"])["singleVisual"]
        fill = single_visual["objects"]["dataPoint"][0]["properties"]["fill"]
        # Measure binding wins; the static '#DC2626' literal must NOT appear.
        self.assertNotIn("Literal", fill["solid"]["color"]["expr"])
        self.assertEqual(fill["solid"]["color"]["expr"]["Measure"]["Property"], "Couleur objectif marge")

    def test_gauge_axis_min_max_target_and_colors(self) -> None:
        response = pbi_add_gauge_tool(
            str(self.extract_folder),
            "Overview",
            "Margin %",
            20,
            20,
            min_value=0.9,
            max_value=1.0,
            target_value=0.92,
            fill_color="#DC2626",
            target_color="#059669",
            title="Marge",
        )
        self.assertTrue(response["ok"], response)
        layout = _read_layout(self.extract_folder)
        single_visual = json.loads(layout["sections"][0]["visualContainers"][0]["config"])["singleVisual"]
        objects = single_visual["objects"]
        # title still present
        self.assertIn("title", objects)
        # axis literals carry the D-suffix decimal format
        axis_props = objects["axis"][0]["properties"]
        self.assertEqual(axis_props["min"]["expr"]["Literal"]["Value"], "0.9D")
        self.assertEqual(axis_props["max"]["expr"]["Literal"]["Value"], "1.0D")
        self.assertEqual(axis_props["target"]["expr"]["Literal"]["Value"], "0.92D")
        # color fills wrap hex in single-quoted literal
        fill_value = objects["dataPoint"][0]["properties"]["fill"]["solid"]["color"]["expr"]["Literal"]["Value"]
        self.assertEqual(fill_value, "'#DC2626'")
        target_value = objects["dataPoint"][0]["properties"]["targetFill"]["solid"]["color"]["expr"]["Literal"]["Value"]
        self.assertEqual(target_value, "'#059669'")

    def test_gauge_rejects_invalid_color(self) -> None:
        from pbi_connection import PowerBIValidationError

        with self.assertRaises(PowerBIValidationError):
            pbi_add_gauge_tool(
                str(self.extract_folder),
                "Overview",
                "Margin %",
                20,
                20,
                fill_color="red",
            )

    def test_tile_slicer_emits_horizontal_orientation(self) -> None:
        response = pbi_add_slicer_tool(
            str(self.extract_folder),
            "Overview",
            "Dim_Date.Year",
            20,
            20,
            slicer_type="tile",
        )
        self.assertTrue(response["ok"], response)
        layout = _read_layout(self.extract_folder)
        single_visual = json.loads(layout["sections"][0]["visualContainers"][0]["config"])["singleVisual"]
        self.assertEqual(single_visual["slicerType"], "list")
        orientation = single_visual["objects"]["general"][0]["properties"]["orientation"]["expr"]["Literal"]["Value"]
        self.assertEqual(orientation, "1L")

    def test_slicer_rejects_unknown_type(self) -> None:
        from pbi_connection import PowerBIValidationError

        with self.assertRaises(PowerBIValidationError):
            pbi_add_slicer_tool(
                str(self.extract_folder),
                "Overview",
                "Dim_Date.Year",
                20,
                20,
                slicer_type="hexagon",
            )

    def test_labelled_card_creates_textbox_above_card(self) -> None:
        response = pbi_add_labelled_card_tool(
            str(self.extract_folder),
            "Overview",
            "CA Total",
            "CA sur 3 ans",
            40,
            50,
            width=200,
            height=110,
            label_height=30,
        )
        self.assertTrue(response["ok"], response)
        self.assertIn("label", response["visuals"])
        self.assertIn("value", response["visuals"])
        layout = _read_layout(self.extract_folder)
        containers = layout["sections"][0]["visualContainers"]
        self.assertEqual(len(containers), 2)
        textbox_cfg = json.loads(containers[0]["config"])
        card_cfg = json.loads(containers[1]["config"])
        self.assertEqual(textbox_cfg["singleVisual"]["visualType"], "textbox")
        self.assertEqual(card_cfg["singleVisual"]["visualType"], "card")
        # label sits above card: same x, card y = label y + label_height
        self.assertEqual(containers[0]["x"], containers[1]["x"])
        self.assertEqual(containers[0]["y"], 50)
        self.assertEqual(containers[1]["y"], 80)
        self.assertEqual(containers[1]["height"], 80)

    def test_labelled_card_rejects_invalid_label_height(self) -> None:
        from pbi_connection import PowerBIValidationError

        with self.assertRaises(PowerBIValidationError):
            pbi_add_labelled_card_tool(
                str(self.extract_folder),
                "Overview",
                "CA Total",
                "CA",
                10,
                10,
                width=200,
                height=80,
                label_height=80,
            )

    def test_set_visual_format_property_writes_canonical_literals(self) -> None:
        added = pbi_add_card_tool(str(self.extract_folder), "Overview", "CA Total", 10, 10, title="CA")
        self.assertTrue(added["ok"], added)
        visual_id = added["visual"]["id"]

        result = pbi_set_visual_format_property_tool(
            str(self.extract_folder),
            "Overview",
            visual_id,
            object_name="title",
            properties={"show": True, "text": "Mon titre", "fontSize": 14, "fontColor": "#1F2937"},
        )
        self.assertTrue(result["ok"], result)

        layout = _read_layout(self.extract_folder)
        config = json.loads(layout["sections"][0]["visualContainers"][0]["config"])
        title_props = config["singleVisual"]["objects"]["title"][0]["properties"]
        self.assertEqual(title_props["show"]["expr"]["Literal"]["Value"], "true")
        self.assertEqual(title_props["text"]["expr"]["Literal"]["Value"], "'Mon titre'")
        self.assertEqual(title_props["fontSize"]["expr"]["Literal"]["Value"], "14L")
        # Hex string in auto mode is detected as a color and wrapped as a solid fill.
        self.assertEqual(
            title_props["fontColor"]["solid"]["color"]["expr"]["Literal"]["Value"],
            "'#1F2937'",
        )

    def test_set_visual_format_property_supports_explicit_type_hints(self) -> None:
        added = pbi_add_card_tool(str(self.extract_folder), "Overview", "CA Total", 10, 10)
        visual_id = added["visual"]["id"]
        # Force an int that happens to look like 1 to be encoded as a decimal literal.
        result = pbi_set_visual_format_property_tool(
            str(self.extract_folder),
            "Overview",
            visual_id,
            object_name="labels",
            properties={"labelDisplayUnits": 1, "labelPrecision": 0},
            property_types={"labelDisplayUnits": "decimal", "labelPrecision": "decimal"},
        )
        self.assertTrue(result["ok"], result)
        layout = _read_layout(self.extract_folder)
        config = json.loads(layout["sections"][0]["visualContainers"][0]["config"])
        props = config["singleVisual"]["objects"]["labels"][0]["properties"]
        self.assertEqual(props["labelDisplayUnits"]["expr"]["Literal"]["Value"], "1.0D")

    def test_set_visual_format_property_rejects_invalid_visual(self) -> None:
        result = pbi_set_visual_format_property_tool(
            str(self.extract_folder),
            "Overview",
            "nonexistent12345",
            object_name="title",
            properties={"show": True},
        )
        self.assertFalse(result["ok"], result)

    def test_disable_card_autoscale_patches_only_cards(self) -> None:
        # Add a card and a non-card visual; only the card should be patched.
        c1 = pbi_add_card_tool(str(self.extract_folder), "Overview", "CA Total", 10, 10)
        c2 = pbi_add_card_tool(str(self.extract_folder), "Overview", "CA Total", 220, 10)
        bar = pbi_add_bar_chart_tool(
            str(self.extract_folder),
            "Overview",
            "Dim_Date.Year",
            "CA Total",
            10,
            150,
        )
        self.assertTrue(c1["ok"] and c2["ok"] and bar["ok"])

        result = pbi_disable_card_autoscale_tool(str(self.extract_folder))
        self.assertTrue(result["ok"], result)
        self.assertEqual(result["patched_count"], 2)

        layout = _read_layout(self.extract_folder)
        for container in layout["sections"][0]["visualContainers"]:
            cfg = json.loads(container["config"])
            sv = cfg["singleVisual"]
            if sv.get("visualType") != "card":
                continue
            props = sv["objects"]["labels"][0]["properties"]
            self.assertEqual(props["labelDisplayUnits"]["expr"]["Literal"]["Value"], "1.0D")
            self.assertEqual(props["labelPrecision"]["expr"]["Literal"]["Value"], "0.0D")

    def test_field_validation_blocks_missing_measure(self) -> None:
        """When a manager is provided, missing measures fail before layout writes."""
        from pbi_connection import PowerBIValidationError

        # Fake manager that returns a model with one measure 'Real Measure' on table 'T'.
        fake_index = {
            "columns": {("dim_date", "year")},
            "measures": {"real measure": {"t"}},
            "measure_tables": {"real measure": {"T"}},
        }
        with patch("tools.visuals._live_model_field_index", return_value=(fake_index, {"status": "available"})):
            fake_manager = SimpleNamespace()
            # Real measure passes
            ok_response = pbi_add_card_tool(
                str(self.extract_folder), "Overview", "Real Measure", 10, 10, manager=fake_manager
            )
            self.assertTrue(ok_response["ok"], ok_response)
            # Typo'd measure raises with structured details
            with self.assertRaises(PowerBIValidationError) as ctx:
                pbi_add_card_tool(str(self.extract_folder), "Overview", "Reel Mesure", 100, 10, manager=fake_manager)
            self.assertIn("Reel Mesure", ctx.exception.message)
            missing = ctx.exception.details.get("missing")
            self.assertEqual(len(missing), 1)
            entry = missing[0]
            self.assertEqual(entry["reference"], "Reel Mesure")
            self.assertEqual(entry["kind"], "measure")

    def test_field_validation_skipped_without_manager(self) -> None:
        """Without a manager, no live check happens — preserves offline scripting."""
        # Anything goes, even a measure that doesn't exist anywhere.
        response = pbi_add_card_tool(str(self.extract_folder), "Overview", "Totally Made Up Measure", 10, 10)
        self.assertTrue(response["ok"], response)

    def test_field_validation_gauge_checks_target_and_color_measure(self) -> None:
        from pbi_connection import PowerBIValidationError

        fake_index = {
            "columns": set(),
            "measures": {"margin %": {"t"}, "couleur margin": {"t"}},
            "measure_tables": {"margin %": {"T"}, "couleur margin": {"T"}},
        }
        with patch("tools.visuals._live_model_field_index", return_value=(fake_index, {"status": "available"})):
            fake_manager = SimpleNamespace()
            # All three references exist
            ok_response = pbi_add_gauge_tool(
                str(self.extract_folder),
                "Overview",
                "Margin %",
                10,
                10,
                target_measure="Margin %",
                fill_color_measure="Couleur Margin",
                manager=fake_manager,
            )
            self.assertTrue(ok_response["ok"], ok_response)
            # fill_color_measure typo bubbles up
            with self.assertRaises(PowerBIValidationError) as ctx:
                pbi_add_gauge_tool(
                    str(self.extract_folder),
                    "Overview",
                    "Margin %",
                    200,
                    10,
                    fill_color_measure="Coleur Marge",
                    manager=fake_manager,
                )
            self.assertIn("Coleur Marge", ctx.exception.message)

    def test_disable_card_autoscale_with_visual_ids_filter(self) -> None:
        c1 = pbi_add_card_tool(str(self.extract_folder), "Overview", "CA Total", 10, 10)
        c2 = pbi_add_card_tool(str(self.extract_folder), "Overview", "CA Total", 220, 10)
        target_id = c1["visual"]["id"]

        result = pbi_disable_card_autoscale_tool(
            str(self.extract_folder),
            visual_ids=[target_id],
        )
        self.assertTrue(result["ok"], result)
        self.assertEqual(result["patched_count"], 1)
        self.assertEqual(result["patched"][0]["visual_id"], target_id)

    def test_auto_grid_layout_places_specs_on_3_column_grid(self) -> None:
        specs = [{"label": f"v{i}"} for i in range(6)]
        result = pbi_auto_grid_layout_tool(
            specs, cols=3, gap=10, start_x=20, start_y=60, page_width=900, cell_height=180
        )
        self.assertTrue(result["ok"], result)
        out = result["specs"]
        self.assertEqual(len(out), 6)
        # First three on row 0, next three on row 1.
        self.assertEqual(out[0]["x"], 20)
        self.assertEqual(out[0]["y"], 60)
        self.assertEqual(out[3]["y"], 60 + 180 + 10)
        # Cell width derived: (900 - 2*20 - 10*2) / 3 = (900 - 40 - 20)/3 = 280
        self.assertEqual(out[0]["width"], 280)

    def test_auto_grid_layout_handles_col_span(self) -> None:
        specs = [{"col_span": 2}, {}, {}, {}]
        result = pbi_auto_grid_layout_tool(
            specs, cols=3, gap=8, start_x=10, start_y=20, page_width=600, cell_height=100
        )
        out = result["specs"]
        self.assertEqual(len(out), 4)
        # First spec spans 2 columns.
        cw = out[0]["width"]
        self.assertGreater(cw, out[1]["width"])

    def test_convert_visual_type_card_to_kpi_preserves_bindings(self) -> None:
        added = pbi_add_card_tool(str(self.extract_folder), "Overview", "CA Total", 10, 10)
        visual_id = added["visual"]["id"]
        result = pbi_convert_visual_type_tool(str(self.extract_folder), "Overview", visual_id, "kpi")
        self.assertTrue(result["ok"], result)
        layout = _read_layout(self.extract_folder)
        cfg = json.loads(layout["sections"][0]["visualContainers"][0]["config"])
        self.assertEqual(cfg["singleVisual"]["visualType"], "kpi")
        self.assertEqual(cfg["singleVisual"]["projections"]["Indicator"][0]["queryRef"], "CA Total")
        self.assertNotIn("Values", cfg["singleVisual"]["projections"])

    def test_convert_visual_type_rejects_incompatible(self) -> None:
        added = pbi_add_donut_chart_tool(str(self.extract_folder), "Overview", "Dim_Date.Year", "CA Total", 10, 10)
        visual_id = added["visual"]["id"]
        result = pbi_convert_visual_type_tool(str(self.extract_folder), "Overview", visual_id, "kpi")
        # Tool wraps via _run, so the exception surfaces as a structured failure payload.
        self.assertFalse(result["ok"], result)
        self.assertEqual(result["error"]["details"]["reason"], "incompatible")

    def test_describe_page_returns_structured_visuals(self) -> None:
        pbi_add_card_tool(str(self.extract_folder), "Overview", "CA Total", 10, 10, title="CA Card")
        pbi_add_bar_chart_tool(str(self.extract_folder), "Overview", "Dim_Date.Year", "CA Total", 250, 10, title="Bar")
        result = pbi_describe_page_tool(str(self.extract_folder), "Overview")
        self.assertTrue(result["ok"], result)
        self.assertEqual(result["page"]["display_name"], "Overview")
        self.assertEqual(len(result["visuals"]), 2)
        types = {v["type"] for v in result["visuals"]}
        self.assertEqual(types, {"card", "clusteredBarChart"})
        for v in result["visuals"]:
            self.assertIn("position", v)
            self.assertIn("bindings", v)
            self.assertIn("binding_health", v)

    def test_projection_role_validator_rejects_unknown_role(self) -> None:
        # Adding a bar chart with an "Unknown" projection role should fail
        # at tool-call time even though it's offline (no manager).
        from pbi_connection import PowerBIValidationError
        from tools.visuals import _validate_projection_roles

        with self.assertRaises(PowerBIValidationError):
            _validate_projection_roles("clusteredBarChart", {"Bogus": [{"queryRef": "X"}]})

    def test_projection_role_validator_with_manager_detects_kind_mismatch(self) -> None:
        from pbi_connection import PowerBIValidationError
        from tools.visuals import _validate_projection_roles

        fake_index = {
            "columns": {("dim_date", "year")},
            "measures": {"ca total": {"facture"}},
            "measure_tables": {"ca total": {"Facture"}},
        }
        with patch("tools.visuals._live_model_field_index", return_value=(fake_index, {"status": "available"})):
            fake_manager = SimpleNamespace()
            # Putting a measure in Category role should be flagged.
            with self.assertRaises(PowerBIValidationError) as ctx:
                _validate_projection_roles(
                    "clusteredBarChart",
                    {"Category": [{"queryRef": "CA Total"}], "Y": [{"queryRef": "CA Total"}]},
                    manager=fake_manager,
                )
            self.assertIn("mismatches", ctx.exception.details)
            kinds = [m["actual_kind"] for m in ctx.exception.details["mismatches"]]
            self.assertIn("measure", kinds)

    def test_format_presets_catalogue_and_apply(self) -> None:
        listing = pbi_list_format_presets_tool(filter_substring="percent")
        self.assertTrue(listing["ok"], listing)
        self.assertGreaterEqual(listing["count"], 4)
        self.assertIn("percent_4dp", listing["presets"])
        # The catalogue should have the same shape as PRESETS.
        self.assertEqual(FORMAT_PRESETS["currency_eur_k"]["format_string"], "#,##0,\\K \\€")

    def test_format_presets_unknown_raises(self) -> None:
        from pbi_connection import PowerBIValidationError

        with self.assertRaises(PowerBIValidationError):
            from tools.formats import _resolve_preset

            _resolve_preset("nonexistent_preset")

    def test_list_pages_and_security_rejection(self) -> None:
        listed = pbi_list_pages_tool(str(self.extract_folder))
        self.assertTrue(listed["ok"], listed)
        self.assertEqual(listed["pages"][0]["display_name"], "Overview")

        with tempfile.TemporaryDirectory() as outside_dir:
            outside = Path(outside_dir) / "outside_report"
            outside.mkdir()
            _write_layout(outside, _base_layout())
            blocked = pbi_list_pages_tool(str(outside))
        self.assertFalse(blocked["ok"])
        self.assertEqual(blocked["error"]["code"], "security_policy_violation")


if __name__ == "__main__":
    unittest.main(verbosity=2)

"""Standalone tests for report layout and visual tools."""

from __future__ import annotations

import json
import sys
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

from security import SECURITY
from tools.visuals import (
    LAYOUT_RELATIVE_PATH,
    _query_ref,
    pbi_add_bar_chart_tool,
    pbi_add_card_tool,
    pbi_apply_theme_tool,
    pbi_build_dashboard_tool,
    pbi_compile_report_tool,
    pbi_create_page_tool,
    pbi_delete_page_tool,
    pbi_extract_report_tool,
    pbi_get_page_tool,
    pbi_list_pages_tool,
    pbi_move_visual_tool,
    pbi_remove_visual_tool,
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
        for index, (entry, reference) in enumerate(zip(select_entries, references)):
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

        with patch("tools.visuals._find_pbi_tools", return_value="pbi-tools"), patch("tools.visuals.subprocess.run", side_effect=_fake_run):
            extracted = pbi_extract_report_tool(str(self.pbix_path))
            compiled = pbi_compile_report_tool(str(self.extract_folder), str(output_pbix))

        self.assertTrue(extracted["ok"], extracted)
        self.assertTrue(compiled["ok"], compiled)
        self.assertEqual(compiled["size_bytes"], len(b"compiled"))

    def test_apply_theme_updates_layout(self) -> None:
        applied = pbi_apply_theme_tool(str(self.extract_folder), str(self.theme_path))
        self.assertTrue(applied["ok"], applied)
        layout = _read_layout(self.extract_folder)
        self.assertEqual(layout["activeTheme"]["name"], "theme")
        target = self.extract_folder / "Report" / "StaticResources" / "Themes" / "theme.json"
        self.assertTrue(target.exists())

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

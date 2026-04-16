"""Standalone tests for report layout and visual tools."""

from __future__ import annotations

import json
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

from security import SECURITY
from tools.visuals import (
    LAYOUT_RELATIVE_PATH,
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
                {"type": "text", "text": "Pilotage", "x": 20, "y": 180, "width": 300, "height": 60},
                {"type": "gauge", "measure": "Marge %", "x": 620, "y": 20},
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

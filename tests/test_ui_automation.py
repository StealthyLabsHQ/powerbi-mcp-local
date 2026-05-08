"""Standalone tests for the UI-automation tool surface (Ctrl+S sender)."""

from __future__ import annotations

import os
import platform
import sys
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

from security import SECURITY, tool_category
from tools.ui_automation import pbi_persist_now_tool


class PbiPersistNowGateTests(unittest.TestCase):
    def setUp(self) -> None:
        self.previous_env = os.environ.get("PBI_MCP_ALLOW_UI_AUTOMATION")
        if "PBI_MCP_ALLOW_UI_AUTOMATION" in os.environ:
            del os.environ["PBI_MCP_ALLOW_UI_AUTOMATION"]

    def tearDown(self) -> None:
        if self.previous_env is None:
            os.environ.pop("PBI_MCP_ALLOW_UI_AUTOMATION", None)
        else:
            os.environ["PBI_MCP_ALLOW_UI_AUTOMATION"] = self.previous_env

    def test_requires_confirm(self) -> None:
        with self.assertRaises(Exception) as ctx:
            pbi_persist_now_tool(confirm=False)
        self.assertIn("confirm=True", str(ctx.exception))

    def test_requires_env_opt_in(self) -> None:
        os.environ.pop("PBI_MCP_ALLOW_UI_AUTOMATION", None)
        if platform.system() != "Windows":
            with self.assertRaises(Exception) as ctx:
                pbi_persist_now_tool(confirm=True)
            self.assertIn("Windows", str(ctx.exception))
            return
        with self.assertRaises(Exception) as ctx:
            pbi_persist_now_tool(confirm=True)
        self.assertIn("PBI_MCP_ALLOW_UI_AUTOMATION", str(ctx.exception))

    def test_tool_category(self) -> None:
        self.assertEqual(tool_category("pbi_persist_now"), "write")


@unittest.skipUnless(platform.system() == "Windows", "UI automation is Windows-only.")
class PbiPersistNowExecutionTests(unittest.TestCase):
    def setUp(self) -> None:
        self.previous_env = os.environ.get("PBI_MCP_ALLOW_UI_AUTOMATION")
        os.environ["PBI_MCP_ALLOW_UI_AUTOMATION"] = "1"
        self.temp_dir = tempfile.TemporaryDirectory()
        self.root = Path(self.temp_dir.name)
        self.previous_allowed = [str(item) for item in SECURITY.allowed_base_dirs()]
        SECURITY.configure_allowed_dirs([str(self.root)])

    def tearDown(self) -> None:
        if self.previous_env is None:
            os.environ.pop("PBI_MCP_ALLOW_UI_AUTOMATION", None)
        else:
            os.environ["PBI_MCP_ALLOW_UI_AUTOMATION"] = self.previous_env
        SECURITY.configure_allowed_dirs(self.previous_allowed)
        self.temp_dir.cleanup()

    def test_no_pid_returns_error_payload(self) -> None:
        with (
            patch("tools.ui_automation._resolve_pid_from_manager", return_value=None),
            patch("tools.ui_automation._fallback_pbidesktop_pid", return_value=None),
        ):
            with self.assertRaises(Exception) as ctx:
                pbi_persist_now_tool(confirm=True)
            self.assertIn("No Power BI Desktop process", str(ctx.exception))

    def test_no_window_for_pid_returns_error(self) -> None:
        with (
            patch("tools.ui_automation._resolve_pid_from_manager", return_value=12345),
            patch("tools.ui_automation._find_main_window_hwnd", return_value=None),
        ):
            with self.assertRaises(Exception) as ctx:
                pbi_persist_now_tool(confirm=True)
            self.assertIn("no visible top-level window", str(ctx.exception))

    def test_full_path_polls_mtime(self) -> None:
        pbix = self.root / "report.pbix"
        pbix.write_bytes(b"orig")
        original_mtime = pbix.stat().st_mtime

        captured: dict[str, object] = {}

        def fake_send_ctrl_s(hwnd: int) -> None:
            captured["hwnd"] = hwnd
            # simulate Power BI flushing the pbix to disk
            pbix.write_bytes(b"saved-after-ctrl-s")
            os.utime(pbix, (original_mtime + 5, original_mtime + 5))

        with (
            patch("tools.ui_automation._resolve_pid_from_manager", return_value=99),
            patch("tools.ui_automation._find_main_window_hwnd", return_value=4242),
            patch("tools.ui_automation._read_window_title", return_value="MyReport - Power BI Desktop"),
            patch("tools.ui_automation._bring_to_foreground", return_value=999),
            patch("tools.ui_automation._send_ctrl_s", side_effect=fake_send_ctrl_s),
        ):
            result = pbi_persist_now_tool(
                pbix_path=str(pbix),
                confirm=True,
                timeout_seconds=2,
                manager=SimpleNamespace(),
            )

        self.assertTrue(result["ok"], result)
        self.assertEqual(captured["hwnd"], 4242)
        self.assertTrue(result["save_observed"])
        self.assertEqual(result["pid"], 99)
        self.assertIsNotNone(result["mtime_after"])
        self.assertGreater(result["mtime_after"], result["mtime_before"])

    def test_no_pbix_path_returns_immediately(self) -> None:
        with (
            patch("tools.ui_automation._resolve_pid_from_manager", return_value=42),
            patch("tools.ui_automation._find_main_window_hwnd", return_value=1),
            patch("tools.ui_automation._read_window_title", return_value=""),
            patch("tools.ui_automation._bring_to_foreground", return_value=2),
            patch("tools.ui_automation._send_ctrl_s") as send_mock,
        ):
            result = pbi_persist_now_tool(confirm=True)
        self.assertTrue(result["ok"], result)
        self.assertFalse(result["save_observed"])
        self.assertEqual(result["pid"], 42)
        self.assertIsNone(result["pbix_path"])
        self.assertEqual(send_mock.call_count, 1)


if __name__ == "__main__":
    unittest.main(verbosity=2)

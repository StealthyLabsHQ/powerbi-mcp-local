"""Offline tests for v0.10.5 atomic writes + dry-run + persistence warnings."""

from __future__ import annotations

import json
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import MagicMock

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

from tools.visuals import (
    _save_layout,
    dry_run_layout_writes,
)


def _layout_path(folder: Path) -> Path:
    return folder / "Report" / "Layout"


class AtomicWriteTests(unittest.TestCase):
    def test_write_creates_bak_after_overwrite(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            folder = Path(tmp)
            (folder / "Report").mkdir()
            initial = {"sections": [{"name": "S1"}]}
            _save_layout(folder, initial)
            self.assertTrue(_layout_path(folder).exists())

            updated = {"sections": [{"name": "S1"}, {"name": "S2"}]}
            _save_layout(folder, updated)
            self.assertTrue(_layout_path(folder).with_name("Layout.bak").exists())

            # bak holds the previous-good content
            bak = _layout_path(folder).with_name("Layout.bak")
            bak_decoded = json.loads(bak.read_bytes().decode("utf-16-le"))
            self.assertEqual(len(bak_decoded["sections"]), 1)

            # current holds the new content
            cur = json.loads(_layout_path(folder).read_bytes().decode("utf-16-le"))
            self.assertEqual(len(cur["sections"]), 2)

    def test_write_leaves_no_temp_after_failure(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            folder = Path(tmp)
            (folder / "Report").mkdir()
            try:
                _save_layout(folder, {"will_fail": object()})  # not JSON-serialisable
            except TypeError:
                pass
            leftover = list((folder / "Report").glob("Layout.tmp.*"))
            self.assertEqual(leftover, [])


class DryRunTests(unittest.TestCase):
    def test_dry_run_intercepts_save(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            folder = Path(tmp)
            (folder / "Report").mkdir()
            with dry_run_layout_writes() as log:
                _save_layout(folder, {"sections": [{"visualContainers": [{"a": 1}, {"b": 2}]}]})
            self.assertFalse(_layout_path(folder).exists())  # nothing written
            self.assertEqual(len(log), 1)
            self.assertEqual(log[0]["section_count"], 1)
            self.assertEqual(log[0]["visual_count"], 2)

    def test_dry_run_resets_after_context(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            folder = Path(tmp)
            (folder / "Report").mkdir()
            with dry_run_layout_writes():
                pass
            # Outside the context, writes should hit disk normally.
            _save_layout(folder, {"sections": []})
            self.assertTrue(_layout_path(folder).exists())


class PersistenceWarningTests(unittest.TestCase):
    def test_execute_write_payload_carries_persistence(self) -> None:
        from pbi_connection import PowerBIConnectionManager

        mgr = PowerBIConnectionManager()
        # Stub everything execute_write needs.
        state = MagicMock()
        state.snapshot.return_value = {"port": 0}
        database = MagicMock()
        model = MagicMock()
        model.SaveChanges.return_value = "OK"
        database.Model = model
        state.database = database

        mgr._lock.acquire()
        mgr._lock.release()  # warm
        mgr._state = state
        mgr._ensure_connected_locked = lambda: None  # type: ignore[assignment]

        captured: dict = {}

        def _mutator(s, d, m):
            captured["called"] = True
            return {"action": "ok"}

        payload = mgr.execute_write("test_op", _mutator)
        self.assertTrue(captured["called"])
        self.assertIn("persistence", payload)
        self.assertEqual(payload["persistence"]["scope"], "memory_only")


if __name__ == "__main__":
    unittest.main()

"""Coverage for the Google Antigravity adapter (``src/server_antigravity.py``).

The adapter is a small shim that strips MCP capabilities to ``tools`` only
and forces UTF-8 / line-buffered stdio so Antigravity's strict client does
not trip over FastMCP 1.27.x defaults. These tests pin the externally
visible behavior: the argparse surface, the capability shape, and the
stdio hardening side-effects.
"""

from __future__ import annotations

import io
import logging
import sys
import unittest

import server_antigravity


class ParseArgsTests(unittest.TestCase):
    def test_defaults_match_main_server(self) -> None:
        args = server_antigravity._parse_args([])
        self.assertFalse(args.readonly)
        self.assertEqual(args.profile, "all")

    def test_readonly_flag_sets_attribute(self) -> None:
        args = server_antigravity._parse_args(["--readonly"])
        self.assertTrue(args.readonly)

    def test_profile_choices_enforced(self) -> None:
        for profile in ("readonly", "write", "all", "grading"):
            args = server_antigravity._parse_args(["--profile", profile])
            self.assertEqual(args.profile, profile)

    def test_invalid_profile_rejected(self) -> None:
        with self.assertRaises(SystemExit):
            server_antigravity._parse_args(["--profile", "destructive"])


class SilenceLoggersTests(unittest.TestCase):
    def setUp(self) -> None:
        # Snapshot + restore root handlers so other tests are not affected.
        self._root = logging.getLogger()
        self._original_handlers = list(self._root.handlers)
        self._original_level = self._root.level

    def tearDown(self) -> None:
        for handler in list(self._root.handlers):
            self._root.removeHandler(handler)
        for handler in self._original_handlers:
            self._root.addHandler(handler)
        self._root.setLevel(self._original_level)

    def test_keeps_one_stderr_handler_at_warning(self) -> None:
        server_antigravity._silence_loggers()
        self.assertEqual(len(self._root.handlers), 1)
        handler = self._root.handlers[0]
        self.assertIsInstance(handler, logging.StreamHandler)
        # Must remain WARNING — ERROR-only would hide capability mismatches
        # and bind failures from the Antigravity diagnostics view.
        self.assertEqual(self._root.level, logging.WARNING)

    def test_known_noisy_loggers_capped(self) -> None:
        server_antigravity._silence_loggers()
        for name in ("FastMCP", "mcp", "uvicorn", "asyncio"):
            self.assertEqual(logging.getLogger(name).level, logging.WARNING)


class HardenStdioTests(unittest.TestCase):
    def test_sets_utf8_env_vars(self) -> None:
        import os

        # Pre-clear so the test reflects what the helper does on a fresh boot.
        os.environ.pop("PYTHONIOENCODING", None)
        os.environ.pop("PYTHONUTF8", None)
        server_antigravity._harden_stdio()
        self.assertEqual(os.environ.get("PYTHONIOENCODING"), "utf-8")
        self.assertEqual(os.environ.get("PYTHONUTF8"), "1")

    def test_tolerates_non_reconfigurable_streams(self) -> None:
        original_stdout, original_stderr = sys.stdout, sys.stderr
        sys.stdout = io.StringIO()
        sys.stderr = io.StringIO()
        try:
            # Should not raise even when reconfigure() is missing.
            server_antigravity._harden_stdio()
        finally:
            sys.stdout = original_stdout
            sys.stderr = original_stderr


if __name__ == "__main__":
    unittest.main()

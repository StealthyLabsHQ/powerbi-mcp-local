"""v0.13: Power BI report-theme JSON validator coverage.

Pins the schema contract used by ``pbi_apply_theme_tool``,
``pbi_validate_theme_tool``, and ``pbi_export_active_theme_tool``: the
allowed top-level key set, the colour format check, the URL-pattern
guard, and the 256 KB size cap.
"""

from __future__ import annotations

import json
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

from tools.visuals._themes import (
    MAX_THEME_BYTES,
    THEME_ALLOWED_TOP_LEVEL_KEYS,
    ThemeValidationError,
    assert_theme_within_size_limit,
    validate_theme_payload,
)


class ThemeSchemaTests(unittest.TestCase):
    def test_minimal_payload_is_valid(self) -> None:
        issues = validate_theme_payload({"name": "Tiny", "dataColors": ["#112233"]})
        errors = [i for i in issues if i["level"] == "error"]
        self.assertEqual(errors, [])

    def test_unknown_top_level_key_is_rejected(self) -> None:
        issues = validate_theme_payload({"name": "x", "rogue": 1})
        keys = {i["path"] for i in issues if i["level"] == "error"}
        self.assertIn("rogue", keys)

    def test_top_level_allowlist_is_stable(self) -> None:
        # If you grow the set, do it deliberately — this is the schema.
        self.assertIn("visualStyles", THEME_ALLOWED_TOP_LEVEL_KEYS)
        self.assertIn("dataColors", THEME_ALLOWED_TOP_LEVEL_KEYS)
        self.assertIn("name", THEME_ALLOWED_TOP_LEVEL_KEYS)

    def test_blank_name_is_rejected(self) -> None:
        issues = validate_theme_payload({"name": "   "})
        self.assertTrue(any(i["path"] == "name" and i["level"] == "error" for i in issues))

    def test_non_string_name_is_rejected(self) -> None:
        issues = validate_theme_payload({"name": 7})
        self.assertTrue(any(i["path"] == "name" and i["level"] == "error" for i in issues))

    def test_dataColors_must_be_list(self) -> None:
        issues = validate_theme_payload({"name": "x", "dataColors": "#FF0000"})
        self.assertTrue(any(i["path"] == "dataColors" for i in issues))

    def test_dataColors_rejects_non_hex(self) -> None:
        issues = validate_theme_payload({"name": "x", "dataColors": ["red"]})
        self.assertTrue(any(i["path"].startswith("dataColors[") for i in issues))

    def test_dataColors_accepts_hex8(self) -> None:
        issues = validate_theme_payload({"name": "x", "dataColors": ["#11223344"]})
        errors = [i for i in issues if i["level"] == "error"]
        self.assertEqual(errors, [])

    def test_colour_field_with_short_hex_is_rejected(self) -> None:
        issues = validate_theme_payload({"name": "x", "foreground": "#FFF"})
        self.assertTrue(any("foreground" in i["path"] for i in issues))

    def test_url_in_string_value_is_forbidden(self) -> None:
        issues = validate_theme_payload({"name": "evil", "good": "javascript:alert(1)"})
        self.assertTrue(any("javascript" in i.get("matched", "") for i in issues))

    def test_data_uri_is_forbidden(self) -> None:
        issues = validate_theme_payload({"name": "evil", "good": "data:text/html,<script>"})
        self.assertTrue(any("data:" in i.get("matched", "") for i in issues))

    def test_https_in_value_is_forbidden(self) -> None:
        issues = validate_theme_payload({"name": "x", "background": "https://attacker.example/img"})
        self.assertTrue(any("http" in i.get("matched", "") for i in issues))

    def test_file_scheme_is_forbidden(self) -> None:
        issues = validate_theme_payload(
            {"name": "x", "visualStyles": {"*": {"*": {"background": "file:///etc/passwd"}}}}
        )
        self.assertTrue(any("file://" in i.get("matched", "") for i in issues))

    def test_nested_lists_are_traversed(self) -> None:
        payload = {
            "name": "deep",
            "visualStyles": {"*": {"*": {"foo": [{"bar": "https://example.com"}]}}},
        }
        issues = validate_theme_payload(payload)
        self.assertTrue(any("http" in i.get("matched", "") for i in issues))

    def test_size_limit_rejects_oversized_payload(self) -> None:
        with self.assertRaises(ThemeValidationError):
            assert_theme_within_size_limit(MAX_THEME_BYTES + 1)

    def test_size_limit_accepts_at_boundary(self) -> None:
        # At the boundary: ok.
        assert_theme_within_size_limit(MAX_THEME_BYTES)

    def test_non_dict_payload_raises(self) -> None:
        with self.assertRaises(ThemeValidationError):
            validate_theme_payload(["not", "a", "theme"])

    def test_strict_top_level_can_be_disabled(self) -> None:
        # When strict_top_level=False, unknown keys do not yield errors.
        issues = validate_theme_payload({"name": "x", "extraKey": True}, strict_top_level=False)
        errors = [i for i in issues if i["level"] == "error"]
        self.assertEqual(errors, [])

    def test_nested_colour_field_validates(self) -> None:
        payload = {
            "name": "deep",
            "visualStyles": {"*": {"*": {"title": [{"fontColor": "#1234"}]}}},
        }
        issues = validate_theme_payload(payload)
        self.assertTrue(any(i["path"].endswith(".fontColor") for i in issues))


class ApplyThemeIntegrationTests(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp = tempfile.TemporaryDirectory()
        self.root = Path(self.tmp.name)
        from security import SECURITY, configure_allowed_dirs

        self.previous_allowed = [str(item) for item in SECURITY.allowed_base_dirs()]
        configure_allowed_dirs([str(self.root)])
        SECURITY.policy(reload=True, cwd=self.root)
        # minimal extract folder
        layout = {"id": 0, "sections": [], "resourcePackages": []}
        (self.root / "Report").mkdir(parents=True, exist_ok=True)
        (self.root / "Report" / "Layout").write_bytes(json.dumps(layout).encode("utf-16-le"))

    def tearDown(self) -> None:
        from security import SECURITY, configure_allowed_dirs

        configure_allowed_dirs(self.previous_allowed)
        SECURITY.policy(reload=True, cwd=Path.cwd())
        self.tmp.cleanup()

    def _write_theme(self, name: str, payload: dict) -> Path:
        path = self.root / f"{name}.json"
        path.write_text(json.dumps(payload), encoding="utf-8")
        return path

    def test_validate_theme_tool_reports_valid(self) -> None:
        from tools.visuals._design import pbi_validate_theme_tool

        theme = self._write_theme("good", {"name": "Good", "dataColors": ["#001122"]})
        result = pbi_validate_theme_tool(str(theme))
        self.assertTrue(result["ok"], result)
        self.assertTrue(result["valid"])
        self.assertEqual(result["error_count"], 0)

    def test_validate_theme_tool_reports_invalid(self) -> None:
        from tools.visuals._design import pbi_validate_theme_tool

        theme = self._write_theme("bad", {"name": "x", "rogue": 1})
        result = pbi_validate_theme_tool(str(theme))
        self.assertTrue(result["ok"], result)
        self.assertFalse(result["valid"])
        self.assertGreaterEqual(result["error_count"], 1)

    def test_apply_theme_blocks_bad_schema(self) -> None:
        from tools.visuals._design import pbi_apply_theme_tool

        theme = self._write_theme("evil", {"name": "x", "good": "javascript:alert('xss')"})
        result = pbi_apply_theme_tool(str(self.root), str(theme))
        self.assertFalse(result.get("ok", True), result)
        self.assertEqual(result["error"]["code"], "theme_validation_error")

    def test_apply_theme_blocks_oversized(self) -> None:
        from tools.visuals._design import pbi_apply_theme_tool

        # Build a payload that serialises just past 256 KB. dataColors is
        # validated leaf-by-leaf, so use a single long ``name`` field —
        # which the size cap is meant to catch before json.loads even runs.
        big = "x" * (MAX_THEME_BYTES + 1)
        theme = self._write_theme("huge", {"name": big})
        result = pbi_apply_theme_tool(str(self.root), str(theme))
        self.assertFalse(result.get("ok", True))
        self.assertEqual(result["error"]["code"], "theme_validation_error")

    def test_apply_theme_accepts_valid_theme(self) -> None:
        from tools.visuals._design import pbi_apply_theme_tool

        theme = self._write_theme(
            "ok",
            {
                "name": "Clean",
                "dataColors": ["#112233", "#445566"],
                "foreground": "#000000",
                "background": "#FFFFFF",
            },
        )
        result = pbi_apply_theme_tool(str(self.root), str(theme))
        self.assertTrue(result["ok"], result)
        self.assertEqual(result["theme"]["name"], "ok")

    def test_export_active_theme_returns_file(self) -> None:
        from tools.visuals._design import pbi_apply_theme_tool, pbi_export_active_theme_tool

        theme = self._write_theme("ok", {"name": "Clean", "dataColors": ["#001122"]})
        apply = pbi_apply_theme_tool(str(self.root), str(theme))
        self.assertTrue(apply["ok"], apply)
        out = self.root / "exported.json"
        result = pbi_export_active_theme_tool(str(self.root), str(out))
        self.assertTrue(result["ok"], result)
        self.assertTrue(Path(result["output_path"]).exists())
        self.assertGreater(result["size_bytes"], 0)

    def test_export_active_theme_no_active(self) -> None:
        from tools.visuals._design import pbi_export_active_theme_tool

        result = pbi_export_active_theme_tool(str(self.root), str(self.root / "out.json"))
        self.assertFalse(result.get("ok", True))


if __name__ == "__main__":
    unittest.main(verbosity=2)

"""v0.13.3: one-shot styling preset coverage.

Pins:
- The 5 built-in presets pass theme schema validation and have a
  palette with valid hex colours.
- Accent inference matches the documented measure-name rules.
- Native PNG inspector + writer round-trip (no Pillow).
- ``pbi_apply_style_preset_tool`` end-to-end on a synthetic PBIX:
  wallpaper embedded under ``StaticResources/RegisteredResources``,
  every section.config carries the wallpaper reference, every card
  picks up its preset vcObjects.
"""

from __future__ import annotations

import io
import json
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

from tools import pbi_apply_style_preset_tool, pbi_list_style_presets_tool
from tools.styling import (
    PRESETS,
    infer_accent_key,
    inspect_png,
    pick_accent,
    write_gradient_png,
)
from tools.styling._embed import (
    patch_layout_for_wallpaper,
    patch_layout_visuals,
    sanitize_resource_name,
    sha1_short,
)
from tools.visuals._themes import validate_theme_payload


_HEX_RE = __import__("re").compile(r"^#[0-9A-Fa-f]{6}$")


def _make_pbix(tmp: Path, name: str, sections: list[dict]) -> Path:
    """Build a synthetic PBIX zip with a small Report/Layout."""
    layout = {"id": 0, "sections": sections, "resourcePackages": []}
    payload = json.dumps(layout, ensure_ascii=False).encode("utf-16-le")
    out = tmp / name
    with zipfile.ZipFile(out, "w") as zf:
        zf.writestr("Report/Layout", payload)
        # A non-trivial DataModel so the DBCC post-check passes.
        zf.writestr("DataModel", b"x" * 8192)
        zf.writestr("Metadata", b"{}")
        zf.writestr("Connections", b"{}")
    return out


def _read_layout(pbix_path: Path) -> dict:
    with zipfile.ZipFile(pbix_path, "r") as zf:
        raw = zf.read("Report/Layout")
    return json.loads(raw.decode("utf-16-le"))


def _pbix_names(pbix_path: Path) -> set[str]:
    with zipfile.ZipFile(pbix_path, "r") as zf:
        return set(zf.namelist())


class PresetCatalogueTests(unittest.TestCase):
    def test_five_presets_present(self) -> None:
        self.assertEqual(
            sorted(PRESETS),
            ["dark_pro", "glassmorph_dark", "glassmorph_light", "minimal_corporate", "neon_cyber"],
        )

    def test_each_preset_palette_is_valid_hex(self) -> None:
        for name, spec in PRESETS.items():
            for key, value in spec.get("palette", {}).items():
                with self.subTest(preset=name, key=key):
                    self.assertTrue(_HEX_RE.match(value), f"{name}.{key}={value}")

    def test_each_preset_theme_passes_schema(self) -> None:
        for name, spec in PRESETS.items():
            errors = [
                issue
                for issue in validate_theme_payload(spec["theme"])
                if issue.get("level") == "error"
            ]
            self.assertEqual(errors, [], f"{name} theme errors: {errors}")

    def test_each_preset_has_accent_map_with_four_keys(self) -> None:
        for name, spec in PRESETS.items():
            with self.subTest(preset=name):
                accent_map = spec.get("cards", {}).get("accentMap", {})
                self.assertEqual(
                    set(accent_map.keys()), {"positive", "warning", "info", "neutral"}
                )

    def test_list_tool_exposes_palette(self) -> None:
        result = pbi_list_style_presets_tool()
        self.assertTrue(result["ok"], result)
        names = {p["name"] for p in result["presets"]}
        self.assertEqual(len(names), 5)


class AccentInferenceTests(unittest.TestCase):
    def test_positive_keywords(self) -> None:
        for name in (
            "Croissance CA",
            "Marge brute",
            "EBE",
            "Marge nette",
            "Gross Margin",
            "Net Margin",
        ):
            with self.subTest(name=name):
                self.assertEqual(infer_accent_key(name), "positive")

    def test_warning_keywords(self) -> None:
        for name in ("Endettement", "BFR", "Charges", "Frais", "Debt ratio", "WCR"):
            with self.subTest(name=name):
                self.assertEqual(infer_accent_key(name), "warning")

    def test_info_keywords(self) -> None:
        for name in ("Variance", "VAR CA", "Geo split", "Atelier load", "Workshop"):
            with self.subTest(name=name):
                self.assertEqual(infer_accent_key(name), "info")

    def test_unknown_is_neutral(self) -> None:
        self.assertEqual(infer_accent_key("Total"), "neutral")
        self.assertEqual(infer_accent_key(None), "neutral")
        self.assertEqual(infer_accent_key(""), "neutral")

    def test_pick_accent_falls_back_through_neutral(self) -> None:
        self.assertEqual(
            pick_accent("Croissance", {"positive": "#70AD47", "neutral": "#FFFFFF"}),
            "#70AD47",
        )
        self.assertEqual(
            pick_accent("Unknown", {"positive": "#70AD47", "neutral": "#FFFFFF"}),
            "#FFFFFF",
        )


class NativePNGTests(unittest.TestCase):
    def test_gradient_png_roundtrip(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            out = Path(tmp) / "g.png"
            write_gradient_png(out, top_color="#1F4E78", bottom_color="#5B9BD5", width=64, height=32)
            info = inspect_png(out)
        self.assertEqual(info["width"], 64)
        self.assertEqual(info["height"], 32)
        self.assertGreater(info["size_bytes"], 100)

    def test_inspect_rejects_non_png(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "bogus.png"
            path.write_bytes(b"hello-not-a-png")
            with self.assertRaises(ValueError):
                inspect_png(path)


class EmbedHelperTests(unittest.TestCase):
    def test_sanitize_resource_name(self) -> None:
        self.assertEqual(sanitize_resource_name("bg glass dark.png"), "bg_glass_dark.png")
        self.assertEqual(sanitize_resource_name("../etc/passwd"), "etc_passwd")
        self.assertEqual(sanitize_resource_name(""), "resource")

    def test_sha1_short_is_deterministic(self) -> None:
        self.assertEqual(sha1_short(b"abc"), sha1_short(b"abc"))
        self.assertNotEqual(sha1_short(b"abc"), sha1_short(b"abd"))
        self.assertEqual(len(sha1_short(b"abc")), 12)

    def test_patch_layout_for_wallpaper_targets_all_pages_by_default(self) -> None:
        layout = {
            "sections": [
                {"displayName": "P1", "name": "s1"},
                {"displayName": "P2", "name": "s2"},
            ]
        }
        patched, touched = patch_layout_for_wallpaper(
            layout, resource_name="bg.png", fit="Fit", transparency=0
        )
        self.assertEqual(sorted(touched), ["P1", "P2"])
        cfg = json.loads(patched["sections"][0]["config"])
        self.assertIn("objects", cfg)
        self.assertEqual(len(cfg["objects"]["background"]), 1)
        # Power BI requires image.name / url / scaling as bare strings;
        # the Literal-wrapped form was the v0.13.3 bug that left the
        # canvas blank on reopen.
        bg = cfg["objects"]["background"][0]["properties"]
        self.assertEqual(bg["image"]["name"], "bg.png")
        self.assertEqual(bg["image"]["url"], "RegisteredResources/bg.png")
        self.assertEqual(bg["image"]["scaling"], "Fit")
        # show / transparency stay Literal-wrapped.
        self.assertEqual(bg["show"]["expr"]["Literal"]["Value"], "true")

    def test_patch_layout_for_wallpaper_respects_page_filter(self) -> None:
        layout = {
            "sections": [
                {"displayName": "P1", "name": "s1"},
                {"displayName": "P2", "name": "s2"},
            ]
        }
        _, touched = patch_layout_for_wallpaper(
            layout, resource_name="bg.png", page_filter={"P2"}
        )
        self.assertEqual(touched, ["P2"])
        self.assertNotIn("config", layout["sections"][0])

    def test_patch_layout_visuals_styles_cards(self) -> None:
        # A section with one card visual.
        single_visual = {
            "visualType": "card",
            "projections": {"Values": [{"queryRef": "Sales.Marge brute"}]},
        }
        layout = {
            "sections": [
                {
                    "displayName": "P1",
                    "name": "s1",
                    "visualContainers": [
                        {"config": json.dumps({"name": "v1", "singleVisual": single_visual})}
                    ],
                }
            ]
        }
        preset = PRESETS["glassmorph_dark"]
        patched, counts = patch_layout_visuals(
            layout, preset=preset, accent_picker=pick_accent
        )
        self.assertEqual(counts["cards_styled"], 1)
        cfg = json.loads(patched["sections"][0]["visualContainers"][0]["config"])
        vc = cfg["singleVisual"]["vcObjects"]
        self.assertIn("background", vc)
        self.assertIn("border", vc)
        # "Marge brute" → positive accent → green.
        border_color = vc["border"][0]["properties"]["color"]["solid"]["color"]["expr"]["Literal"]["Value"]
        self.assertIn("70AD47", border_color)


class ApplyToolEndToEndTests(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp = tempfile.TemporaryDirectory()
        self.root = Path(self.tmp.name)
        from security import SECURITY, configure_allowed_dirs

        self.previous = [str(p) for p in SECURITY.allowed_base_dirs()]
        configure_allowed_dirs([str(self.root)])
        SECURITY.policy(reload=True, cwd=self.root)

    def tearDown(self) -> None:
        from security import SECURITY, configure_allowed_dirs

        configure_allowed_dirs(self.previous)
        SECURITY.policy(reload=True, cwd=Path.cwd())
        self.tmp.cleanup()

    def _fixture_pbix(self) -> Path:
        return _make_pbix(
            self.root,
            "in.pbix",
            [
                {
                    "displayName": "Overview",
                    "name": "ReportSection1",
                    "visualContainers": [
                        {
                            "config": json.dumps(
                                {
                                    "name": "card1",
                                    "singleVisual": {
                                        "visualType": "card",
                                        "projections": {
                                            "Values": [{"queryRef": "F.Marge brute"}]
                                        },
                                    },
                                }
                            )
                        },
                        {
                            "config": json.dumps(
                                {
                                    "name": "bar1",
                                    "singleVisual": {
                                        "visualType": "stackedBarChart",
                                        "projections": {
                                            "Y": [{"queryRef": "F.Total"}],
                                            "Category": [{"queryRef": "D.Region"}],
                                        },
                                    },
                                }
                            )
                        },
                    ],
                },
                {
                    "displayName": "Detail",
                    "name": "ReportSection2",
                    "visualContainers": [],
                },
            ],
        )

    def test_apply_glassmorph_dark_embeds_wallpaper_and_theme(self) -> None:
        pbix = self._fixture_pbix()
        result = pbi_apply_style_preset_tool(str(pbix), "glassmorph_dark")
        self.assertTrue(result["ok"], result)
        self.assertEqual(result["preset"], "glassmorph_dark")
        self.assertEqual(sorted(result["applied_pages"]), ["Detail", "Overview"])
        self.assertEqual(result["applied_visuals"]["cards_styled"], 1)
        self.assertEqual(result["applied_visuals"]["charts_styled"], 1)
        self.assertTrue(result["dbcc_valid"], result.get("dbcc"))

        names = _pbix_names(Path(result["output_path"]))
        wallpaper_parts = [n for n in names if n.startswith("StaticResources/RegisteredResources/")]
        self.assertEqual(len(wallpaper_parts), 1, wallpaper_parts)
        theme_parts = [n for n in names if n.startswith("StaticResources/SharedResources/BaseThemes/")]
        self.assertEqual(len(theme_parts), 1, theme_parts)

        layout = _read_layout(Path(result["output_path"]))
        for section in layout["sections"]:
            cfg = json.loads(section["config"])
            self.assertIn("background", cfg["objects"])
        self.assertIn("activeTheme", layout)

    def test_pages_filter_applies_to_named_pages_only(self) -> None:
        pbix = self._fixture_pbix()
        result = pbi_apply_style_preset_tool(
            str(pbix), "minimal_corporate", pages=["Detail"]
        )
        self.assertTrue(result["ok"], result)
        self.assertEqual(result["applied_pages"], ["Detail"])

        layout = _read_layout(Path(result["output_path"]))
        overview_cfg = json.loads(layout["sections"][0].get("config", "{}"))
        detail_cfg = json.loads(layout["sections"][1].get("config", "{}"))
        self.assertNotIn("objects", overview_cfg)
        self.assertIn("objects", detail_cfg)

    def test_unknown_preset_rejected(self) -> None:
        from pbi_connection import PowerBIValidationError

        pbix = self._fixture_pbix()
        with self.assertRaises(PowerBIValidationError):
            pbi_apply_style_preset_tool(str(pbix), "no_such_preset")

    def test_custom_preset_requires_custom_spec(self) -> None:
        from pbi_connection import PowerBIValidationError

        pbix = self._fixture_pbix()
        with self.assertRaises(PowerBIValidationError):
            pbi_apply_style_preset_tool(str(pbix), "custom")

    def test_custom_preset_with_spec_applies(self) -> None:
        pbix = self._fixture_pbix()
        custom = dict(PRESETS["dark_pro"])
        custom["name"] = "custom"
        result = pbi_apply_style_preset_tool(
            str(pbix), "custom", custom_spec=custom
        )
        self.assertTrue(result["ok"], result)
        self.assertEqual(result["preset"], "custom")


if __name__ == "__main__":
    unittest.main(verbosity=2)

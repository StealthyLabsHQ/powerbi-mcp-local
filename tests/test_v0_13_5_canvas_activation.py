"""v0.13.5: canvas wallpaper + theme activation regressions.

v0.13.3 / v0.13.4 embedded the PNG and the theme JSON correctly but
left the Layout JSON un-referencing both. Power BI Desktop opened with
a blank canvas. This suite pins:

- ``image.name`` / ``image.url`` / ``image.scaling`` are bare strings
  (Literal-wrapped values are silently dropped by the renderer).
- ``image.url`` is prefixed with ``RegisteredResources/``.
- ``objects.wallpaper`` is also written so the chrome around the canvas
  matches the canvas itself.
- ``layout.activeTheme`` AND ``layout.reportSettings.activeTheme`` are
  set to the embedded theme.
- A pre-existing custom title on a visual is preserved across the
  styling pass.
- The ``wallpaper_fit`` parameter overrides the preset default.
- The post-write validation gate raises ``PowerBIValidationError``
  when any of the above is missing.
"""

from __future__ import annotations

import json
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

from tools import pbi_apply_style_preset_tool
from tools.styling import (
    CONTENT_TYPES_PART,
    LAYOUT_PART,
    patch_layout_for_wallpaper,
    patch_layout_visuals,
)
from tools.styling._presets import PRESETS


def _make_pbix(tmp: Path, name: str, sections: list[dict]) -> Path:
    layout = {"id": 0, "sections": sections, "resourcePackages": []}
    payload = json.dumps(layout, ensure_ascii=False).encode("utf-16-le")
    out = tmp / name
    with zipfile.ZipFile(out, "w") as zf:
        zf.writestr(LAYOUT_PART, payload)
        zf.writestr("DataModel", b"x" * 8192)
        zf.writestr("Metadata", b"{}")
        zf.writestr("Connections", b"{}")
        zf.writestr(
            CONTENT_TYPES_PART,
            (
                b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                b'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                b'<Default Extension="xml" ContentType="application/xml"/>'
                b"</Types>"
            ),
        )
    return out


def _read_layout(path: Path) -> dict:
    with zipfile.ZipFile(path, "r") as zf:
        return json.loads(zf.read(LAYOUT_PART).decode("utf-16-le"))


def _fixture_sections() -> list[dict]:
    return [
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
                                "objects": {
                                    "title": [
                                        {
                                            "properties": {
                                                "text": {
                                                    "expr": {
                                                        "Literal": {"Value": "'KPI principal'"}
                                                    }
                                                }
                                            }
                                        }
                                    ]
                                },
                            },
                        }
                    )
                }
            ],
        }
    ]


class WallpaperSchemaTests(unittest.TestCase):
    def test_image_block_uses_bare_strings(self) -> None:
        layout = {"sections": [{"displayName": "P1", "name": "s1"}]}
        patched, _ = patch_layout_for_wallpaper(
            layout, resource_name="bg.png", fit="Stretch"
        )
        cfg = json.loads(patched["sections"][0]["config"])
        bg = cfg["objects"]["background"][0]["properties"]
        # Bare-string contract: the v0.13.3 Literal-wrapped variant left
        # the canvas blank because PBI silently drops non-string image
        # values. Pin every key.
        self.assertEqual(bg["image"]["name"], "bg.png")
        self.assertEqual(bg["image"]["url"], "RegisteredResources/bg.png")
        self.assertEqual(bg["image"]["scaling"], "Stretch")
        self.assertEqual(bg["show"]["expr"]["Literal"]["Value"], "true")
        self.assertEqual(bg["transparency"]["expr"]["Literal"]["Value"], "0D")

    def test_wallpaper_layer_written_by_default(self) -> None:
        layout = {"sections": [{"displayName": "P1", "name": "s1"}]}
        patched, _ = patch_layout_for_wallpaper(layout, resource_name="bg.png")
        cfg = json.loads(patched["sections"][0]["config"])
        self.assertIn("wallpaper", cfg["objects"])
        wp = cfg["objects"]["wallpaper"][0]["properties"]
        self.assertEqual(wp["image"]["url"], "RegisteredResources/bg.png")

    def test_wallpaper_layer_skipped_when_disabled(self) -> None:
        layout = {"sections": [{"displayName": "P1", "name": "s1"}]}
        patched, _ = patch_layout_for_wallpaper(
            layout, resource_name="bg.png", apply_wallpaper_layer=False
        )
        cfg = json.loads(patched["sections"][0]["config"])
        self.assertNotIn("wallpaper", cfg["objects"])

    def test_fit_rejects_unknown_value(self) -> None:
        layout = {"sections": [{"displayName": "P1", "name": "s1"}]}
        with self.assertRaises(ValueError):
            patch_layout_for_wallpaper(layout, resource_name="bg.png", fit="Cover")

    def test_glassmorph_presets_default_to_stretch(self) -> None:
        for name in ("glassmorph_dark", "glassmorph_light"):
            self.assertEqual(
                PRESETS[name]["page"]["wallpaper"]["fit"], "Stretch", name
            )


class CustomTitlePreservationTests(unittest.TestCase):
    def test_existing_title_survives_styling_pass(self) -> None:
        layout = {"sections": _fixture_sections()}
        patched, counts = patch_layout_visuals(
            layout,
            preset=PRESETS["glassmorph_dark"],
            accent_picker=None,
        )
        self.assertGreaterEqual(counts["titles_preserved"], 1)
        cfg = json.loads(patched["sections"][0]["visualContainers"][0]["config"])
        title_text = cfg["singleVisual"]["objects"]["title"][0]["properties"]["text"][
            "expr"
        ]["Literal"]["Value"]
        self.assertIn("KPI principal", title_text)

    def test_empty_title_is_not_counted_as_preserved(self) -> None:
        sections = [
            {
                "displayName": "P1",
                "name": "s1",
                "visualContainers": [
                    {
                        "config": json.dumps(
                            {
                                "name": "card1",
                                "singleVisual": {
                                    "visualType": "card",
                                    "projections": {"Values": [{"queryRef": "F.X"}]},
                                    "objects": {
                                        "title": [
                                            {
                                                "properties": {
                                                    "text": {
                                                        "expr": {"Literal": {"Value": "''"}}
                                                    }
                                                }
                                            }
                                        ]
                                    },
                                },
                            }
                        )
                    }
                ],
            }
        ]
        _, counts = patch_layout_visuals(
            {"sections": sections},
            preset=PRESETS["glassmorph_dark"],
            accent_picker=None,
        )
        self.assertEqual(counts["titles_preserved"], 0)


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

    def test_apply_injects_wallpaper_reference_in_layout(self) -> None:
        pbix = _make_pbix(self.root, "in.pbix", _fixture_sections())
        result = pbi_apply_style_preset_tool(str(pbix), "glassmorph_dark")
        self.assertTrue(result["ok"], result)
        self.assertEqual(result["validation_errors"], [])
        self.assertGreaterEqual(len(result["wallpaper_applied_pages"]), 1)

        layout = _read_layout(Path(result["output_path"]))
        section_cfg = json.loads(layout["sections"][0]["config"])
        image = section_cfg["objects"]["background"][0]["properties"]["image"]
        self.assertTrue(image["url"].startswith("RegisteredResources/"))
        # The image's filename must resolve to a real archive part.
        with zipfile.ZipFile(Path(result["output_path"]), "r") as zf:
            archive_names = set(zf.namelist())
        expected_part = f"StaticResources/{image['url']}"
        self.assertIn(expected_part, archive_names)

    def test_apply_activates_theme_in_layout_root_and_settings(self) -> None:
        pbix = _make_pbix(self.root, "in.pbix", _fixture_sections())
        result = pbi_apply_style_preset_tool(str(pbix), "dark_pro")
        self.assertTrue(result["theme_activated"], result)

        layout = _read_layout(Path(result["output_path"]))
        self.assertIn("activeTheme", layout)
        active = layout["activeTheme"]
        self.assertTrue(
            active["path"].startswith("StaticResources/SharedResources/BaseThemes/")
        )
        # Also expose it via reportSettings.activeTheme for newer
        # Power BI builds that read the theme from that key.
        self.assertIn("reportSettings", layout)
        self.assertEqual(layout["reportSettings"]["activeTheme"], active)

    def test_apply_preserves_custom_titles(self) -> None:
        pbix = _make_pbix(self.root, "in.pbix", _fixture_sections())
        result = pbi_apply_style_preset_tool(str(pbix), "glassmorph_light")
        self.assertGreaterEqual(result["custom_titles_preserved"], 1)

        layout = _read_layout(Path(result["output_path"]))
        container_cfg = json.loads(
            layout["sections"][0]["visualContainers"][0]["config"]
        )
        title = container_cfg["singleVisual"]["objects"]["title"][0]["properties"][
            "text"
        ]["expr"]["Literal"]["Value"]
        self.assertIn("KPI principal", title)

    def test_wallpaper_fit_override(self) -> None:
        pbix = _make_pbix(self.root, "in.pbix", _fixture_sections())
        result = pbi_apply_style_preset_tool(
            str(pbix), "minimal_corporate", wallpaper_fit="Fill"
        )
        self.assertEqual(result["wallpaper_fit"], "Fill")
        layout = _read_layout(Path(result["output_path"]))
        cfg = json.loads(layout["sections"][0]["config"])
        self.assertEqual(
            cfg["objects"]["background"][0]["properties"]["image"]["scaling"], "Fill"
        )

    def test_wallpaper_fit_rejects_bad_value(self) -> None:
        from pbi_connection import PowerBIValidationError

        pbix = _make_pbix(self.root, "in.pbix", _fixture_sections())
        with self.assertRaises(PowerBIValidationError):
            pbi_apply_style_preset_tool(
                str(pbix), "dark_pro", wallpaper_fit="Cover"
            )

    def test_validation_gate_fires_when_layout_lacks_wallpaper(self) -> None:
        # Simulate a corrupted apply by patching out the wallpaper
        # patcher so the layout never gets the image reference. The
        # post-write validation gate must then reject the file.
        from pbi_connection import PowerBIValidationError
        from tools.styling import _apply as apply_mod
        from tools.styling import _embed as embed_mod

        original = embed_mod.patch_layout_for_wallpaper

        def _no_op(layout, **_):
            return layout, ["Overview"]

        apply_mod.patch_layout_for_wallpaper = _no_op
        try:
            pbix = _make_pbix(self.root, "in.pbix", _fixture_sections())
            with self.assertRaises(PowerBIValidationError):
                pbi_apply_style_preset_tool(str(pbix), "glassmorph_dark")
        finally:
            apply_mod.patch_layout_for_wallpaper = original


if __name__ == "__main__":
    unittest.main(verbosity=2)

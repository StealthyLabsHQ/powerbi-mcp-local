"""v0.13.4: PBIX [Content_Types].xml must declare every embedded part's
extension.

Power BI rejects a PBIX with ``MashupValidationError`` ("This file is
corrupted or was created by an unrecognized version of Power BI
Desktop") when a part is present whose extension has no matching
``Default`` entry in ``[Content_Types].xml``. The styling layer now
rewrites the manifest on every repack and the apply tool re-reads the
output as a fail-loud safety net.
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
    patch_content_types,
    validate_content_types_declarations,
)
from tools.styling._embed import _DEFAULT_CONTENT_TYPES, _required_extensions


def _make_pbix(tmp: Path, name: str, *, with_content_types: bool = True) -> Path:
    layout = {
        "id": 0,
        "sections": [
            {
                "displayName": "P1",
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
                    }
                ],
            }
        ],
        "resourcePackages": [],
    }
    out = tmp / name
    with zipfile.ZipFile(out, "w") as zf:
        zf.writestr("Report/Layout", json.dumps(layout, ensure_ascii=False).encode("utf-16-le"))
        zf.writestr("DataModel", b"x" * 8192)
        zf.writestr("Metadata", b"{}")
        zf.writestr("Connections", b"{}")
        if with_content_types:
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


class ContentTypesPatchTests(unittest.TestCase):
    def test_default_content_types_catalogue(self) -> None:
        self.assertEqual(_DEFAULT_CONTENT_TYPES["png"], "image/png")
        self.assertEqual(_DEFAULT_CONTENT_TYPES["jpg"], "image/jpeg")
        self.assertEqual(_DEFAULT_CONTENT_TYPES["jpeg"], "image/jpeg")
        self.assertEqual(_DEFAULT_CONTENT_TYPES["json"], "application/json")

    def test_required_extensions_from_part_names(self) -> None:
        exts = _required_extensions(
            [
                "Report/Layout",
                "StaticResources/RegisteredResources/bg.png",
                "StaticResources/SharedResources/BaseThemes/Theme.json",
                "DataModel",
            ]
        )
        self.assertEqual(exts, {"png", "json"})

    def test_patch_adds_missing_png_default(self) -> None:
        original = (
            b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            b'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            b'<Default Extension="xml" ContentType="application/xml"/>'
            b"</Types>"
        )
        patched, added = patch_content_types(original, {"png", "xml"})
        self.assertEqual(added, ["png"])
        self.assertIn(b'Extension="png"', patched)
        self.assertIn(b'ContentType="image/png"', patched)
        # Original xml entry must survive.
        self.assertIn(b'Extension="xml"', patched)

    def test_patch_is_no_op_when_all_declared(self) -> None:
        already = (
            b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            b'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            b'<Default Extension="png" ContentType="image/png"/>'
            b"</Types>"
        )
        patched, added = patch_content_types(already, {"png"})
        self.assertEqual(added, [])
        self.assertEqual(patched, already)

    def test_patch_handles_missing_content_types(self) -> None:
        patched, added = patch_content_types(b"", {"png", "json"})
        self.assertEqual(sorted(added), ["json", "png"])
        self.assertIn(b'<Default Extension="json" ContentType="application/json"/>', patched)
        self.assertIn(b'<Default Extension="png" ContentType="image/png"/>', patched)
        self.assertTrue(patched.endswith(b"</Types>"))

    def test_validate_reports_missing_declarations(self) -> None:
        xml = (
            b'<?xml version="1.0"?><Types xmlns="x">'
            b'<Default Extension="xml" ContentType="application/xml"/>'
            b"</Types>"
        )
        self.assertEqual(
            validate_content_types_declarations(xml, {"xml", "png", "json"}),
            ["json", "png"],
        )

    def test_validate_returns_empty_when_all_declared(self) -> None:
        xml = (
            b'<?xml version="1.0"?><Types xmlns="x">'
            b'<Default Extension="png" ContentType="image/png"/>'
            b'<Default Extension="xml" ContentType="application/xml"/>'
            b"</Types>"
        )
        self.assertEqual(
            validate_content_types_declarations(xml, {"png", "xml"}),
            [],
        )

    def test_validate_treats_missing_xml_as_all_missing(self) -> None:
        self.assertEqual(
            validate_content_types_declarations(b"", {"png", "json"}),
            ["json", "png"],
        )


class ApplyToolUpdatesContentTypesTests(unittest.TestCase):
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

    def _content_types(self, pbix_path: Path) -> bytes:
        with zipfile.ZipFile(pbix_path, "r") as zf:
            return zf.read(CONTENT_TYPES_PART)

    def test_apply_writes_png_default_entry(self) -> None:
        pbix = _make_pbix(self.root, "in.pbix", with_content_types=True)
        result = pbi_apply_style_preset_tool(str(pbix), "glassmorph_dark")
        self.assertTrue(result["ok"], result)
        self.assertEqual(result["content_types_missing"], [])
        self.assertIn("png", result["content_types_required"])
        self.assertIn("json", result["content_types_required"])

        xml = self._content_types(Path(result["output_path"]))
        self.assertIn(b'Extension="png"', xml)
        self.assertIn(b'ContentType="image/png"', xml)
        self.assertIn(b'Extension="json"', xml)

    def test_apply_creates_content_types_when_missing(self) -> None:
        pbix = _make_pbix(self.root, "no_ct.pbix", with_content_types=False)
        result = pbi_apply_style_preset_tool(str(pbix), "minimal_corporate")
        self.assertTrue(result["ok"], result)
        self.assertEqual(result["content_types_missing"], [])
        # Output must now contain a valid Content_Types part.
        xml = self._content_types(Path(result["output_path"]))
        self.assertTrue(xml.strip().startswith(b"<?xml"))
        self.assertIn(b"<Types", xml)
        self.assertIn(b'Extension="png"', xml)

    def test_apply_is_idempotent_on_content_types(self) -> None:
        pbix = _make_pbix(self.root, "in.pbix")
        first = pbi_apply_style_preset_tool(str(pbix), "dark_pro")
        second = pbi_apply_style_preset_tool(str(first["output_path"]), "dark_pro")
        self.assertEqual(second["content_types_missing"], [])

        xml = self._content_types(Path(second["output_path"]))
        # Each extension must appear exactly once — no duplicate Default
        # entries from a second apply.
        self.assertEqual(xml.count(b'Extension="png"'), 1)
        self.assertEqual(xml.count(b'Extension="json"'), 1)


if __name__ == "__main__":
    unittest.main(verbosity=2)

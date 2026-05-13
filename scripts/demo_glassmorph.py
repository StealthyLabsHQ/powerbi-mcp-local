"""Demo: apply glassmorph_dark to a fixture PBIX and print the result.

Usage:

    .venv\\Scripts\\python.exe scripts\\demo_glassmorph.py [path\\to\\fixture.pbix]

If no PBIX is supplied, a tiny synthetic one is generated under
``tmp_glassmorph_demo/fixture.pbix``. The script prints the apply tool's
full return payload (wallpaper applied pages, theme activation,
preserved titles, validation errors, DBCC) so the styling round-trip is
visible without opening Power BI Desktop.
"""

from __future__ import annotations

import argparse
import json
import sys
import zipfile
from pathlib import Path
from pprint import pprint

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "src"))

from tools import pbi_apply_style_preset_tool  # noqa: E402
from tools.styling import CONTENT_TYPES_PART, LAYOUT_PART  # noqa: E402


def _make_fixture(target: Path) -> Path:
    layout = {
        "id": 0,
        "sections": [
            {
                "displayName": "Synthese",
                "name": "ReportSection1",
                "visualContainers": [
                    {
                        "config": json.dumps(
                            {
                                "name": "kpi_marge",
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
                                                            "Literal": {
                                                                "Value": "'Marge brute (custom)'"
                                                            }
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
        ],
        "resourcePackages": [],
    }
    target.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(target, "w") as zf:
        zf.writestr(LAYOUT_PART, json.dumps(layout, ensure_ascii=False).encode("utf-16-le"))
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
    return target


def main() -> int:
    parser = argparse.ArgumentParser(description="Apply glassmorph_dark to a PBIX.")
    parser.add_argument("pbix", nargs="?", help="Input PBIX (generates a fixture if absent).")
    parser.add_argument("--preset", default="glassmorph_dark")
    parser.add_argument("--wallpaper-fit", default=None)
    parser.add_argument("--out-dir", default=str(ROOT / "tmp_glassmorph_demo"))
    args = parser.parse_args()

    out_dir = Path(args.out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    if args.pbix:
        source = Path(args.pbix).resolve()
    else:
        source = _make_fixture(out_dir / "fixture.pbix")
        print(f"[fixture] generated {source}")

    out_path = out_dir / f"{source.stem}_{args.preset}.pbix"
    out_path.write_bytes(source.read_bytes())

    result = pbi_apply_style_preset_tool(
        str(out_path),
        args.preset,
        wallpaper_fit=args.wallpaper_fit,
    )

    print("\n=== apply result ===")
    pprint({k: v for k, v in result.items() if k != "dbcc"})
    print(f"\nopen this file in Power BI Desktop: {result['output_path']}")
    return 0 if not result.get("validation_errors") else 1


if __name__ == "__main__":
    raise SystemExit(main())

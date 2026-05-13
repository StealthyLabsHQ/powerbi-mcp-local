"""Demo script for the v0.13.3 styling layer.

Materialises each built-in preset's default wallpaper, applies every
preset to a fixture PBIX, and prints a summary table. Intended for local
inspection — not part of the CI suite.

Usage:

    .venv\\Scripts\\python.exe scripts\\demo_styling.py path\\to\\fixture.pbix
"""

from __future__ import annotations

import argparse
import json
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "src"))

from tools import pbi_apply_style_preset_tool, pbi_list_style_presets_tool  # noqa: E402
from tools.styling._apply import _ensure_default_wallpaper  # noqa: E402
from tools.styling._presets import PRESETS  # noqa: E402


def _make_fixture_pbix(target: Path) -> Path:
    layout = {
        "id": 0,
        "sections": [
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
                    }
                ],
            },
            {
                "displayName": "Detail",
                "name": "ReportSection2",
                "visualContainers": [
                    {
                        "config": json.dumps(
                            {
                                "name": "card2",
                                "singleVisual": {
                                    "visualType": "card",
                                    "projections": {
                                        "Values": [{"queryRef": "F.Endettement"}]
                                    },
                                },
                            }
                        )
                    }
                ],
            },
        ],
        "resourcePackages": [],
    }
    target.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(target, "w") as zf:
        zf.writestr("Report/Layout", json.dumps(layout, ensure_ascii=False).encode("utf-16-le"))
        zf.writestr("DataModel", b"x" * 8192)
        zf.writestr("Metadata", b"{}")
        zf.writestr("Connections", b"{}")
    return target


def main() -> int:
    parser = argparse.ArgumentParser(description="Apply every style preset to a PBIX.")
    parser.add_argument("pbix", nargs="?", help="Input PBIX path (omit to use a generated fixture).")
    parser.add_argument("--out-dir", default=str(ROOT / "tmp_styling_demo"))
    args = parser.parse_args()

    out_dir = Path(args.out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    # Pre-materialise every preset's default wallpaper so the listing is
    # immediately visible on disk.
    for name, spec in PRESETS.items():
        wallpaper = _ensure_default_wallpaper(spec)
        print(f"[wallpaper] {name:<22} -> {wallpaper}")

    if args.pbix:
        source = Path(args.pbix).resolve()
    else:
        source = out_dir / "fixture.pbix"
        _make_fixture_pbix(source)
        print(f"[fixture]  generated {source}")

    catalogue = pbi_list_style_presets_tool()
    print(f"[presets]  {len(catalogue['presets'])} available")

    print()
    print(f"{'preset':<22} {'pages':<8} {'cards':<6} {'charts':<6} {'dbcc':<6} {'output':<60}")
    for preset_name in PRESETS:
        out_path = out_dir / f"{source.stem}_{preset_name}.pbix"
        # Copy the source PBIX bytes so each preset writes a fresh file.
        out_path.write_bytes(source.read_bytes())
        try:
            result = pbi_apply_style_preset_tool(str(out_path), preset_name, output_path=str(out_path))
        except Exception as exc:
            print(f"{preset_name:<22} ERROR: {exc}")
            continue
        print(
            f"{preset_name:<22} "
            f"{len(result['applied_pages']):<8} "
            f"{result['applied_visuals']['cards_styled']:<6} "
            f"{result['applied_visuals']['charts_styled']:<6} "
            f"{'OK' if result['dbcc_valid'] else 'FAIL':<6} "
            f"{result['output_path']}"
        )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

"""Native PNG helpers — no Pillow dependency.

We only need two operations:

- **inspect**: read a PNG file's IHDR chunk to validate width/height/size
  before embedding it in the PBIX.
- **synthesize**: write a small vertical-gradient PNG that ships as the
  default wallpaper for each preset, so the styling tool has something
  to embed when the caller doesn't supply ``wallpaper_path``.

PIL/Pillow is not in the runtime dependency set (would pull a ~3 MB
binary wheel) so both paths are hand-rolled around ``zlib`` and the
PNG chunk format described in the W3C spec.
"""

from __future__ import annotations

import struct
import zlib
from pathlib import Path

PNG_SIGNATURE = b"\x89PNG\r\n\x1a\n"


def _crc(chunk_type: bytes, data: bytes) -> int:
    return zlib.crc32(chunk_type + data) & 0xFFFFFFFF


def _chunk(chunk_type: bytes, data: bytes) -> bytes:
    return struct.pack(">I", len(data)) + chunk_type + data + struct.pack(">I", _crc(chunk_type, data))


def inspect_png(path: Path) -> dict:
    """Return ``{"width", "height", "size_bytes"}`` for a PNG file.

    Raises ``ValueError`` if the file is not a PNG. Avoids loading the
    pixel data — only the 24-byte IHDR is parsed.
    """
    raw = Path(path).read_bytes()
    if not raw.startswith(PNG_SIGNATURE):
        raise ValueError(f"Not a PNG file: {path}")
    # IHDR layout: 4 length + 4 'IHDR' + 13 data + 4 CRC, starting at byte 8.
    if len(raw) < 33:
        raise ValueError(f"PNG truncated: {path}")
    width, height = struct.unpack(">II", raw[16:24])
    return {"width": int(width), "height": int(height), "size_bytes": len(raw)}


def _hex_to_rgb(hex_color: str) -> tuple[int, int, int]:
    hex_color = hex_color.strip().lstrip("#")
    if len(hex_color) != 6:
        raise ValueError(f"Expected #RRGGBB, got '{hex_color}'")
    return (
        int(hex_color[0:2], 16),
        int(hex_color[2:4], 16),
        int(hex_color[4:6], 16),
    )


def write_gradient_png(
    path: Path,
    *,
    top_color: str,
    bottom_color: str,
    width: int = 1920,
    height: int = 1080,
) -> Path:
    """Write a vertical-gradient PNG to ``path``.

    Used to materialize per-preset default wallpapers. Output is RGB
    (24-bit), zlib-compressed, ~50–150 KB at 1920×1080 for typical
    palette pairs. Returns the same path for chainability.
    """
    top = _hex_to_rgb(top_color)
    bottom = _hex_to_rgb(bottom_color)

    rows = bytearray()
    denom = max(1, height - 1)
    for y in range(height):
        t = y / denom
        r = round(top[0] * (1 - t) + bottom[0] * t)
        g = round(top[1] * (1 - t) + bottom[1] * t)
        b = round(top[2] * (1 - t) + bottom[2] * t)
        rows.append(0)  # filter type = None
        row_pixel = bytes((r, g, b))
        rows.extend(row_pixel * width)

    idat = zlib.compress(bytes(rows), level=6)
    ihdr_data = struct.pack(
        ">IIBBBBB",
        width,
        height,
        8,
        2,
        0,
        0,
        0,  # bit_depth=8, color_type=2 (RGB)
    )

    buf = bytearray(PNG_SIGNATURE)
    buf += _chunk(b"IHDR", ihdr_data)
    buf += _chunk(b"IDAT", idat)
    buf += _chunk(b"IEND", b"")

    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(bytes(buf))
    return path


__all__ = ["PNG_SIGNATURE", "inspect_png", "write_gradient_png"]

#!/usr/bin/env python3
"""Generate placeholder PNG icons for the Synamail manifest.

This is intentionally tiny — real artwork lands in Sprint 4 (see
planning/APPSOURCE_CHECKLIST.md). For sideload / dev / CI we only need
valid PNGs at the correct dimensions so office-addin-manifest validates.
"""

import os
import struct
import sys
import zlib
from pathlib import Path

# Synaplan-ish brand blue (oklch ~ 0.55, hue ~250 → approximately).
BRAND = (74, 144, 226, 255)

ICON_SIZES = [16, 32, 64, 80, 128]
HERO_SIZES = [256, 512]


def png_solid(width: int, height: int, color: tuple[int, int, int, int]) -> bytes:
    """Produce a minimal PNG byte stream of solid colour."""

    def chunk(typ: bytes, data: bytes) -> bytes:
        return (
            struct.pack(">I", len(data))
            + typ
            + data
            + struct.pack(">I", zlib.crc32(typ + data) & 0xFFFFFFFF)
        )

    sig = b"\x89PNG\r\n\x1a\n"
    ihdr = chunk(b"IHDR", struct.pack(">IIBBBBB", width, height, 8, 6, 0, 0, 0))
    row = bytes(color) * width
    raw = b"".join(b"\x00" + row for _ in range(height))
    idat = chunk(b"IDAT", zlib.compress(raw, level=9))
    iend = chunk(b"IEND", b"")
    return sig + ihdr + idat + iend


def main() -> int:
    repo_root = Path(__file__).resolve().parent.parent
    assets = repo_root / "assets"
    store = assets / "store"
    assets.mkdir(parents=True, exist_ok=True)
    store.mkdir(parents=True, exist_ok=True)

    for size in ICON_SIZES:
        (assets / f"icon-{size}.png").write_bytes(png_solid(size, size, BRAND))
        print(f"  wrote assets/icon-{size}.png ({size}x{size})")

    for size in HERO_SIZES:
        (store / f"hero-{size}.png").write_bytes(png_solid(size, size, BRAND))
        print(f"  wrote assets/store/hero-{size}.png ({size}x{size})")

    return 0


if __name__ == "__main__":
    sys.exit(main())

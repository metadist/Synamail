#!/usr/bin/env python3
"""Generate Synamail PNG icons from the Synaplan parrot SVG.

Renders the parrot from `assets/source/parrot.svg` onto a cream-coloured
square "button" at every manifest-required size. The placeholders this
script used to emit (solid brand-blue squares) are now replaced by branded
artwork suitable for sideload, OWA "Get add-ins" listings, and AppSource
submission (Sprint 4 will still want store hero retouching, but the icons
themselves are real).

Re-run after editing the source SVG, the background colour, or the padding.

    python3 scripts/generate-placeholder-icons.py

Dependencies: ImageMagick `convert` (commonly `apt install imagemagick`).
"""

import shutil
import subprocess
import sys
from pathlib import Path

# Warm off-white, like card stock. Provides a friendly "button" feel against
# Outlook's white/grey chrome without being pure white (which disappears).
BG_COLOR = "#F5F2EA"

# Fraction of canvas height the parrot occupies. 64% leaves a comfortable
# margin all around without making the bird tiny at 16 px.
BIRD_HEIGHT_PCT = 64

ICON_SIZES = (16, 32, 64, 80, 128)
HERO_SIZES = (256, 512)


def render(svg: Path, png: Path, size: int) -> None:
    """Rasterize `svg` onto a `size`×`size` cream-coloured PNG at `png`."""
    inner = max(int(size * BIRD_HEIGHT_PCT / 100), min(size - 2, 12))
    cmd = [
        "convert",
        "-background", "none",
        "-density", "384",
        str(svg),
        "-resize", f"x{inner}",
        "-background", BG_COLOR,
        "-gravity", "center",
        "-extent", f"{size}x{size}",
        f"PNG32:{png}",
    ]
    subprocess.run(cmd, check=True)


def main() -> int:
    if shutil.which("convert") is None:
        print(
            "error: ImageMagick `convert` not found. Install with `apt install imagemagick`.",
            file=sys.stderr,
        )
        return 1

    repo_root = Path(__file__).resolve().parent.parent
    src = repo_root / "assets" / "source" / "parrot.svg"
    if not src.exists():
        print(f"error: source SVG missing at {src}", file=sys.stderr)
        return 1

    assets = repo_root / "assets"
    store = assets / "store"
    assets.mkdir(parents=True, exist_ok=True)
    store.mkdir(parents=True, exist_ok=True)

    for size in ICON_SIZES:
        out = assets / f"icon-{size}.png"
        render(src, out, size)
        print(f"  wrote {out.relative_to(repo_root)} ({size}x{size}, {out.stat().st_size} B)")

    for size in HERO_SIZES:
        out = store / f"hero-{size}.png"
        render(src, out, size)
        print(f"  wrote {out.relative_to(repo_root)} ({size}x{size}, {out.stat().st_size} B)")

    return 0


if __name__ == "__main__":
    sys.exit(main())

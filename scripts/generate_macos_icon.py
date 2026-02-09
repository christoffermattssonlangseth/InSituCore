#!/usr/bin/env python3
"""Generate macOS-style icon assets for InSituCore.

Primary behavior:
- If `assets/icon.png` exists, use it as design source and apply a macOS polish pass.
- Otherwise, fall back to a simple generated vector-style icon.
"""

from __future__ import annotations

import argparse
import shutil
import subprocess
from pathlib import Path

from PIL import Image, ImageChops, ImageDraw, ImageEnhance, ImageFilter


ROOT = Path(__file__).resolve().parents[1]
ASSETS = ROOT / "assets"
SOURCE_PNG = ASSETS / "icon.png"
BASE_PNG = ASSETS / "app_icon_1024.png"
ICNS_PATH = ASSETS / "InSituCore.icns"
ICONSET_DIR = ASSETS / "InSituCore.iconset"


def _lerp(a: int, b: int, t: float) -> int:
    return int(round(a + (b - a) * t))


def _rounded_mask(size: int, radius: int) -> Image.Image:
    mask = Image.new("L", (size, size), 0)
    draw = ImageDraw.Draw(mask)
    draw.rounded_rectangle((0, 0, size - 1, size - 1), radius=radius, fill=255)
    return mask


def _apply_macos_shell(icon: Image.Image) -> Image.Image:
    size = icon.width
    mask = _rounded_mask(size, int(size * 0.225))
    shell = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    shell.paste(icon, (0, 0), mask)

    # Soft top sheen.
    sheen = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    d = ImageDraw.Draw(sheen)
    d.ellipse(
        (-int(size * 0.1), -int(size * 0.42), int(size * 1.1), int(size * 0.56)),
        fill=(255, 255, 255, 20),
    )
    sheen = sheen.filter(ImageFilter.GaussianBlur(radius=size * 0.03))
    shell.alpha_composite(sheen)

    # Inner edge vignette only, avoids any outside halo/rectangle artifacts.
    inner = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    idr = ImageDraw.Draw(inner)
    idr.rounded_rectangle(
        (int(size * 0.012), int(size * 0.012), int(size * 0.988), int(size * 0.988)),
        radius=int(size * 0.225),
        outline=(0, 0, 0, 32),
        width=max(2, int(size * 0.006)),
    )
    inner = inner.filter(ImageFilter.GaussianBlur(radius=size * 0.008))
    inner_masked = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    inner_masked.paste(inner, (0, 0), mask)
    shell.alpha_composite(inner_masked)
    return shell


def _fallback_icon(size: int = 1024) -> Image.Image:
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    px = img.load()
    top = (53, 71, 87)
    bottom = (29, 40, 53)
    for y in range(size):
        t = y / (size - 1)
        for x in range(size):
            px[x, y] = (_lerp(top[0], bottom[0], t), _lerp(top[1], bottom[1], t), _lerp(top[2], bottom[2], t), 255)

    draw = ImageDraw.Draw(img)
    cx, cy = size // 2, size // 2
    ring_r = int(size * 0.235)
    draw.ellipse((cx - ring_r, cy - ring_r, cx + ring_r, cy + ring_r), outline=(236, 243, 247, 230), width=int(size * 0.06))
    core_r = int(size * 0.105)
    draw.ellipse((cx - core_r, cy - core_r, cx + core_r, cy + core_r), fill=(223, 236, 242, 232))
    return _apply_macos_shell(img)


def _trim_near_black(src: Image.Image, threshold: int = 10) -> Image.Image:
    rgb = src.convert("RGB")
    r, g, b = rgb.split()
    max_channel = ImageChops.lighter(ImageChops.lighter(r, g), b)
    mask = max_channel.point(lambda p: 255 if p > threshold else 0)
    bbox = mask.getbbox()
    if bbox:
        return src.crop(bbox)
    return src


def _icon_from_source(source_path: Path, size: int = 1024) -> Image.Image:
    src = Image.open(source_path).convert("RGBA")
    src = _trim_near_black(src, threshold=8)

    # Tone down neon look toward a softer native app icon profile.
    src = ImageEnhance.Color(src).enhance(0.78)
    src = ImageEnhance.Contrast(src).enhance(0.92)
    src = ImageEnhance.Brightness(src).enhance(0.95)

    # Fill the tile directly to avoid adding dark framing bars around the source art.
    src = src.resize((size, size), Image.Resampling.LANCZOS)

    # Very light soft-focus pass to calm high-frequency glow artifacts.
    blur = src.filter(ImageFilter.GaussianBlur(radius=size * 0.0025))
    src = Image.blend(src, blur, alpha=0.2)

    canvas = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    canvas.alpha_composite(src, dest=(0, 0))

    # Light unifying tint and gentle milk layer for a macOS-like polish.
    tint = Image.new("RGBA", (size, size), (24, 35, 56, 28))
    canvas.alpha_composite(tint)
    milk = Image.new("RGBA", (size, size), (255, 255, 255, 12))
    canvas.alpha_composite(milk)

    return _apply_macos_shell(canvas)


def write_iconset(base_img: Image.Image) -> None:
    if ICONSET_DIR.exists():
        shutil.rmtree(ICONSET_DIR)
    ICONSET_DIR.mkdir(parents=True, exist_ok=True)

    icon_sizes = [16, 32, 128, 256, 512]
    for size in icon_sizes:
        base_img.resize((size, size), Image.Resampling.LANCZOS).save(ICONSET_DIR / f"icon_{size}x{size}.png")
        double = size * 2
        base_img.resize((double, double), Image.Resampling.LANCZOS).save(ICONSET_DIR / f"icon_{size}x{size}@2x.png")


def build_icns() -> None:
    # Try iconutil first; fallback to Pillow ICNS writer.
    if shutil.which("iconutil") is not None:
        if ICNS_PATH.exists():
            ICNS_PATH.unlink()
        try:
            subprocess.run(
                ["iconutil", "--convert", "icns", str(ICONSET_DIR), "--output", str(ICNS_PATH)],
                check=True,
            )
            return
        except subprocess.CalledProcessError:
            pass

    base = Image.open(BASE_PNG).convert("RGBA")
    if ICNS_PATH.exists():
        ICNS_PATH.unlink()
    base.save(
        ICNS_PATH,
        format="ICNS",
        sizes=[(16, 16), (32, 32), (64, 64), (128, 128), (256, 256), (512, 512), (1024, 1024)],
    )


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Generate macOS app icon assets.")
    parser.add_argument(
        "--source",
        default=str(SOURCE_PNG),
        help="Path to source PNG to riff on (default: assets/icon.png).",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    ASSETS.mkdir(parents=True, exist_ok=True)

    source_path = Path(args.source).expanduser().resolve()
    if source_path.exists():
        base = _icon_from_source(source_path, size=1024)
        print(f"Using source icon: {source_path}")
    else:
        base = _fallback_icon(1024)
        print(f"Source icon not found at {source_path}; using fallback icon.")

    base.save(BASE_PNG)
    write_iconset(base)
    build_icns()
    print(f"Wrote {BASE_PNG}")
    print(f"Wrote {ICONSET_DIR}")
    if ICNS_PATH.exists():
        print(f"Wrote {ICNS_PATH}")


if __name__ == "__main__":
    main()

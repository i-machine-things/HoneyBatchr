#!/usr/bin/env python3
"""
Honey Batchr - Icon Generator
Amber circle with white "HB" initials.
"""

import os
from PIL import Image, ImageDraw, ImageFont

OUTPUT_DIR = "resources"

AMBER = (230, 160, 30, 255)
WHITE = (255, 255, 255, 255)

# Try to load a bold system font; fall back to PIL default
FONT_CANDIDATES = [
    "C:/Windows/Fonts/arialbd.ttf",
    "C:/Windows/Fonts/segoeuib.ttf",
    "C:/Windows/Fonts/calibrib.ttf",
]


def get_font(size):
    for path in FONT_CANDIDATES:
        if os.path.exists(path):
            return ImageFont.truetype(path, size)
    return ImageFont.load_default()


def draw_icon(size: int) -> Image.Image:
    img  = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)

    # Amber circle
    pad = max(1, size // 16)
    draw.ellipse([pad, pad, size - pad - 1, size - pad - 1], fill=AMBER)

    # "HB" text centred in the circle
    font_size = int(size * 0.38)
    font = get_font(font_size)

    bbox = draw.textbbox((0, 0), "HB", font=font)
    tw = bbox[2] - bbox[0]
    th = bbox[3] - bbox[1]
    tx = (size - tw) // 2 - bbox[0]
    ty = (size - th) // 2 - bbox[1]

    draw.text((tx, ty), "HB", fill=WHITE, font=font)
    return img


def main() -> None:
    print("Generating Honey Batchr icons...")
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    native_sizes = [256, 128, 64, 48, 32]
    images = {}

    for size in native_sizes:
        images[size] = draw_icon(size)
        path = os.path.join(OUTPUT_DIR, f"badger_{size}.png")
        images[size].save(path, "PNG")
        print(f"  [{size:>3}x{size:<3}] {path}")

    images[16] = images[32].resize((16, 16), Image.LANCZOS)
    images[16].save(os.path.join(OUTPUT_DIR, "badger_16.png"), "PNG")
    print(f"  [ 16x16 ] resources/badger_16.png")

    images[256].save(os.path.join(OUTPUT_DIR, "badger.png"), "PNG")
    print(f"  [main   ] resources/badger.png")

    images[256].save(
        os.path.join(OUTPUT_DIR, "badger.ico"),
        format="ICO",
        sizes=[(16, 16), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)],
    )
    print(f"  [ico    ] resources/badger.ico")
    print("\nDone!")


if __name__ == "__main__":
    main()

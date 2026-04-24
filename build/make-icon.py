"""
Generate build/icon.png for electron-builder.

Design:
- Rounded-square blue gradient background (Material Blue 600 → 900)
- Centered pointy-top benzene hexagon in white, with an inner aromatic
  ring (Kekulé circle) so it reads as "chemistry" at any size.

Strategy: render at 4× resolution, then downsample with LANCZOS so lines
stay crisp and antialiased without needing a vector rasterizer.

Run: python3 build/make-icon.py
Output: build/icon.png (1024×1024). electron-builder picks this up
automatically and generates .icns / .ico / Linux PNGs on package.
"""

import math
import os
from PIL import Image, ImageDraw

HERE = os.path.dirname(os.path.abspath(__file__))
OUT_PATH = os.path.join(HERE, "icon.png")

OUT_SIZE = 1024
SCALE = 4
SIZE = OUT_SIZE * SCALE

# Palette — Material Blue 600 → 900, strong but neutral "serious app" blue
TOP_COLOR = (30, 136, 229)
BOT_COLOR = (13, 71, 161)
WHITE = (255, 255, 255, 255)

# Geometry (in final-resolution pixels; multiply by SCALE for supersampled canvas)
CORNER_RADIUS = 220          # matches iOS / macOS "squircle-ish" feel
HEX_RADIUS = 300             # distance from center to hexagon vertex
HEX_STROKE = 36              # thickness of hexagon edges
RING_RADIUS = 185            # inner aromatic ring
RING_STROKE = 24             # thickness of the inner ring


def render():
    # --- 1. Vertical gradient filling the full canvas -------------------
    gradient = Image.new("RGBA", (SIZE, SIZE))
    gdraw = ImageDraw.Draw(gradient)
    for y in range(SIZE):
        t = y / (SIZE - 1)
        r = round(TOP_COLOR[0] * (1 - t) + BOT_COLOR[0] * t)
        g = round(TOP_COLOR[1] * (1 - t) + BOT_COLOR[1] * t)
        b = round(TOP_COLOR[2] * (1 - t) + BOT_COLOR[2] * t)
        gdraw.line([(0, y), (SIZE, y)], fill=(r, g, b, 255))

    # --- 2. Mask to a rounded square ------------------------------------
    mask = Image.new("L", (SIZE, SIZE), 0)
    ImageDraw.Draw(mask).rounded_rectangle(
        [0, 0, SIZE - 1, SIZE - 1],
        radius=CORNER_RADIUS * SCALE,
        fill=255,
    )
    bg = Image.new("RGBA", (SIZE, SIZE), (0, 0, 0, 0))
    bg.paste(gradient, (0, 0), mask)

    # --- 3. Pointy-top benzene hexagon ----------------------------------
    cx, cy = SIZE // 2, SIZE // 2
    R = HEX_RADIUS * SCALE
    vertices = []
    for i in range(6):
        angle = math.radians(90 + i * 60)   # top, upper-left, lower-left, bottom, lower-right, upper-right
        x = cx + R * math.cos(angle)
        y = cy - R * math.sin(angle)
        vertices.append((x, y))

    draw = ImageDraw.Draw(bg)
    draw.polygon(vertices, outline=WHITE, width=HEX_STROKE * SCALE)

    # --- 4. Inner aromatic ring (Kekulé circle) -------------------------
    rr = RING_RADIUS * SCALE
    draw.ellipse(
        [cx - rr, cy - rr, cx + rr, cy + rr],
        outline=WHITE,
        width=RING_STROKE * SCALE,
    )

    # --- 5. Downsample to 1024×1024 -------------------------------------
    final = bg.resize((OUT_SIZE, OUT_SIZE), Image.LANCZOS)
    final.save(OUT_PATH, "PNG", optimize=True)
    print(f"Wrote {OUT_PATH}  ({OUT_SIZE}×{OUT_SIZE})")


if __name__ == "__main__":
    render()

"""Generate icon design variants for review.
Each variant is saved at 128px for easy comparison.
"""
import os
from PIL import Image, ImageDraw, ImageFont

NAVY      = (31, 58, 110)   # #1F3A6E  FINOVATE primary
NAVY_DK   = (15, 33, 70)
NAVY_LT   = (45, 80, 150)
GREEN     = (38, 166, 91)   # #26A65B
GREEN_LT  = (76, 217, 100)
WHITE     = (255, 255, 255)
TEAL      = (20, 130, 130)
GOLD      = (230, 165, 35)
CHARCOAL  = (38, 50, 56)

SIZE = 128
OUT_DIR = os.path.join(os.path.dirname(__file__), "..", "assets", "variants")

FONT_CANDIDATES = [
    "C:/Windows/Fonts/arialbd.ttf",
    "C:/Windows/Fonts/segoeuib.ttf",
    "C:/Windows/Fonts/calibrib.ttf",
]
def get_font(px):
    for p in FONT_CANDIDATES:
        if os.path.exists(p):
            return ImageFont.truetype(p, px)
    return ImageFont.load_default()


def base_canvas(bg=WHITE):
    img = Image.new("RGBA", (SIZE, SIZE), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    radius = int(SIZE * 0.14)
    draw.rounded_rectangle([(0, 0), (SIZE - 1, SIZE - 1)], radius=radius, fill=bg)
    return img, draw


def draw_diamond(draw, color, line_w, scale=0.34):
    cx, cy = SIZE / 2, SIZE / 2
    r = SIZE * scale
    points = [(cx, cy - r), (cx + r, cy), (cx, cy + r), (cx - r, cy)]
    draw.polygon(points, outline=color, width=line_w)


def draw_bars(draw, color, baseline_y_pct=0.62, max_h_pct=0.26, bar_w_pct=0.075, count=3):
    cx = SIZE / 2
    bar_w = int(SIZE * bar_w_pct)
    bar_gap = int(SIZE * 0.035)
    total_w = count * bar_w + (count - 1) * bar_gap
    start_x = cx - total_w / 2
    baseline_y = int(SIZE * baseline_y_pct)
    max_h = int(SIZE * max_h_pct)
    radius_bar = max(1, int(bar_w / 2.5))
    for i in range(count):
        h = int(max_h * (i + 1) / count)
        x = int(start_x + i * (bar_w + bar_gap))
        draw.rounded_rectangle(
            [(x, baseline_y - h), (x + bar_w, baseline_y)],
            radius=radius_bar, fill=color
        )


def draw_letter(draw, letter, color, size_pct=0.7, dy=0):
    font = get_font(int(SIZE * size_pct))
    bbox = draw.textbbox((0, 0), letter, font=font)
    tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
    x = (SIZE - tw) / 2 - bbox[0]
    y = (SIZE - th) / 2 - bbox[1] + dy
    draw.text((x, y), letter, fill=color, font=font)


# ─────────── VARIANTS ───────────
variants = {}

# V1: Current — navy bg + white diamond + green bars
img, draw = base_canvas(NAVY)
draw_diamond(draw, WHITE, line_w=7, scale=0.34)
draw_bars(draw, GREEN, baseline_y_pct=0.62)
variants["v1_navy_diamond_bars"] = img

# V2: Navy bg + bold white "P" letter
img, draw = base_canvas(NAVY)
draw_letter(draw, "P", WHITE, size_pct=0.72, dy=-2)
variants["v2_navy_P"] = img

# V3: Navy bg + bold white "F" letter (FINOVATE)
img, draw = base_canvas(NAVY)
draw_letter(draw, "F", WHITE, size_pct=0.72, dy=-2)
variants["v3_navy_F"] = img

# V4: White bg + navy diamond + green bars (logo-faithful)
img, draw = base_canvas(WHITE)
draw_diamond(draw, NAVY, line_w=8, scale=0.36)
draw_bars(draw, GREEN, baseline_y_pct=0.62)
variants["v4_white_bg_navy_mark"] = img

# V5: Navy gradient bg + white diamond + green bars
img = Image.new("RGBA", (SIZE, SIZE), (0, 0, 0, 0))
draw = ImageDraw.Draw(img)
# gradient by drawing horizontal lines
for y in range(SIZE):
    t = y / SIZE
    r = int(NAVY_DK[0] * (1 - t) + NAVY_LT[0] * t)
    g = int(NAVY_DK[1] * (1 - t) + NAVY_LT[1] * t)
    b = int(NAVY_DK[2] * (1 - t) + NAVY_LT[2] * t)
    draw.line([(0, y), (SIZE, y)], fill=(r, g, b))
# clip to rounded rect via mask
mask = Image.new("L", (SIZE, SIZE), 0)
ImageDraw.Draw(mask).rounded_rectangle([(0, 0), (SIZE - 1, SIZE - 1)], radius=int(SIZE * 0.14), fill=255)
img.putalpha(mask)
draw = ImageDraw.Draw(img)
draw_diamond(draw, WHITE, line_w=7, scale=0.34)
draw_bars(draw, GREEN_LT, baseline_y_pct=0.62)
variants["v5_navy_gradient"] = img

# V6: Filled white diamond + navy bars + green outer accent
img, draw = base_canvas(NAVY)
# filled white diamond
cx, cy = SIZE / 2, SIZE / 2
r = SIZE * 0.36
draw.polygon([(cx, cy - r), (cx + r, cy), (cx, cy + r), (cx - r, cy)], fill=WHITE)
# green bars inside white diamond
draw_bars(draw, GREEN, baseline_y_pct=0.62, max_h_pct=0.24)
variants["v6_filled_diamond"] = img

# V7: White bg + navy filled diamond + WHITE bars (clean inverse)
img, draw = base_canvas(WHITE)
cx, cy = SIZE / 2, SIZE / 2
r = SIZE * 0.36
draw.polygon([(cx, cy - r), (cx + r, cy), (cx, cy + r), (cx - r, cy)], fill=NAVY)
draw_bars(draw, GREEN, baseline_y_pct=0.62, max_h_pct=0.24)
variants["v7_white_bg_filled"] = img

# V8: Charcoal bg (dark grey) + gold diamond + green bars (premium feel)
img, draw = base_canvas(CHARCOAL)
draw_diamond(draw, GOLD, line_w=7, scale=0.34)
draw_bars(draw, GREEN_LT, baseline_y_pct=0.62)
variants["v8_charcoal_gold"] = img

# V9: Teal bg + white diamond + bars (alt color scheme)
img, draw = base_canvas(TEAL)
draw_diamond(draw, WHITE, line_w=7, scale=0.34)
draw_bars(draw, GOLD, baseline_y_pct=0.62)
variants["v9_teal_bars"] = img

# V10: Navy bg + "F" letter + small green chart underline accent
img, draw = base_canvas(NAVY)
draw_letter(draw, "F", WHITE, size_pct=0.65, dy=-10)
# tiny green bars underneath F
draw_bars(draw, GREEN, baseline_y_pct=0.85, max_h_pct=0.10, bar_w_pct=0.05)
variants["v10_F_with_bars"] = img


os.makedirs(OUT_DIR, exist_ok=True)
for name, img in variants.items():
    out = os.path.join(OUT_DIR, f"{name}.png")
    img.save(out, optimize=True)
    print(f"OK: {out}")

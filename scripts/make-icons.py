"""Generate Project Update add-in icons (FINOVATE branded).

Design: navy rounded square + bold white 'P' letter + small green
ascending bars below the letter (the FINOVATE chart accent).
At very small sizes (16 / 32) the bars are dropped — only the bold P
stays so it doesn't turn into unreadable noise.
"""
import os
from PIL import Image, ImageDraw, ImageFont

NAVY  = (31, 58, 110)   # #1F3A6E  — FINOVATE primary
GREEN = (38, 166, 91)   # #26A65B  — chart accent
WHITE = (255, 255, 255)

SIZES = [16, 32, 64, 80, 128]
LETTER = "P"
OUT_DIR = os.path.join(os.path.dirname(__file__), "..", "assets")

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


def make_icon(size: int) -> Image.Image:
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)

    # Rounded navy square
    radius = max(2, int(size * 0.14))
    draw.rounded_rectangle([(0, 0), (size - 1, size - 1)], radius=radius, fill=NAVY)

    cx = size / 2

    # Small icons: just a big centred letter (bars would be invisible specks)
    if size < 48:
        font = get_font(int(size * 0.72))
        bbox = draw.textbbox((0, 0), LETTER, font=font)
        tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
        x = (size - tw) / 2 - bbox[0]
        y = (size - th) / 2 - bbox[1]
        draw.text((x, y), LETTER, fill=WHITE, font=font)
        return img

    # Larger icons: letter shifted up + tiny green ascending bars beneath
    letter_size = int(size * 0.62)
    font = get_font(letter_size)
    bbox = draw.textbbox((0, 0), LETTER, font=font)
    tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
    x = (size - tw) / 2 - bbox[0]
    y = (size - th) / 2 - bbox[1] - int(size * 0.085)  # nudge up to make room for bars
    draw.text((x, y), LETTER, fill=WHITE, font=font)

    bar_count = 3
    bar_w = max(2, int(size * 0.055))
    bar_gap = max(1, int(size * 0.025))
    total_w = bar_count * bar_w + (bar_count - 1) * bar_gap
    start_x = cx - total_w / 2
    baseline_y = int(size * 0.86)
    max_h = int(size * 0.13)
    radius_bar = max(1, int(bar_w / 2.5))
    for i in range(bar_count):
        h = int(max_h * (i + 1) / bar_count)
        xb = int(start_x + i * (bar_w + bar_gap))
        draw.rounded_rectangle(
            [(xb, baseline_y - h), (xb + bar_w, baseline_y)],
            radius=radius_bar, fill=GREEN
        )

    return img


def main() -> None:
    os.makedirs(OUT_DIR, exist_ok=True)
    for size in SIZES:
        out_path = os.path.join(OUT_DIR, f"icon-{size}.png")
        make_icon(size).save(out_path, optimize=True)
        print(f"OK: {out_path}  ({size}x{size})")


if __name__ == "__main__":
    main()

"""Draws the CarWatch app icons (site/icon-*.png): a gauge whose centre is the centre of the icon, so it stays
centred under any launcher mask (circle, squircle, rounded square). Run: python3 tools/make_icons.py"""
import math
import os
from PIL import Image, ImageDraw

BG, GREEN, GREY, WHITE = (15, 19, 24), (60, 207, 142), (43, 51, 62), (238, 241, 244)
SITE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "site")


def gauge(size, scale=1.0, ss=4):
    """The icon at `size` px; `scale` shrinks the gauge (maskable icons keep it well inside the safe zone)."""
    n = size * ss
    im = Image.new("RGB", (n, n), BG)
    d = ImageDraw.Draw(im)
    c = n / 2                                           # the gauge's centre is the icon's centre
    r_out, stroke = 0.33 * n * scale, 0.084 * n * scale  # outer radius and arc width (as the original design)
    box = [c - r_out, c - r_out, c + r_out, c + r_out]
    split = 341                                         # where green turns grey: 19 degrees above 3 o'clock
    d.arc(box, 135, split, fill=GREEN, width=round(stroke))
    d.arc(box, split, 405, fill=GREY, width=round(stroke))
    a = math.radians(16.5)                              # the needle, just short of the arc
    length = 0.254 * n * scale
    d.line([c, c, c + length * math.cos(a), c - length * math.sin(a)], fill=WHITE, width=round(0.024 * n * scale))
    pr = 0.042 * n * scale
    d.ellipse([c - pr, c - pr, c + pr, c + pr], fill=WHITE)
    return im.resize((size, size), Image.LANCZOS)


if __name__ == "__main__":
    for s in (192, 512):
        gauge(s).save(os.path.join(SITE, f"icon-{s}.png"), optimize=True)
        gauge(s, 0.9).save(os.path.join(SITE, f"icon-maskable-{s}.png"), optimize=True)
    print("icons written to", os.path.normpath(SITE))

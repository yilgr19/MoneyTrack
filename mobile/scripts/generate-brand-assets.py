"""
Genera icon.png, adaptive-icon.png, splash-icon.png y notification-icon.png
con la marca minimalista MoneyTrack (barras + línea oro). No depende del PNG antiguo.
"""
from __future__ import annotations

import os
import sys

try:
    from PIL import Image, ImageDraw
except ImportError:
    print("Instala Pillow: pip install Pillow", file=sys.stderr)
    sys.exit(1)

ROOT = os.path.join(os.path.dirname(__file__), "..", "assets")
CANVAS = 1024
NOTIF_SIZE = 96

# theme.js
BG_ADAPTIVE = (30, 11, 46, 255)  # #1e0b2e
BG_SPLASH = (12, 8, 18, 255)  # #0c0812
MINT = (125, 193, 145, 255)
CHART_BLUE = (167, 216, 222, 255)
GOLD = (217, 180, 74, 255)


def _draw_mark(draw: ImageDraw.ImageDraw, u: float, ox: float, oy: float) -> None:
    """Coordenadas en espacio 0..100, escaladas por u y desplazadas (ox, oy)."""
    y_mid = 71 * u + oy
    th = max(2, int(1.25 * u))
    xa, xb = int(20 * u + ox), int(80 * u + ox)
    draw.rectangle([xa, int(y_mid - th / 2), xb, int(y_mid + th / 2)], fill=GOLD)

    bars = [
        (23, 50, 11, 21, MINT),
        (44.5, 40, 11, 31, CHART_BLUE),
        (66, 30, 11, 41, MINT),
    ]
    r = max(2, int(2.2 * u))
    for bx, by, bw, bh, fill in bars:
        x0 = int(bx * u + ox)
        y0 = int(by * u + oy)
        x1b = int((bx + bw) * u + ox)
        y1b = int((by + bh) * u + oy)
        draw.rounded_rectangle([x0, y0, x1b, y1b], radius=r, fill=fill)


def draw_icon_canvas(bg: tuple[int, int, int, int]) -> Image.Image:
    img = Image.new("RGBA", (CANVAS, CANVAS), bg)
    d = ImageDraw.Draw(img)
    u = CANVAS / 100.0
    _draw_mark(d, u, 0, 0)
    return img


def draw_notification_icon() -> Image.Image:
    """Silueta blanca sobre transparente (Android status bar)."""
    n = Image.new("RGBA", (NOTIF_SIZE, NOTIF_SIZE), (0, 0, 0, 0))
    d = ImageDraw.Draw(n)
    u = NOTIF_SIZE / 100.0
    white = (255, 255, 255, 255)
    y_mid = 71 * u
    th = max(1, int(1.25 * u))
    xa, xb = int(20 * u), int(80 * u)
    d.rectangle([xa, int(y_mid - th / 2), xb, int(y_mid + th / 2)], fill=white)
    bars = [(23, 50, 11, 21), (44.5, 40, 11, 31), (66, 30, 11, 41)]
    r = max(1, int(2.2 * u))
    for bx, by, bw, bh in bars:
        x0 = int(bx * u)
        y0 = int(by * u)
        x1b = int((bx + bw) * u)
        y1b = int((by + bh) * u)
        d.rounded_rectangle([x0, y0, x1b, y1b], radius=r, fill=white)
    return n


def main() -> None:
    os.makedirs(ROOT, exist_ok=True)

    adaptive = draw_icon_canvas(BG_ADAPTIVE)
    splash = draw_icon_canvas(BG_SPLASH)
    icon = adaptive.copy()

    adaptive.save(os.path.join(ROOT, "adaptive-icon.png"), "PNG")
    icon.save(os.path.join(ROOT, "icon.png"), "PNG")
    splash.save(os.path.join(ROOT, "splash-icon.png"), "PNG")

    notif = draw_notification_icon()
    notif.save(os.path.join(ROOT, "notification-icon.png"), "PNG")

    print("OK: icon.png, adaptive-icon.png, splash-icon.png, notification-icon.png (marca minimalista)")


if __name__ == "__main__":
    main()

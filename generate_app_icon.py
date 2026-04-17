from __future__ import annotations

from pathlib import Path
from PIL import Image, ImageDraw, ImageFilter, ImageFont


ROOT = Path(__file__).resolve().parent
ICON_DIR = ROOT / "assets" / "icon.iconset"
ICNS_PATH = ROOT / "assets" / "audio_maintenance_tool.icns"
PNG_PATH = ROOT / "assets" / "audio_maintenance_tool_1024.png"
ICO_PATH = ROOT / "assets" / "audio_maintenance_tool.ico"


def _rounded_panel(draw: ImageDraw.ImageDraw, box: tuple[int, int, int, int], radius: int, fill: tuple[int, int, int, int]) -> None:
    draw.rounded_rectangle(box, radius=radius, fill=fill)


def _make_base_canvas(size: int = 1024) -> Image.Image:
    image = Image.new("RGBA", (size, size), (14, 18, 28, 255))
    draw = ImageDraw.Draw(image)

    for y in range(size):
        blend = y / max(size - 1, 1)
        r = int(20 + 16 * blend)
        g = int(28 + 22 * blend)
        b = int(46 + 40 * blend)
        draw.line((0, y, size, y), fill=(r, g, b, 255))

    glow = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    glow_draw = ImageDraw.Draw(glow)
    glow_draw.ellipse((90, 120, 910, 850), fill=(34, 180, 255, 70))
    glow_draw.ellipse((250, 260, 960, 980), fill=(255, 170, 72, 60))
    glow = glow.filter(ImageFilter.GaussianBlur(80))
    image.alpha_composite(glow)

    card_box = (96, 96, 928, 928)
    shadow = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    shadow_draw = ImageDraw.Draw(shadow)
    shadow_draw.rounded_rectangle((110, 126, 930, 946), radius=180, fill=(0, 0, 0, 170))
    shadow = shadow.filter(ImageFilter.GaussianBlur(28))
    image.alpha_composite(shadow)

    draw = ImageDraw.Draw(image)
    _rounded_panel(draw, card_box, 168, (32, 39, 55, 245))
    _rounded_panel(draw, (128, 128, 896, 896), 150, (40, 49, 69, 235))
    return image


def _draw_waveform(draw: ImageDraw.ImageDraw, box: tuple[int, int, int, int]) -> None:
    left, top, right, bottom = box
    mid_y = (top + bottom) // 2
    width = right - left
    center = left + width * 0.5

    draw.rounded_rectangle(box, radius=52, outline=(79, 92, 120, 180), width=4, fill=(24, 29, 42, 150))
    draw.line((left + 28, mid_y, right - 28, mid_y), fill=(81, 92, 120, 90), width=2)

    points: list[tuple[float, float]] = []
    for x in range(left + 30, right - 30, 8):
        normalized = (x - center) / max(width * 0.34, 1)
        envelope = max(0.16, 1.0 - min(1.0, abs(normalized)) ** 1.4)
        pulse = abs((normalized * 7.2) % 2 - 1)
        ridge = 0.28 + 0.72 * (1.0 - pulse)
        amplitude = envelope * ridge
        y = mid_y - (bottom - top) * 0.34 * amplitude
        points.append((x, y))

    mirrored = [(x, 2 * mid_y - y) for x, y in reversed(points)]
    polygon = points + mirrored
    draw.polygon(polygon, fill=(53, 150, 255, 210))
    draw.line(points, fill=(150, 217, 255, 235), width=4)
    draw.line(list(reversed(mirrored)), fill=(150, 217, 255, 160), width=4)


def _draw_number_badges(image: Image.Image, box: tuple[int, int, int, int]) -> None:
    draw = ImageDraw.Draw(image)
    left, top, right, bottom = box
    pill_w = right - left
    pill_h = (bottom - top - 48) // 3
    font = ImageFont.load_default(size=54)
    sub_font = ImageFont.load_default(size=32)
    rows = [("001", "OK", (67, 170, 255), (23, 82, 146)),
            ("002", "OK", (67, 170, 255), (23, 82, 146)),
            ("003", "NG", (255, 166, 76), (138, 81, 19))]

    for idx, (number, label, accent, bg) in enumerate(rows):
        y0 = top + idx * (pill_h + 24)
        y1 = y0 + pill_h
        draw.rounded_rectangle((left, y0, right, y1), radius=40, fill=(34, 42, 59, 230))
        draw.rounded_rectangle((left + 18, y0 + 18, left + 164, y1 - 18), radius=32, fill=(*bg, 255))
        draw.text((left + 44, y0 + 44), number, font=font, fill=(240, 245, 252, 255))
        tag_w = 106
        draw.rounded_rectangle((right - tag_w - 24, y0 + 26, right - 24, y0 + 74), radius=24, fill=(*accent, 255))
        draw.text((right - tag_w, y0 + 38), label, font=sub_font, fill=(255, 255, 255, 255))
        draw.line((left + 196, y0 + pill_h // 2, right - tag_w - 48, y0 + pill_h // 2), fill=(97, 108, 136, 160), width=6)


def render_icon(size: int = 1024) -> Image.Image:
    image = _make_base_canvas(size)
    draw = ImageDraw.Draw(image)

    draw.text((136, 156), "WAV", font=ImageFont.load_default(size=44), fill=(151, 167, 198, 220))

    _draw_waveform(draw, (150, 248, 874, 560))
    _draw_number_badges(image, (184, 612, 840, 850))
    return image


def save_icon_assets() -> None:
    ICON_DIR.mkdir(parents=True, exist_ok=True)
    base = render_icon(1024)
    PNG_PATH.parent.mkdir(parents=True, exist_ok=True)
    base.save(PNG_PATH)
    base.save(ICO_PATH, sizes=[(256, 256), (128, 128), (64, 64), (48, 48), (32, 32), (16, 16)])

    icon_specs = {
        "icon_16x16.png": 16,
        "icon_16x16@2x.png": 32,
        "icon_32x32.png": 32,
        "icon_32x32@2x.png": 64,
        "icon_128x128.png": 128,
        "icon_128x128@2x.png": 256,
        "icon_256x256.png": 256,
        "icon_256x256@2x.png": 512,
        "icon_512x512.png": 512,
        "icon_512x512@2x.png": 1024,
    }

    for filename, size in icon_specs.items():
        resized = base.resize((size, size), Image.Resampling.LANCZOS)
        resized.save(ICON_DIR / filename)


if __name__ == "__main__":
    save_icon_assets()
    print(PNG_PATH)
    print(ICNS_PATH)
    print(ICO_PATH)

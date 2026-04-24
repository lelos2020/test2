"""
Image generation for Instagram posts.
Creates minimalist, aesthetic quote cards using PIL/Pillow.
"""

import logging
import os
import random
import textwrap
from datetime import datetime
from pathlib import Path

from PIL import Image, ImageDraw, ImageFilter, ImageFont

import config

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Font helpers
# ---------------------------------------------------------------------------

def _load_font(size: int, bold: bool = False) -> ImageFont.ImageFont:
    """
    Try to load a TTF font; gracefully fall back to PIL's built-in bitmap font.
    Drop custom .ttf files into the fonts/ directory to use them.
    """
    font_dir = Path(config.FONT_DIR) if hasattr(config, "FONT_DIR") else Path(config.BASE_DIR) / "fonts"
    candidates = []

    if bold:
        candidates += list(font_dir.glob("*[Bb]old*.ttf")) + list(font_dir.glob("*[Bb]lack*.ttf"))
    candidates += list(font_dir.glob("*.ttf")) + list(font_dir.glob("*.otf"))

    # System font fallbacks
    system_candidates = [
        "/usr/share/fonts/truetype/dejavu/DejaVuSerif.ttf",
        "/usr/share/fonts/truetype/liberation/LiberationSerif-Regular.ttf",
        "/System/Library/Fonts/Georgia.ttf",
        "C:/Windows/Fonts/georgia.ttf",
    ]
    candidates += [Path(p) for p in system_candidates]

    for path in candidates:
        if path.exists():
            try:
                return ImageFont.truetype(str(path), size)
            except Exception:
                continue

    logger.warning("No TTF font found — using PIL default bitmap font.")
    return ImageFont.load_default()


# ---------------------------------------------------------------------------
# Gradient background
# ---------------------------------------------------------------------------

def _make_gradient(size: tuple[int, int], palette: dict) -> Image.Image:
    """Create a subtle two-color diagonal gradient."""
    w, h = size
    img = Image.new("RGB", size)
    draw = ImageDraw.Draw(img)

    bg = _hex_to_rgb(palette["bg"])
    accent = _hex_to_rgb(palette["accent"])

    for y in range(h):
        ratio = y / h * 0.25          # very subtle — only 25 % accent contribution
        r = int(bg[0] + (accent[0] - bg[0]) * ratio)
        g = int(bg[1] + (accent[1] - bg[1]) * ratio)
        b = int(bg[2] + (accent[2] - bg[2]) * ratio)
        draw.line([(0, y), (w, y)], fill=(r, g, b))

    # Soft diagonal noise overlay (cheaply achieved by a tiny blur trick)
    img = img.filter(ImageFilter.GaussianBlur(radius=2))
    return img


def _hex_to_rgb(hex_color: str) -> tuple[int, int, int]:
    h = hex_color.lstrip("#")
    return tuple(int(h[i: i + 2], 16) for i in (0, 2, 4))


# ---------------------------------------------------------------------------
# Text layout helpers
# ---------------------------------------------------------------------------

def _wrap_text(text: str, font: ImageFont.ImageFont, max_width: int, draw: ImageDraw.ImageDraw) -> list[str]:
    """Word-wrap text to fit within max_width pixels."""
    words = text.split()
    lines, current = [], ""
    for word in words:
        test = f"{current} {word}".strip()
        bbox = draw.textbbox((0, 0), test, font=font)
        if bbox[2] <= max_width:
            current = test
        else:
            if current:
                lines.append(current)
            current = word
    if current:
        lines.append(current)
    return lines


def _text_block_height(lines: list[str], font: ImageFont.ImageFont, line_spacing: int, draw: ImageDraw.ImageDraw) -> int:
    total = 0
    for line in lines:
        bbox = draw.textbbox((0, 0), line, font=font)
        total += (bbox[3] - bbox[1]) + line_spacing
    return total


# ---------------------------------------------------------------------------
# Decorative elements
# ---------------------------------------------------------------------------

def _draw_divider(draw: ImageDraw.ImageDraw, cx: int, y: int, palette: dict, width: int = 80) -> None:
    """Draw a thin centered horizontal rule."""
    color = _hex_to_rgb(palette["accent"])
    draw.line([(cx - width // 2, y), (cx + width // 2, y)], fill=color, width=2)


def _draw_quote_marks(draw: ImageDraw.ImageDraw, x: int, y: int, size: int, palette: dict) -> None:
    font = _load_font(size)
    color = _hex_to_rgb(palette["accent"])
    draw.text((x, y), "“", font=font, fill=color)


# ---------------------------------------------------------------------------
# Layout strategies
# ---------------------------------------------------------------------------

def _layout_centered(draw, img, quote, palette, w, h):
    """Classic centered layout."""
    margin = int(w * 0.12)
    text_w = w - 2 * margin
    cx = w // 2

    quote_font = _load_font(int(h * 0.038), bold=False)
    attr_font = _load_font(int(h * 0.022))

    lines = _wrap_text(quote["text"], quote_font, text_w, draw)
    line_h = int(h * 0.048)
    block_h = _text_block_height(lines, quote_font, line_h - int(h * 0.038), draw)

    start_y = (h - block_h) // 2 - int(h * 0.04)

    # Opening quote mark
    _draw_quote_marks(draw, margin, start_y - int(h * 0.06), int(h * 0.065), palette)

    # Quote text
    y = start_y
    text_color = _hex_to_rgb(palette["text"])
    for line in lines:
        bbox = draw.textbbox((0, 0), line, font=quote_font)
        x = cx - (bbox[2] - bbox[0]) // 2
        draw.text((x, y), line, font=quote_font, fill=text_color)
        y += line_h

    # Divider
    _draw_divider(draw, cx, y + int(h * 0.025), palette)

    # Attribution
    source = quote.get("source", "")
    if source and source != "Original":
        attr_text = f"— {source}"
        bbox = draw.textbbox((0, 0), attr_text, font=attr_font)
        ax = cx - (bbox[2] - bbox[0]) // 2
        draw.text((ax, y + int(h * 0.045)), attr_text, font=attr_font,
                  fill=_hex_to_rgb(palette["accent"]))


def _layout_top_heavy(draw, img, quote, palette, w, h):
    """Quote starts at ~20 % from top, attribution at bottom."""
    margin = int(w * 0.10)
    text_w = w - 2 * margin
    cx = w // 2

    quote_font = _load_font(int(h * 0.040))
    attr_font = _load_font(int(h * 0.022))

    lines = _wrap_text(quote["text"], quote_font, text_w, draw)
    line_h = int(h * 0.050)

    start_y = int(h * 0.20)
    _draw_quote_marks(draw, margin, start_y - int(h * 0.07), int(h * 0.065), palette)

    y = start_y
    text_color = _hex_to_rgb(palette["text"])
    for line in lines:
        bbox = draw.textbbox((0, 0), line, font=quote_font)
        x = cx - (bbox[2] - bbox[0]) // 2
        draw.text((x, y), line, font=quote_font, fill=text_color)
        y += line_h

    source = quote.get("source", "")
    if source and source != "Original":
        attr_text = f"— {source}"
        attr_font2 = _load_font(int(h * 0.024))
        bbox = draw.textbbox((0, 0), attr_text, font=attr_font2)
        ax = cx - (bbox[2] - bbox[0]) // 2
        draw.text((ax, int(h * 0.82)), attr_text, font=attr_font2,
                  fill=_hex_to_rgb(palette["accent"]))


def _layout_bottom_heavy(draw, img, quote, palette, w, h):
    """Quote sits in the lower half, plenty of white space above."""
    margin = int(w * 0.10)
    text_w = w - 2 * margin
    cx = w // 2

    quote_font = _load_font(int(h * 0.038))
    attr_font = _load_font(int(h * 0.022))

    lines = _wrap_text(quote["text"], quote_font, text_w, draw)
    line_h = int(h * 0.048)
    block_h = _text_block_height(lines, quote_font, line_h - int(h * 0.038), draw)

    start_y = int(h * 0.55) - block_h // 2
    _draw_quote_marks(draw, margin, start_y - int(h * 0.06), int(h * 0.065), palette)

    y = start_y
    text_color = _hex_to_rgb(palette["text"])
    for line in lines:
        bbox = draw.textbbox((0, 0), line, font=quote_font)
        x = cx - (bbox[2] - bbox[0]) // 2
        draw.text((x, y), line, font=quote_font, fill=text_color)
        y += line_h

    source = quote.get("source", "")
    if source and source != "Original":
        attr_text = f"— {source}"
        bbox = draw.textbbox((0, 0), attr_text, font=attr_font)
        ax = cx - (bbox[2] - bbox[0]) // 2
        draw.text((ax, y + int(h * 0.03)), attr_text, font=attr_font,
                  fill=_hex_to_rgb(palette["accent"]))


def _layout_split(draw, img, quote, palette, w, h):
    """Accent bar on the left, text on the right."""
    bar_w = int(w * 0.015)
    margin_left = int(w * 0.10) + bar_w + int(w * 0.03)
    margin_right = int(w * 0.08)
    text_w = w - margin_left - margin_right
    cy = h // 2

    # Accent bar
    bar_color = _hex_to_rgb(palette["accent"])
    bar_x = int(w * 0.08)
    bar_top = int(h * 0.25)
    bar_bot = int(h * 0.75)
    draw.rectangle([bar_x, bar_top, bar_x + bar_w, bar_bot], fill=bar_color)

    quote_font = _load_font(int(h * 0.036))
    attr_font = _load_font(int(h * 0.021))

    lines = _wrap_text(quote["text"], quote_font, text_w, draw)
    line_h = int(h * 0.046)
    block_h = _text_block_height(lines, quote_font, line_h - int(h * 0.036), draw)

    y = cy - block_h // 2
    text_color = _hex_to_rgb(palette["text"])
    for line in lines:
        draw.text((margin_left, y), line, font=quote_font, fill=text_color)
        y += line_h

    source = quote.get("source", "")
    if source and source != "Original":
        draw.text((margin_left, y + int(h * 0.025)), f"— {source}",
                  font=attr_font, fill=_hex_to_rgb(palette["accent"]))


LAYOUTS = {
    "centered": _layout_centered,
    "top_heavy": _layout_top_heavy,
    "bottom_heavy": _layout_bottom_heavy,
    "split": _layout_split,
}


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def create_image(quote: dict, layout: Optional[str] = None, palette_index: Optional[int] = None) -> str:
    """
    Generate a quote image and save it to GENERATED_IMAGES_DIR.
    Returns the file path.
    """
    from typing import Optional  # local to avoid circular

    os.makedirs(config.GENERATED_IMAGES_DIR, exist_ok=True)

    # Choose palette and layout
    palette = config.COLOR_PALETTES[palette_index % len(config.COLOR_PALETTES)] if palette_index is not None \
        else random.choice(config.COLOR_PALETTES)
    chosen_layout = layout if layout in LAYOUTS else random.choice(config.LAYOUT_OPTIONS)

    w, h = config.IMAGE_SIZE
    img = _make_gradient((w, h), palette)
    draw = ImageDraw.Draw(img)

    layout_fn = LAYOUTS.get(chosen_layout, _layout_centered)
    layout_fn(draw, img, quote, palette, w, h)

    # Timestamp-based filename
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"quote_{quote['id']}_{timestamp}.jpg"
    filepath = os.path.join(config.GENERATED_IMAGES_DIR, filename)
    img.save(filepath, "JPEG", quality=95)

    logger.info("Image saved: %s (layout=%s)", filepath, chosen_layout)
    return filepath


# Allow Optional import at module level for the function signature
from typing import Optional  # noqa: E402

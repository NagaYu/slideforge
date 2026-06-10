"""Layout detection, geometry computation, and text auto-fit.

All geometry is computed dynamically for a 16:9 canvas (13.333" x 7.5")
and returned as plain float inches, so the renderer just wraps values in
``Inches()``.  Nothing here imports python-pptx — the engine is pure math
and therefore trivially unit-testable.
"""

from __future__ import annotations

import math
from dataclasses import dataclass

from .md_parser import Slide

# ---------------------------------------------------------------- canvas ---
SLIDE_W = 13.333
SLIDE_H = 7.5
MARGIN = 0.6                 # outer margin
TITLE_TOP = 0.42
TITLE_H = 0.85
CONTENT_TOP = 1.55           # where body content starts on titled slides
CONTENT_W = SLIDE_W - 2 * MARGIN


@dataclass
class Box:
    left: float
    top: float
    width: float
    height: float


# ------------------------------------------------------- layout detection ---
LAYOUT_STATEMENT = "statement"        # title-only → full-bleed statement
LAYOUT_TWO_COLUMN = "two_column"      # two ## headings → left/right split
LAYOUT_THREE_CARDS = "three_cards"    # three parallel bullets → card grid
LAYOUT_TIMELINE = "timeline"          # numbered list → step timeline
LAYOUT_BULLETS = "bullets"            # fallback: classic bullet slide


def detect_layout(slide: Slide) -> str:
    """Map a parsed slide onto one of the dynamic layouts."""
    if slide.is_statement:
        return LAYOUT_STATEMENT
    if len(slide.sections) == 2:
        return LAYOUT_TWO_COLUMN
    if slide.steps:
        return LAYOUT_TIMELINE
    if len(slide.bullets) == 3 and all(b.lead for b in slide.bullets):
        return LAYOUT_THREE_CARDS
    return LAYOUT_BULLETS


# ------------------------------------------------------------- geometries ---
def title_box() -> Box:
    return Box(MARGIN, TITLE_TOP, CONTENT_W, TITLE_H)


def two_column_boxes(gap: float = 0.45) -> list[Box]:
    """Two equal cards, side by side."""
    col_w = (CONTENT_W - gap) / 2
    h = SLIDE_H - CONTENT_TOP - 0.55
    return [
        Box(MARGIN, CONTENT_TOP, col_w, h),
        Box(MARGIN + col_w + gap, CONTENT_TOP, col_w, h),
    ]


def card_boxes(n: int, gap: float = 0.4) -> list[Box]:
    """``n`` equal cards in one row (used for the 3-column layout)."""
    top = CONTENT_TOP + 0.15
    card_w = (CONTENT_W - gap * (n - 1)) / n
    h = SLIDE_H - top - 0.7
    return [Box(MARGIN + i * (card_w + gap), top, card_w, h) for i in range(n)]


def timeline_geometry(n: int) -> dict:
    """Horizontal step timeline: connector line, numbered circles, and a
    text box under each circle.  Spacing adapts to the step count."""
    n = max(n, 1)
    circle_d = 0.62
    line_y = 2.85
    slot_w = CONTENT_W / n
    centers = [MARGIN + slot_w * (i + 0.5) for i in range(n)]
    text_w = min(slot_w - 0.25, 3.2)
    return {
        "line": Box(centers[0], line_y + circle_d / 2 - 0.015,
                    centers[-1] - centers[0], 0.03),
        "circles": [Box(c - circle_d / 2, line_y, circle_d, circle_d)
                    for c in centers],
        "texts": [Box(c - text_w / 2, line_y + circle_d + 0.3, text_w,
                      SLIDE_H - (line_y + circle_d + 0.3) - 0.5)
                  for c in centers],
    }


def bullets_box() -> Box:
    return Box(MARGIN + 0.1, CONTENT_TOP + 0.1, CONTENT_W - 0.2,
               SLIDE_H - CONTENT_TOP - 0.65)


# ------------------------------------------------------------- auto-fit ----
def _char_width_pt(ch: str, size_pt: float) -> float:
    """Rough advance width of one character at ``size_pt``.

    CJK glyphs are full-width (~1.0 em); Latin averages ~0.52 em with a few
    narrow/wide exceptions.  Estimation errs slightly wide on purpose so the
    fitter shrinks a step too early rather than a step too late.
    """
    o = ord(ch)
    if o >= 0x2E80:                      # CJK, kana, full-width forms
        return size_pt * 1.02
    if ch in "iIljft.,:;!|'`[]()" :
        return size_pt * 0.32
    if ch.isupper() or ch in "wmWM@%&":
        return size_pt * 0.72
    return size_pt * 0.52


def estimate_width_pt(text: str, size_pt: float) -> float:
    return sum(_char_width_pt(c, size_pt) for c in text)


def wrapped_line_count(text: str, size_pt: float, avail_w_pt: float) -> int:
    if not text:
        return 1
    return max(1, math.ceil(estimate_width_pt(text, size_pt) / avail_w_pt))


def fit_font_size(
    lines: list[str],
    box_w_in: float,
    box_h_in: float,
    start_pt: float,
    min_pt: float = 10,
    line_spacing: float = 1.18,
    space_after_pt: float = 6,
    pad_w_in: float = 0.2,
    pad_h_in: float = 0.12,
) -> float:
    """Largest font size (stepping down 2pt at a time) at which ``lines``
    fit inside the box.  Never errors: bottoms out at ``min_pt``.

    Each entry in ``lines`` is one paragraph; wrapping inside the box width
    is estimated per paragraph.
    """
    avail_w = (box_w_in - pad_w_in) * 72
    avail_h = (box_h_in - pad_h_in) * 72
    if avail_w <= 0 or avail_h <= 0:
        return min_pt

    size = float(start_pt)
    while size > min_pt:
        wrapped = sum(wrapped_line_count(ln, size, avail_w) for ln in lines)
        needed = wrapped * size * line_spacing + max(len(lines) - 1, 0) * space_after_pt
        if needed <= avail_h:
            return size
        size -= 2
    return min_pt

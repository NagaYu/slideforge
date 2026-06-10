"""Layout detection + dynamic geometry (Inches) for SlideForge.

The engine looks at the *shape of the content* and picks a composition:

========================  =============================================
content pattern           layout
==========================  ===========================================
title only (first/last)    ``title`` / ``closing``  (hero slide)
two ``##`` headings         ``two_column``  (left / right split)
3-4 parallel top bullets    ``cards``       (column card grid)
numbered list               ``timeline``    (step circles + connector)
anything else               ``content``     (title + bullets)
==========================  ===========================================

All geometry is computed dynamically from the slide size so the same
engine works for 16:9 and 4:3 decks.  Every function returns plain
``Rect`` tuples in **inches**; only the renderer touches python-pptx.
"""

from __future__ import annotations

from dataclasses import dataclass

from .parser import Slide

# 16:9 default canvas (inches)
SLIDE_W = 13.333
SLIDE_H = 7.5

MARGIN = 0.6        # outer margin (skill guideline: >= 0.5")
GAP = 0.4           # gap between cards / columns
TITLE_TOP = 0.45
TITLE_H = 0.95
BODY_TOP = TITLE_TOP + TITLE_H + 0.25


@dataclass(frozen=True)
class Rect:
    left: float
    top: float
    width: float
    height: float


# --------------------------------------------------------------------------
# layout detection
# --------------------------------------------------------------------------

def detect_layout(slide: Slide, index: int, total: int) -> str:
    """Pick a layout name for one parsed slide."""
    has_body = bool(slide.blocks)

    if not has_body:
        if index == 0:
            return "title"
        if index == total - 1:
            return "closing"
        return "section"

    if len(slide.steps) >= 2:
        return "timeline"

    if len(slide.headings) == 2 and all(
        h.level == 2 for h in slide.headings
    ):
        return "two_column"

    bullets = slide.top_bullets
    if len(bullets) in (3, 4) and not slide.headings:
        return "cards"

    return "content"


# --------------------------------------------------------------------------
# geometry helpers (all values in inches)
# --------------------------------------------------------------------------

def title_rect() -> Rect:
    return Rect(MARGIN, TITLE_TOP, SLIDE_W - 2 * MARGIN, TITLE_H)


def body_rect() -> Rect:
    return Rect(MARGIN, BODY_TOP,
                SLIDE_W - 2 * MARGIN, SLIDE_H - BODY_TOP - MARGIN)


def column_rects(n: int = 2, header_h: float = 0.6) -> list[tuple[Rect, Rect]]:
    """Return ``n`` (header, body) rect pairs side by side."""
    area = body_rect()
    col_w = (area.width - GAP * (n - 1)) / n
    pairs = []
    for i in range(n):
        left = area.left + i * (col_w + GAP)
        header = Rect(left, area.top, col_w, header_h)
        body = Rect(left, area.top + header_h + 0.1,
                    col_w, area.height - header_h - 0.1)
        pairs.append((header, body))
    return pairs


def card_rects(n: int, reserve_bottom: float = 0.0) -> list[Rect]:
    """Equal-width card rects across the body area (max 4 per row).

    ``reserve_bottom`` shrinks the grid to leave room for a trailing
    quote / paragraph band below the cards.
    """
    area = body_rect()
    if reserve_bottom:
        area = Rect(area.left, area.top, area.width,
                    area.height - reserve_bottom)
    per_row = min(n, 4) if n != 4 else 2          # 4 cards -> 2x2 grid
    rows = -(-n // per_row)                       # ceil division
    card_w = (area.width - GAP * (per_row - 1)) / per_row
    card_h = (area.height - GAP * (rows - 1)) / rows
    rects = []
    for i in range(n):
        row, col = divmod(i, per_row)
        rects.append(Rect(
            area.left + col * (card_w + GAP),
            area.top + row * (card_h + GAP),
            card_w, card_h,
        ))
    return rects


def timeline_positions(n: int) -> dict:
    """Geometry for a horizontal step timeline.

    Returns dict with the connector line rect, circle rects and
    label rects for ``n`` steps.  Falls back to two rows when more
    than 5 steps are present.
    """
    area = body_rect()
    per_row = n if n <= 5 else -(-n // 2)
    rows = 1 if n <= 5 else 2
    circle_d = 0.85
    # cap the row height and center the block vertically so a single-row
    # timeline does not cling to the top of an otherwise empty slide
    row_h = min(area.height / rows, 2.9)
    top_offset = max((area.height - row_h * rows) / 2 - 0.2, 0.0)
    area = Rect(area.left, area.top + top_offset, area.width, area.height)
    slot_w = area.width / per_row

    circles, labels, connectors = [], [], []
    for i in range(n):
        row, col = divmod(i, per_row)
        cx = area.left + slot_w * col + slot_w / 2
        cy = area.top + row * row_h + 0.55
        circles.append(Rect(cx - circle_d / 2, cy - circle_d / 2,
                            circle_d, circle_d))
        labels.append(Rect(area.left + slot_w * col + 0.12,
                           cy + circle_d / 2 + 0.15,
                           slot_w - 0.24,
                           row_h - circle_d - 0.75))
    for row in range(rows):
        in_row = min(per_row, n - row * per_row)
        if in_row >= 2:
            first = circles[row * per_row]
            last = circles[row * per_row + in_row - 1]
            y = first.top + circle_d / 2
            connectors.append(Rect(first.left + circle_d, y - 0.015,
                                   last.left - first.left - circle_d, 0.03))
    return {"circles": circles, "labels": labels, "connectors": connectors}


def hero_rects() -> dict:
    """Centered composition for title / closing / section slides."""
    return {
        "kicker": Rect(MARGIN + 0.2, 2.35, SLIDE_W - 2 * (MARGIN + 0.2), 0.5),
        "title": Rect(MARGIN + 0.2, 2.9, SLIDE_W - 2 * (MARGIN + 0.2), 1.6),
        "subtitle": Rect(MARGIN + 0.2, 4.6, SLIDE_W - 2 * (MARGIN + 0.2), 0.9),
    }

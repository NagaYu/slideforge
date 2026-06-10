"""Text auto-fit: shrink font sizes so text never spills out of its box.

python-pptx cannot measure rendered text (that happens inside PowerPoint),
so SlideForge uses a deterministic width estimate:

* Latin characters average ~0.50 em, CJK characters a full ~1.0 em.
* A line of width ``W`` inches at ``size`` pt therefore fits roughly
  ``W * 72 / (size * avg_em)`` characters.

``fit_font_size`` walks down from the requested size in 2 pt steps until
the estimated wrapped line count fits the box height (or the minimum
size is reached -- it never raises).
"""

from __future__ import annotations

import unicodedata

LINE_SPACING = 1.18     # approximate single line height multiplier
MIN_SIZE = 9            # never go below this (pt)
STEP = 2                # shrink 2pt at a time, per spec


def _char_em(ch: str) -> float:
    """Approximate width of one character in em units."""
    if unicodedata.east_asian_width(ch) in ("W", "F"):
        return 1.0          # CJK / full-width
    if ch in "iljI.,:;'|!()[] ":
        return 0.32
    if ch.isupper() or ch in "mwMW@%&":
        return 0.66
    return 0.50


def text_width_em(text: str) -> float:
    return sum(_char_em(c) for c in text)


def estimate_lines(text: str, width_in: float, size_pt: float) -> int:
    """Estimate the number of wrapped lines for ``text`` in a box."""
    if not text:
        return 0
    em_per_line = max(width_in * 72.0 / size_pt, 1.0)
    lines = 0
    # respect explicit newlines, wrap each segment independently
    for segment in text.split("\n"):
        seg_em = text_width_em(segment)
        lines += max(1, int(seg_em / em_per_line) + (seg_em % em_per_line > 0))
    return lines


def fits(paragraphs: list[tuple[str, float]], width_in: float,
         height_in: float, scale: float = 1.0,
         para_spacing_pt: float = 6.0) -> bool:
    """True when all (text, size_pt) paragraphs fit the box at ``scale``."""
    used_pt = 0.0
    for text, size in paragraphs:
        size = max(size * scale, MIN_SIZE)
        n = estimate_lines(text, width_in, size)
        used_pt += n * size * LINE_SPACING + para_spacing_pt
    return used_pt <= height_in * 72.0


def fit_font_size(text: str, width_in: float, height_in: float,
                  start_pt: float, min_pt: float = MIN_SIZE) -> int:
    """Largest size <= start_pt (2pt steps) whose wrapped text fits the box."""
    size = start_pt
    while size > min_pt:
        n = estimate_lines(text, width_in, size)
        if n * size * LINE_SPACING <= height_in * 72.0:
            return int(size)
        size -= STEP
    return int(max(min_pt, MIN_SIZE))


def fit_scale(paragraphs: list[tuple[str, float]], width_in: float,
              height_in: float) -> float:
    """Uniform scale factor for a whole box of mixed-size paragraphs.

    Keeps the visual hierarchy (relative sizes) intact while shrinking
    everything together until the content fits.
    """
    scale = 1.0
    while scale > 0.4:
        if fits(paragraphs, width_in, height_in, scale):
            return scale
        scale -= 0.08
    return 0.4

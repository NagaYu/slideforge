"""Markdown → slide model parser.

Slides are separated by ``---`` lines.  Inside a slide:

* ``# Heading``   slide title (first slide's ``#`` becomes the deck title)
* ``## Heading``  column heading (two of them triggers the split layout)
* ``- item``      bullet (``**bold lead:** description`` becomes a card)
* ``1. item``     numbered step (triggers the timeline layout)
* plain text      paragraph (on the title slide: the subtitle)
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field

_BOLD_LEAD = re.compile(r"^\*\*(.+?)\*\*\s*[::]?\s*(.*)$", re.S)
_NUMBERED = re.compile(r"^(\d+)[.)]\s+(.*)$")


@dataclass
class Item:
    """One bullet / step. ``lead`` is the bold lead-in if the author wrote
    ``**lead:** body``, otherwise None and ``body`` holds the whole text."""

    body: str
    lead: str | None = None


@dataclass
class Section:
    """A ``##`` heading and the items beneath it."""

    heading: str
    items: list[Item] = field(default_factory=list)


@dataclass
class Slide:
    title: str = ""
    sections: list[Section] = field(default_factory=list)
    bullets: list[Item] = field(default_factory=list)
    steps: list[Item] = field(default_factory=list)
    paragraphs: list[str] = field(default_factory=list)

    @property
    def is_statement(self) -> bool:
        """Title-only slide → rendered as a full-bleed statement slide."""
        return bool(self.title) and not (
            self.sections or self.bullets or self.steps or self.paragraphs
        )


@dataclass
class Deck:
    title: str = ""
    subtitle: str = ""
    slides: list[Slide] = field(default_factory=list)


def _parse_item(text: str) -> Item:
    m = _BOLD_LEAD.match(text.strip())
    if m and m.group(2).strip():
        # the colon may sit inside the bold span: "**lead:** body"
        lead = m.group(1).strip().rstrip(":：")
        return Item(body=m.group(2).strip(), lead=lead)
    return Item(body=text.strip())


def _parse_slide(block: str) -> Slide:
    slide = Slide()
    current: Section | None = None
    for raw in block.splitlines():
        line = raw.rstrip()
        stripped = line.strip()
        if not stripped:
            continue
        if stripped.startswith("## "):
            current = Section(heading=stripped[3:].strip())
            slide.sections.append(current)
        elif stripped.startswith("# "):
            slide.title = stripped[2:].strip()
            current = None
        elif stripped.startswith(("- ", "* ")):
            item = _parse_item(stripped[2:])
            (current.items if current else slide.bullets).append(item)
        elif _NUMBERED.match(stripped):
            slide.steps.append(_parse_item(_NUMBERED.match(stripped).group(2)))
        else:
            slide.paragraphs.append(stripped)
    return slide


def parse_markdown(text: str) -> Deck:
    deck = Deck()
    blocks = [b for b in re.split(r"^\s*---\s*$", text, flags=re.M) if b.strip()]
    for i, block in enumerate(blocks):
        slide = _parse_slide(block)
        if i == 0 and slide.title and not (slide.sections or slide.bullets or slide.steps):
            deck.title = slide.title
            deck.subtitle = " ".join(slide.paragraphs)
        else:
            deck.slides.append(slide)
    if not deck.title and deck.slides:
        deck.title = deck.slides[0].title or "Untitled"
    return deck

"""Markdown -> slide model parser.

Dependency-free, line oriented.  The grammar SlideForge understands:

* ``---`` on its own line          -> slide separator
* ``# Title``                      -> slide title (a new ``#`` also starts
                                      a new slide when no ``---`` was given)
* ``## Heading``                   -> section heading inside a slide
* ``- item`` / ``* item``          -> bullet (2-space indent = sub bullet)
* ``1. item``                      -> numbered step
* ``> text``                       -> quote / tagline
* plain text                       -> paragraph
* ``**bold**`` inline              -> bold run

The parser is intentionally forgiving: anything it does not recognise is
treated as a paragraph so the user never gets a hard failure.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field


# --------------------------------------------------------------------------
# slide model
# --------------------------------------------------------------------------

@dataclass
class Block:
    """One content block inside a slide."""

    kind: str            # heading | bullet | step | quote | para
    text: str            # raw text (inline markers stripped)
    level: int = 0       # bullet nesting / heading level
    runs: list = field(default_factory=list)  # [(text, bold), ...]
    children: list = field(default_factory=list)  # sub bullets for cards


@dataclass
class Slide:
    title: str = ""
    subtitle: str = ""
    blocks: list = field(default_factory=list)

    @property
    def headings(self) -> list:
        return [b for b in self.blocks if b.kind == "heading"]

    @property
    def steps(self) -> list:
        return [b for b in self.blocks if b.kind == "step"]

    @property
    def top_bullets(self) -> list:
        return [b for b in self.blocks if b.kind == "bullet" and b.level == 0]


# --------------------------------------------------------------------------
# inline markdown
# --------------------------------------------------------------------------

_BOLD_RE = re.compile(r"\*\*(.+?)\*\*")


def parse_runs(text: str) -> list:
    """Split ``**bold**`` markers into [(text, bold)] runs."""
    runs, pos = [], 0
    for m in _BOLD_RE.finditer(text):
        if m.start() > pos:
            runs.append((text[pos:m.start()], False))
        runs.append((m.group(1), True))
        pos = m.end()
    if pos < len(text):
        runs.append((text[pos:], False))
    return runs or [(text, False)]


def strip_inline(text: str) -> str:
    return _BOLD_RE.sub(r"\1", text)


# --------------------------------------------------------------------------
# document parsing
# --------------------------------------------------------------------------

_HR_RE = re.compile(r"^\s*-{3,}\s*$")
_H_RE = re.compile(r"^(#{1,4})\s+(.*)$")
_BULLET_RE = re.compile(r"^(\s*)[-*]\s+(.*)$")
_STEP_RE = re.compile(r"^\s*(\d+)[.)]\s+(.*)$")
_QUOTE_RE = re.compile(r"^\s*>\s?(.*)$")


def parse_markdown(text: str) -> list[Slide]:
    """Parse a markdown document into a list of :class:`Slide`."""
    slides: list[Slide] = []
    current: Slide | None = None

    def ensure_slide() -> Slide:
        nonlocal current
        if current is None:
            current = Slide()
            slides.append(current)
        return current

    for raw in text.splitlines():
        line = raw.rstrip()
        if not line.strip():
            continue

        if _HR_RE.match(line):
            current = None
            continue

        m = _H_RE.match(line)
        if m:
            level, content = len(m.group(1)), strip_inline(m.group(2).strip())
            if level == 1:
                # a new H1 always starts a new slide
                current = Slide(title=content)
                slides.append(current)
            else:
                slide = ensure_slide()
                if not slide.title and level == 2 and not slide.blocks:
                    # "## only" documents: promote first H2 to slide title
                    slide.title = content
                else:
                    slide.blocks.append(
                        Block("heading", content, level=level,
                              runs=parse_runs(m.group(2).strip()))
                    )
            continue

        m = _STEP_RE.match(line)
        if m:
            ensure_slide().blocks.append(
                Block("step", strip_inline(m.group(2)),
                      runs=parse_runs(m.group(2)))
            )
            continue

        m = _BULLET_RE.match(line)
        if m:
            indent = len(m.group(1).expandtabs(4))
            level = 1 if indent >= 2 else 0
            block = Block("bullet", strip_inline(m.group(2)), level=level,
                          runs=parse_runs(m.group(2)))
            slide = ensure_slide()
            if level == 1 and slide.top_bullets:
                slide.top_bullets[-1].children.append(block)
            slide.blocks.append(block)
            continue

        m = _QUOTE_RE.match(line)
        if m:
            ensure_slide().blocks.append(
                Block("quote", strip_inline(m.group(1)),
                      runs=parse_runs(m.group(1)))
            )
            continue

        # plain paragraph -- right after the title it becomes the subtitle
        slide = ensure_slide()
        if slide.title and not slide.blocks and not slide.subtitle:
            slide.subtitle = strip_inline(line.strip())
        else:
            slide.blocks.append(
                Block("para", strip_inline(line.strip()),
                      runs=parse_runs(line.strip()))
            )

    return slides

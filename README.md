# SlideForge ⚒️

**A CLI tool that turns Markdown into professional PowerPoint (`.pptx`) decks in one shot.**

Just write Markdown and get 16:9 slides with theme colors, fonts, and composition
(two-column / cards / timeline) automatically detected and applied. Built on
python-pptx, with a single dependency.

```
# Proposal title            ┐
A one-line subtitle.        │──▶  dark hero cover
                            ┘
## Before … / ## After …    ──▶  two-column split layout
- A: … / - B: … / - C: …    ──▶  3-up card grid
1. Analyze 2. PoC 3. Roll out … ──▶  step timeline
```

| | |
|---|---|
| ![Cover (TechBlue)](docs/images/title-techblue.jpg) | ![Two-column (TechBlue)](docs/images/two-column-techblue.jpg) |
| ![Cards (WarmCreative)](docs/images/cards-warmcreative.jpg) | ![Timeline (MinimalGray)](docs/images/timeline-minimalgray.jpg) |

> A full set of generated samples lives in [examples/](examples/), including a
> Japanese deck that demonstrates CJK support.

## Install

```bash
git clone https://github.com/NagaYu/slideforge.git
cd slideforge
pip install .          # or: pipx install .
```

Python 3.10+ / dependency: `python-pptx`

## Usage

```bash
# Print a sample proposal in Markdown
slideforge sample > proposal.md

# Build with a specific theme
slideforge build proposal.md --theme TechBlue

# Build every theme at once (handy for comparing themes)
slideforge build proposal.md --all-themes -o build/

# List available themes
slideforge themes
```

Output is named `<input-name>_<theme>.pptx`. Pass a `.pptx` path to `-o` to set the
filename explicitly.

## Writing Markdown

| Syntax | Result |
|--------|--------|
| `# Heading` | A new slide (title) |
| `---` | Slide separator |
| Line right after a title | Subtitle (centered italic on cover slides) |
| `## Heading` × 2 | Auto-converts to a **two-column split layout** |
| 3–4 parallel bullets | **Card layout** (3-up / 2×2 grid) |
| `1.` `2.` … numbered list | **Step timeline** (6+ items wrap to two rows) |
| `- item` / 2-space indent | Bullet / sub-bullet |
| `> quote` | Italic accent text in the accent color |
| `**emphasis**` | Bold |

Split a card heading with `：` or `:` — e.g. `- **Title**: body` — to auto-split it
into a title plus body.

### Layout auto-detection rules

1. Two or more numbered lists → `timeline`
2. Exactly two `##` headings → `two_column`
3. 3–4 parallel bullets (no headings) → `cards`
4. No body, first/last slide → `title` / `closing` (hero)
5. Otherwise → `content` (title + bullets)

## Themes

| Theme | Mood | Title font | Traits |
|-------|------|------------|--------|
| `TechBlue` | Navy × mint | Trebuchet MS | Dark cover, rounded cards |
| `MinimalGray` | Achromatic charcoal | Georgia | Light throughout, square cards |
| `WarmCreative` | Terracotta × sage | Palatino | Large rounded corners, warm feel |

### Text auto-fit

When text doesn't fit its box, the font size **shrinks automatically in 2pt steps**
(minimum 9pt, never raises an error). Boxes that mix multiple font sizes are scaled
down uniformly so the visual hierarchy is preserved. CJK characters are measured at
full width, so Japanese text never overflows either.

## Extending

### Add a theme

Just add one entry to the `THEMES` dict in
[slideforge/themes.py](slideforge/themes.py).

```python
THEMES["ForestGreen"] = {
    "display_name": "Forest Green",
    "colors": {
        "primary": (44, 95, 45),     # titles / emphasis
        "secondary": (151, 188, 98), # sub-headings
        "accent": (245, 245, 245),   # number circles, etc.
        "bg": (255, 255, 255), "bg_dark": (24, 48, 24),
        "text": (40, 50, 40), "text_inverse": (240, 245, 240),
        "muted": (130, 140, 130),
        "card_bg": (243, 248, 240), "card_border": (210, 225, 200),
    },
    "fonts": {"title": "Cambria", "body": "Calibri"},
    "rules": {
        "dark_title_slide": True,   # dark background on the cover
        "card_corner_radius": 0.15, # card corner rounding (0 = square)
        "bullet_char": "▸",
    },
}
```

### Add a layout

1. [slideforge/layout_engine.py](slideforge/layout_engine.py) — add a detection
   condition to `detect_layout()` and write a geometry function (returns `Rect`s in
   inches)
2. [slideforge/renderer.py](slideforge/renderer.py) — add a `render_<name>()` method
   and one dispatch line in `render()`

The layout engine is pure coordinate math with no python-pptx dependency, so it's
easy to unit-test.

### Architecture

```
parser.py        Markdown → Slide/Block model (zero dependencies)
layout_engine.py composition detection + geometry in inches (zero dependencies)
autofit.py       overflow-safe font sizing (zero dependencies)
themes.py        color / font / rule dictionaries (zero dependencies)
renderer.py      ties it all together and draws with python-pptx
cli.py           argparse-based CLI
```

## Testing

```bash
pip install -e ".[dev]"
pytest
```

Covers the parser, layout detection, geometry, auto-fit, end-to-end generation for
every theme, and a **zero-content-loss guarantee** (every line of the input Markdown
is present in the output `.pptx`).

## License

MIT

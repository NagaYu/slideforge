"""Automated checks for the SlideForge core engine.

Run with:  python -m pytest tests/ -q   (or plain `python tests/test_slideforge.py`)
"""

import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from pptx import Presentation

from slideforge import layout_engine as le
from slideforge.cli import SAMPLE_MD
from slideforge.md_parser import parse_markdown
from slideforge.renderer import render_deck
from slideforge.themes import THEMES, get_theme, resolve_fonts


def test_parser_extracts_deck_structure():
    deck = parse_markdown(SAMPLE_MD)
    assert deck.title.startswith("DX")
    assert deck.subtitle
    assert len(deck.slides) == 5


def test_layout_detection():
    deck = parse_markdown(SAMPLE_MD)
    layouts = [le.detect_layout(s) for s in deck.slides]
    assert layouts == [
        le.LAYOUT_BULLETS,      # agenda: 4 bullets
        le.LAYOUT_TWO_COLUMN,   # two ## headings
        le.LAYOUT_THREE_CARDS,  # three bold-lead bullets
        le.LAYOUT_TIMELINE,     # numbered list
        le.LAYOUT_STATEMENT,    # title-only closing
    ]


def test_two_column_geometry_symmetric():
    a, b = le.two_column_boxes()
    assert abs(a.width - b.width) < 1e-9
    assert abs((a.left - le.MARGIN) -
               (le.SLIDE_W - le.MARGIN - (b.left + b.width))) < 1e-9


def test_card_boxes_fit_canvas():
    for n in (2, 3, 4):
        boxes = le.card_boxes(n)
        assert len(boxes) == n
        right = boxes[-1].left + boxes[-1].width
        assert right <= le.SLIDE_W - le.MARGIN + 1e-9


def test_timeline_geometry_scales():
    for n in (3, 4, 6):
        geo = le.timeline_geometry(n)
        assert len(geo["circles"]) == n
        for t in geo["texts"]:
            assert t.left >= 0 and t.left + t.width <= le.SLIDE_W


def test_autofit_shrinks_long_text():
    short = le.fit_font_size(["短い"], 4.0, 2.0, 18)
    long = le.fit_font_size(["こちらは非常に長い説明文で、" * 12], 4.0, 2.0, 18)
    assert short == 18
    assert long < short
    assert long >= 10  # never below the floor, never an error


def test_autofit_never_errors_on_extreme_input():
    assert le.fit_font_size(["x" * 5000], 1.0, 0.5, 32) == 10
    assert le.fit_font_size([], 0.0, 0.0, 32) == 10


def test_font_resolution_per_platform():
    theme = get_theme("MinimalGray")           # head: mincho, body: gothic
    win = resolve_fonts(theme, "win")
    mac = resolve_fonts(theme, "mac")
    assert win["fonts"]["head_ea"] == "Yu Mincho"
    assert win["fonts"]["body_ea"] == "Yu Gothic"
    assert mac["fonts"]["head_ea"] == "Hiragino Mincho ProN"
    auto = resolve_fonts(theme, "auto")
    assert auto["fonts"]["head_ea"] in ("Yu Mincho", "Hiragino Mincho ProN")
    # concrete font names pass through untouched
    custom = resolve_fonts({"fonts": {"head_ea": "My Font", "body_ea": "gothic"}}, "win")
    assert custom["fonts"]["head_ea"] == "My Font"
    # original theme dict is not mutated
    assert theme["fonts"]["head_ea"] == "mincho"


def test_render_all_themes_produces_valid_pptx():
    deck = parse_markdown(SAMPLE_MD)
    with tempfile.TemporaryDirectory() as tmp:
        for name in THEMES:
            out = Path(tmp) / f"{name}.pptx"
            layouts = render_deck(deck, get_theme(name), str(out))
            assert out.stat().st_size > 20_000
            assert layouts[0] == "title"
            prs = Presentation(str(out))   # re-open: valid OOXML
            assert len(prs.slides) == 6


if __name__ == "__main__":
    fns = [v for k, v in sorted(globals().items()) if k.startswith("test_")]
    for fn in fns:
        fn()
        print(f"ok  {fn.__name__}")
    print(f"\n{len(fns)} tests passed")

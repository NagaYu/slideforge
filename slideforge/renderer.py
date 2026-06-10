"""python-pptx rendering layer.

Takes a parsed ``Deck``, a theme dict, and turns layout-engine geometry
into actual shapes.  All text passes through ``layout_engine.fit_font_size``
so nothing ever overflows its box.
"""

from __future__ import annotations

import re

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

from . import layout_engine as le
from .md_parser import Deck, Item, Slide
from .themes import resolve_fonts

_INLINE_BOLD = re.compile(r"\*\*(.+?)\*\*")


# ------------------------------------------------------------- text utils ---
def _set_run(run, theme: dict, *, role: str = "body", size: float = 16,
             bold: bool = False, color=None, italic: bool = False) -> None:
    f = run.font
    f.name = theme["fonts"][role]
    f.size = Pt(size)
    f.bold = bold
    f.italic = italic
    f.color.rgb = RGBColor(*(color or theme["colors"]["text"]))
    # python-pptx only sets the Latin typeface; mirror it onto <a:ea> so CJK
    # text picks up the theme's East-Asian font instead of a fallback.
    rPr = run._r.get_or_add_rPr()
    for tag in ("a:ea", "a:cs"):
        el = rPr.find(qn(tag))
        if el is None:
            el = rPr.makeelement(qn(tag), {})
            rPr.append(el)
        el.set("typeface", theme["fonts"][f"{role}_ea"])


def _add_runs(para, text: str, theme: dict, **kw) -> None:
    """Split ``**bold**`` spans into bold runs; everything else as-is."""
    pos = 0
    for m in _INLINE_BOLD.finditer(text):
        if m.start() > pos:
            r = para.add_run()
            r.text = text[pos:m.start()]
            _set_run(r, theme, **kw)
        r = para.add_run()
        r.text = m.group(1)
        _set_run(r, theme, **{**kw, "bold": True})
        pos = m.end()
    if pos < len(text):
        r = para.add_run()
        r.text = text[pos:]
        _set_run(r, theme, **kw)


def _textbox(slide, box: le.Box, *, anchor=MSO_ANCHOR.TOP):
    tb = slide.shapes.add_textbox(Inches(box.left), Inches(box.top),
                                  Inches(box.width), Inches(box.height))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Inches(0.06)
    tf.margin_top = tf.margin_bottom = Inches(0.03)
    return tb, tf


def _para(tf, first: bool):
    return tf.paragraphs[0] if first else tf.add_paragraph()


# ------------------------------------------------------------ shape utils ---
def _fill(shape, rgb, line_rgb=None) -> None:
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor(*rgb)
    if line_rgb is None:
        shape.line.fill.background()
    else:
        shape.line.color.rgb = RGBColor(*line_rgb)
        shape.line.width = Pt(1)
    shape.shadow.inherit = False


def _rect(slide, box: le.Box, rgb, *, rounded=True, line_rgb=None, radius=0.06):
    kind = MSO_SHAPE.ROUNDED_RECTANGLE if rounded else MSO_SHAPE.RECTANGLE
    sp = slide.shapes.add_shape(kind, Inches(box.left), Inches(box.top),
                                Inches(box.width), Inches(box.height))
    if rounded:
        try:
            sp.adjustments[0] = radius
        except (IndexError, ValueError):
            pass
    _fill(sp, rgb, line_rgb)
    return sp


def _background(slide, rgb) -> None:
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = RGBColor(*rgb)


# --------------------------------------------------------------- elements ---
def _slide_title(slide, theme: dict, text: str) -> None:
    box = le.title_box()
    _, tf = _textbox(slide, box, anchor=MSO_ANCHOR.MIDDLE)
    size = le.fit_font_size([text], box.width, box.height, 30, min_pt=20)
    p = _para(tf, True)
    _add_runs(p, text, theme, role="head", size=size, bold=True,
              color=theme["colors"]["primary"])


def _footer(slide, theme: dict, deck_title: str, page: int) -> None:
    _, tf = _textbox(slide, le.Box(le.MARGIN, le.SLIDE_H - 0.42,
                                   le.CONTENT_W, 0.3))
    p = _para(tf, True)
    p.alignment = PP_ALIGN.RIGHT
    r = p.add_run()
    r.text = f"{deck_title}   |   {page:02d}"
    _set_run(r, theme, size=9, color=theme["colors"]["muted"])


def _items_into(tf, items: list[Item], theme: dict, box: le.Box,
                start_pt: float, *, lead_color, marker="●") -> None:
    """Write bullets (with optional bold leads) into a text frame,
    auto-fitting font size to the box."""
    lines = []
    for it in items:
        lines.append((f"{it.lead} — " if it.lead else "") + it.body)
    size = le.fit_font_size(lines, box.width, box.height, start_pt)
    first = True
    for it in items:
        p = _para(tf, first)
        first = False
        p.space_after = Pt(max(size * 0.45, 4))
        p.line_spacing = 1.12
        r = p.add_run()
        r.text = f"{marker}  "
        _set_run(r, theme, size=size * 0.62, color=theme["colors"]["accent"])
        if it.lead:
            r = p.add_run()
            r.text = f"{it.lead}　"
            _set_run(r, theme, role="head", size=size, bold=True,
                     color=lead_color)
        _add_runs(p, it.body, theme, size=size)


# ---------------------------------------------------------------- layouts ---
def _render_statement(slide, theme: dict, title: str, subtitle: str = "",
                      footer: str = "") -> None:
    """Dark full-bleed slide for the deck title and closing statements."""
    c = theme["colors"]
    _background(slide, c["title_bg"])
    # small brand glyph above the title
    _rect(slide, le.Box(0.95, 2.18, 0.3, 0.3), c["title_sub"], rounded=True,
          radius=0.5)
    box = le.Box(0.9, 2.7, le.SLIDE_W - 2.4, 1.9)
    _, tf = _textbox(slide, box)
    size = le.fit_font_size([title], box.width, box.height, 40, min_pt=24)
    p = _para(tf, True)
    _add_runs(p, title, theme, role="head", size=size, bold=True,
              color=c["title_text"])
    if subtitle:
        # place the subtitle below however many lines the title wrapped to
        n_lines = le.wrapped_line_count(title, size, (box.width - 0.2) * 72)
        sbox = le.Box(0.92, box.top + n_lines * size / 72 * 1.3 + 0.3,
                      le.SLIDE_W - 2.4, 1.0)
        _, tf2 = _textbox(slide, sbox)
        ssize = le.fit_font_size([subtitle], sbox.width, sbox.height, 18)
        p2 = _para(tf2, True)
        _add_runs(p2, subtitle, theme, size=ssize, color=c["title_sub"])
    if footer:
        _, tf3 = _textbox(slide, le.Box(0.92, le.SLIDE_H - 0.75,
                                        le.SLIDE_W - 2.4, 0.4))
        p3 = _para(tf3, True)
        r = p3.add_run()
        r.text = footer
        _set_run(r, theme, size=11, color=c["title_sub"])


def _render_two_column(slide, theme: dict, s: Slide) -> None:
    c = theme["colors"]
    for box, section in zip(le.two_column_boxes(), s.sections):
        _rect(slide, box, c["card_bg"], line_rgb=c["card_line"])
        head_box = le.Box(box.left + 0.3, box.top + 0.28, box.width - 0.6, 0.55)
        _, tf = _textbox(slide, head_box)
        hsize = le.fit_font_size([section.heading], head_box.width,
                                 head_box.height, 19, min_pt=14)
        p = _para(tf, True)
        _add_runs(p, section.heading, theme, role="head", size=hsize,
                  bold=True, color=c["primary"])
        body_box = le.Box(box.left + 0.3, box.top + 1.0,
                          box.width - 0.6, box.height - 1.3)
        _, tf2 = _textbox(slide, body_box)
        _items_into(tf2, section.items, theme, body_box, 15,
                    lead_color=c["text"])


def _render_three_cards(slide, theme: dict, s: Slide) -> None:
    c = theme["colors"]
    for i, (box, item) in enumerate(zip(le.card_boxes(3), s.bullets), 1):
        _rect(slide, box, c["card_bg"], line_rgb=c["card_line"])
        d = 0.52
        circle = le.Box(box.left + 0.3, box.top + 0.32, d, d)
        sp = _rect(slide, circle, c["accent"], rounded=True, radius=0.5)
        tf = sp.text_frame
        tf.word_wrap = False
        tf.margin_left = tf.margin_right = 0
        tf.margin_top = tf.margin_bottom = 0
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = str(i)
        _set_run(r, theme, role="head", size=20, bold=True,
                 color=(255, 255, 255))
        head_box = le.Box(box.left + 0.3, box.top + 1.05, box.width - 0.6, 0.9)
        _, tfh = _textbox(slide, head_box)
        hsize = le.fit_font_size([item.lead or ""], head_box.width,
                                 head_box.height, 17, min_pt=13)
        ph = _para(tfh, True)
        _add_runs(ph, item.lead or "", theme, role="head", size=hsize,
                  bold=True, color=c["primary"])
        body_box = le.Box(box.left + 0.3, box.top + 1.95,
                          box.width - 0.6, box.height - 2.25)
        _, tfb = _textbox(slide, body_box)
        bsize = le.fit_font_size([item.body], body_box.width,
                                 body_box.height, 14)
        pb = _para(tfb, True)
        pb.line_spacing = 1.2
        _add_runs(pb, item.body, theme, size=bsize)


def _render_timeline(slide, theme: dict, s: Slide) -> None:
    c = theme["colors"]
    geo = le.timeline_geometry(len(s.steps))
    _rect(slide, geo["line"], c["card_line"], rounded=False)
    for i, (circle, tbox, step) in enumerate(
            zip(geo["circles"], geo["texts"], s.steps), 1):
        sp = _rect(slide, circle, c["accent"] if i < len(s.steps)
                   else c["primary"], rounded=True, radius=0.5)
        tf = sp.text_frame
        tf.margin_left = tf.margin_right = 0
        tf.margin_top = tf.margin_bottom = 0
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = str(i)
        _set_run(r, theme, role="head", size=22, bold=True,
                 color=(255, 255, 255))
        _, tft = _textbox(slide, tbox)
        lead = step.lead or step.body
        body = step.body if step.lead else ""
        lsize = le.fit_font_size([lead], tbox.width, 0.6, 16, min_pt=12)
        pl = _para(tft, True)
        pl.alignment = PP_ALIGN.CENTER
        _add_runs(pl, lead, theme, role="head", size=lsize, bold=True,
                  color=c["primary"])
        if body:
            bsize = le.fit_font_size([body], tbox.width, tbox.height - 0.6, 13)
            pb = tft.add_paragraph()
            pb.alignment = PP_ALIGN.CENTER
            pb.space_before = Pt(6)
            pb.line_spacing = 1.15
            _add_runs(pb, body, theme, size=bsize, color=c["text"])


def _render_bullets(slide, theme: dict, s: Slide) -> None:
    c = theme["colors"]
    box = le.bullets_box()
    _, tf = _textbox(slide, box)
    lines: list[Item] = list(s.bullets)
    if s.paragraphs:
        lines = [Item(body=p) for p in s.paragraphs] + lines
    for sec in s.sections:                     # 1 or >2 sections: flatten
        lines.append(Item(body=sec.heading, lead=None))
        lines.extend(sec.items)
    _items_into(tf, lines, theme, box, 17, lead_color=c["text"])


# ------------------------------------------------------------------ deck ---
_RENDERERS = {
    le.LAYOUT_TWO_COLUMN: _render_two_column,
    le.LAYOUT_THREE_CARDS: _render_three_cards,
    le.LAYOUT_TIMELINE: _render_timeline,
    le.LAYOUT_BULLETS: _render_bullets,
}


def render_deck(deck: Deck, theme: dict, output_path: str,
                footer_note: str = "", font_target: str = "auto") -> list[str]:
    """Render the deck and return the layout name used for each slide."""
    theme = resolve_fonts(theme, font_target)
    prs = Presentation()
    prs.slide_width = Emu(int(le.SLIDE_W * 914400))
    prs.slide_height = Emu(int(le.SLIDE_H * 914400))
    blank = prs.slide_layouts[6]
    used: list[str] = []

    title_slide = prs.slides.add_slide(blank)
    _render_statement(title_slide, theme, deck.title, deck.subtitle,
                      footer=footer_note)
    used.append("title")

    for s in deck.slides:
        slide = prs.slides.add_slide(blank)
        layout = le.detect_layout(s)
        used.append(layout)
        if layout == le.LAYOUT_STATEMENT:
            _render_statement(slide, theme, s.title)
            continue
        _background(slide, theme["colors"]["bg"])
        if s.title:
            _slide_title(slide, theme, s.title)
        _RENDERERS[layout](slide, theme, s)
        _footer(slide, theme, deck.title, len(used))

    prs.save(output_path)
    return used

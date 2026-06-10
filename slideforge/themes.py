"""Theme definitions for SlideForge.

Each theme is a plain dict so users can add their own without touching the
rendering code.  Colors are (R, G, B) tuples.  Fonts come in pairs: a Latin
typeface plus an East-Asian (``*_ea``) typeface so Japanese/Chinese/Korean
text renders with a proper CJK font instead of a fallback.

East-Asian fonts are written as *logical* names (``gothic``, ``round_gothic``,
``mincho``) and resolved to a concrete typeface per target OS at render time
(see ``FONT_SCHEMES`` / ``resolve_fonts``), because a .pptx can only embed one
typeface name per run and Windows/macOS ship different Japanese fonts.

Color roles
-----------
bg          slide background for content slides
text        default body text
muted       captions, footers, de-emphasized text
primary     slide titles / dominant brand color (60-70% visual weight)
accent      sharp accent for numbers, card titles, step circles
card_bg     fill for column/card containers
card_line   subtle outline for cards (None = no outline)
title_bg    background of the title / closing "statement" slides
title_text  text color on title_bg
title_sub   subtitle color on title_bg
"""

THEMES = {
    # Trustworthy navy for tech / SaaS proposals.
    "TechBlue": {
        "colors": {
            "bg": (255, 255, 255),
            "text": (31, 36, 48),
            "muted": (107, 114, 128),
            "primary": (30, 39, 97),       # midnight navy
            "accent": (45, 108, 223),      # electric blue
            "card_bg": (238, 243, 252),    # ice blue tint
            "card_line": (202, 220, 252),
            "title_bg": (30, 39, 97),
            "title_text": (255, 255, 255),
            "title_sub": (202, 220, 252),
        },
        "fonts": {
            "head": "Trebuchet MS",
            "body": "Calibri",
            "head_ea": "gothic",
            "body_ea": "gothic",
        },
    },

    # Quiet charcoal monochrome for formal / executive material.
    "MinimalGray": {
        "colors": {
            "bg": (255, 255, 255),
            "text": (33, 33, 33),
            "muted": (138, 143, 152),
            "primary": (54, 69, 79),       # charcoal
            "accent": (33, 33, 33),        # near-black
            "card_bg": (242, 242, 242),
            "card_line": (224, 224, 224),
            "title_bg": (54, 69, 79),
            "title_text": (255, 255, 255),
            "title_sub": (208, 213, 219),
        },
        "fonts": {
            "head": "Georgia",
            "body": "Calibri",
            "head_ea": "mincho",
            "body_ea": "gothic",
        },
    },

    # Terracotta & sand for creative / brand-flavored decks.
    "WarmCreative": {
        "colors": {
            "bg": (255, 255, 255),
            "text": (59, 47, 42),
            "muted": (154, 143, 136),
            "primary": (184, 80, 66),      # terracotta
            "accent": (80, 128, 142),      # slate teal counterpoint
            "card_bg": (243, 239, 228),    # light sand
            "card_line": (231, 232, 209),
            "title_bg": (122, 59, 49),     # deep terracotta
            "title_text": (255, 248, 243),
            "title_sub": (231, 232, 209),
        },
        "fonts": {
            "head": "Trebuchet MS",
            "body": "Calibri",
            "head_ea": "round_gothic",
            "body_ea": "gothic",
        },
    },
}


# Logical East-Asian font names → concrete typefaces per target OS.
# "win" uses fonts bundled with Windows 8.1+; "mac" uses macOS-bundled
# Hiragino.  PowerPoint substitutes sensibly when a deck crosses platforms,
# but baking the right name for the audience's OS gives the intended look.
FONT_SCHEMES = {
    "win": {
        "gothic": "Yu Gothic",
        "round_gothic": "Yu Gothic",   # Windows has no bundled round gothic
        "mincho": "Yu Mincho",
    },
    "mac": {
        "gothic": "Hiragino Kaku Gothic ProN",
        "round_gothic": "Hiragino Maru Gothic ProN",
        "mincho": "Hiragino Mincho ProN",
    },
}


def resolve_fonts(theme: dict, target: str = "auto") -> dict:
    """Return a copy of ``theme`` with logical EA font names replaced by
    concrete typefaces for ``target`` ("win", "mac", or "auto" = the OS the
    generator is running on).  Concrete names pass through untouched, so
    custom themes may pin an exact font."""
    if target == "auto":
        import sys
        target = "mac" if sys.platform == "darwin" else "win"
    try:
        scheme = FONT_SCHEMES[target]
    except KeyError:
        known = ", ".join(sorted(FONT_SCHEMES) + ["auto"])
        raise KeyError(f"Unknown font target '{target}'. Use one of: {known}") from None
    fonts = dict(theme["fonts"])
    for key in ("head_ea", "body_ea"):
        fonts[key] = scheme.get(fonts[key], fonts[key])
    return {**theme, "fonts": fonts}


def get_theme(name: str) -> dict:
    """Return a theme dict, raising a friendly error on unknown names."""
    try:
        return THEMES[name]
    except KeyError:
        known = ", ".join(sorted(THEMES))
        raise KeyError(f"Unknown theme '{name}'. Available themes: {known}") from None

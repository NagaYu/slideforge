"""SlideForge theme definitions.

Each theme is a plain dict so that users can add their own theme by simply
appending an entry to ``THEMES`` (or passing a dict of the same shape to the
renderer).  All colors are ``(R, G, B)`` tuples.

Theme schema
------------
colors:
    primary        dominant brand color (titles, emphasis)
    secondary      supporting tone (card fills, sub headers)
    accent         sharp accent (numbers, highlights)
    bg             content-slide background
    bg_dark        title/closing-slide background ("sandwich" structure)
    text           body text on light background
    text_inverse   body text on dark background
    muted          captions / footers
    card_bg        background fill for cards & boxes
    card_border    border for cards & boxes
fonts:
    title / body   font names (header font with personality + clean body)
rules:
    dark_title_slide   render the first and last slide on ``bg_dark``
    card_corner_radius 0.0-1.0 relative corner rounding for cards
    bullet_char        character used for custom bullets
"""

THEMES = {
    # -- Cool, confident navy + ice blue. For tech / SaaS proposals. ------
    "TechBlue": {
        "display_name": "Tech Blue",
        "colors": {
            "primary": (30, 39, 97),        # deep navy
            "secondary": (28, 114, 147),    # teal blue
            "accent": (2, 195, 154),        # mint
            "bg": (255, 255, 255),
            "bg_dark": (16, 22, 58),
            "text": (40, 45, 60),
            "text_inverse": (236, 242, 252),
            "muted": (120, 130, 150),
            "card_bg": (240, 245, 252),
            "card_border": (202, 220, 252),
        },
        "fonts": {"title": "Trebuchet MS", "body": "Calibri"},
        "rules": {
            "dark_title_slide": True,
            "card_corner_radius": 0.12,
            "bullet_char": "▪",  # small square
        },
    },
    # -- Quiet charcoal minimalism. For reports / formal documents. -------
    "MinimalGray": {
        "display_name": "Minimal Gray",
        "colors": {
            "primary": (54, 69, 79),        # charcoal
            "secondary": (110, 120, 130),
            "accent": (33, 33, 33),
            "bg": (255, 255, 255),
            "bg_dark": (33, 37, 41),
            "text": (50, 55, 60),
            "text_inverse": (242, 242, 242),
            "muted": (150, 155, 160),
            "card_bg": (244, 245, 246),
            "card_border": (220, 222, 224),
        },
        "fonts": {"title": "Georgia", "body": "Calibri"},
        "rules": {
            "dark_title_slide": False,       # stays light throughout
            "card_corner_radius": 0.0,       # sharp corners
            "bullet_char": "–",         # en dash
        },
    },
    # -- Warm terracotta + sage. For creative pitches / branding decks. ---
    "WarmCreative": {
        "display_name": "Warm Creative",
        "colors": {
            "primary": (184, 80, 66),       # terracotta
            "secondary": (167, 190, 174),   # sage
            "accent": (47, 60, 126),        # navy pop
            "bg": (255, 255, 255),
            "bg_dark": (94, 41, 35),
            "text": (60, 50, 45),
            "text_inverse": (250, 244, 240),
            "muted": (155, 140, 130),
            "card_bg": (250, 243, 238),
            "card_border": (231, 215, 203),
        },
        "fonts": {"title": "Palatino", "body": "Calibri"},
        "rules": {
            "dark_title_slide": True,
            "card_corner_radius": 0.25,      # soft, friendly corners
            "bullet_char": "●",         # filled circle
        },
    },
}


def get_theme(name: str) -> dict:
    """Return a theme dict by name (case-insensitive). Raises KeyError."""
    for key, theme in THEMES.items():
        if key.lower() == name.lower():
            return {**theme, "name": key}
    raise KeyError(
        f"Unknown theme '{name}'. Available: {', '.join(THEMES)}"
    )


def list_themes() -> list[str]:
    return list(THEMES.keys())

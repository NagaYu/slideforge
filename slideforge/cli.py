"""SlideForge command line interface.

Examples
--------
    slideforge build proposal.md --theme TechBlue
    slideforge build proposal.md --all-themes -o build/
    slideforge themes
    slideforge sample > proposal.md
"""

from __future__ import annotations

import argparse
import pathlib
import sys

from .parser import parse_markdown
from .renderer import Renderer
from .sample import SAMPLE_MARKDOWN
from .themes import get_theme, list_themes


def _build(md_path: str, theme_name: str, out: str | None) -> str:
    text = pathlib.Path(md_path).read_text(encoding="utf-8")
    slides = parse_markdown(text)
    if not slides:
        sys.exit("error: no slides found in markdown input")
    theme = get_theme(theme_name)

    if out and out.endswith(".pptx"):
        out_path = pathlib.Path(out)
    else:
        stem = pathlib.Path(md_path).stem
        out_dir = pathlib.Path(out) if out else pathlib.Path(".")
        out_path = out_dir / f"{stem}_{theme['name']}.pptx"
    out_path.parent.mkdir(parents=True, exist_ok=True)

    Renderer(theme).render(slides, str(out_path))
    return str(out_path)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        prog="slideforge",
        description="Generate professional .pptx decks from Markdown.")
    sub = parser.add_subparsers(dest="command", required=True)

    p_build = sub.add_parser("build", help="build a .pptx from markdown")
    p_build.add_argument("input", help="markdown file")
    p_build.add_argument("-t", "--theme", default="TechBlue",
                         help="theme name (see `slideforge themes`)")
    p_build.add_argument("-o", "--output", default=None,
                         help="output .pptx path or directory")
    p_build.add_argument("--all-themes", action="store_true",
                         help="render one deck per available theme")

    sub.add_parser("themes", help="list available themes")
    sub.add_parser("sample", help="print a sample markdown document")

    args = parser.parse_args(argv)

    if args.command == "themes":
        for name in list_themes():
            theme = get_theme(name)
            print(f"{name:14s} {theme['display_name']}  "
                  f"(title font: {theme['fonts']['title']})")
        return 0

    if args.command == "sample":
        print(SAMPLE_MARKDOWN)
        return 0

    themes = list_themes() if args.all_themes else [args.theme]
    for name in themes:
        path = _build(args.input, name, args.output)
        print(f"✓ {name}: {path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

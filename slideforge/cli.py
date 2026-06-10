"""SlideForge CLI.

Examples
--------
    slideforge deck.md --theme TechBlue -o deck.pptx
    slideforge deck.md --all-themes -o out/
    slideforge deck.md --fonts win        # bake Windows Japanese fonts
    slideforge --write-sample samples/proposal.md
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from . import __version__
from .md_parser import parse_markdown
from .renderer import render_deck
from .themes import THEMES, get_theme

SAMPLE_MD = """\
# DXアクセラレーション・プログラムのご提案

業務プロセスの自動化とデータ活用基盤の構築による、次世代の働き方改革

---

# 本日のアジェンダ

- **背景と目的:** 市場環境の変化と、貴社が直面する業務効率化の課題を整理します
- 現状業務の分析から導いた、解決すべき3つのボトルネック
- 提案ソリューションの全体像と、それぞれの導入効果の試算
- 段階的な導入ロードマップと、初期フェーズの体制・スケジュール

---

# 現状の課題と解決アプローチ

## 現状の課題

- **属人化:** 主要業務の手順が担当者の経験に依存しており、引き継ぎコストが増大しています
- **二重入力:** 部門間でシステムが分断され、同じデータを複数回入力する作業が常態化しています
- **可視性の欠如:** 経営層がリアルタイムの業務指標を把握できず、意思決定が遅延しています

## 解決アプローチ

- **プロセス標準化:** 業務フローを棚卸しし、自動化可能な定型作業を特定します
- **システム連携基盤:** APIハブを構築し、既存システム間のデータを自動同期します
- **ダッシュボード:** 主要KPIをリアルタイムに可視化し、週次の意思決定サイクルを実現します

---

# 3つの提供価値

- **業務時間の削減:** 定型業務の自動化により、対象部門の作業時間を年間約4,200時間削減。創出された時間を企画・分析業務へ再配分できます
- **データ品質の向上:** 入力の一元化と自動検証により、手作業に起因する転記ミスを排除。監査対応の工数も大幅に軽減します
- **意思決定の高速化:** 経営ダッシュボードで全社指標を即時把握。月次だった振り返りサイクルを週次へ短縮します

---

# 導入ロードマップ

1. **現状分析:** 業務ヒアリングとデータ棚卸しを実施(第1〜4週)
2. **設計:** 自動化対象の選定と基盤アーキテクチャ設計(第5〜8週)
3. **構築:** API連携基盤とダッシュボードの段階的構築(第9〜16週)
4. **定着化:** 利用トレーニングと運用ルールの整備(第17〜20週)

---

# ご清聴ありがとうございました
"""


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="slideforge",
        description="Generate professional .pptx decks from Markdown.",
    )
    p.add_argument("input", nargs="?", help="input Markdown file")
    p.add_argument("-o", "--output",
                   help="output .pptx path (or directory with --all-themes)")
    p.add_argument("-t", "--theme", default="TechBlue",
                   help=f"theme name ({', '.join(sorted(THEMES))})")
    p.add_argument("--all-themes", action="store_true",
                   help="render one deck per available theme")
    p.add_argument("--fonts", default="auto", choices=("auto", "win", "mac"),
                   help="Japanese font target: 'win' bakes Yu Gothic/Mincho, "
                        "'mac' bakes Hiragino, 'auto' matches this machine "
                        "(default). Pick the OS your audience opens the "
                        "deck on.")
    p.add_argument("--version", action="version",
                   version=f"%(prog)s {__version__}")
    p.add_argument("--footer", default="",
                   help="footer note shown on the title slide")
    p.add_argument("--write-sample", metavar="PATH",
                   help="write the bundled sample proposal Markdown and exit")
    return p


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)

    if args.write_sample:
        path = Path(args.write_sample)
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(SAMPLE_MD, encoding="utf-8")
        print(f"sample written: {path}")
        return 0

    if not args.input:
        print("error: input Markdown file is required", file=sys.stderr)
        return 2

    md_path = Path(args.input)
    deck = parse_markdown(md_path.read_text(encoding="utf-8"))

    if args.all_themes:
        out_dir = Path(args.output or ".")
        out_dir.mkdir(parents=True, exist_ok=True)
        for name in THEMES:
            out = out_dir / f"{md_path.stem}_{name}.pptx"
            layouts = render_deck(deck, get_theme(name), str(out), args.footer,
                                  font_target=args.fonts)
            print(f"{out}  [{name}]  slides: {len(layouts)} "
                  f"({', '.join(layouts)})")
        return 0

    out = Path(args.output or md_path.with_suffix(".pptx"))
    layouts = render_deck(deck, get_theme(args.theme), str(out), args.footer,
                          font_target=args.fonts)
    print(f"{out}  [{args.theme}]  slides: {len(layouts)} "
          f"({', '.join(layouts)})")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

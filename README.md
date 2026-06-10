# SlideForge ⚒️

**Markdown → プロ品質の PowerPoint (.pptx) を一発生成する CLI ツール**

Markdown を書くだけで、テーマカラー・フォント・構図（2カラム / カード型 / タイムライン）を自動で判定・適用した 16:9 のスライドが手に入ります。python-pptx ベース、依存は1つだけ。

```
# 業務効率化のご提案          ┐
チームの時間を、創造的な仕事へ。 │──▶  ダークヒーロー表紙
                              ┘
## 従来 … / ## 導入後 …       ──▶  左右分割レイアウト
- A：… / - B：… / - C：…      ──▶  3カラムカード
1. 分析 2. PoC 3. 展開 …      ──▶  ステップ・タイムライン
```

| | |
|---|---|
| ![表紙 (TechBlue)](docs/images/title-techblue.jpg) | ![2カラム (TechBlue)](docs/images/two-column-techblue.jpg) |
| ![カード (WarmCreative)](docs/images/cards-warmcreative.jpg) | ![タイムライン (MinimalGray)](docs/images/timeline-minimalgray.jpg) |

生成済みのサンプル一式は [examples/](examples/) にあります。

## インストール

```bash
git clone https://github.com/NagaYu/slideforge.git
cd slideforge
pip install .          # または pipx install .
```

Python 3.10+ / 依存: `python-pptx`

## 使い方

```bash
# サンプルの提案書 Markdown を出力（samples/汎用ビジネス提案書.md と同内容）
slideforge sample > proposal.md

# テーマを指定してビルド
slideforge build proposal.md --theme TechBlue

# 全テーマ分を一括生成（テーマ比較に便利）
slideforge build proposal.md --all-themes -o build/

# 利用可能テーマの一覧
slideforge themes
```

出力は `<入力ファイル名>_<テーマ名>.pptx`。`-o` に `.pptx` パスを渡せばファイル名も指定できます。

## Markdown の書き方

| 記法 | 結果 |
|------|------|
| `# 見出し` | 新しいスライド（タイトル） |
| `---` | スライド区切り |
| タイトル直後の行 | サブタイトル（表紙では中央寄せイタリック） |
| `## 見出し` ×2 | **左右分割レイアウト**に自動変換 |
| 並列な箇条書き ×3〜4 | **カード型レイアウト**（3カラム / 2×2グリッド） |
| `1.` `2.` … 番号付きリスト | **ステップ・タイムライン**（6個以上は2段組） |
| `- 項目` / 2スペース字下げ | 箇条書き / サブ箇条書き |
| `> 引用` | アクセントカラーのイタリック表示 |
| `**強調**` | 太字 |

カードの見出しは `- **タイトル**：本文` のように「：」や `:` で区切ると、タイトル＋本文に自動分割されます。

### レイアウト自動判定のルール

1. 番号付きリストが2つ以上 → `timeline`
2. `##` がちょうど2つ → `two_column`
3. 並列箇条書きが3〜4個（見出しなし） → `cards`
4. 本文なし＋先頭/末尾スライド → `title` / `closing`（ヒーロー）
5. それ以外 → `content`（タイトル＋箇条書き）

## テーマ

| テーマ | 雰囲気 | タイトルフォント | 特徴 |
|--------|--------|------------------|------|
| `TechBlue` | ネイビー×ミント | Trebuchet MS | ダーク表紙、角丸カード |
| `MinimalGray` | チャコール無彩色 | Georgia | 全編ライト、直角カード |
| `WarmCreative` | テラコッタ×セージ | Palatino | 大きな角丸、温かみ |

### 文字溢れ防止（Auto-fit）

テキストがボックスに収まらない場合、フォントサイズを **2pt ずつ自動で縮小**します（最小9pt、エラーは出しません）。複数の文字サイズが混在するボックスは、視覚的な階層を保ったまま全体を均等にスケールダウンします。CJK文字は全角幅で計算されるため、日本語でも溢れません。

## 拡張方法

### テーマを追加する

[slideforge/themes.py](slideforge/themes.py) の `THEMES` 辞書にエントリを1つ追加するだけです。

```python
THEMES["ForestGreen"] = {
    "display_name": "Forest Green",
    "colors": {
        "primary": (44, 95, 45),     # タイトル・強調
        "secondary": (151, 188, 98), # サブ見出し
        "accent": (245, 245, 245),   # 番号サークル等
        "bg": (255, 255, 255), "bg_dark": (24, 48, 24),
        "text": (40, 50, 40), "text_inverse": (240, 245, 240),
        "muted": (130, 140, 130),
        "card_bg": (243, 248, 240), "card_border": (210, 225, 200),
    },
    "fonts": {"title": "Cambria", "body": "Calibri"},
    "rules": {
        "dark_title_slide": True,   # 表紙をダーク背景に
        "card_corner_radius": 0.15, # カード角丸 (0=直角)
        "bullet_char": "▸",
    },
}
```

### レイアウトを追加する

1. [slideforge/layout_engine.py](slideforge/layout_engine.py) — `detect_layout()` に判定条件を追加し、ジオメトリ関数（インチ単位の `Rect` を返す）を書く
2. [slideforge/renderer.py](slideforge/renderer.py) — `render_<name>()` メソッドを追加し、`render()` のディスパッチに1行追加

レイアウトエンジンは python-pptx に依存しない純粋な座標計算なので、単体テストが容易です。

### アーキテクチャ

```
parser.py        Markdown → Slide/Block モデル（依存ゼロ）
layout_engine.py 構図判定 + インチ単位のジオメトリ計算（依存ゼロ）
autofit.py       文字溢れ防止のフォントサイズ計算（依存ゼロ）
themes.py        配色・フォント・ルールの辞書（依存ゼロ）
renderer.py      ↑すべてを束ねて python-pptx で描画
cli.py           argparse ベースの CLI
```

## テスト

```bash
pip install -e ".[dev]"
pytest
```

パーサ・レイアウト判定・ジオメトリ・Auto-fit・全テーマのエンドツーエンド生成・**コンテンツ欠落ゼロ保証**（入力Markdownの全行が出力pptxに存在すること）を検証します。

## ライセンス

MIT

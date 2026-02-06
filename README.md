# excelmd

高忠実度で `.xlsx` を単一Markdownに変換し、LLMが解釈しやすい形にするPythonライブラリです。

## 特徴

- `.xlsx` 専用（初版）
- デフォルトは「作業ビュー向け」のMarkdownを出力
- `--sheetview` でシート見た目再現向け（HTMLテーブル）を出力
- `--full` で従来の全量・高忠実度ダンプを出力
- セル情報（座標、値、数式、計算済み値、スタイルID）を保持
- 結合セル、データ入力規則、印刷設定、定義名を保持
- 図形・画像・コネクタを抽出
- コネクタは幾何推定でノード接続を推定
- Mermaid (`flowchart TD`) を併記
- 画像は data URI でMarkdown内に埋め込み
- hiddenシートを含む全シート出力
- 未対応要素は警告しつつ raw XML を退避

## 要件

- Python 3.12
- [uv](https://github.com/astral-sh/uv)

## セットアップ

```bash
uv sync --python 3.12 --extra dev
```

## CLI

```bash
uv run --python 3.12 excel-md 入力.xlsx -o 出力.md
```

上記は作業ビュー（推奨）です。シートの内容を人が追いやすいように行ベースで展開します。

strictモード（未対応要素があると失敗）:

```bash
uv run --python 3.12 excel-md 入力.xlsx -o 出力.md --strict-unsupported
```

シート見た目再現モード（Markdown内にHTMLテーブルを埋め込み）:

```bash
uv run --python 3.12 excel-md 入力.xlsx -o 出力.md --sheetview
```

独立HTML出力モード（シートビュー再現HTMLを直接生成）:

```bash
uv run --python 3.12 excel-md 入力.xlsx -o 出力.html --html
```

全量モード（サイズ大、完全ダンプ）:

```bash
uv run --python 3.12 excel-md 入力.xlsx -o 出力.md --full
```

## Python API

```python
from excelmd.api import convert_xlsx_to_markdown, load_xlsx

markdown = convert_xlsx_to_markdown("入力.xlsx")
doc = load_xlsx("入力.xlsx")
print(doc.summary)
```

## テスト

```bash
uv run --python 3.12 pytest
```

## Docker

```bash
docker build -t excelmd .
docker run --rm -v "$PWD":/work -w /work excelmd 入力.xlsx -o 出力.md
```

## 出力構造

- `# Workbook: <filename>`
- `## Source Metadata`
- `## Styles (XML-equivalent)`
- `## Defined Names`
- `## Sheet: <name> [visible|hidden]`
- `### Sheet Metadata`
- `### Print Metadata`
- `### Data Validations`
- `### Cell Regions`
- `### Drawings Raw Objects`
- `### Connectors (Raw + Inferred)`
- `### Mermaid`
- `### Embedded Images`
- `### Unsupported Elements`
- `## Extraction Summary`
- `## Warnings`

## 制約

- `.xls`, `.xlsm` は未対応
- OCRは未実装（画像内文字は抽出しない）
- 出力Markdownは大きくなることがある（忠実性優先）

# pptx-template

[![Test](https://github.com/m3dev/pptx-template/actions/workflows/test.yml/badge.svg)](https://github.com/m3dev/pptx-template/actions/workflows/test.yml)
[![Python](https://img.shields.io/badge/python-3.10%20%7C%203.11%20%7C%203.12%20%7C%203.13-blue)](https://www.python.org/)
[![License](https://img.shields.io/badge/license-Apache%202.0-green)](LICENSE)

## 概要

pptx-template は pptx のテンプレートを元に、別途用意した JSON 中の文字列や CSV データを差し込んだ pptx を生成するツールです。

定型レポートなどで大量のグラフ付きスライドを作成する際の作業を代行してくれます。

  - テンプレートには "{sales.0.june.us}" のような形で JSON内の値を指す id を記入できます
  - Python 3.10+, pandas, python-pptx に依存しています
  - 扱う json や csv の 文字コードは utf-8 前提です

### テキスト置換

<img src="docs/01.png?raw=true" width="80%" />

### CSVインポート

<img src="docs/02.png?raw=true" width="80%" />

## はじめに

```bash
pip install pptx-template
echo '{ "slides": [ { "greeting" : "Hello!!" } ] }' > model.json

# "{greeting}" を含むテンプレートファイル (test.pptx) を用意してください

pptx_template --out out.pptx --template test.pptx --model model.json
```

## 開発

### 必要条件

- Python 3.10, 3.11, 3.12, または 3.13
- [uv](https://docs.astral.sh/uv/)（推奨）または pip

### インストール

```bash
git clone https://github.com/m3dev/pptx-template.git
cd pptx-template

# uvを使う場合（推奨）
uv sync --extra dev

# pipを使う場合
pip install -e ".[dev]"
```

### テスト実行

```bash
# uvを使う場合
uv run --extra dev pytest

# 特定のPythonバージョンで実行
uv run --python 3.13 --extra dev pytest

# pip インストール後
pytest
```

### コマンドラインで実行

```bash
# uvを使う場合
uv run pptx_template \
  --template test/data3/in.pptx \
  --model test/data3/in.xlsx \
  --out test/data3/out.pptx \
  --debug

# pip インストール後
pptx_template \
  --template test/data3/in.pptx \
  --model test/data3/in.xlsx \
  --out test/data3/out.pptx \
  --debug
```

### REPLで実行（開発時）

```bash
uv run python
```

```python
import sys
from importlib import reload
import pptx_template.cli as cli

# 実行引数設定
sys.argv = ['dummy.py', '--out', 'test/data3/out.pptx', '--template', 'test/data3/in.pptx', '--model', 'test/data3/in.xlsx', '--debug']

# 実行
cli.main()

# ソースコード変更後、リロードして再実行
reload(sys.modules.get('pptx_template.xlsx_model'))
reload(sys.modules.get('pptx_template.text'))
reload(sys.modules.get('pptx_template.table'))
reload(sys.modules.get('pptx_template.chart'))
reload(sys.modules.get('pptx_template.core'))
reload(sys.modules.get('pptx_template.cli'))
cli.main()
```

### リリース手順

1. featureブランチを作成する
2. 実装する
3. 全Pythonバージョン（3.10, 3.11, 3.12, 3.13）でテストを実行する
4. pushする
5. GitHub上でPull Requestを作成する
6. コードレビューを依頼する
7. QAを実施する
8. Pull Requestをマージする
9. PyPIにアップロードする（PyPIのリポジトリ管理者のみ可）

### PyPIへのアップロード

```bash
# ビルド
uv build

# PyPIにアップロード
uv publish
```

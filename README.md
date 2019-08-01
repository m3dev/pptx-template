# pptx-template [![Build Status](https://travis-ci.org/m3dev/pptx-template.svg?branch=master)](https://travis-ci.org/m3dev/pptx-template)

## Overview

pptx-template is a PowerPoint presentation builder.

This helps your routine reporting work that have many manual copy-paste from excel chart to powerpoint, or so.

  - Building a new powerpoint presentation file from a "template" pptx file which contains "id"
  - Import some strings and CSV data which is defined in a JSON config file or a Python dict
  - "id" in pptx template is expressed as a tiny DSL, like "{sales.0.june.us}"
  - requires python envirionment (2 or 3), pandas, python-pptx
  - for now, only UTF-8 encoding is supported for json, csv

### Text substitution
!<img src="docs/01.png?raw=true" width="80%" />

### CSV Import
!<img src="docs/02.png?raw=true" width="80%" />

## Japanese translation

pptx-template は pptx のテンプレートを元に、別途用意した JSON 中の文字列や CSV データを差し込んだ pptx を生成するツールです。

定型レポートなどで大量のグラフ付きスライドを作成する際の作業を代行してくれます。

  - テンプレートには "{sales.0.june.us}" のような形で JSON内の値を指す id を記入できます
  - python 2 または 3, pandas, pptx に依存しています
  - 扱う json や csv の 文字コードは utf-8 前提です

## Getting started

TBD

```
$ pip install pptx-template
$ echo '{ "slides": [ { "greeting" : "Hello!!" } ] }' > model.json

# prepare your template file (test.pptx) which contains "{greeting}" in somewhere

$ pptx-template --out out.pptx --template test.pptx --model model.json
```

## Development (Japanese)

ローカル開発環境構築（Mac）

1. pythonインストール

```
brew install pyenv

pyenv install 3.5.6
pyenv install 3.7.1
```

2. 必要なパッケージをインストール

```
# インストール済みバージョン確認
pyenv versions


# pythonバージョン切り替え ※各バージョンで実行必要
pyenv shell 3.5.6

# インストール
pip install numpy
pip install pytest
pip install -r requirements.txt
```

3. git clone

```
git clone https://github.com/m3dev/pptx-template.git
```

4. 環境変数設定

```
export PYTHONPATH={プロジェクトフォルダ}
```

5. REPLで実行 ※開発時はこの方法

pythonのREPLを起動

```
cd {プロジェクトフォルダ}
pyenv shell 3.7.1
python
```

REPL内で実行

```
import sys
from importlib import reload
import pptx_template.cli as cli


## 実行引数設定
## sys.argv = ['{pyファイル名}', '--out', '{出力pptxファイルパス}', '--template', '{テンプレートpptxファイルパス}', '--model', '{設定xlsxファイルパス}', '--debug']
## 以下の設定でテストファイルにて実行できます
sys.argv = ['dummy.py', '--out', 'test/data3/out.pptx', '--template', 'test/data3/in.pptx', '--model', 'test/data3/in.xlsx', '--debug']

## 実行
cli.main()

## 変更したソースをリロードして実行
reload(sys.modules.get('pptx_template.xlsx_model'))
reload(sys.modules.get('pptx_template.text'))
reload(sys.modules.get('pptx_template.table'))
reload(sys.modules.get('pptx_template.chart'))
reload(sys.modules.get('pptx_template.core'))
reload(sys.modules.get('pptx_template.cli'))
cli.main()
```

6. コマンドラインで実行 ※githubに上がっているものの動作確認をしたい場合はこの方法

```
## pptx_template --out {出力pptxファイルパス} --template {テンプレートpptxファイルパス} --model {設定xlsxファイルパス}  --debug
pptx_template --out test/data3/out.pptx --template test/data3/in.pptx --model test/data3/in.xlsx  --debug
```

7. テスト実行

```
pytest
```

## ロールアウト手順
1. featureブランチを作成する
2. 実装する
3. 全pythonバージョンでtestが動くようにする
4. pushする
5. github上でpull requestを作成する
6. コードレビューを依頼する
7. QAを実施する（QAする人は、上記ローカル環境構築が必要）
8. pll requestをマージする
9. PyPIにアップロードする（PyPIのリポジトリ管理者のみ可）

## PyPIへのアップロード手順
1. パッケージインストール

```
pip install wheel
pip install twine
```

2. コンパイル

```
python setup.py bdist_wheel
```

3. PyPIにアップロード

```
twine upload dist/*
```
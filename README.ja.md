# pptx-template [![Build Status](https://travis-ci.org/m3dev/pptx-template.svg?branch=master)](https://travis-ci.org/m3dev/pptx-template)

## Overview

pptx-template は pptx のテンプレートを元に、別途用意した JSON 中の文字列や CSV データを差し込んだ pptx を生成するツールです。

定型レポートなどで大量のグラフ付きスライドを作成する際の作業を代行してくれます。

  - テンプレートには "{sales.0.june.us}" のような形で JSON内の値を指す id を記入できます
  - python 3, pandas, pptx に依存しています
  - 扱う json や csv の 文字コードは utf-8 前提です

### Text substitution

<img src="docs/01.png?raw=true" width="80%" />

### CSV Import

<img src="docs/02.png?raw=true" width="80%" />

## Getting started

TBD

```
$ pip install pptx-template
$ echo '{ "slides": [ { "greeting" : "Hello!!" } ] }' > model.json

# prepare your template file (test.pptx) which contains "{greeting}" in somewhere

$ pptx-template --out out.pptx --template test.pptx --model model.json
```

## 開発工程

### インストールの流れ

pyenvをインストールしておく

```
git clone https://github.com/m3dev/pptx-template.git

pyenv install 3.7.1 # Pythonをインストール
pyenv shell 3.7.1 # シェルで使うPython

venv .venv # 開発用の仮想環境を作成
source .venv/bin/activate # 仮想環境を使用する

python setup.py develop        # パッケージを開発用にインストール＆依存パッケージをインストール
pip install -r requirements.txt # 開発用のパッケージをインストール
```

### REPLで実行 ※開発時はこの方法

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

### コマンドラインで実行 ※githubに上がっているものの動作確認をしたい場合はこの方法

```
## pptx_template --out {出力pptxファイルパス} --template {テンプレートpptxファイルパス} --model {設定xlsxファイルパス}  --debug
pptx_template --out test/data3/out.pptx --template test/data3/in.pptx --model test/data3/in.xlsx  --debug
```

### テスト実行

```
pytest
```

### ロールアウト手順

1. featureブランチを作成する
2. 実装する
3. 全pythonバージョンでtestが動くようにする
4. pushする
5. github上でpull requestを作成する
6. コードレビューを依頼する
7. QAを実施する（QAする人は、上記ローカル環境構築が必要）
8. pll requestをマージする
9. PyPIにアップロードする（PyPIのリポジトリ管理者のみ可）

### PyPIへのアップロード手順

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

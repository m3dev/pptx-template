# ppt-template

ppt-template is an PowerPoint presentation builder.

This helps your routine reporting work that have many manual copy-paste from excel chart to powerpoint, or so.

  - Building a new powerpoint presentation file from a "template" pptx file which contains "id"
  - Import some strings and CSV data which is defined in a JSON config file or a Python dict
  - "id" in pptx template is expressed as a tiny DSL, like "{sales.0.june.us}"
  - requires python envirionment (2 or 3), pandas, python-pptx
  - for now, only UTF-8 encoding is supported for json, csv

below is Japanese transation:

ppt-template は pptx のテンプレートを元に、別途用意した JSON 中の文字列や CSV データを差し込んだ pptx を生成するツールです。

定型レポートなどで大量のグラフ付きスライドを作成する際の作業を代行してくれます。

  - テンプレートには "{sales.0.june.us}" のような形で JSON内の値を指す id を記入できます
  - python 2 または 3, pandas, pptx に依存しています
  - 扱う json や csv の 文字コードは utf-8 前提です

## Getting started

TBD

```
$ pip install pptx-template
$ echo '{ "slides": [ { "greeting" : "Hello!!" } ] }' > model.json
$ # prepare your template pptx which contains "{greering}" in somewhere
$ pptx-template --out out.pptx --template test.pptx --model model.json
```

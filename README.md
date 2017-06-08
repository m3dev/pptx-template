# ppt-template

ppt-template ver. 2 は pptx テンプレート内に直接記入した id に対して、別途用意した JSON オブジェクト中の文字列やCSVデータを差し込んだ pptx を生成するツールです

レポートなどで大量のグラフ付きスライドを作成する際の作業を代行してくれます。

  - python2.7|3.6, pandas, pptx に依存しています
  - 扱う json や csv の 文字コードは utf-8 前提です

## Python コードからの使い方の例

  - git clone する (以降は、仮に ~/git/pptx-template に clone したとします)
  - 自身が開発したいレポートに pptx-template というディレクトリを作成する
  - ``$ cp ~/git/pptx-template/pptx-template/*.py pptx-template``
  - テンプレートとしたい .pptx  ファイルを template.pptx という名前で用意します
    - 中に ``{greet}`` という内容のテキストエリアを作成しておきます
  - 自作のコードに以下のコードを追加

```
from pptx-template as pt
from pptx

ppt = pptx.Presentation('template.pptx')
slide = ppt.slides[0]
pt.edit_slide(slide, { "slides": [ { "greet": "Hello" } ] })
ppt.save('out.pptx')
```
  - pptx 導入済みの python2.7 環境から上記のコードを実行


## コマンドラインからの使い方

  - git clone する (以降は、仮に ~/git/pptx-template に clone したとします)
  - テンプレpptx を template.pptx という名前で用意します
    - 中に ``{greet}`` という内容のテキストエリアを作成しておきます
  - model.json という名前で以下のようなファイルを作ります
```
  $ cat > model.json
  { "slides": [ { "greet": "Hello" } ] }
  ^D
```
  
  - コマンドを実行
```
  python ~/git/pptx-template/pptx-template/cli.py --template template.pptx --out out.pptx --model model.json
```
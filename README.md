# csv2pptx ver.2

csv2pptx ver. 2 は pptx テンプレート内に直接記入した id に対して、別途用意した JSON オブジェクト中の文字列やCSVデータを差し込んだ pptx を生成するツールです

  - python2.7, pandas, pptx に依存しています
  - 扱う json や csv の 文字コードは utf-8 前提です

## Python コードからの使い方

  - git clone する (以降は、仮に ~/git/csv2pptx に clone したとします)
  - 開発したいレポートに csv2pptx というディレクトリを作成する
  - ``$ cp ~/git/csv2pptx/ver2/*.py csv2pptx``
  - テンプレpptx を template.pptx という名前で用意します
    - 中に ``{greet}`` という内容のテキストエリアを作成しておきます
  - 自作のコードに以下のコードを追加

```
from csv2pptx.ver2 as c2p
from pptx

ppt = pptx.Presentation('template.pptx')
slide = ppt.slides[0]
c2p.edit_slide(slide, { "slides": [ { "greet": "Hello" } ] })
ppt.save('out.pptx')
```
  - pptx 導入済みの python2.7 環境から上記のコードを実行


## 帳票環境コマンドラインからの使い方

  - git clone する (以降は、仮に ~/git/csv2pptx に clone したとします)
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
  /apps/python27/bin/python ~/git/csv2pptx/ver2/cli.py --template template.pptx --out out.pptx --model model.json
```
# model.json 仕様

## 例
```
{
  "slides:" [    // すべての要素は この slides 配下に 配列 として設定します。
                 // テンプレの先頭スライドから順に、一つずつ適用されます

                 // スライド#0 の設定
      "remove",  //  "remove" と文字列を書くと、そのスライドを削除します

                 // スライド#1 の設定
      {
        "greeting": "Hello",      // 任意のキー名で、流し込みたい文字列を指定できます
        "num": [ "100", "200" ]   // 設定したい文字列は配列やハッシュで構造を持つことも出来ます。 
                                  // この場合は ``{num.0}`` とテンプレ側に記入すれば最初の要素の値が入ります
      },
      
                 // スライド#2 の設定
      {
        "greeting": "Hola",
        "chart0": {                           // チャートタイトルに設定されたIDに対してはチャート用の設定を記入してください
          "file_name": "data-for-chart-csv"   // file_name : 読み込み対象 csv ファイル名
          "body": "Year,Sales,Cost\n2001,200,150",
                                              // body: CSVの内容を直接記入できます。file_name と両方指定した場合は body が優先します
          "value_axis_max": 100,              // value_axis_max: Y軸の最大値
          "value_axis_min": 200               // value_axis_max: Y軸の最小値
        }
      }
  ]
}

```

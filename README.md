# trans-excel
Translate English excel to Japanese excel file by using Amazon Translate.

以下、プログラム内で実施しているタスク概要です。
１）　オリジナルファイルを指定されたファイル名にコピー
２）　コピーされたファイルが、日本語訳を格納するターゲットファイル
３）　ターゲットファイルの各シートを１つづつ処理していきます。
４）　最初のシートの各セルを１つ１つREAD
５）　Amazon Translateを呼び出して各セル内の英文を翻訳
６）　翻訳した日本語文をもとのセルに上書き
７）　すべてのシートを日本語で書き換えたら、ファイルを保存して終了

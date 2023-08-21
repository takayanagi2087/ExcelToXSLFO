# ExcelToXSLFO

## Description
ExcelファイルをXSL-FOファイルに変換します。
Excelの印刷の完璧な再現はできませんが、とりあえずXSL-SOファイルのひな形を作る程度には使えるものになっています。
dataforms.jarでApache-FOPを利用したpdf出力やサーバからの直接印刷をサポートする予定なので、前段階で作成したツールになります。

## Install
Java8がインストールされたPCで、[リリース](https://github.com/takayanagi2087/ExcelToXSLFO/releases)からExcelToXSLFOxxx.zipをダウンロードし、適切なフォルダに展開してください。
[Apache POI](https://poi.apache.org/)をダウンロードし、「<ExcelToXSLFOxxx.zipのディレクトリ>/lib」にPOIに含まれる*.jarファイルをコピーします。
excel2xslfo.batまたはexcel2xslfo.shを使用し、ExcelファイルをXSL-FOファイル変換を行います。
コマンドラインは以下のようになります。

excel2xslfo [options] excelfile fofile
options:
-s sheetidx

## Demo
ExcelToXSLFOxxx.zip中のsample.xlsxとsample.foは以下のコマンドの実行結果です。

excel2xlsfo sample.xlsx sample.fo

## Licence
[MIT](https://github.com/takayanagi2087/dataforms/blob/master/LICENSE)






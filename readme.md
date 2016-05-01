Xls2XlsTemplator
===

#概要
Excel VBAにて開発したテンプレートエンジンです。<br>
テンプレートの各種設定とデータソースを設定することで、容易テンプレート及びデータが設定されたワークシートを出力することができます。<br>
テンプレートはヘッダー、アイテム、フッター領域でそれぞれ指定することができます。<br>

#詳細
本ライブラリは1つのクラスライブラリで完結できるよう簡易な設計になっています。<br>
そのため、ヘッダーおよびフッターにはデータを設定することができません。<br>
ヘッダーおよびフッターにデータを設定したい場合、本ライブラリの処理完了後に処理を行ってください。<br>

データの適用はアイテム領域のみ行えます。<br>
データソースにADODB.Recordsetを使用するため、VBAの参照設定で'Microsoft ActiveX Data Object 6.1'を参照するよう設定してください。

#インストール方法
Srcフォルダ内CXls2XlsTemplatorをダウンロードし、VBEでインポートしてください。

#動作確認環境<br>
* Windows10（x64)<br>
* Microsoft Excel2016(Office365)<br>
    
#ライセンス
Copyright (c) 2016 TskGunGun<br>
Released under the MIT license<br>
[MIT](https://github.com/tcnksm/tool/blob/master/LICENCE)

本ソフトウェアは自由な変更及び再配布を認めます。商用利用など自由に使ってください。<br>
再配布する場合、本ライセンス全文をソースコード、もしくはライセンス表示ファイルなどに掲載してください。<br>



                 　　　ICS to CSV コンバータ
                           ics2gacsv


                                                            2025/9/24
                                                (最終更新) 2025/12/12

                                                          by 松元隆二

*はじめに

スケジュールの標準フォーマットであるICS形式(別名iCalendar形式)をCSV形
式に変換するツール「ics2gacsv」です。

配布元は以下になります。

    https://qiita.com/qiitamatumoto/items/ab9e0cb9a6da257597a4
    https://github.com/githubmatumoto/ics2gacsv

ICS形式については下記を参照ください。

　　https://ja.wikipedia.org/wiki/ICalendar

以下のアプリが生成するICSファイルに対応していますが、主にWeb版 Outlook
を主眼に開発を行ってます。

       Cybozu Garoon(Version 5.0.2)
       Web版 Outlook
       Windowsアプリ版 Outlook(classic)
       いずれも2025年12月ごろに生成されたICSファイルの出力で確認。

出力するCSVは、グループウエアCybozu Garoon(Version 5.0.2)で作成したス
ケジュール表が生成するCSVと同じ形式にのみ対応しています。

**開発動機

職場ではグループウエアCybozu Garoonのスケジュールが生成するCSV形式で業
務記録簿を提出することになってます。グループウエアがOffice365に変更に
なる事になりました。移行期のため、CSV形式に変換する必要がありました。

内部向けツールですが、ICS to CSVを行いたい需要があるかもしれないので、
公開します。

資料内で「業務記録」という単語が多数出てきますが、同じ職場以外の皆様は、
本ツールで作成したCSV形式のファイルとお考えください。

なお、まだ実運用は行ってない模様。。これからバグ出しですね(^^;

*依存ライブラリ
ライブラリ vobjectを使わせていただいてます。バージョンは0.99で動作確認
しています。

 https://vobject.readthedocs.io/latest/
 https://py-vobject.github.io/
 https://github.com/py-vobject/vobject/releases

*ライセンス

ライセンスは Apache License 2.0です。

依存しているライブラリvobjectと同様のライセンスになります。

*動作環境およびインストール方法

以下から最新版をzipファイルもしくはtar.gzファイルでダウンロードして
展開してください。

    https://github.com/githubmatumoto/ics2gacsv/releases/

展開したファイルに含まれる

  INSTALL.txt

を参照して、Pythonの初期設定を行ってください。

ソフトウエアを展開したフォルダを覚えておいてください。ファイル
「INSTALL.txt」や「isc2gacsv.py」が含まれるフォルダです。

* 使い方

Web版のOutlookでいろいろスケジュールを作成してください。そのあと、ICS
形式でダウンロードしてファイルに保存してください。ICSのダウンロード手
順は下記の資料を参考にしてください。

  https://qiita.com/qiitamatumoto/items/24343d860ccc065b4cc8

を参照ください。

* CSVに変換する方法。
(事前にINSTALL.txtにそってインストールを行ってください。)

コマンドを利用する前に、``毎回''必要な初期や確認事項があります。

Linux/macOSは以下を実行してください。Pythonの初期化になります。

  $ source ~/.ics2gacsv/bin/activate

Windowsはライブラリvobjectを導入したpythonと同じであるか確認ください。

  > python3 --version

ICSをCSVに変換コマンドは下記の形式になります。

  $ cd "ソフトウエアを展開したフォルダ"
  $ python3 ics2gacsv.py 期間 入力ICSファイル 出力CSVファイル

実行例:

ソフトウエアを展開したフォルダにICSファイルをcalendar.icsという名前
で置いてください。

CSVに出力する「期間」は年4桁と数字2桁で指定してください。下記を実行を
すると、2025年9月分のスケジュールをschedules202509.csvに出力します。

  $ python3 ics2gacsv.py 202509 calendar.ics schedules202509.csv

変換に成功した場合

  INFO: 変換に成功しました: 'calendar.ics' to 'schedules202509.csv'

と表示されます。CSVファイルの文字コードはShiftJISになっています。

サンプルでCSVを生成するスクリプトを以下の名前で作成してます。

Windows用 : Windowsバッチファイル

  > samplegen.bat

  文字コードがShiftJISのCSVが生成されます。

Linux,macOS用: Shellスクリプト

  $ sh samplegen.sh

  文字コードがUTF-8とShiftJISのCSVが生成されます。Excelで閲覧や編集す
  る場合は文字コードがShiftJISのファイル(sjis-*.csv)を使ってください。

※セキュリティエラーが出たら中身を1行づつコピペして実行してください

その他、引数の詳細については引数「-h」で確認ください。

*ICSとCSVの要素の対応について
ICSとCSVの要素の対応についてはフォルダmiscに置いてある

   TECH-MEMO.txt

を参照ください。

*Known bugs
既知のバグがいろいろあります。フォルダmiscに置いてある

    TECH-MEMO.txt

を参照ください。

特に、業務で使う場合は充分な事前テストをお願いします。ほぼ同等のスケ
ジュールをWeb版OutlookとCybozu Garoonに記載し、比較調査お願いします。

またCybozu GaroonにはCSV以外にICSを出力する機能があります。GaroonでICS
を出力し、ics2gacsvでCSVに変換して、比較していただくのも良いかもしれま
せん。

特に

 - 終日スケジュール (時間指定がないスケジュール)
 - 繰り返しスケジュール
 - 繰り返しスケジュールの一部上書(RECURRENCE-ID)

でいろいろバグが出ました。一応既知の例はデバックしましたが、未知のバグ
があるかもしれません。

                                                              以上です。

            　　　ICS to CSV コンバータ
                       ics2gacsv


                                                            2025/9/24
                                                (最終更新) 2025/12/11

                                                          by 松元隆二

*はじめに

スケジュールの標準フォーマットであるICS形式(別名iCalendar形式)をCSV形
式に変換するツール「ics2gacsv」です。ICS形式については下記を参照くだ
さい。

　　https://ja.wikipedia.org/wiki/ICalendar

ICSファイルは下記に対応していますが、主にWeb版 Outlookを主眼に開発
を行ってます。
       Cybozu Garoon(Version 5.0.2)
       Web版 Outlook
       Windowsアプリ版 Outlook(classic)
       いずれも2025年12月ごろに生成されたICSファイルの出力で確認。

出力のCSVは、グループウエアCybozu Garoon(Version 5.0.2)で作成したスケ
ジュール表が生成するCSVと同じ形式にのみ対応しています。

**開発動機

職場はグループウエアCybozu GaroonのCSV形式で業務記録簿を提出することに
なってます。グループウエアがOffice365に変更になる事になりました。移行
期のため、CSV形式に変換する必要がありました。

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

*ライセンス

ライセンスは Apache License 2.0です。

依存しているライブラリvobjectと同様のライセンスになります。

*動作環境およびインストール方法

動作環境およびインストール方法は

  INSTALL.txt

を参照ください。

* 使い方

Web版のOutlookでいろいろスケジュールを作成してください。そのあと、ICS
形式でダウンロードしてファイルに保存してください。ICSのダウンロード手
順は下記の資料を参考にしてください。

  https://qiita.com/qiitamatumoto/items/24343d860ccc065b4cc8

を参照ください。

* CSVに変換する方法。

事前にINSTALL.txtにそってインストールを行ってください。

コマンドを実行する前に毎回必要な初期設定があります。

  (Linux/macOS)
  $ source ~/.ics2gacsv/bin/activate
  pythonのvenv環境の初期化になります。

  (Windows)
  $ python3 --version
  別資料INSTALL.txtに沿ってライブラリvobjectを導入したpythonであるか確
  認ください。

変換コマンドは下記の形式になります。

  $ python3 ics2gacsv.py 期間 入力ICSファイル 出力CSVファイル

実行例:

ICSファイルをダウンロードして、ファイルics2gacsv.pyが置いてある場所に
calendar.icsという名前で置いてください。CSVに出力する「期間」は年4桁と
数字2桁で指定してください。下記を実行をすると、2025年9月分のスケジュー
ルをschedules202509.csvに出力します。

  $ python3 ics2gacsv.py 202509 calendar.ics schedules202509.csv

成功した場合「変換に成功しました: 'calendar.ics' to 'schedules202509.csv'」
と表示されます。出力CSVファイルはShiftJISになっています。

引数「期間」は必須になっています。Web版のOutlookが生成するICSファイル
は出力する範囲の指定ができません。しかし業務記録としては生成するCSVの
日付の範囲の指定が必要になります。そのため、年と月を指定する引数を必須
としています。

サンプルでCSVを生成するスクリプトを以下の名前で作成して置いてます。

Linux,macOS用: samplegen.sh (Shellスクリプト)
Windows用    : samplegen.bat (Windowsバッチファイル)

Excelで閲覧や編集する場合は文字コードがShiftJISのファイル(sjis-*.csv)
を使ってください。その他、引数の詳細については引数「-h」で確認ください。

*ICSとCSVの要素の対応について
ICSとCSVの要素の対応については

   TECH-MEMO.txt

を参照ください。

*Known bugs
既知のバグがいろいろあります。

    TECH-MEMO.txt

を参照ください。

特に、業務で使う場合は充分な事前テストをお願いします。ほぼ同等のスケ
ジュールをOutlook(Web版)とCybozu Garoonに記載し、比較調査お願いします。

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

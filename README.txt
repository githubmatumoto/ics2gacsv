                 　　　ICS to CSV コンバータ


                                                         2025/9/24
                                                          by 松元隆二

*はじめに

スケジュールの標準フォーマットであるICS形式(別名iCalendar形式)をCSV形
式に変換するツールです。

　　https://ja.wikipedia.org/wiki/ICalendar

入力のICSは、Microsoft365のWeb版Outlookのスケジュール表が出力するICSで
のみ確認しています。

出力のCSVは、グループウエアCybozu Garoon(Version 5.0.2)で作成したスケ
ジュール表が生成するCSVと同じ形式にのみ対応しています。

補足1: Web版Outlookはスケジュールをファイルでダウンロードする場合はICS
形式のみ対応してます。(2025年9月24日現在)

補足2: Windowsのアプリ版Outlookの古い版はスケジュールをCSVで出力できま
す。

補足3: 有償版のMicrosoft365のWeb版Outlook)で主に開発しましたが、おそら
く無償版のWeb版Outlookでも動作します。

**開発動機

職場はグループウエアCybozu GaroonのCSV形式で業務記録簿を提出することに
なってます。グループウエアがOffice365に変更になる事になりました。移行
期のため、CSV形式に変換する必要がありました。

内部向けツールですが、ICS to CSVを行いたい需要があるかもしれないので、
公開します。

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
形式でダウンロードしてファイルに保存してください。

Web版のOutlookでICS形式でダウンロードの方法は

  DOWNLOAD-Outlook-ICS.pdf

を参照ください。

* CSVに変換する方法。

以下のような感じで変換できます。(コマンドpythonは環境によってはpython3
の場合があります)

 $ python ics2gacsv.py 期間指定 入力ICSファイル 出力CSVファイル

例:
DOWNLOAD-Outlook-ICS.pdfを参照の上、ICSファイルをダウンロードし、
calendar.icsという名前で置いてください。そして下記を実行をすると、
2025年9月分のスケジュールをkiroku202509.csvに出力します。

 $ python ics2gacsv.py 202509 calendar.ics kiroku202509.csv

Web版のOutlookが生成するICSファイルは出力する範囲の指定ができません。
しかし業務記録簿としては生成するCSVの日付の範囲の指定が必要になります。
そのため、年と月を指定する引数を必須としています。

引数の詳細については「-h」オプションで確認ください。

サンプルでCSVを生成するスクリプトを以下の名前で作成して置いてます。

 Linux用:   samplegen.sh
 Windows用: samplegen.bat

*ICSとCSVの要素の対応について
ICSとCSVの要素の対応については

   TECH-MEMO.txt
  
を参照ください。

*バグ
既知のバグがいろいろあります。

    TECH-MEMO.txt

を参照ください。

特に、業務で使う場合は充分な事前テストをお願いします。ほぼ同等のスケ
ジュールをWeb版のOutlookとCybozu Garoonに記載し、比較調査お願いします。

またCybozu GaroonにはCSV以外にICSを出力する機能があります。GaroonでICS
を出力し、本プログラムでCSVに変換して、比較していただくのも良いかもし
れません。

特に

 - 終日スケジュール (時間指定がないスケジュール)
 - 繰り返しスケジュール

でいろいろバグが出ました。具体例として2025年9月の毎週水曜日の終日スケ
ジュールは9月3,10,17,24日に行う事になりますが、最終水曜日9月24日が欠落
するという事が多々あります。一応既知の例はデバックしましたが、未知のバ
グがあるかもしれません。

                                                              以上です。

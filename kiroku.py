#!/usr/bin/env python3
# -*- python -*-
# -*- coding: utf-8 -*-
#
# Copyright (c) 2025 MATSUMOTO Ryuji.
# License: Apache License 2.0
#
import os
import sys
import re
import datetime
import libics2gacsv

__doc__=f"""ICS(iCalendar)をCSVに変換。作者の職場の業務記録提出用。

SYNOPSIS:

     $ python3 {sys.argv[0]} YoureName

     引数: YoureName: 業務記録提出者の氏名。

DESCRIPTION:
   0. 初めて利用する場合は、README.txt および INSTALL.txt を参照して、
      ライブラリの初期設定を行ってください。

   1. ソフトウエアics2gacsvを展開したフォルダにICSファイルを置く。
      ファイル名は「calendar.ics」で置いてください。

      Outlook(Web)からダウンロードするのを勧めます。Outlook(classic)か
      らダウンロードすると、大量の個人情報が含まれます。

   2. コマンドプロンプトで毎回必要な初期設定や確認事項の確認を行う。

     Linux/macOSは以下を実行してください。Pythonの初期化になります。

     $ source ~/.ics2gacsv/bin/activate

     Windowsはライブラリvobjectを導入したpythonと同じであるか確認する。

     > python3 --version

   3. コマンドプロンプトでソフトウエアを展開したフォルダに移動する。

     $ cd "ソフトウエアを展開したフォルダ"

   4. 下記コマンドを実行する。YoureNameには氏名を記載する

     $ python3 {sys.argv[0]} YoureName

     成功すると、CSV形式の業務記録が2個生成されます。実行した当月と、
     その前の月になります。

     例: 引数YoureNameに工大太郎を指定して2026年1月に実行すると、出力
     ファイル名は以下になります。

      schedules202512工大太郎.csv
      schedules202601工大太郎.csv

   5. CSVをExcelで確認し、個人のプライベートスケジュールが含まれてない
     か確認し、必要に応じて削除を行ってください。

注意事項:
　 ※出力ファイルと同じ名前のファイルがある場合は、確認なしに上書きします。
　　
   ※同梱されている「ics2gacsv.py」の簡易版になります。
   ※作者の職場の業務記録提出用のため頻繁に仕様が変更になります。

以上です。

=====================================================================

ライセンス:
Apache License 2.0

配布元:
{libics2gacsv.G_HAIFU_URL}
{libics2gacsv.G_GITHUB_URL}

修正履歴およびKnown bugsは misc/CHANGELOG.md 参照ください。
"""

########################################

def __myhelp(fname):
    help(fname)
    sys.exit()

if __name__ == '__main__':

    if libics2gacsv.G_VERSION != "2.0":
        print("ERROR: ファイルが古いです。最新のics2gacsv.pyとlibics2gacsv.pyをダウンロードしてください。",file=sys.stderr)
        sys.exit()

    #CSVを出力する期間指定。
    TIMERANGE=0

    exec_filename = os.path.basename(__file__)
    exec_filename = re.sub(r'\.py$', "", exec_filename)
    if len(sys.argv) != 2:
        __myhelp(exec_filename)

    YoureName = sys.argv[1]
    if YoureName == '-h':
        __myhelp(exec_filename)

    if (len(YoureName)) == 0:
        print("ERROR: 何らかの理由で名前の取得に失敗しました。",file=sys.stderr)
        sys.exit()

    # ファイル名に使えない記号検出
    # すでに同じような関数がありそうな気が若干するのだが(^_^;
    bad_char = False
    hava_space = False
    for i in list(YoureName):
        if i.isspace():
            bad_char = True
            hava_space = True
        if not i.isprintable():
            bad_char = True
        if i == '\\' or i == '/' or i == '[' or  i == ']':
            bad_char = True
        if i == ':' or  i == '<' or i == '>' or i == '?':
            bad_char = True
        if i == '"' or  i == "'" or i == "*" or i == '-':
            bad_char = True
    #

    if bad_char :
        print("ERROR: 引数YoureNameにファイル名として使えない文字が指定されました。", file=sys.stderr)
        if hava_space :
            print("ERROR: 引数YoureNameに空白が含まれてます。", file=sys.stderr)
        print(f"ERROR: 引数YoureName = '{YoureName}'", file=sys.stderr)

        sys.exit()
    #

    dt = datetime.datetime.now()

    if dt.month > 1:
        TIMERANGE = dt.year * 100 + (dt.month-1)
    else:
        TIMERANGE = (dt.year-1) * 100 + 12
    OUTPUT_CSV_FILENAME = 'schedules%06d%s.csv' % (TIMERANGE, YoureName)
    libics2gacsv.ics2csv("calendar.ics", OUTPUT_CSV_FILENAME, TIMERANGE)

    TIMERANGE = dt.year * 100 + dt.month
    OUTPUT_CSV_FILENAME = 'schedules%06d%s.csv' % (TIMERANGE, YoureName)
    libics2gacsv.ics2csv("calendar.ics", OUTPUT_CSV_FILENAME, TIMERANGE)

#End of main()

#!/usr/bin/env python3
# -*- python -*-
# -*- coding: utf-8 -*-
#
# Copyright (c) 2025 MATSUMOTO Ryuji.
# License: Apache License 2.0
#
import os
import sys
import getopt
import re
import time
import datetime
import libics2gacsv

__doc__=f"""ICS(iCalendar)をCSVに変換。作者の職場の業務記録提出用。

SYNOPSIS:

     $ python3 {sys.argv[0]} YoureName

     引数: YoureName: 業務記録提出者の氏名。

DESCRIPTION:

   ソフトウエアics2gacsvに同梱されるスクリプトの一つです。作者の職場の
   業務記録提出用のスクリプトです。入力/出力ファイル名が決め打ちになっ
   ています。

   本プログラムを初めて利用する場合は、README.txt および INSTALL.txt
   を参照して、ソフトウエアics2gacsvの初期設定を行ってください。

   以下は初期設定を行った後の手順書になります。

   1. ソフトウエアics2gacsvを展開したフォルダにICSファイルを置く。
      ファイル名は「calendar.ics」で置いてください。

      ICSファイルの取扱いには注意ください。特にTeams会議のパスワードが
      含まれます。

      ICSファイルはOutlook(Web)からダウンロードするのを勧めます。
      Outlook(classic)からダウンロードすると、Teams会議参加者のメール
      アドレスが含まれます。漏洩すると個人情報の流出となります。

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

    if libics2gacsv.G_VERSION != "2.1":
        print("ERROR: ファイルが古いです。最新のics2gacsv.pyとlibics2gacsv.pyをダウンロードしてください。",file=sys.stderr)
        sys.exit()

    #####################################################################
    # 入力ファイル名
    INPUT_ICS_FILENAME="calendar.ics"

    # key: 出力期間, value:出力ファイル名
    OUTPUT_CSV_FILENAME  = {}
    #
    # 出力ファイルがすでに存在する場合、上書きするかの確認をするかしないか。
    # False: 上書き確認を行う
    # True: 上書き確認を行わない。
    flag_overwrite= False
    # 入力ファイルの日付確認、古いファイルへの警告を行うかいなか。
    # False:  日付確認を行わない。
    # True: 日付確認を行う
    flag_old_file_check = True

    #####################################################################
    # 引数解析
    exec_filename = os.path.basename(__file__)
    exec_filename = re.sub(r'\.py$', "", exec_filename)


    opts, argv = getopt.getopt(sys.argv[1:], 'hw')
    try:
        for o, a in opts:
            if o == "-w":  # 出力ファイルの上書き確認/入力ファイルの日付確認を行わない
                flag_overwrite = True
                flag_old_file_check = False
            elif o == "-h":
                __myhelp(exec_filename)


        if len(argv) != 1:
            raise ValueError("ERROR: 引数が足りません。")
        YoureName = argv[0]
        if (len(YoureName)) == 0:
            raise ValueError("ERROR: 何らかの理由で引数YoureNameの取得に失敗しました。")

        # ファイル名に使えない記号検出
        # すでに同じような関数がありそうな気が若干するのだが(^_^;
        bad_char = ['\\', '/', '[', ']', ':', '<', '>', '?', '"', "'" ,"*",'-', '@']
        for i in list(YoureName):
            if i.isspace():
                raise ValueError("ERROR: 引数YoureNameに空白が含まれてます。")
            if not i.isprintable():
                raise ValueError("ERROR: 引数YoureNameにファイル名として使えない文字が含まれます。")

            if i in bad_char:
                raise ValueError(f"ERROR: 引数YoureNameにファイル名として使えない記号「{i}」が含まれます。")
    except ValueError as e:
        print("ERROR: ", e,  file=sys.stderr)
        print("ERROR:  引数 -h でヘルプが表示されます。", file=sys.stderr)
        sys.exit()

    #####################################################################
    #　業務記録番号の拡張仕様。
    libics2gacsv.flag_enhanced_gyoumunum = True

    #####################################################################
    # 出力ファイル名生成

    dt = datetime.datetime.now()

    if dt.month > 1:
        t = dt.year * 100 + (dt.month-1)
    else:
        t = (dt.year-1) * 100 + 12
    OUTPUT_CSV_FILENAME[t] = f'./schedules{t:06}{YoureName}.csv'

    t = dt.year * 100 + dt.month
    OUTPUT_CSV_FILENAME[t] = f'./schedules{t:06}{YoureName}.csv'

    #####################################################################
    # 出力ファイルに書き込み権限があるか確認。
    for fout in OUTPUT_CSV_FILENAME.values():
        if (not flag_overwrite) and os.path.exists(fout):
            print(f"WARNING: CSVファイル 「{fout}」 がすでに存在します。")
            print("WARNING: 上書きしますか?")
            inp = input('WARNING: [Y]es/[N]o? >> ').lower()
            if not (inp in ('y', 'yes')):
                print("WARNING: 処理を中断します。")
                sys.exit()

    for fout in OUTPUT_CSV_FILENAME.values():
        with open(fout, 'w',  encoding='utf-8') as f:
            f.write("WRITE TEST")
        os.remove(fout)

    #####################################################################
    # 入力ファイルの存在確認
    if not os.path.exists(INPUT_ICS_FILENAME):
        print(f"ERROR: 入力元のICSファイル「{INPUT_ICS_FILENAME}」が存在しません", file=sys.stderr)
        sys.exit()

    #####################################################################
    # 入力ファイルの日付確認。日付が古いファイルはファイル名が間違えたり
    # 最新のファイルが「calendar(2).ics」とかスペルミスの「calender.ics」
    # とかになってる可能性あるため。

    if flag_old_file_check:
        stat_info = os.stat(INPUT_ICS_FILENAME)
        current_time = time.time()
        diff_time = current_time - stat_info.st_mtime

        #print("mtime=", stat_info.st_mtime)
        #print("curre=", current_time)
        #print("diff =", diff_time)

        # 日付が未来。プログラム停止。
        if diff_time <= -60:
            print(f"ERROR: 入力元のICSファイル「{INPUT_ICS_FILENAME}」の日付が未来です。", file=sys.stderr)
            print("ERROR: パソコンの時計が正しくない可能性があります。", file=sys.stderr)
            sys.exit()

            # 日付が3時間以上前。プログラム停止。
        if diff_time >= 60*60*3:
            h = int(diff_time/3600 + 0.5)
            print(f"ERROR: 入力元のICSファイル「{INPUT_ICS_FILENAME}」の日付が古いです。", file=sys.stderr)
            print(f"ERROR: {h}時間前のファイルです。最新のファイルをダウンロードしてください。", file=sys.stderr)
            print("ERROR: 最新の場合はファイル名が間違えてないか確認ください。", file=sys.stderr)
            sys.exit()

        old_m = 15
        # 日付が15分以上前。警告のみ。
        if diff_time >= (old_m*60):
            print(f"WARNING: 入力元のICSファイル「{INPUT_ICS_FILENAME}」の日付が古いです。", file=sys.stderr)
            print(f"WARNING: {old_m}分以上前のファイルです。もし{old_m}分以内にダウンロードした場合は、", file=sys.stderr)
            print("WARNING: ファイル名が間違えてないか確認ください。", file=sys.stderr)

    #####################################################################
    # CSV変換
    # TODO: 単発で動かした場合と本プログラムで差分がないか確認
    # TODO: Windowsで確認。とくにファイルの日付確認。
    for key, value in OUTPUT_CSV_FILENAME.items():
        libics2gacsv.ics2csv(INPUT_ICS_FILENAME, value, key)

#End of main()

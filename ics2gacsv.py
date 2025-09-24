#!/usr/bin/env python3
# -*- python -*-
# -*- coding: utf-8 -*-
import os
import sys
import getopt
import re
import libics2gacsv

__doc__=f"""ICS(iCalendar)をCSVに変換する。CSVの出力形式はガルーンと同じ。

usage:  {sys.argv[0]} [-hubstodpm] 期間指定 入力.ics 出力.csv

期間指定  : CSVの出力年月の期間指定。

           例えば2025年8月分を出力する場合は「202508」と指定する。
           有効範囲は1970年1月以降「197001」。JSTで指定する。

           「all」もしくは「0」だと全部変換する。

          「guessin」だと入力ファイル名から期間を推測する。

          「guess」だと出力ファイル名から期間を推測する。例えば出力ファ
           イル名が「kiroku202509.csv」なら2025年9月と推測する。

           年月指定は月を超えてるスケジュールは翌月分も含まれます。例
           えば11月30日23:00に開始で12月1日02:00に終了の場合は、11月分
           に12月1日02:00終了のスケジュールが入ります。12月分には入り
           ません。

入力.ics    :　変換元のICSファイル名を指定。「stdin」を指定すると標準入力。

出力.csv    :　変換先のCSVファイル名を指定。「stdout」を指定すると標準出力。

-h         :  ヘルプを出力する。

-u         :  CSV出力をUTF-8にする。defaultはShiftJIS。(Garoonと同じ)

-b         : RRULE/EXDATEのバグフィックスを無効にする。
             defaultは有効。
             ※詳細は関数bug_fix_rrule()をみよ

-s         : ICSのsummaryの分割を無効にする。defaultは有効。
             分割するとCSVの「予定」/「予定詳細」がGaroonと同じ形式になる。
             ※詳細は関数split_garoon_style_summary()をみよ。

-t         : CSVの出力時刻にTimeZone情報を表示する。
             defaultは表示しない。(Garoonと同じ)
             ※詳細は関数ics_time_to_csv()をみよ。

-o         : CSVの出力時刻をWindowsの旧版Outlookが出力する形式にする。
             defaultの時間情報はGaroon形式。
             ※詳細は関数ics_time_to_csv()をみよ。

-d         : メモ欄(description)の4行目以降を消す。
　　　　　　defaultでは消さない。ただし、後述の「-p」が消す場合がある。
             ※詳細は関数modify_description()をみよ。

-p         : Teamsの会議インフォメーションを消さない。
             defaultでは消す。パスワードが入ってるため。
             ※詳細は関数modify_description()をみよ。

-m         : ICSのsummaryの分割で作者用の修正。defaultは無効。
             ※詳細は関数ics_time_to_csv()をみよ
"""

########################################

def myhelp():
    help(EXEC_FILENAME)
    help("libics2gacsv")
    sys.exit()
    
if __name__ == '__main__':
    #CSVを出力する期間指定。
    TIMERANGE=0

    EXEC_FILENAME=os.path.basename(__file__)
    EXEC_FILENAME = re.sub(r'\.py$', "", EXEC_FILENAME)
    opts, argv = getopt.getopt(sys.argv[1:], 'ubtsodpmh')

    try:
        for o, a in opts:
            if o == "-u":
                libics2gacsv.flag_output_sjis = False
                libics2gacsv.G_CSV_ENCODING = "utf-8"
            elif o == "-b":
                libics2gacsv.flag_rrule_bugfix = False
            elif o == "-s":
                libics2gacsv.flag_split_summary = False
            elif o == "-t":
                libics2gacsv.flag_csv_remove_timezone = False
            elif o == "-o":
                libics2gacsv.flag_add_am12_time = True
            elif o == "-d":
                libics2gacsv.flag_description_delete_4th_line_onwards = True
            elif o == "-p":
                libics2gacsv.flag_remove_teams_infomation = False
            elif o == "-m":
                libics2gacsv.flag_matumoto_modify = True
            elif o == "-h":
                myhelp()

        if(len(argv)) < 3:
            raise ValueError("引数が足りません。")

        TIMERANGE = argv[0]
        INPUT_ICS_FILENAME=argv[1]
        OUTPUT_CSV_FILENAME=argv[2]

        # CHECK TIMERANGE
        t_bak = TIMERANGE
        if TIMERANGE == "all":
            TIMERANGE = 0
        elif TIMERANGE == "guessin":
            TIMERANGE = libics2gacsv.guess_timerange(INPUT_ICS_FILENAME)
            if TIMERANGE is None:
                raise ValueError(f"入力ファイル名からCSVの期間の推測に失敗しました: {INPUT_ICS_FILENAME}")
        elif TIMERANGE == "guess":
            TIMERANGE = libics2gacsv.guess_timerange(OUTPUT_CSV_FILENAME)
            if TIMERANGE is None:
                raise ValueError(f"出力ファイル名からCSVの期間の推測に失敗しました: {OUTPUT_CSV_FILENAME}")
        elif TIMERANGE.isdecimal():
            TIMERANGE = int(TIMERANGE) + 0
        else:
            raise ValueError(f"期間指定の誤り: {t_bak}")

        if not libics2gacsv.format_check_timerange(TIMERANGE):
            raise ValueError(f"期間指定の誤り: {t_bak}")

    except ValueError as e:
        print("ERROR: ", e,  file=sys.stderr)
        print("ERROR:  引数 -h でヘルプが表示されます。", file=sys.stderr)
        sys.exit()

    #Ref: https://qiita.com/dai_zamurai/items/d42a85c3bd42a847c35c
    #os.environ['TZ'] = "JST-9"
    #time.tzset()

    libics2gacsv.ics2csv(INPUT_ICS_FILENAME, OUTPUT_CSV_FILENAME, TIMERANGE)
#End of main()

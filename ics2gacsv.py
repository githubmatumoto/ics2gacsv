#!/usr/bin/env python3
# -*- python -*-
# -*- coding: utf-8 -*-
import os
import sys
import getopt
import re
import libics2gacsv

__doc__=f"""ICS(iCalendar)をCSVに変換する。CSVの出力形式はGaroonと同じ。

使用方法:
   $ python3 {sys.argv[0]} [-hubstoOdpmT] 期間 入力.ics 出力.csv

   例
   $ python3 {sys.argv[0]} 202512 calendar.ics schedules.csv
   $ python3 {sys.argv[0]} guess calendar.ics schedules202509.csv

必須引数:
  「期間」: CSVの出力年月の期間指定。有効範囲は2000年1月から2099年
   12月まで。

   例えばCSVに2025年8月分を出力する場合は「202508」と指定する。

   特殊な値として、「all」もしくは「0」だと全期間を変換する。
  「guessin」だと入力ファイル名から期間を推測する。
  「guess」だと出力ファイル名から期間を推測する。例えば出力ファイル名が
  「schedules.csv」なら2025年9月と推測する。

   期間指定を行った場合、月末のスケジュールで月を超えてる場合は翌月分
   も含まれます。例えば11月30日23:00に開始で12月1日02:00に終了の場合は、
   11月分に12月1日02:00終了のスケジュールが入ります。12月分には入りま
   せん。

「入力.ics」: 変換元のICSファイル名を指定。「stdin」を指定すると標準入力。
    ICSファイルは規格で文字コードがUTF-8と決まってます。そのため、必ず
    UTF-8のICSファイルを指定してください。

    ※対応ICSファイル:
       Cybozu Garoon(Version 5.0.2)
       Web版 Outlook
       Windowsアプリ版 Outlook(classic)
       いずれも2025年9月から12月ごろに生成されたICSファイルの出力で確認。

「出力.csv」: 変換先のCSVファイル名を指定。「stdout」を指定すると標準出力。
    文字コードのdefaultはShiftJIS。

オプション引数:

-h : ヘルプを出力する。

-u : CSV出力の文字コードをUTF-8に変更する。defaultはShiftJIS。
    Garoonの生成するCSVのデフォルトの文字列はShiftJIS

-T"文字列" : 一部のICSファイルはTimeZoneの定義が複数行われている場合が
    ある。採用するTimeZoneを指定する。日本標準時なら -T"Asia/Tokyo"な
    どを定義すればよい。

    ※本オプションをWindows環境で使う場合はライブラリ「tzdata」をイン
    ストールしてください。

以下は滅多に利用しない引数です。開発者向け。

-b : RRULE/EXDATEのバグフィックスを無効にする。defaultは有効。
    ※詳細は関数bug_fix_rrule()をみよ。

-s : ICSのsummaryの分割を無効にする。defaultは有効。分割するとCSVの
   「予定」/「予定詳細」がGaroonと同じ形式になる。
    ※詳細は関数split_garoon_style_summary()をみよ。

-t : CSVの出力時刻にTimeZone情報を表示する。

   default(TimeZone情報なし)/Garoonと同じ:
        12:34:55
        12:34:55

   オプション「-t」指定(TimeZone情報あり):
        12:34:55+09:00
        12:34:55-03:00

    ※TimeZoneが複数定義が行われてる場合は正常に動作しない。
    ※詳細は関数ics_time_to_csv()をみよ。

-o : 終日スケジュール(時刻なし)のCSV出力形式を終日スケジュール
     内部表現(時刻あり)「0:00開始翌日0:00終了」とする。

    ※ICSでは終日スケジュールで時刻がある「0:00開始翌日0:00終了」の場
    合と、時刻がない終日スケジュールがあり、区別される。

    defaultの終日スケジュール(時刻なし)の出力:
      "2025/09/05","","2025/09/05","" (1日間の終日スケジュール)
      "2025/09/05","","2025/09/06","" (2日間の終日スケジュール)

    オプション「-o」指定時の終日スケジュール(時刻なし)の出力
      "2025/09/05","00:00:00","2025/09/06","00:00:00" (1日間の終日スケジュール)
      "2025/09/05","00:00:00","2025/09/07","00:00:00" (2日間の終日スケジュール)

    ※オプション「-g」と比較してください。
    ※Outlook(classic)はこの出力形式になります。
    ※詳細は関数ics_time_to_csv()をみよ。

-g : CSVの終日スケジュール(時刻あり)の出力形式を
    終日スケジュール内部表現(時刻なし1)とする。

    ※ICSでは終日スケジュールで時刻がある「0:00開始翌日0:00終了」の場
    合と、時刻がない終日スケジュールがあり、区別される。

    defaultの 0:00開始,0:00終了スケジュールの表示
      "2025/09/05","00:00:00","2025/09/06","00:00:00" (24時間の終日スケジュール)
      "2025/09/05","00:00:00","2025/09/07","00:00:00" (48時間の終日スケジュール)

    オプション「-g」指定時の終日スケジュール(時刻あり)の出力:
      "2025/09/05","","2025/09/05","" (24時間の終日スケジュール)
      "2025/09/05","","2025/09/06","" (48時間の終日スケジュール)

    ※オプション「-g」と「-o」が同時指定された場合は「-g」が優先される。
    ※詳細は関数ics_time_to_csv()をみよ。


-d : ICSメモ欄(description)の4行目以降を消してCSVに出力する。
    defaultでは消さない。

    ※詳細は関数modify_description()をみよ。

-p : defaultではICSメモ欄(description)のTeamsの会議インフォメーションを消
     してCSVに出力する。パスワードが入ってるため。

    本オプションを指定すると、本気が無効になり、Teamsの会議インフォメー
    ションを消さない。

    ※詳細は関数modify_description()をみよ。

-m : ICSのsummaryの分割で作者用の修正。defaultは無効。

    ※詳細は関数ics_time_to_csv()をみよ

-r : タイトル(SUMMARY)やメモ欄(description)の最後の改行や空白を除去する。
    defaultでは除去しない。

    ※詳細は関数modify_description()をみよ。

-w : 繰返しスケジュール(RRULE)の一部上書き(RECURRENCE_ID)を行った場合
     であっても、ICS上は上書きされる前の情報が残っている場合がある。
     (アプリによって異なる。)

    例えば、月から金の昼休みにスケジュールとして「昼食(社食)」を記載し
    たあとに、水曜日を修正し「昼食(外食)」とした場合、内部的には水曜日
    のスケジュールは「昼食(社食)」が残っている。

    「-w」を指定すると一部修正する前のスケジュールも表示される。ただし
     上書きを行ったスケジュールと区別するため、Summaryにprefixとして
     「Hidden: 」が挿入される。具体的には水曜は「昼食(社食)」と
     「Hidden: 昼食(社食)」となる。

    ※詳細はRFC5545のRECURRENCE_ID参照。

-x : RFC5545のRECURRENCE_ID関連の処理を一切おこなわない。
"""

########################################

def myhelp():
    help(EXEC_FILENAME)
    help("libics2gacsv")
    sys.exit()

if __name__ == '__main__':

    if libics2gacsv.G_VERSION != "2.0":
        print("ERROR: ファイルが古いです。最新のics2gacsv.pyとlibics2gacsv.pyをダウンロードしてください。",file=sys.stderr)
        sys.exit()

    #CSVを出力する期間指定。
    TIMERANGE=0

    EXEC_FILENAME=os.path.basename(__file__)
    EXEC_FILENAME = re.sub(r'\.py$', "", EXEC_FILENAME)
    opts, argv = getopt.getopt(sys.argv[1:], 'ubtsogdpmrwxhT:')

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
            elif o == "-g":
                libics2gacsv.flag_remove_am12_time = True
            elif o == "-d":
                libics2gacsv.flag_description_delete_4th_line_onwards = True
            elif o == "-p":
                libics2gacsv.flag_remove_teams_infomation = False
            elif o == "-m":
                libics2gacsv.flag_matumoto_modify = True
            elif o == "-r":
                libics2gacsv.flag_remove_tail_cr = True
            elif o == "-w":
                libics2gacsv.flag_override_recurrence_id = False
            elif o == "-x":
                libics2gacsv.flag_support_recurrence_id = False
            elif o == "-T":
                libics2gacsv.G_OVERRIDE_TIMEZONE=a
            elif o == "-h":
                myhelp()

        if(len(argv)) < 3:
            raise ValueError("ERROR: 引数が足りません。")

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
                raise ValueError(f"ERROR: 入力ファイル名からCSVの期間の推測に失敗しました: {INPUT_ICS_FILENAME}")
        elif TIMERANGE == "guess":
            TIMERANGE = libics2gacsv.guess_timerange(OUTPUT_CSV_FILENAME)
            if TIMERANGE is None:
                raise ValueError(f"ERROR: 出力ファイル名からCSVの期間の推測に失敗しました: {OUTPUT_CSV_FILENAME}")
        elif TIMERANGE.isdecimal():
            TIMERANGE = int(TIMERANGE) + 0
        else:
            raise ValueError(f"ERROR: 期間指定の誤り: {t_bak}")

        if not libics2gacsv.format_check_timerange(TIMERANGE):
            raise ValueError(f"ERROR: 期間指定の誤り: {t_bak}")

    except ValueError as e:
        print("ERROR: ", e,  file=sys.stderr)
        print("ERROR:  引数 -h でヘルプが表示されます。", file=sys.stderr)
        sys.exit()

    #Ref: https://qiita.com/dai_zamurai/items/d42a85c3bd42a847c35c
    #os.environ['TZ'] = "JST-9"
    #time.tzset()

    #デバグ用のコード。代入したUIDのログが実行中に表示される。
    #libics2gacsv.G_DEBUG_UID="UIDを指定する"
    libics2gacsv.ics2csv(INPUT_ICS_FILENAME, OUTPUT_CSV_FILENAME, TIMERANGE)
#End of main()

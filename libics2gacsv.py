# -*- python -*-
# -*- coding: utf-8 -*-
#
# Copyright (c) 2025 MATSUMOTO Ryuji.
# License: Apache License 2.0
#
import sys
import io
import re
import datetime
import zoneinfo
import csv
import dateutil
import vobject

G_HAIFU_URL="https://qiita.com/qiitamatumoto/items/ab9e0cb9a6da257597a4"
G_GITHUB_URL="https://github.com/githubmatumoto/ics2gacsv"

__doc__=f"""ICS to CSV コンバータ ライブラリ


ICS to CSV コンバータ isc2gacsv.py が利用するライブラリ。

ライセンス:
Apache License 2.0

配布元:
{G_HAIFU_URL}
{G_GITHUB_URL}

修正履歴およびKnown bugsは misc/CHANGELOG.md 参照ください。
"""

#######################################
G_VERSION = "2.1"
#########################################################################
def get_ics_val(ics_parts, name, default_val=None, exit_none=True):
    """
    vobjectのicsの要素を取り出す。

    引数:
    ics_parts: ICSをよみこんだvobjectのcomponetオブジェクト。VEVENTが一つだけ
    入ってる。

    name: 取り出す要素。例えば(summaryとかdtstartなど)
    default_val: nameで指定した要素が存在しない場合に返す値。
                 もしここにNoneを指定したときにnameで指定した
                 要素が存在しない場合は例外を送出する。

    exit_none  : 通常はdefault_valがNoneで要素が存在しない場合、
                 defaultでは停止するが、Falseだと停止せずに
                 Noneを返す。
    """

    if hasattr(ics_parts, name):
        return getattr(ics_parts, name, default_val).valueRepr()

    if (default_val is None) and exit_none:
        print(f"ERROR: 不適切なICSファイルです。必須パラメータがありません: {name}", file=sys.stderr)
        ics_parts.prettyPrint()
        raise ValueError(f"ERROR: 不適切なICSファイルです。必須パラメータがありません: {name}")
    return default_val

###############################################################
# timerange
def format_check_timerange(timerange: int) -> bool:
    """
    CSVの出力範囲を指定するtimerangeの値が異常な値でないかを判断する。

    引数:
       timerange (int): CSVの出力範囲を指定するtimerangeの値。

    返り値:
       正常ならTrue, 異常ならFalse
"""
    #timerange無効。全出力する。
    if timerange == 0:
        return True

    #異常値。
    if timerange < 0:
        return False

    #2000年未満/2100年以降は無効。
    y = timerange // 100
    if (y < 2000) or (y >= 2100):
        return False

    m = timerange % 100
    # 0月は無効/13月以上は無効。
    if (m < 1) or (m > 12):
        return False

    #異常なし。
    return True

###
def is_collect_timerange(ics_time, timerange: int)->bool:
    """
    引数で渡した時刻ics_timeがCSVへの出力対象か判断します。

    引数:
       ics_time: datetime.datetime型もしくはdatetime.date型
       timerange: CSVの出力範囲を指定するtimerangeの値。

    返り値:
       出力対象ならTrue, それ以外はFalse

    """
    #print(f"DEBUG: timerange = {timerange}")
    #print(f"DEBUG: ics_time.year = {ics_time.year}")
    #print(f"DEBUG: ics_time.month = {ics_time.month}")
    if timerange == 0:
        return True

    if(ics_time.year == timerange//100) and (ics_time.month == timerange%100):
        return True

    return False

###
def guess_timerange(FILENAME: str):
    """
ファイル名をもとにCSVが出力する期間の推測を行います。

    引数:
       str: ファイル名

    返り値:
       推測したCSV出力期間。推測に失敗した場合はNoneを返します。

"""
    if re.search("all", FILENAME):
        return 0

    t = re.search(r"\d{6}", FILENAME)
    if not t:
        return None
    #print (f"DEBUG: {t}")
    ret = int(t.group())
    if not format_check_timerange(ret):
        return None
    return ret

#########################################################################
# 時間関係のis関数
###
def is_aware(d) -> bool:
    """
    引数で渡した datetime オブジェクトd が aware(timezoneあり) かどうかを判定する
"""

    # Pythonマニュアルより:
    # date 型のオブジェクトは常に naive です。
    if type(d) is datetime.date:
        return False

    # Pythonマニュアルより:
    # 次の条件を両方とも満たす場合、 time オブジェクト t は aware です:
    # t.tzinfo が None でない
    # t.tzinfo.utcoffset(None) が None を返さない
    #どちらかを満たさない場合は、 t は naive です。

    if type(d) is datetime.datetime:
        if d.tzinfo is None:
            return False
        if d.tzinfo.utcoffset(None) is None:
            return False
        return True

    raise RuntimeError(f"ERROR: 想定外の型が渡されました: type={type(d)}")

###
def is_native(d) -> bool:
    """
    引数で渡した datetime オブジェクトd が native(timezoneなし) かどうかを判定する

"""
    return not is_aware(d)

###
def is_am12(t) -> bool:
    """
    引数で渡した datetime オブジェクトt に 時/分/秒の時刻情報があり、
    深夜12時"00:00:00"ならTrue

"""
    if not type(t) is datetime.datetime:
        raise ValueError(f"ERROR: 想定外の型が渡されました。{type(t)}")
    return (t.hour + t.minute + t.second) == 0
###
def hava_time(t) -> bool:
    """
    引数で渡した datetime オブジェクトt に 時刻情報があるかないか。

    MEMO: pylintに isinstance() を使うように指示されたが、
    なぜか動かなくなったので、もとに戻した
"""
    if type(t) is datetime.date:
        return False

    if type(t) is datetime.datetime:
        return True

    raise RuntimeError(f"ERROR: 想定外の型が渡された: = {type(t)}")

#########################################################################
# TimeZone関係の関数

###
# ZoneInfoで扱える文字列を指定する。
# Ref: https://zenn.dev/fujimotoshinji/scraps/f9c25aeb00a716
#
# dtstart/dtendなどにTimeZoneが指定されていたらそれを優先。
#
G_OVERRIDE_TIMEZONE = None
#G_OVERRIDE_TIMEZONE = "Asia/Tokyo" # 文字列
#
# 推測した値を保存。
__GUESS_TIMEZONE = None # 値は文字列ではない。TimeZoneオブジェクト
__flag_init_guess_timezone = False
#
#
def init_guess_timezone(cal_tz: dateutil.tz.tz.tzical, override_timezone: str = None):
    """
    TimeZoneを推測する関数の初期化

    制御変数:
    G_OVERRIDE_TIMEZONE:ただし ics2csv()で使ってます。

"""
    global __GUESS_TIMEZONE
    global __flag_init_guess_timezone

    __flag_init_guess_timezone = False

    n = []
    if not n is None:
        n = cal_tz.keys()

    if not override_timezone is None:
        print("INFO: 引数でデフォルトのTimeZoneが指定されています。", file=sys.stderr)

        # VTIMEZONEで定義されていたTimeZoneから探す。
        if override_timezone in n:
            __GUESS_TIMEZONE = cal_tz.get(override_timezone)
        else:
            # OS定義のTimeZOneから探す。
            try:
                __GUESS_TIMEZONE = zoneinfo.ZoneInfo(override_timezone)
            except zoneinfo.ZoneInfoNotFoundError as e:
                raise ValueError(f"ERROR: 無効なTimeZoneとして[{override_timezone}]が指定されました。") from e

        print(f"INFO: TimeZoneとして[{override_timezone}]を採用します。", file=sys.stderr)
        __flag_init_guess_timezone = True
        return


    if len(n) == 0:
        print("INFO: ICSデータにTimeZoneデータがありません。", file=sys.stderr)
        print("INFO: Floating Timeのデータです。(Ref: RFC5545, 3.3.12. TIME)", file=sys.stderr)
        __GUESS_TIMEZONE = None
        __flag_init_guess_timezone = True
        return

    if len(n) == 1:
        __GUESS_TIMEZONE = cal_tz.get()
        __flag_init_guess_timezone = True

        return

    if len(n) > 1:
        print("INFO: ICSファイルにTimzeZoneが複数定義されています。", file=sys.stderr)
        print(f"INFO: 現在定義されているTimeZone一覧: {cal_tz.keys()}", file=sys.stderr)
        print(f"INFO: TimeZoneとして1番目に定義されている[{n[0]}]を採用します。", file=sys.stderr)
        print("WARNING: 採用したTimeZoneが不適切な場合、繰返しスケジュールの最終日(UNTIL)の計算に失敗し、", file=sys.stderr)
        print("WARNING: スケジュールが欠落する可能性あります。不適切な場合は引数で指定してください。", file=sys.stderr)
        __GUESS_TIMEZONE = cal_tz.get(n[0])
        __flag_init_guess_timezone = True
        return

    raise ValueError("ERROR: ここには来ないはずだが。。")
    #

def guess_timezone():
    """
    TimeZoneを推測する。事前にinit_guess_timezone()で初期化する必要あり。

"""
    if not __flag_init_guess_timezone:
        raise ValueError("ERROR: 初期化されていません")

    if __GUESS_TIMEZONE is None:
        raise ValueError("ERROR: 壊れたICSファイルです。恐らくFloating Timeだが世界標準時が使われているためローカルタイムに変換できない。")

    return __GUESS_TIMEZONE

#########################################################################
# TimeZoneの変換関係の関数。
###
def convert_native2aware(d) -> datetime.datetime:
    """
    RFC5545ではfloating timeという概念があるdatetimeのnative timeとほぼ同等。

    引数で渡したdatetimeオブジェクトdはnativeな
    (datetime.datetimedate or datetime.date)とする。

    引数dをawareなdatetimeに変換する。

    時刻情報が無い日付のみであるdatetime.date型を渡された時は0時0分0秒とする。

"""
    #print(f"DEBUG(conver_aware/pre):{d}", file=sys.stderr)
    #print(f"DEBUG type(d) = {type(d)}", file=sys.stderr)

    if is_aware(d):
        #print(f"DEBUG(conver_aware/aft):{d}", file=sys.stderr)
        return d

    y = d.year
    m = d.month
    day = d.day

    s = minute = h = 0
    if type(d) is datetime.datetime:
        s = d.second
        minute = d.minute
        h = d.hour
        #print(f"DEBUG s={s}, min={minute}, h={h}", file=sys.stderr)

    d = datetime.datetime(y, m, day, h, minute, s, tzinfo=guess_timezone())
    #print(f"DEBUG(conver_aware/aft):{d}", file=sys.stderr)
    return d

def convert_localtime(d, exit_none=True, exit_native=False):
    """

    引数で渡したdatetimeオブジェクトdに TimeZone情報があれば、ローカルタイ
    ムに変換する。

    TimeZone情報がなければ、何もしない。

    exit_none: Noneが渡された時の挙動。
               True: 例外を送出。
               False : Noneを返す。

    exit_native: Timezoneがないnative timeを渡された時の挙動。
               True: 例外を送出。
               False : なにもせずに渡された時刻を返す。

    """
    if d is None:
        if exit_none:
            raise RuntimeError("ERROR: 引数にNoneが渡されました")
        return None

    if is_aware(d):
        return d.astimezone(guess_timezone()) # ローカルタイムに変換。

    if exit_native:
        raise RuntimeError(f"ERROR: floatingtimeをローカルタイムへの変換しようとしました, time={d}")
    return d
###
#########################################################################
# vobjectがRRULEのEXDATE関連で例外を履く記述の修正関数
#
# RRULEのbugfixを有効にする
flag_rrule_bugfix = True

###
def find_ics_data(data: list, key: str, stop=True) -> int:
    """
    文字列型で渡されたVEVENTのデータdataからkeyで指示された
    要素がある行を探します。stop=Trueの場合は例外を送出します。
"""
    for i, d in enumerate(data):
        if re.match(key, d):
            return i
    if stop:
        raise RuntimeError(f"ICSのアイテム「{key}」が無いVEVENTが渡された")
    return -1

###
def bug_fix_exdate_aux(data: list) -> list:
    """
bug_fix_rruleの補助関数1

BEGIN:VEVENTからEND:VEVENTの間のデータをSTRING型のlistで渡して、
EXDATE関連を修正する。

"""
    flag_exdate = find_ics_data(data, 'EXDATE')

    # EXDATEの時刻指定は複数ある場合があるので、修正時は要注意。
    # 「EXDATE:20250909,20250915」など。

    exdate = data[flag_exdate].split(':')
    if len(exdate) != 2:
        raise ValueError(f"ERROR: DXDATEの書式異常: {data[flag_exdate]}")

    # 時刻情報(T)が一つでもあれば何もしない。
    #「EXDATE:20250909T112233」など。
    if re.search('T', exdate[1]):
        return data

    # EXDATEに何らかのオプションがあれば何もしない。
    if not re.fullmatch('EXDATE', exdate[0]):
        return data

    # maybe Garoon. 時刻情報で(時・分)なし。
    # search: 「EXDATE:20250909」で後ろにTなし。
    exdate[0], count = re.subn('^EXDATE$', 'EXDATE;VALUE=DATE', exdate[0])
    if count == 0:
        raise RuntimeError(f"ERROR: 多分バグ(Garoonfix): {exdate[0]}")
    data[flag_exdate] = exdate[0] + ":" + exdate[1]
    return data

###
def bug_fix_rrule(data: str) -> str:
    """EXDATE関連のbugfix。ICSのファイルをすべて読み込んだ
string型のdataを渡して、修正して返却する。

RRULEのEXDATEが下記形式だとライブラリvobject-0.99では例外を
送出します。

  EXDATE:20251128

本関数は、下記形式に修正します。

  EXDATE;VALUE=DATE:20251128

  バグの詳細についてはmisc/TECH-MEMO.txt 参照ください。


    """
    flag_debug = False
    #flag_debug = True
    #print(f"DEBUG: arg_type:  {type(data)}")
    # typeはstrを想定
    if not type(data) is str:
        raise RuntimeError(f"ERROR: 想定外の型が渡されました: type={type(data)}")

    # https://maku77.github.io/python/numstr/split-lines.html
    # 文字列を改行で分割する。
    lines = data.splitlines()
    org_line_num = len(lines)

    flag_in_vevent = False
    flag_hava_exdate = False
    lines_in_vevent = []

    lines_ret = []

    for i in lines:
        if re.match('BEGIN:VEVENT', i):
            if flag_in_vevent is True:
                raise RuntimeError("ERROR: 「BEGIN:VEVENT」が二重に現れました")
            if len(lines_in_vevent) != 0:
                raise RuntimeError("ERROR: おそらくバグ")
            flag_in_vevent = True
            lines_in_vevent.append(i)
            continue

        if re.match('END:VEVENT', i):
            if flag_in_vevent is False:
                raise RuntimeError("ERROR: 「END:VEVENT」が二重に現れました")
            flag_in_vevent = False
            lines_in_vevent.append(i)
            if flag_hava_exdate:
                if flag_debug:
                    print("DEBUG: PRE:--EXDATE--\n"+'\n'.join(lines_in_vevent) + "-----\n", file=sys.stderr)
                lines_in_vevent = bug_fix_exdate_aux(lines_in_vevent)
                if flag_debug:
                    print("DEBUG: AFT:----\n"+'\n'.join(lines_in_vevent) + "-----\n", file=sys.stderr)
                flag_hava_exdate = False

            lines_ret += lines_in_vevent
            lines_in_vevent = []
            continue

        if not flag_in_vevent:
            lines_ret.append(i)
            continue

        if re.match('EXDATE', i):
            flag_hava_exdate = True

        lines_in_vevent.append(i)
    ##
    if len(lines_ret) != org_line_num:
        raise RuntimeError("ERROR: 行数が変化している。たぶんバグ")

    return "\n".join(lines_ret) + "\n"

##########################################################################
# CSVに出力する時の各種処理関数
#
flag_csv_remove_timezone = True
flag_add_am12_time = False
flag_remove_am12_time = False
#
###
def ics_parts_to_csv_time(ics_parts, rrule_start) -> tuple:
    """
   VEVENTのDTSTART&DTENDをCSV出力用の文字列(tuple)に変換する。

   引数:

    ics_parts: ICSをよみこんだvobjectのcomponetオブジェクト。VEVENTが一つだけ
               入ってる。

    rrule_start: 繰返しスケジュールの時の開始時刻。VEVENTのDTSTARTとDTENDを置き換える。
                 Noneの場合は、VEVENTのDTSTARTとDTENDがそのまま使われる。
                 datetime.datetime型もしくはdatetime.date型

   返り値:
    タプルで文字列４つ
      "開始日","開始時刻","終了日","終了時刻"

    例:
      "2025/09/05","14:40:00","2025/09/05","16:10:00"

    通日スケジュールの場合は時刻情報がないため、下記のようになる。
      "2025/09/05","","2025/09/05","" (1日間の通日スケジュール)
      "2025/09/05","","2025/09/06","" (2日間の通日スケジュール)

   外部制御変数:
    flag_remove_timezone: Bool型
         False: TimeZone情報が有る場合、TimeZone情報込で出力。
            TimeZone情報あり: 12:34:55+09:00, 12:34:55-03:00
            TimeZone情報なし: 12:34:55, 12:34:55

         True: TimeZone情報がある場合、JSTならJSTの表示「+09:00」をけずる。
            JST以外ならJSTに修正の上、TimeZone情報を削って表示
            12:34:55, 12:34:55

   flag_remove_am12_time: Bool型
        Trueを指定すると、終日スケジュール(時刻あり)の時刻情報を除去し、
        終日スケジュール内部表現(時刻なし1)とする。
        後述のflag_add_am12_timeと相反する引数だが、flag_add_am12_time
        の方が優先される。

         FlatingTimeの終日スケジュール(時刻なし)のICSをOutlook(Web)に
        インポートすると終日スケジュール(時刻あり)に変更される挙動の対策。

         例: 2025年3月31日の終日スケジュール(時刻あり)
            "開始日","開始時刻","終了日","終了時刻"
            "2025/03/31", "", "2025/03/31", ""　(Trueのとき)
            "2025/03/31", "00:00:00", "2025/04/01", "00:00:00" (False(default)のとき)

         ただし0分間のスケジュールの場合は、常に時刻情報を表示する。
         例: 2025年3月31日の0:00開始, 3月31日0:00終了の時。
            "2025/03/31", "00:00:00", "2025/03/31", "00:00:00"

    flag_add_am12_time: Bool型
　　　　Trueを指定すると、終日スケジュール(時刻なし)であっても、
        終日スケジュール内部表現(時刻あり)とする。
         常に「0:00開始翌日0:00終了」のスケジュールとして出力する。
         前述のflag_remove_am12_timeは無視される。

         Outlook(classic)は常に上記の挙動となるため、Outlook(classic)
         と同等とする対策。

         例: 2025年9月11日の終日スケジュール(時刻なし)
            "開始日","開始時刻","終了日","終了時刻"
            "2025/09/11", "00:00:00", "2025/09/12", "00:00:00" (Trueのとき)
            "2025/09/11", "", "2025/09/11", "" (False(default)のとき)

         ただし0分間のスケジュールの場合は、常に時刻情報を表示する。
         例: 2025年3月31日の0:00開始, 3月31日0:00終了の時。
            "2025/03/31", "00:00:00", "2025/03/31", "00:00:00"

    """
    start = get_ics_val(ics_parts, 'x-org-dtstart', None, exit_none=False)
    if start is None:
        start = get_ics_val(ics_parts, 'dtstart')

    end = get_ics_val(ics_parts, 'x-org-dtend', None, exit_none=False)
    if end is None:
        end = get_ics_val(ics_parts, 'dtend')

    if False:
        print(f"DEBUG: rrule_start = {rrule_start}", file=sys.stderr)
        print(f"DEBUG: type(rrule_start) = {type(rrule_start)}", file=sys.stderr)
        print(f"DEBUG: start = {start}", file=sys.stderr)
        print(f"DEBUG: type(start) = {type(start)}", file=sys.stderr)
        print(f"DEBUG: end = {end}", file=sys.stderr)
        print(f"DEBUG: type(end) = {type(end)}", file=sys.stderr)
        print(f"DEBUG: flag_csv_remove_timezone = {flag_csv_remove_timezone}", file=sys.stderr)
        print(f"DEBUG: flag_add_am12_time = {flag_add_am12_time}", file=sys.stderr)

    if not rrule_start is None:
        if is_aware(start) != is_aware(rrule_start):
            #この状態になる場合は、事前処理にミスしてる可能性大
            raise ValueError("BUG: timezoneありなし混在")

        # rrule_startは常に時刻情報ありのdatetime型
        # start/endは時刻情報ありのdatetime型と時刻情報なしのdate型のときがある。
        # 足し算でdate型がdetatime型に昇格するとまずいので、rrule_startの
        # 時刻が00:00:00のときはdete型に降格させる。

        """
        if is_am12(rrule_start):
            #rrule_start時刻が00:00:00のときは時刻情報が無いとみなしdate型に降格
            rrule_start = rrule_start.date()
            if False:
                print(f"DEBUG: rrule_start(conv) = {rrule_start}", file=sys.stderr)
                print(f"DEBUG: type(rrule_start)(conv) = {type(rrule_start)}", file=sys.stderr)
        """
        end = rrule_start+(end-start)
        start = rrule_start


    if False:
        print(f"DEBUG: start = {start}", file=sys.stderr)
        print(f"DEBUG: type(start) = {type(start)}", file=sys.stderr)
        print(f"DEBUG: end = {end}", file=sys.stderr)
        print(f"DEBUG: type(end) = {type(end)}", file=sys.stderr)

    # TimeZone情報あるなら無条件にLocalTimeに変換。
    start = convert_localtime(start)
    end = convert_localtime(end)

    if False:
        print(f"DEBUG: start = {start}", file=sys.stderr)
        print(f"DEBUG: type(start) = {type(start)}", file=sys.stderr)
        print(f"DEBUG: end = {end}", file=sys.stderr)
        print(f"DEBUG: type(end) = {type(end)}", file=sys.stderr)

    s_d = "N/A"
    s_t = "N/A"
    e_d = "N/A"
    e_t = "N/A"

    if flag_remove_am12_time and hava_time(start):
        if is_am12(start) and is_am12(end) and start != end:
            start = start.date()
            end = end.date()

    if hava_time(start):
        s_d = start.strftime("%Y/%m/%d")
        e_d = end.strftime("%Y/%m/%d")
        if flag_csv_remove_timezone:
            s_t = start.strftime("%H:%M:%S")
            e_t = end.strftime("%H:%M:%S")
        else:
            s_t = start.strftime("%H:%M:%S%z")
            e_t = end.strftime("%H:%M:%S%z")
    elif flag_add_am12_time:
        #以下時刻情報がない場合の処理
        #終日スケジュール
        # Outlook(classic) style CSV
        s_d = start.strftime("%Y/%m/%d")
        e_d = end.strftime("%Y/%m/%d")
        s_t = "00:00:00"
        e_t = "00:00:00"
        return s_d, s_t, e_d, e_t
    else:
        # Garoon style CSV
        yesterday = end - datetime.timedelta(days=1)
        s_d = start.strftime("%Y/%m/%d")
        e_d = yesterday.strftime("%Y/%m/%d")
        s_t = ""
        e_t = ""

    return s_d, s_t, e_d, e_t

###
# ICSのsummaryの分割を試みる。
flag_split_summary = True

# 予定の選択肢を追加
flag_split_summary_enhance = False
# 追加項目。増やす場合はライブラリ呼び出し側から追加してください。
G_SPLIT_SUMMARY_ENHANCE = []
#
def split_garoon_style_summary(summary: str) -> str:
    """
Garoonはタイトルは二種類の入力があり、
  タイトルの選択肢:'出張', '往訪', '来訪', '会議', '休み'
  タイトルの本文: 「東京特許許可局」

などの入力がある。CSVではそれぞれ"予定","予定詳細"と出力される。

ICSのSUMMARYにはこのような区別はないが、GaroonのCSVに変換するときにタ
イトルの選択肢とコロンがあった場合は分割処理を行っている。

例
ICSの"SUMMARY":「出張:東京特許許可局」
CSVの"予定":「出張」
CSVの"予定詳細":「予定詳細」


    引数:
    summary: タイトルstr型で入った変数。

    返り値:
    summaryを加工後返す。

   外部制御変数:
    flag_split_summary = True
    flag_split_summary_enhance = False
    G_SPLIT_SUMMARY_ENHANCE
"""
    # 以下正規表現文字列になるので変な記号は入れない
    # Garoonの選択肢のデフォルト。
    head = ['出張', '往訪', '来訪', '会議', '休み']
    #SUMMARY分割の拡張
    if flag_split_summary_enhance:
        head += G_SPLIT_SUMMARY_ENHANCE
    # コロンの半角。
    splitter = [':']
    # コロンの全角。
    splitter += ['：']

    for s in splitter:
        ret = summary.split(s, 1)
        if len(ret) == 1:
            continue
        ret[0] = ret[0].strip()
        ret[1] = ret[1].strip()
        for h in head:
            if re.fullmatch(h, ret[0]):
                return ret[0], ret[1]
    #分割失敗
    return "", summary

###
#メモ欄(description)の4行目以降を消す。
flag_description_delete_4th_line_onwards = False
# Teamsの会議インフォメーションを消す。パスワードが入ってる。
flag_remove_teams_infomation = True
#

def modify_description(description: str) -> str:
    """
    メモ欄(description)の加工を行う。長いと見にくいのと、Teamsのパスワード
    が入ってることがあるので。

    引数:
    description: メモ欄がstr型で入った変数。

    返り値:
    descriptionを加工後返す。


   外部制御変数:
    flag_description_delete_4th_line_onwards
    flag_remove_teams_infomation

   Known bugs:
    Teamsの会議インフォメーションの削除は、フォーマットが変わったら無効です。
    2025/9現在のフォーマットをもとに削除を行います。

    """
    if description is None:
        return description

    if flag_remove_teams_infomation:
        lines = description.splitlines()
        new_line = []
        for i in lines:
            if re.search("Microsoft Teams ヘルプが必要ですか", i):
                new_line.append("(REMOVE TEAMS INFOMATION)")
                break
            if re.search(r"\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.", i):
                new_line.append("(REMOVE TEAMS INFOMATION)")
                break
            if re.search("___________________________", i):
                new_line.append("(REMOVE TEAMS INFOMATION)")
                break
            new_line.append(i)
        description = "\n".join(new_line)
        description += "\n"

    if flag_description_delete_4th_line_onwards:
        r = 4
        lines = description.splitlines()
        if len(lines) > r:
            description = "\n".join(lines[:r])
        else:
            description = "\n".join(lines)
        description += "\n"

    return description

###
# 業務番号記入の拡張仕様
def modify_enhanced_gyoumunum(description: str, summary: str) -> str:
    """Summary分割で、Summaryの最後尾に「-数字」もしくは「g数字」があった場合は、
    業務番号と見なし、DESCRIPTIONと置き換える。

    ※仕様検討中。

    スケジュールは技術部のみが使うものではなく他の教職員も使う。そのた
    め、謎の記述は極力避けるべきである。タイトル欄に謎の数字をいれるの
    は好ましくない。

    しかしながら、メモ欄にはTeams会議のパスワードなどセキュリティ情報がか
    かれる可能性があり、業務番号の摘出のためとはいえ、不必要に見える状
    態にするのは好ましくない。

　　そのため、目立たない形で、タイトルの最後尾に記載する案とした。

    業務番号のための区切り文字の検討

    「&」「#」はxml生成時に別の意味が出るので不可
    「:」はGaroon形式に変換するときの「会議」「出張」とかの区切りと区別出来ないので不可
    「,」「;」はGaroonのICS生成時のバグの引き金になるので不可。
    「-」日付情報(12-31)やOS名(Windows-11)などを記載した場合誤認する
         それを許容したとしても、常に半角でいれてくれればいいが全角だ
         と似たようなのが多数あり面倒。全角ハイフン/全角マイナスなどい
         ろいろ。
    「$」はお金を書いてるのと区別つかないのと、一部プログラミング言語の変数
         表記なのでやめた方がよいかも。
    「/」は日付情報(12/31)などを記載した場合誤認する。また会議などの
         部屋番号等の記述でも見られる。
    「%」このあたりで試行する。
         誤用が無いとは言いきれないが。

    Known bugs:
    業務番号を4桁の数字としている。5桁以上なら要修正


    デバグコード:
      debug_modify_enhanced_gyoumunum.py

    処理の流れ:

    1 Summary: 文字列
      Summaryに業務番号がなければ、Descriptionは一切さわらない。無変。

    2. Summary: abcd %数字A or
       Summary: abcd g数字A

    SUMMARYに業務番号があった場合は、SUMMARY側が優先されて、
    DESCRIPTIONの変換を行う。

    Description 変換規則
     方針1:行数の変化は可能な限り避ける。
     方針2:1行めにある業務番号とSUMMARYの業務番号が異なれば置き換える。
     方針3:1行めに空行があれば業務番号と置き換える。
     方針4:1行目に業務番号と空行以外がある場合は、行数を増やす。
     方針5:1行めに「可」「急」があれば、先頭に空行を1行追加して、1行めに業務番号を書き込む
     方針6:1行めに「可」「急」以外があれば、先頭に空行2行追加して、1行めに業務番号を書き込む


    以下
     Pre: DESCRIPTION変換前
     Aft: DESCRIPTION変換後
     数字A: SUMMARYに記載があった業務番号
     数字B: DESCRIPTIONに記載があった業務番号

    "(N/A)": CHANGELOG.mdにも記載したが、DESCRIPTIONが未定義という事をしめす
         特殊な文字。改行のみがあった場合は未定義ではなく""が入る。

    (空白文字): TAB, 全角SPACE (改行は含まない。)
    (改行文字): \\n  (改行の正規化をしているので\\rは出ないはず。)
    「\\s」: 改行を含む空白文字(SPACE, TAB, 全角SPACE, \\n)
    「.」: 任意の一文字

    Description-Type1-1:
    注: 行数が減る可能正あり。
    -(Pre)-------
    (空白文字)*数字B\\s*
    -------------
    -(Aft)-------
    数字A
    -------------

    Description-Type1-2:
    注:"(N/A)"の後ろの文字は削除。
    注: 行数が減る可能正あり。
    -(Pre)-------
    (N/A).*
    -------------
    -(Aft)-------
    数字A
    -------------

    Description-Type1-3:
    注: 行数が減る可能正あり。
    -(Pre)-------
    \\s*
    -------------
    -(Aft)-------
    数字A
    -------------

    Description-Type2:
    行数は変化しない。
    -(Pre)-------
    (空白文字)*数字B(空白文字)*(改行文字).*
    -------------
    -(Aft)-------
    数字A(改行文字).*
    -------------

    Description-Type3:
    行数は変化しない。
    -(Pre)-------
    (空白文字)*(改行文字).*
    -------------
    -(Aft)-------
    数字A*(改行文字).*
    -------------

    Description-Type4:
    行数が1行増える
    -(Pre)-------
    [可|急](空白文字)*(改行文字).*
    -------------
    -(Aft)-------
    (空白文字)*数字A(改行文字)[可|急](空白文字)*(改行文字).*
    -------------

    Description-Type5:
    行数が2行増える
    -(Pre)-------
    .*
    -------------
    -(Aft)-------
    数字A(改行文字)(改行文字).*
    -------------

    """
    # 業務番号記入の拡張仕様: SUMMARYの「g」と「%」
    m = re.search(r"[ｇg%％]([0-9０-９]{1,4})[　 \t]*$", summary)
    if m is None:
        return None
    # 全角数字を半角にするため、0を足してる。
    #print(m.groups())
    #print(m.group())
    gyoumunum = str(int(m.groups()[0])+0)

    if (int(gyoumunum) < 0) or (int(gyoumunum) > 9999):
        # 負の数は業務番号としては無効
        # 5桁の業務番号は無効(正規表現的にないはずだが。)
        raise RuntimeError("ERROR: Summaryの業務番号の取得に失敗しました")

    if re.search(r"\r", description, flags=re.DOTALL):
        raise RuntimeError("ERROR: 改行の正規化が行われてません。「\\n」のみ有効です。")

    # re.matchは1行めのみ検索する。2行目は見ない。
    #print(f"DEBUG: found enhance gyoumunum = {gyoumunum}")

    # Description-Type1-1:
    # 正規表現の空白のところに全角スペース入ってる
    # 改行の正規化をしているので改行に「\r」は出ないはず。
    m1 = re.search(r"^[　 \t]*[0-9０-９]{1,4}[　 \t\n]*$", description, flags=re.DOTALL)

    # Description-Type1-2:
    # 注:"(N/A)"の後ろの文字は無視。
    m2 = re.search(r"^\(N\/A\).*$", description, flags=re.DOTALL)

    # Description-Type1-3:
    # 空白と改行のみ。
    m3 = re.search(r"^[　 \t\r\n]*$", description, flags=re.DOTALL)
    if m1 or m2 or m3:
        return gyoumunum

    # Description-Type2:
    # 行数は変化しない。
    m1 = re.search(r"^[　 \t]*[0-9０-９]{1,4}[　 \t]*\n.*", description, flags=re.DOTALL)
    if m1:
        ret = re.sub(r"^[　 \t]*[0-9０-９]{1,4}[　 \t]*", gyoumunum, description)
        if ret is None:
            raise RuntimeError("ERROR: 正規表現の想定外のエラー")
        return ret

    # Description-Type3:
    # 行数は変化しない。
    m1 = re.search(r"^[　 \t]*\n.*", description, flags=re.DOTALL)
    if m1:
        return re.sub(r"^[　 \t]*", gyoumunum, description)
    # Description-Type4:
    # 冒頭が「可急」の場合は業務番号を行頭に差し込む。行数が1行増える。
    lines = description.splitlines()
    m1 = re.search(r"^[　 \t]*[可急][　 \t]*$", lines[0])
    if m1:
        if re.search("可", lines[0]):
            lines[0] = "可"
        if re.search("急", lines[0]):
            lines[0] = "急"
        return gyoumunum + "\n" + "\n".join(lines) + "\n"

    # Description-Type5:
    #それ以外は業務番号を行頭に差し込み改行を2個差し込む。行数が2行増える。
    return gyoumunum + "\n\n" + description

###
#各要素の最後の改行と空白をすべて取り除く。
flag_remove_tail_cr = False

def ics_parts_to_csv_buffer(ics_parts, rrule_start=None) -> list:
    """
   VEVENTをCSV出力用の文字列(list)に変換する。

    引数:
    ics_parts: ICSをよみこんだvobjectのcomponetオブジェクト。VEVENTが一つだけ
               入ってる。

    rrule_start: 繰返しスケジュールの時の開始時刻。VEVENTのDTSTARTとDTENDを置き換える。
                 Noneの場合は、VEVENTのDTSTARTとDTENDがそのまま使われる。
                 datetime.datetime型もしくはdatetime.date型

    返り値: CSV出力用の文字列に変換してLISTにいれて返す。

   外部制御変数:
    flag_remove_tail_cr
"""
    row = []

    s_d, s_t, e_d, e_t = ics_parts_to_csv_time(ics_parts, rrule_start)

    row.append(s_d) # "開始日"
    row.append(s_t) # "開始時刻"
    row.append(e_d) # "終了日"
    row.append(e_t) # "終了時刻"

    #予定(選択肢のところ)
    summary_h = ""
    #予定本文
    summary_t = get_ics_val(ics_parts, 'summary', None, exit_none=False)

    if flag_split_summary and (not summary_t is None):
        #ICS形式の場合は予定(選択肢のところ)が無いので生成試みる
        summary_h, summary_t = split_garoon_style_summary(summary_t)

    row.append(summary_h) # 予定(選択肢のところ)
    row.append(summary_t) # 予定本文

    description = modify_description(get_ics_val(ics_parts, 'description', None, exit_none=False))
    row.append(description) # メモ

    #各要素の最後の改行と空白をすべて取り除く。
    if flag_remove_tail_cr:
        for i in range(len(row)):
            if not row[i] is None:
                row[i].rstrip()
    return row

##########################################################################
# GaroonのCSVの最初のヘッダを、出力する(True)、しない(False)
flag_print_csv_header = True

# VERSION1.3追加
# 上書スケジュール(RECURRENCE-ID)の対応する。
flag_support_recurrence_id = True
# 上書スケジュール(RECURRENCE-ID)で、
# 基のスケジュールを隠す:True
# 基のスケジュールを表示する: False. 基のスケジュールのSummaryに"Hidden: "と追加します。
flag_override_recurrence_id = True

# CSVの出力の日付ソートを行う(True)。しない(False)
flag_output_sort = True

# 業務記録の拡張フォーマットを使うか
# True: 使う
# False: 使わない
flag_enhanced_gyoumunum = False
# 上書スケジュール(RECURRENCE-ID)対応のためバッファリングを行う。
# CSVの要素は4番目から。冒頭にUIDとDTSTARTとを差し込む。
# ["UID", "DTSTART", "RECURRENCE-ID", "開始日","開始時刻","終了日","終了時刻","予定","予定詳細","メモ"]
# UIDはVEVENTのUID。もしくはヘッダなどは'N/A', RECURRENCE-IDで不可視化された場合はNone
# Noneの場合はファイルへの出力対象外。
# DTSTARTは datetime.datetime型もしくはdatetime.date型
# RECURRENCE-IDは要素に含まれるならその値が入る。datetime.datetime型もしくはdatetime.date型
# なければNone
# datetime.datetimeの時はJSTに変換する。
# それ以外は文字列。
G_CSV_HEADER = ["N/A", None, None, "開始日", "開始時刻", "終了日", "終了時刻", "予定", "予定詳細", "メモ"]
#ヘッダ部分を除いた実際の開始地点
G_CSV_B_OFFSET = 3
G_CSV_H_UID = 0
G_CSV_H_DTSTART = 1
G_CSV_H_RECURRENCE_ID = 2
#
G_CSV_SUMMARY_H = 4 # 実際の位置はG_CSV_B_OFFSET+G_CSV_SUMMARY_H
G_CSV_SUMMARY_T = 5
G_CSV_DESCRIPTION = 6
#
G_DEBUG_UID = None
#
def csv_buffer_dump(buff: list, prefix="DEBUG:", uid=None, file=sys.stderr):
    """
    csv_buffeをdump
"""
    print("----", file=file)
    for i in range(len(buff)):
        if uid is None:
            print(f"{prefix}{i}:{buff[i]}", file=file)
        else:
            if buff[i][G_CSV_H_UID] == uid:
                print(f"{prefix}{i}:{buff[i]}", file=file)
    print("----", file=file)

def recurrence_id_list_dump(l: dict, prefix="DEBUG:", file=sys.stderr):
    """
    recurrence_id_listをdump
"""
    print("----", file=file)
    for uuid in l.keys():
        print(f'{prefix}uid = {uuid}', file=file)
        for vv in l[uuid]:
            print(f'{prefix}\tRECURRENCE-ID = {vv}', file=file)
    print("----", file=file)

#End of func()

def file2str(fname: str) -> str:
    """
    ファイルから読み込み
"""
    ret = ""
    if fname == "stdout" or fname[0] == "-":
        raise RuntimeError(f"入力ファイル名指定エラー: {fname}")

    if fname == "stdin":
        ret = sys.stdin.read()
    else:
        with open(fname, 'r', encoding='utf-8') as f:
            ret = f.read()
    return ret
# end of func

#######################################
# 試してないが、改行コードの話。
# Ref: https://qiita.com/tatsuya-miyamoto/items/f57408064b803f55cf99

# 試してないが、入力文字コードの変更はこちらが参考になる。
# Ref: https://techblog.asahi-net.co.jp/entry/2021/10/04/162109
#出力するCSVの文字コード
G_CSV_ENCODING = "shift_jis"

def open_csv_object(fname: str):
    """
    出力先のファイルを開く。 文字列で"stdout"を指定すると標準出力となる。
"""
    if fname == "stdin"  or fname[0] == "-":
        raise RuntimeError(f"ファイル名指定エラー: {fname}")

    # errors='xmlcharrefreplace' utfからsjisに変換時にsjis未定義コードが出たときに "&#xxxx;"に変換する。
    # Ref: https://docs.python.org/ja/3/howto/unicode.html
    # Ref: https://zenn.dev/hassaku63/articles/f7ca587b86398c
    #
    #escale_type = 'backslashreplace'
    escale_type = 'xmlcharrefreplace'
    if fname == "stdout":
        #https://geroforce.hatenablog.com/entry/2018/12/05/114633
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding=G_CSV_ENCODING, errors=escale_type, newline="")
        csv_out = sys.stdout
    else:
        csv_out = open(fname, 'w', encoding=G_CSV_ENCODING, errors=escale_type, newline="")

    # Pythonライブラリの仕様でCSVの最後の改行はCR+LF。
    # Ref: https://docs.python.org/ja/3/library/csv.html
    # ->「Dialect.lineterminator」
    # -> 「writer が作り出す各行を終端する際に用いられる文字列です。デフォルトでは '\r\n' です。」

    return csv.writer(csv_out, quoting=csv.QUOTE_ALL)

def vobject2csv(calendar: vobject.base.Component, timerange: int):
    """
    補助関数。 vobjectを読み込んで、csv出力用のbufferにいれていく。

    引数timerangeは使ってないが、念のため残しておく。
"""
    # 返り値
    # VERSION1.3追加
    # 旧版1.2ではCSVの要素を生成したらすぐ出力してたが、
    # 上書スケジュール(RECURRENCE-ID)対応のためバッファリングを行う。
    csv_buffer = []

    # key: UID, value: RECURRENCE-IDをリストで収納。
    # 処理段階で基になるVEVENTと読み替え成功したら、消していく。
    # 最終的に残るのが読み替えに失敗したRECURRENCE-ID。
    recurrence_id_list = {}

    for component in calendar.components():
        if component.name == 'VEVENT':
            dtstart = get_ics_val(component, 'dtstart')
            dtend = get_ics_val(component, 'dtend')
            uid = get_ics_val(component, 'uid', "N/A")
            rrule = get_ics_val(component, 'rrule', None, exit_none=False)
            # VERSION1.3追加: RECURRENCE-IDコード。
            recurrence_id = get_ics_val(component, 'recurrence-id', None, exit_none=False)
            recurrence_id = convert_localtime(recurrence_id, exit_none=False)

            # debugコード
            if (G_DEBUG_UID is not None) and G_DEBUG_UID != uid:
                continue

            # データの検査
            if (not rrule is None) and (not recurrence_id is None):
                raise ValueError("ERROR: ICSデータ不整合: 同一VEVENTにRECURRENCE-IDとRRULEがあります。")

            # TODO: RDATE対応作業
            # RDATEはOutlook(classic)でICSファイル出力時の詳細情報の設定で
            # 「詳細情報の一部」を選んだ時に生成される事があります。
            #
            if get_ics_val(component, 'rdate', "N/A") != "N/A":
                raise RuntimeError("ERROR: 本プログラム未実装のICS命令RDATEが使われています。")
            #
            if is_aware(dtstart) != is_aware(dtend):
                raise ValueError("ERROR: ICSデータ不整合: 同一VEVENTにdtstart/dtendにtimezone有り/無しが混在。")

            if hava_time(dtstart) != hava_time(dtend):
                raise ValueError("ERROR: ICSデータ不整合: 同一VEVENTにdtstart/dtendに時刻情報の有り/無しが混在。")

            if G_DEBUG_UID == uid:
                print(f"STEP1: uid = {uid}", file=sys.stderr)
                print(f"STEP1: dtstart = {dtstart}", file=sys.stderr)
                print(f"STEP1: dtend = {dtend}", file=sys.stderr)

            # CSV用のlist生成開始。
            buff_pre = [uid, convert_localtime(dtstart), recurrence_id]
            buff_aft = ics_parts_to_csv_buffer(component)

            # ICSのRRULE命令が未使用ならそのまま出力する。
            if rrule is None:
                csv_buffer.append(buff_pre + buff_aft)

                if flag_support_recurrence_id and (not recurrence_id is None):
                    if uid not in recurrence_id_list:
                        recurrence_id_list[uid] = []
                    recurrence_id_list[uid].append(recurrence_id)
                continue

            #ICSのRRULE命令の処理。
            #t_c = component
            #if type(dtstart) is datetime.date:
            #print(type(rrule))
            #print(rrule)
            #ignoretz = (not isinstance(dtstart, datetime.datetime) or dtstart.tzinfo is None)

            # rruleの繰返し回数はcountで指定と最終日時のuntilの場合がある。
            # untilの場合はいろいろ大変。バグが非常に出やすい。

            # Known bugs: 内部変数「_until」にアクセスしているため、ライブラリの仕様が
            # 変わったら動かない。
            until = dateutil.rrule.rrulestr(rrule)._until

            org_dtstart = dtstart
            if not until is None:
                if is_aware(dtstart) and is_native(until):
                    # dtstartがaware(timezone有)であり、untilがnative(floatingtime)。
                    raise ValueError("ERROR: ICSデータ不整合: 同一VEVENTのdtstartはtimezone有りで、rruleのuntilがtimezone有無し。")

                if is_aware(until) and is_native(dtstart):
                    # dtstartがnative(floatingtime)だが、untilがaware(timezoneあり)の場合。
                    # 本来はICSデータの不整合なのだが、あまりにこの事例が多いため対処。
                    component.add('x-org-dtstart').value = dtstart
                    component.dtstart.value = convert_native2aware(dtstart)
                    component.add('x-org-dtend').value = dtend
                    component.dtend.value = convert_native2aware(dtend)

                    if G_DEBUG_UID == uid:
                        t = get_ics_val(component, 'dtstart')
                        print(f"STEP2: aware dtstart  = {t}", file=sys.stderr)
                        t = get_ics_val(component, 'dtend')
                        print(f"STEP2: aware dtend  = {t}", file=sys.stderr)

            #component.prettyPrint()

            rrule_set = component.getrruleset(addRDate=True)
            if G_DEBUG_UID == uid:
                print(f"STEP2.5: RRULE = {rrule}", file=sys.stderr)
                for s in rrule_set:
                    print(f"RRULE_PARTS={s}", file=sys.stderr)

            for s in rrule_set:
                # getrrulesetがdatetime.dateからdatetime.datetimeに拡張する事がある。
                if hava_time(s) and (not hava_time(org_dtstart)):
                    if not is_am12(s):
                        raise ValueError("BUG: getrrulesetの計算がおかしい")
                    s = s.date()


                if G_DEBUG_UID == uid:
                    print(f"STEP3: s   = {s}", file=sys.stderr)

                if G_DEBUG_UID == uid:
                    print(f"STEP4: PASS(timerange={timerange})", file=sys.stderr)

                t = ics_parts_to_csv_time(component, s)
                buff_pre[G_CSV_H_DTSTART] = convert_localtime(s)
                buff_aft[0:4] = t
                csv_buffer.append(buff_pre + buff_aft)
                if G_DEBUG_UID == uid:
                    print(f"STEP4: normailize(s) = {buff_pre[G_CSV_H_DTSTART]}", file=sys.stderr)
    # end for()
    return csv_buffer, recurrence_id_list
#end of func.

def modify_reference_id_data(csv_buffer: list, recurrence_id_list: dict, outlook_bugfix=False) -> int:
    """
    TODO: 関数コメントが意味不明だから書き換える。

    RECURRENCE-IDで上書きするVEVENTはSUMMARYやDESCRIPTIONが未定義の場合が
    ある。復元する。なお、SUMMARYやDESCRIPTIONは各種加工が入っているので、
    csv_bufferに入ってるのを使った方がよい。

    空文字の場合と未定義の場合を区別するため、未定義の場合は"N/A"が入ってる。
    バグ: 仕様: RECURRENCE-IDで時刻を上書きした場合、上書きの基となるVEVENTが
    出力対象としている月の外になった場合、復元できない事がある。
    全期間データをバッファリングしないといけないので仕様とする。

    RECURRENCE-IDで上書きするVEVENTを消す。もしくは印をつける。

"""
    bad_count = 0

    uid2line = {}
    for i in range(len(csv_buffer)):
        uid = csv_buffer[i][G_CSV_H_UID]
        if uid in recurrence_id_list:
            if not uid in uid2line:
                uid2line[uid] = []
            uid2line[uid].append(i)

    key_list = list(recurrence_id_list.keys())
    for key in key_list:
        line_list = uid2line[key]
        for i in line_list:
            b = csv_buffer[i]
            uid = b[G_CSV_H_UID]
            recurrence_id = b[G_CSV_H_RECURRENCE_ID]

            if recurrence_id is None:
                continue

            flag_found_j = -1
            for j in line_list:
                if i == j:
                    continue
                if not csv_buffer[j][G_CSV_H_RECURRENCE_ID] is None:
                    continue

                if recurrence_id == csv_buffer[j][G_CSV_H_DTSTART]:
                    flag_found_j = j
                    break

                # 修正前のdtstartがdatetime.date(日付のみ)なのに、
                # 修正先のrecurrence_idがdatetime.datetime(日時情報あり)になっとる。
                if outlook_bugfix:
                    dd = convert_native2aware(csv_buffer[j][G_CSV_H_DTSTART])
                    if recurrence_id == dd:
                        flag_found_j = j
                        break

            UNREF = "(REFERENCE DATA DOES NOT EXIST)"
            org_summary_h = ""
            org_summary_t = UNREF
            org_description = UNREF

            k = -1
            if flag_found_j < 0:
                bad_count = bad_count + 1
            else:
                k = G_CSV_B_OFFSET+G_CSV_SUMMARY_H
                org_summary_h = csv_buffer[flag_found_j][k]
                k = G_CSV_B_OFFSET+G_CSV_SUMMARY_T
                org_summary_t = csv_buffer[flag_found_j][k]
                k = G_CSV_B_OFFSET+G_CSV_DESCRIPTION
                org_description = csv_buffer[flag_found_j][k]
                if flag_override_recurrence_id:
                    csv_buffer[flag_found_j][G_CSV_H_UID] = None
                else:
                    k = G_CSV_B_OFFSET+G_CSV_SUMMARY_T
                    csv_buffer[flag_found_j][k] = "Hidden: " + csv_buffer[flag_found_j][k]

                recurrence_id_list[uid].remove(recurrence_id)
                if len(recurrence_id_list[uid]) == 0:
                    del recurrence_id_list[uid]

            k = G_CSV_B_OFFSET+G_CSV_SUMMARY_T
            if (csv_buffer[i][k] is None) or (outlook_bugfix and csv_buffer[i][k] == UNREF):
                csv_buffer[i][k] = org_summary_t
                k = G_CSV_B_OFFSET+G_CSV_SUMMARY_H
                csv_buffer[i][k] = org_summary_h

            k = G_CSV_B_OFFSET+G_CSV_DESCRIPTION
            if (csv_buffer[i][k] is None) or (outlook_bugfix and csv_buffer[i][k] == UNREF):
                csv_buffer[i][k] = org_description

            #end for j
        #end for i
    #end for key
    return bad_count

# end of func()

def load_ics_vtimezone(ics_data: str):
    """
    TimeZoneデータ読み込み
    https://dateutil.readthedocs.io/en/stable/tz.html
    vobjectのサンプルtests.py

    返り値: None -> VTIMEZONEのデータが無かった。
    データ異常などの場合は、例外raiseで停止する。
"""
    ret = None

    try:
        ret = dateutil.tz.tzical(io.StringIO(ics_data))
    except ValueError as e:
        # Outlook(classic)が作ったICSを一気よみすると途中のVEVENTで
        # 解析に失敗する。
        new = []
        found_vtimezone = False
        found_vcalendar = False
        found_vevent = False
        for i in ics_data.splitlines():
            if re.match('BEGIN:VTIMEZONE', i):
                found_vtimezone = True
            if re.match('BEGIN:VCALENDAR', i):
                found_vcalendar = True
            if re.match('BEGIN:VEVENT', i):
                if found_vevent:
                    raise ValueError("ERROR: BEGIN:VEVENTからEND:VEVENTの対応が壊れてる(1)") from e
                found_vevent = True
                continue
            if re.match('END:VEVENT', i):
                if not found_vevent:
                    raise ValueError("ERROR: BEGIN:VEVENTからEND:VEVENTの対応が壊れてる(2)") from e
                found_vevent = False
                continue
            if found_vevent:
                continue
            new.append(i)
        # end for i:

        if found_vevent:
            raise ValueError("ERROR: BEGIN:VEVENTからEND:VEVENTの対応が壊れてる(3)") from e

        if not found_vcalendar:
            raise ValueError("ERROR: 入力して渡されたICSファイルにBEGIN:VCALENDARが無い") from e

        if not found_vtimezone:
            # VTIMEZONEが無い。おそらくFloating timeモード。ただし時刻がUTC(世界標準時)の可能性あり。
            ret = None
        else:
            #print("\n".join(new))
            ret = dateutil.tz.tzical(io.StringIO("\n".join(new) + "\n"))

    #print(type(ret))
    #if len(ret.keys()) == 1:
        #t = datetime.datetime(2003, 9, 27, 12, 40, 12, tzinfo=ret.get())
        #t = datetime.datetime(2003, 9, 27, 12, 40, 12, tzinfo=dateutil.tz.tzutc())
        #print(t, file=sys.stderr)

    #print(tzs.get('Tokyo Standard Time'))
    #jp = vobject.icalendar.TimezoneComponent(tzs.get('Tokyo Standard Time'))
    #print(jp)

    return ret
#####

def ics2csv(ics_file_path: str, csv_file_path: str, timerange: int = 0) -> None:
    """
    ICS(iCalendar)ファイルをCSVファイルに変換する。

    引数:
        ics_file_path (str): 変換元のICS(iCalendar)ファイル。"stdin"を指定すると標準入力。
        csv_file_path (str): 変換先のCSVファイル。"stdout"を指定すると標準出力。
        timerange (int): CSVに変換する日時を限定する場合は、指定する。
                         2025年8月分がほしい場合は「202508」と指定する。
                         未指定や「0」だと全部変換する。
    返り値:
        None。失敗したら停止する。
    """
    ######################
    if vobject.VERSION != "0.9.9":
        print("ERROR: 依存ライブラリvobjectのバージョンが開発環境と異なります。", file=sys.stderr)
        print(f"ERROR: githubよりプルリクエストください。 vobject.VERSION={vobject.VERSION}", file=sys.stderr)
        sys.exit()

    ######################
    #ファイルから読み込み
    ics_data = file2str(ics_file_path)
    # あまりに小さい。
    if len(ics_data) < 10:
        raise RuntimeError(f"ERROR: ファイル読み込みエラー: ファイル行数: {len(ics_data)}")

    ######################
    # ライブラリvobjectがのicsファイルの読み込む時に例外を履く
    # 記述の修正を行う。
    # liics2gacsv(v1.4)では RRULEのEXDATEの記述の修正のみ。
    if flag_rrule_bugfix:
        ics_data = bug_fix_rrule(ics_data)
    else: # 処理の都合で改行の正規化が必要
        ics_data = "\n".join(ics_data.splitlines()) + "\n"

    ######################
    # 読み込んだデータstrをvobjectに変換。
    calendar = vobject.readOne(ics_data)

    ######################
    # TimeZoneデータ読み込み
    cal_tz = load_ics_vtimezone(ics_data)
    # TimeZoneが明記されてない文脈でTimeZoneの推測を行う関数
    init_guess_timezone(cal_tz, G_OVERRIDE_TIMEZONE)

    ######################
    # vobjectのオブジェクトをCSVに変換。
    csv_buffer, recurrence_id_list = vobject2csv(calendar, timerange)

    if not G_DEBUG_UID is None:
        csv_buffer_dump(csv_buffer, prefix="D1:", uid=G_DEBUG_UID)

    # RECURRENCE-IDで上書きするVEVENTはSUMMARYやDESCRIPTIONが未定義の場合が
    # ある。復元する。
    # RECURRENCE-IDで上書きするVEVENTを消す。もしくは印をつける。

    bad_recurrence_id_count = 0
    if flag_support_recurrence_id:
        bad_recurrence_id_count = modify_reference_id_data(csv_buffer, recurrence_id_list)
        if bad_recurrence_id_count > 0:
            bad_recurrence_id_count = modify_reference_id_data(csv_buffer, recurrence_id_list, outlook_bugfix=True)
            if bad_recurrence_id_count > 0:
                print(f"WARNING: 繰返し命令の{bad_recurrence_id_count}個の復元に失敗。", file=sys.stderr)
                print("WARNING: 失敗したスケジュールは「(REFERENCE DATA DOES NOT EXIST)」と記載されています。", file=sys.stderr)
                print("WARNING: 手作業で修正してください。", file=sys.stderr)
                if not G_DEBUG_UID is None:
                    print(f"D2: recurrence_id_list = {recurrence_id_list}", file=sys.stderr)

        # 浮いてたrecurrence_id_list。
        # 例えば4月のスケジュール生成時に3月のデータを引用してたら浮く。
        # Outlook(classic)の場合は「(REFERENCE DATA DOES NOT EXIST)」となる。
        # Outlook(Web)の場合は特に問題なし。
        if len(recurrence_id_list) > 0:
            print("WARNING: ICSファイルの異常です。参照先が無い繰返し命令があります。", file=sys.stderr)
            recurrence_id_list_dump(recurrence_id_list, "BROKEN VEVENT:")


    if not G_DEBUG_UID is None:
        csv_buffer_dump(csv_buffer, prefix="D2:", uid=G_DEBUG_UID)
        print(f"D2: recurrence_id_list = {recurrence_id_list}", file=sys.stderr)

    # ICSのデータで指定の要素がなかった場合はNoneが入っている。
    # 適切な用語に書き換える。
    for i in range(len(csv_buffer)):
        for j in range(G_CSV_B_OFFSET, len(G_CSV_HEADER)):
            if csv_buffer[i][j] is None:
                csv_buffer[i][j] = "(N/A)"


    # 日付でsortする. index sort.
    # csv_indexのG_CSV_B_OFFSET番目以降の要素は文字列なので、次でよい。
    # BUG: 月と日が2桁でないとバグる。また日の順番が y/d/mだったらNG.
    csv_index = list(range(len(csv_buffer)))
    if flag_output_sort:
        csv_index.sort(key=lambda x: csv_buffer[x][G_CSV_B_OFFSET:])

    if not G_DEBUG_UID is None:
        csv_buffer_dump(csv_buffer, prefix="D3:", uid=G_DEBUG_UID)

    #GaroonのHeader出力
    if flag_print_csv_header:
        csv_buffer.insert(0, G_CSV_HEADER)
        # CSV_INDEXを1個づつ増やして
        csv_index = [i+1 for i in csv_index]
        # CSV_INDEXの頭に0をつっこむ。
        csv_index.insert(0, 0)

    # CSVに出力。
    csv_writer = open_csv_object(csv_file_path)

    for i in csv_index:
        if csv_buffer[i][G_CSV_H_UID] is None:
            continue
        if csv_buffer[i][G_CSV_H_DTSTART] is None:
            continue
        if not is_collect_timerange(csv_buffer[i][G_CSV_H_DTSTART], timerange):
            continue

        # 業務番号記入の拡張仕様
        # SUMMARYに記載された業務番号をDESCRIPTIONに差し込む。
        if flag_enhanced_gyoumunum:
            summ = csv_buffer[i][G_CSV_B_OFFSET + G_CSV_SUMMARY_T]
            des = csv_buffer[i][G_CSV_B_OFFSET + G_CSV_DESCRIPTION]
            des = modify_enhanced_gyoumunum(des, summ)
            if des:
                csv_buffer[i][G_CSV_B_OFFSET + G_CSV_DESCRIPTION] = des

        csv_writer.writerow(csv_buffer[i][G_CSV_B_OFFSET:])

    if not G_DEBUG_UID is None:
        csv_buffer_dump(csv_buffer, prefix="D4:", uid=G_DEBUG_UID)


    # 終了ステータス表示。
    if bad_recurrence_id_count == 0:
        print(f"INFO: 変換に成功しました: '{ics_file_path}' to '{csv_file_path}'", file=sys.stderr)
    else:
        print(f"WARNING: 変換に*概ね*成功しました: '{ics_file_path}' to '{csv_file_path}'", file=sys.stderr)

#end func

############################################
__v= tuple(sys.version_info)

# ライブラリvobjectはPython3.8以上が必要。
# (2025/10現在)Python3.9はすでにEOLのため本来なら3.10以上にしたいが
# macOSの標準Pythonが3.9なので3.9以上としている。
#
if __v < (3, 9):
    print(f"ERROR: Pythonはversion3.9以上が必要です。現在version {__v[0]}.{__v[1]}'", file=sys.stderr)
    sys.exit()
###
if __name__ == '__main__':
    help("libics2gacsv")
#EOF

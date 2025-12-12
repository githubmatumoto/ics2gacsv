# -*- python -*-
# -*- coding: utf-8 -*-
import sys
import io
import re
import datetime
import zoneinfo
import csv
import dateutil
import vobject

__doc__ = """補足:
ソースコードに加えて別途 README.txt および TECH-MEMO.txt　参照ください。

*Cangelog:
2025/9/24: Version:1.0 (公開終了)
初版公開

2025/9/25: Version:1.1 (公開終了)

変数名等でicalとicsが混ざってたので、icsに統一。

2025/9/26: Version:1.2 (公開終了)
(内部メモ:subversion revision 2070, フォルダv1.2)

CSVの改行コードの扱いが悪く無駄な空行が入るのを改善。生成される改行コー
ドについては TECH-MEMO.txt　参照ください。

CSVでダブルクオートで囲まれた文字列の最後に改行が空白があった場合除去
するコードを追加。

出力文字コードがShiftJISの時にUTF-8からShiftJISに変換する作業に失敗し
た時のエスケープ手段がstdoutに出力した時とファイルに出力した時で異なっ
たので、統一。

macOS 13.7.8/Intel/日本語環境動作確認しました。

ライブラリのバージョン管理用の変数を追加しています。

その他用語の統一など行ってます。

2025/9/30: Version:1.3 (非公開)
(内部メモ:subversion revision 2079, フォルダv1.3)

上書スケジュール(RECURRENCE-ID)へ限定対応。細かいRECURRENCE-ID命令には
対応していない。挙動がおかしい場合はオプション「-w」で無効化された命令
をCSVに出力するようにして欲しい。

上書スケジュール対応のため、CSVに出力する前に一度バッファリングを行う
ように変更。

CSVに変更したの特殊な値を定義
- 「(N/A)」: 指定の要素がVEVENTに無かった。SUMMARY/DESCRIPTONのみ。
   カレンダーアプリにより、上記要素が無い場合、下記の挙動がある。
    要素を未定義とする場合は「(N/A)」が入る。
　  改行をいれる場合は空行が入る。

- 「((REFERENCE DATA DOES NOT EXIST))」: 滅多にないはずだが、
 RECURRENCE-IDで上書をするVEVENTにはSUMMARY/DESCRIPTONが無い場合がある。
 基のVEVENTからコピーする処理を行っているが、基のVEVENTが見つからなかっ
 た。

- 「Hidden: 」オプション「-w」を指定した時のみ。
   RECURRENCE-IDで上書をした基のVEVENTを「Hidden: 」を付けて出力する。

暗にTimeZoneはJSTを想定しているコードに 「*DEPEND ON JST*」という印を
つけた。(把握している分のみ)

2025/12/10: Version:1.4(非公開)
(内部メモ:subversion revision 2100, フォルダv1.4。間違えてv1.4版の
「libics2gacsv.py」を更新してしまったので2105で差し戻しを行っている。
なので、「libics2gacsv.py」は現在2105だが、2100と同じファイルのはず)

- 用語統一
   Windowsアプリ版の「Outlook(classic)」を「Outlook(legacy)」と誤記。
    →正しい「Outlook(classic)」に修正。

  「日程」→「スケジュール」に統一。

  日本語の「時間」と「時刻」の誤用→正しい日本語の意味に修正。

- 概ね下記のICSファイルに対応。
    Cybozu Garoon(Version 5.0.2)
    Web版 Outlook
    Windowsアプリ版 Outlook(classic)
  いずれも2025年12月ごろに生成されたICSファイルの出力で確認。

- 警告やメッセージはv1.3まではSTDOUTに出力してたが、STDERRに出力するよ
  うに修正。

- 上書スケジュール(RECURRENCE-ID)の対応強化。概ね確認できた範囲では正
  常に変換できる。

- 開始時刻が0:00、終了時刻が翌日0:00のスケジュールを終日スケジュールと
  みなすオプション「-g」を追加。

- (不完全)旧版では常にTimeZoneとしてJST(Asia/Tokyo)を想定していたが、
  他のTimeZoneに対応。真面目に調査していない。

  v1.3で記載した「*DEPEND ON JST*」削除。

- TimeZone情報はICSファイルのVTIMEZONEというエリアに記載されているが、
  VTIMEZONEが無いICSファイルは「Flating Time」と呼ばれる。
  その場合、以下の情報をSTDERRに出力する。

  "INFO: ICSデータにTimeZoneデータがありません。"
  "INFO: Floating Timeのデータです。(Ref: RFC5545, 3.3.12. TIME)"

- 一部ICSファイルはVTIMEZONEがにTimeZoneが2個以上定義されている。
  その場合、以下の情報をSTDERRに出力する。

   "INFO: ICSファイルにTimzeZoneが複数定義されています。"
   "INFO: 現在定義されているTimeZone一覧: {一覧表示}"
   "INFO: TimeZoneとして1番目に定義されている[初出のTimeZone]を採用します。"
   "WARNING: 採用したTimeZoneが不適切な場合、繰返しスケジュールの最終日(UNTIL)の計算に失敗し、"
   "WARNING: スケジュールが欠落する可能性あります。不適切な場合は引数で指定してください。"

   UNTILの時刻はGMTで記載されていることがあります。そのため、文脈によっ
   てはTimeZone情報が必須となります。

- マニュアル「DOWNLOAD-Outlook-ICS.pdf」加筆。

- サポートOSからAlmalinux8.10を削除しました。OS標準のPythonが3.6のため。
  依存しているライブラリvobjectがPython3.8以降対応になります。

- macOS(Intel)の動作確認環境が無くなったため、サポートOSから削除
- macOS 26.0.1/ARM/日本語環境をサポートOSに追加。

2025/12/XX: Version:2.0 (公開版)
(内部メモ:subversion revision 未定, フォルダv2.0)

- 公開用にversion打ち直し。

- マニュアル「DOWNLOAD-Outlook-ICS.pdf」を削除して、Qiitaに移動。
  https://qiita.com/qiitamatumoto/items/24343d860ccc065b4cc8


*Known bugs:

西暦を判断する基準の正規用言が「[^\\d]20[\\d]{6}」などになってるので、
西暦2000年から2099年までしか動作しない。

上書スケジュール(RECURRENCE-ID)で上書きされる元のスケジュールが1999年
以前もしくは2100年以降だと動作しない。

繰り返し命令でRDATEというのがあるらしいが、私の環境では生成されないの
で、未対応。ICSに含まれる場合は例外を送出します。

(未調査)サマータイムの切り替えがあるTimeZoneの場合、サマータイムの前後
で1日の長さが23時間もしくは25時間の場合はおそらく正常に動作しない。

Teamsの会議インフォーメーションの削除はフォーマットが変わったら無効です。
2025年9月のフォーマットを元に削除を行います。

ICSの各要素になにも入ってないというのを示すのに"N/A"という文字列を使っ
ています。SUMMARYやDESCRIPTIONに最初から"N/A"と入っていた場合は誤動作
する。

"""
#######################################
G_VERSION = "2.0"
#########################################################################
def get_ics_val(ics_parts, name, default_val=None, exit_none=True):
    """
    vobjectのicsの要素を取り出す。

    引数:
    ics_parts: ICSをよみこんだvobjectのcomponetオブジェクト。
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
    引数:
    ics_time: datetime.datetime型もしくはdatetime.date型
    timerange: CSVの出力範囲を指定するtimerangeの値。
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
### 関数中のコメント参照
flag_csv_remove_timezone = True
flag_add_am12_time = False
flag_remove_am12_time = False
#
###
def is_aware(d) -> bool:
    """datetime オブジェクトが aware(timezoneあり) かどうかを判定する"""
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
    """datetime オブジェクトが native(timezoneなし) かどうかを判定する"""
    return not is_aware(d)

###
def is_am12(t) -> bool:
    """
    時/分/秒の時刻情報があり、深夜12時"00:00:00"ならTrue
"""
    if not type(t) is datetime.datetime:
        raise ValueError(f"ERROR: 想定外の型が渡されました。{type(t)}")
    return (t.hour + t.minute + t.second) == 0
###
def hava_time(t) -> bool:
    """
    時刻情報があるかないか。

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
def init_guess_timezone(cal_tz: dateutil.tz.tz.tzical, override_timezone: str = None):
    """TimeZoneを推測する関数の初期化"""
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
    """TimeZoneを推測する。事前にinit_guess_timezone()で初期化する。"""
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
    nativeな(datetime.datetimedate,datetime.date)をawareなdatetimeに変換する。
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
    TimeZone情報があれば、ローカルタイムに変換する。
    TimeZone情報がなければ、何もしない。

    exit_none: Noneが渡された時の挙動。
               True: raiseを渡して停止。
               False : Noneを返す。

    exit_native: Timezoneがないnative timeを渡された時の挙動。
               True: raiseを渡して停止。
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
def concat_ics_date(h: str, t: list) -> str:
    """
    分解したICSのアイテムを再結合します。
"""
    return h+":"+",".join(t)

###
def find_ics_data(data: list, key: str, stop=True) -> int:
    """
    文字列型で渡されたVEVENTのデータからkeyで指示された
    要素がある行を探します。stop=Trueの場合は例外を送出します。
"""
    for i, d in enumerate(data):
        if re.match(key, d):
            return i
    if stop:
        raise RuntimeError(f"ICSのアイテム「{key}」が無いVEVENTが渡された")
    return -1

###
def bug_fix_pre_check_dtstart_and_dtend(start: str, end: str) -> bool:
    """
    DTSTART/DTENDに時刻情報が無い/TimeZoneが無い場合のみ処理する。
    OK: 「DTSTART:20250912」-> True
    NG: 「DTSTART:20250912T123344」-> False
    NG: 「DTSTART;TZID...」-> False
    """
    if re.search(r'[^\d]20[\d]{6}T', start) or re.search('TZID', start):
        return False

    if re.search(r'[^\d]20[\d]{6}T', end) or re.search('TZID', end):
        raise RuntimeError(f"ERROR: DTSTARTとDTENDの書式が異なる(本来DTSTARTで検出されるべき) DTSTART={start},DTEND={end}")

    return True

###
def bug_fix_exdate_aux(data: list) -> list:
    """
bug_fix_rruleの補助関数1

BEGIN:VEVENTからEND:VEVENTの間のデータをSTRING型のlistで渡して、
EXDATE関連を修正する。

処理の流れはbug_fix_recurrence_id_auxと同じなので、修正時は両方修正する。

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

一部のカレンダーアプリケーションが生成するRRULEのEXDATEのバグを回避するため、
vObjectの関数にICSを読み込ませる前に修正作業を行います。

バグの詳細については TECH-MEMO.txt 参照ください。

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

###
def ics_time_to_csv(ics_parts, rrule_start):
    """返り値:
    タプルで文字列４つ
      "開始日","開始時刻","終了日","終了時刻"

    例:
      "2025/09/05","14:40:00","2025/09/05","16:10:00"

    通日スケジュールの場合は時刻情報がないため、下記のようになる。
      "2025/09/05","","2025/09/05","" (1日間の通日スケジュール)
      "2025/09/05","","2025/09/06","" (2日間の通日スケジュール)

   引数:
    start: ICSデータのDTSTART(datetime.datetime型もしくはdatetime.date型)

    end:   ICSデータのDTEND(datetime.datetime型もしくはdatetime.date型)

   外部制御変数:
    flag_remove_timezone: Bool型
         False: TimeZone情報が有る場合、TimeZone情報込で出力。
            TimeZone情報あり: 12:34:55+09:00, 12:34:55-03:00
            TimeZone情報なし: 12:34:55, 12:34:55

         True: TimeZone情報がある場合、JSTならJSTの表示「+09:00」をけずる。
            JST以外ならJSTに修正の上、TimeZone情報を削って表示
            12:34:55, 12:34:55

   flag_remove_am12_time: Bool型
        Trueを指定すると、終日スケジュール(時刻あり)の時刻情報を除去する。
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
    start = get_ics_val(ics_parts, 'x-org-dtstart',  None, exit_none=False)
    if start is None:
        start = get_ics_val(ics_parts, 'dtstart')

    end = get_ics_val(ics_parts, 'x-org-dtend',  None, exit_none=False)
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

#松元用の予定の選択肢を追加
flag_matumoto_modify = False
#
def split_garoon_style_summary(summary: str):
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
"""
    # 以下正規表現文字列になるので変な記号は入れない
    # Garoonの選択肢を転載。他にも必要なら適時追加する。
    head = ['出張', '往訪', '来訪', '会議', '休み']
    #松元ローカル定義。
    # 2025/10/7: add 'TEST': ICS生成のテスト用の選択肢。
    if flag_matumoto_modify:

        head += ['TODO', 'MEMO', '授業', 'TEST']
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
#各要素の最後の改行と空白をすべて取り除く。
flag_remove_tail_cr = False
#

def modify_description(description: str) -> str:
    """
    メモ欄(description)の加工を行う。長いと見にくいのと、
    Teamsのパスワードが入ってることがあるので。

    Known bugs: Teamsの会議インフォメーションの削除は、フォーマットが変わったら無効です。
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
def ics_parts_to_csv_buffer(ics_parts, rrule_start=None) -> list:
    """内部関数"""
    row = []

    s_d, s_t, e_d, e_t = ics_time_to_csv(ics_parts, rrule_start)

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

# VERSION1.3内部版にあった、recurrence_idのバグを再現する。
# これを有効にしないと、recurrence_idのエラー検出コードがほぼ動かない。
flag_enable_recurrence_id_bug = False

# CSVの出力の日付ソートを行う(True)。しない(False)
flag_output_sort = True

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
#出力するCSVの文字コード
flag_output_sjis = True
# 試してないが、改行コードの話。
# Ref: https://qiita.com/tatsuya-miyamoto/items/f57408064b803f55cf99

# 試してないが、入力文字コードの変更はこちらが参考になる。
# Ref: https://techblog.asahi-net.co.jp/entry/2021/10/04/162109
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

            # ICSのRDATE命令はGaroonやWeb/Outlookで出現例を作れなかったため、
            # 未対応。
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
                if flag_enable_recurrence_id_bug:
                    if not is_collect_timerange(dtstart, timerange):
                        continue
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

                if flag_enable_recurrence_id_bug:
                    if not is_collect_timerange(s, timerange):
                        continue

                if G_DEBUG_UID == uid:
                    print("STEP4: PASS(timerange)", file=sys.stderr)

                t = ics_time_to_csv(component, s)
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

            org_summary_h = ""
            org_summary_t = "(REFERENCE DATA DOES NOT EXIST)"
            org_description = "(REFERENCE DATA DOES NOT EXIST)"

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
            if csv_buffer[i][k] is None:
                csv_buffer[i][k] = org_summary_t
                k = G_CSV_B_OFFSET+G_CSV_SUMMARY_H
                csv_buffer[i][k] = org_summary_h

            k = G_CSV_B_OFFSET+G_CSV_DESCRIPTION
            if csv_buffer[i][k] is None:
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

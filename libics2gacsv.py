#!/usr/bin/env python3
# -*- python -*-
# -*- coding: utf-8 -*-

import sys
import io
import re
import datetime
import pytz
import csv
import vobject

__doc__="""補足:
ソースコードに加えて別途 README.txt および TECH-MEMO.txt　参照ください。

*Cangelog:
2025/9/24: Version:1.0
初版公開

*2025/9/25: Version:1.1

変数名等でicalとicsが混ざってたので、icsに統一。

*2025/9/26: Version:1.2

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
"""
VERSION="1.2"

#######################################

#######################################
#出力するCSVの文字コード/ホントは改行コードも変更が必要だけ未対応。
flag_output_sjis = True
# 試してないが、改行コードの話。
# Ref: https://qiita.com/tatsuya-miyamoto/items/f57408064b803f55cf99

# 試してないが、入力文字コードの変更はこちらが参考になる。
# Ref: https://techblog.asahi-net.co.jp/entry/2021/10/04/162109
G_CSV_ENCODING = "shift_jis"


#########################################################################
def get_ics_val(ics_parts, name, default_val=None):
    """
    vobjectのicsの要素を取り出す。

    引数:
    ics_parts: ICSをよみこんだvobjectのcomponetオブジェクト。
    name: 取り出す要素。例えば(summaryとかdtstartなど)
    default_val: nameで指定した要素が存在しない場合に返す値。
                 もしここにNoneを指定したときにnameで指定した
    　　　　　　　　要素が存在しない場合は例外履いて終了する。
    """

    if hasattr(ics_parts, name):
        return getattr(ics_parts, name, default_val).valueRepr()

    if default_val is None:
        print(f"不適切なICSファイルです。必須パラメータがありません: {name}")
        ics_parts.prettyPrint()
        raise RuntimeError("不適切なICSファイルです。")
    return default_val

#########################################################################

def format_check_timerange(timerange:int) -> bool:
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

    #1970年未満は無効。
    if timerange//100 < 1970:
        return False

    #2100年以降は無効。
    if timerange//100 >= 2100:
        return False

    m = timerange % 100
    # 0月は無効
    if m < 1:
        return False
    # 13月以上は無効。
    if m > 12:
        return False

    #異常なし。
    return True

#########################################################################
def is_collect_timerange(ics_time, timerange:int)->bool:
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

#########################################################################
# vobjectがEXDATE関連で例外を履く記述の修正・RRULEの日付計算のミスを修正
#
# RRULEのbugfixを有効にする
flag_rrule_bugfix = True
########

def bug_fix_exdate_aux(data:list) -> list:
    """
bug_fix_rruleの補助関数1

BEGIN:VEVENTからEND:VEVENTの間のデータをSTRING型のlistで渡して、
EXDATE関連を修正する。

"""
    def concat_exdate(h:str, t:list) -> str:
        return h+":"+",".join(t)
    #
    flag_dtstart = -1
    flag_dtend = -1
    flag_exdate = -1
    flag_rrule = -1

    for i in range(len(data)):
        if re.match('RRULE', data[i]):
            flag_rrule = i
        if re.match('DTSTART', data[i]):
            flag_dtstart = i
        if re.match('DTEND', data[i]):
            flag_dtend = i
        if re.match('EXDATE', data[i]):
            flag_exdate = i

    if flag_exdate == -1:
        raise RuntimeError("EXDATEが無いVEVENTが渡された")

    if flag_dtstart == -1:
        raise RuntimeError("DTSTARTが無いVEVENTが渡された")

    if flag_dtend == -1:
        raise RuntimeError("DTENDが無いVEVENTが渡された")

    # DTSTART/DTENDに時間情報が無い/TimeZoneが無い場合のみ処理する。
    # OK: 「DTSTART:20250912」-> 先に進む
    # NG: 「DTSTART:20250912T123344」-> RETURN
    # NG: 「DTSTART;TZID...」-> RETURN
    if re.search('[^\d]20[\d]{6}T', data[flag_dtstart]):
        return data

    if re.search('TZID', data[flag_dtstart]):
        return data

    if re.search('[^\d]20[\d]{6}T', data[flag_dtend]):
        raise RuntimeError(f"DTSTARTとDTENDの書式が異なる(本来DTSTARTで検出されるべき) DTSTART={data[flag_dtstart]},DTEND={data[flag_dtend]}, ")

    if re.search('TZID', data[flag_dtend]):
        raise RuntimeError(f"DTSTARTとDTENDの書式が異なる(本来DTSTARTで検出されるべき) DTSTART={data[flag_dtstart]},DTEND={data[flag_dtend]}, ")


    # EXDATEの時間指定は複数ある場合があるので、修正時は要注意。
    # 「EXDATE:20250909,20250915」など。

    exdate = data[flag_exdate].split(':')
    if len(exdate) != 2:
        raise RuntimeError(f"DXDATEの書式異常: {data[flag_exdate]}" )

    exdate_tail = exdate[1].split(',')
    # 以降、exdata[1]に対する検索はsearch
    # exdate_tailに対する検索はfullmatch & match

    for i in range(len(exdate_tail)):
        # EXDATEのtailにZがついてる場合はTimeZone付きのため除外
        # NG: 「20250912T123344Z」-> RETURN
        if re.fullmatch('20[\d]{6}T[\d]{6}Z', exdate_tail[i]):
            return data

        # EXDATEに時間があり、時間が0時0分0秒以外の場合は除外
        # NG:「20250912T12345」-> RETURN
        if re.fullmatch('20[\d]{6}T[\d]{6}', exdate_tail[i]):
            if not re.fullmatch('20[\d]{6}T0{6}', exdate_tail[i]):
                return data

    # 各ICSの出力アプリ別の処理。
    if re.fullmatch('EXDATE', exdate[0]) and (not re.search('T', exdate[1])):
        #print(f"DEBUG exdate[1] = : {exdate[1]}")
        # maybe Garoon. 時間情報なし。
        # search: 「EXDATE:20250909」で後ろにTなし。
        exdate[0], count = re.subn('^EXDATE$', 'EXDATE;VALUE=DATE', exdate[0])
        if count == 0:
            raise RuntimeError(f"多分バグ(Garoonfix): {exdate[0]}")
        data[flag_exdate] = concat_exdate(exdate[0], exdate_tail)
        return data

    if re.fullmatch('EXDATE;TZID.+', exdate[0]) and re.search('T0{6}', exdate[1]):
        # maybe outlook/web
        # search: EXDATE;TZID=Tokyo Standard Time:20251215T000000
        #         TZIDがあって、時間が0時0分0秒。
        exdate[0],count  = re.subn('^EXDATE;TZID.+$', 'EXDATE;VALUE=DATE', exdate[0])
        if count == 0:
            raise RuntimeError(f"多分バグ(outlook/web fix-1): {exdate[0]}")
        for i in range(len(exdate_tail)):
            # 前処理で時間が000000以外のを抜いてるはずだから必ず成功するはず。
            exdate_tail[i], count = re.subn('T0{6}', '', exdate_tail[i])
            if count == 0:
                raise RuntimeError(f"多分バグ(outlook/web fix-2): {exdate_tail[i]}")
        data[flag_exdate] = concat_exdate(exdate[0], exdate_tail)
        return data

    # 修正の必要がない。もしくは想定してない仕様のICS
    return data
####

def bug_fix_until_aux(data:list) -> list:
    """
bug_fix_rruleの補助関数2. 

BEGIN:VEVENTからEND:VEVENTの間のデータをSTRING型のlistで渡して、
UNTIL関連を修正する。

バグ: TimeZoneはJSTを想定している。

"""
    flag_debug = False
    #
    flag_dtstart = -1
    flag_dtend = -1
    flag_exdate = -1
    flag_rrule = -1

    for i in range(len(data)):
        if re.match('RRULE', data[i]):
            flag_rrule = i
        if re.match('DTSTART', data[i]):
            flag_dtstart = i
        if re.match('DTEND', data[i]):
            flag_dtend = i

    if flag_dtstart == -1:
        raise RuntimeError("DTSTARTが無いVEVENTが渡された")

    if flag_dtend == -1:
        raise RuntimeError("DTENDが無いVEVENTが渡された")

    # DTSTART/DTENDに時間情報が無い/TimeZoneが無い場合のみ処理する。
    # OK: 「DTSTART:20250912」-> 先に進む
    # NG: 「DTSTART:20250912T123344」-> RETURN
    # NG: 「DTSTART;TZID...」-> RETURN
    if re.search('[^\d]20[\d]{6}T', data[flag_dtstart]):
        return data

    if re.search('TZID', data[flag_dtstart]):
        return data

    if re.search('[^\d]20[\d]{6}T', data[flag_dtend]):
        raise RuntimeError(f"DTSTARTとDTENDの書式が異なる(本来DTSTARTで検出されるべき) DTSTART={data[flag_dtstart]},DTEND={data[flag_dtend]}, ")

    if re.search('TZID', data[flag_dtend]):
        raise RuntimeError(f"DTSTARTとDTENDの書式が異なる(本来DTSTARTで検出されるべき) DTSTART={data[flag_dtstart]},DTEND={data[flag_dtend]}, ")

    if flag_debug:
        print(f"DEBUG: orig rrule: {data[flag_rrule]}")

    rruledata = data[flag_rrule].split(':')
    if len(rruledata) != 2:
        raise RuntimeError(f"RRULEの書式異常: {data[flag_rrule]}")

    rrule_tail = rruledata[1].split(';')

    for i in range(len(rrule_tail)):
        if not re.match('UNTIL', rrule_tail[i]):
            continue

        if not re.fullmatch('UNTIL=20[\d]{6}T[\d]{6}Z', rrule_tail[i]):
            return data

        m = re.fullmatch('UNTIL=([\d]{4})([\d]{2})([\d]{2})T([\d]{2})([\d]{2})([\d]{2})Z', rrule_tail[i])
        t = m.groups()
        # 暗にJSTを想定したコード
        org_date = datetime.datetime(int(t[0]),int(t[1]),int(t[2]),int(t[3]),int(t[4]),int(t[5]), tzinfo=datetime.timezone.utc)
        jst_date = org_date.astimezone(pytz.timezone('Asia/Tokyo'))
        if flag_debug:
            print (f"DEBIG: org_date = {org_date}")
            print (f"DEBIG: jst_date = {jst_date}")
        if not is_am12(jst_date) :
            raise RuntimeError(f"UNTILがAM12:00以外。許容されるべきだろうが、何がおきるかわからない。。 jst_date={jst_date}")
        rrule_tail[i] = "UNTIL=" + jst_date.strftime('%Y%m%d')
        if flag_debug:
            print(f"DEBUG: replace until: {rrule_tail[i]}")

    data[flag_rrule] = rruledata[0] + ":" + ";".join(rrule_tail)
    if flag_debug:
        print(f"DEBUG: replace rrule: {data[flag_rrule]}")
    return data
####

def bug_fix_rrule(data:str) -> str:
    """EXDATE関連のbugfix。ICSのファイルをすべて読み込んだ
string型のdataを渡して、修正して返却する。

RRULEのEXDATEとUNTILのバグを回避するため、vObjectの関数にICSを読み込ま
せる前に修正作業を行います。

バグの詳細については TECH-MEMO.txt 参照ください。
    """
    flag_debug = False
    #flag_debug = True
    #print(f"DEBUG: arg_type:  {type(data)}")
    # typeはstrを想定
    if not type(data) is str:
        raise RuntimeError(f"想定外の型が渡されました: type={type(data)}")

    # https://maku77.github.io/python/numstr/split-lines.html
    # 文字列を改行で分割する。
    lines = data.splitlines()
    org_line_num = len(lines)

    flag_in_vevent = False
    flag_hava_exdate = False
    flag_hava_rrule = False
    lines_in_vevent = []

    lines_ret = []

    for i in lines:
        if re.match('BEGIN:VEVENT', i):
            if flag_in_vevent is True:
                raise RuntimeError("BEGIN:VEVENTが二重に現れました")
            if len(lines_in_vevent) != 0:
                raise RuntimeError("Logic Error")
            flag_in_vevent = True
            lines_in_vevent.append(i)
            continue

        if re.match('END:VEVENT', i):
            if flag_in_vevent is False:
                raise RuntimeError("END:VEVENTが二重に現れました")
            flag_in_vevent = False
            lines_in_vevent.append(i)
            if flag_hava_exdate:
                if flag_debug:
                    print("DEBUG PRE:--EXDATE--\n"+'\n'.join(lines_in_vevent) + "-----\n")
                lines_in_vevent = bug_fix_exdate_aux(lines_in_vevent)
                if flag_debug:
                    print("DEBUG AFT:----\n"+'\n'.join(lines_in_vevent) + "-----\n")
                flag_hava_exdate = False

            if flag_hava_rrule:
                if flag_debug:
                    print("DEBUG PRE:--UNTIL--\n"+'\n'.join(lines_in_vevent) + "-----\n")
                lines_in_vevent = bug_fix_until_aux(lines_in_vevent)
                if flag_debug:
                    print("DEBUG AFT:----\n"+'\n'.join(lines_in_vevent) + "-----\n")
                flag_hava_rrule = False

            lines_ret += lines_in_vevent
            lines_in_vevent = []
            continue

        if not flag_in_vevent:
            lines_ret.append(i)
            continue

        if re.match('EXDATE', i):
            flag_hava_exdate = True

        if re.match('RRULE', i):
            flag_hava_rrule = True

        lines_in_vevent.append(i)
    ##
    if len(lines_ret) != org_line_num:
        raise RuntimeError("行数が変化している。たぶんバグ")

    return "\n".join(lines_ret) + "\n"

#########################################################################
### 関数中のコメント参照
flag_csv_remove_timezone = True
flag_add_am12_time = False
####

def is_am12(t) -> bool:
    """
        時間/分/秒の時間情報があり、深夜12時"00:00:00"ならTrue
    """
    return (t.hour + t.minute + t.second) == 0
#
def hava_time(t) -> bool:
    """
        時間情報があるかないか。
    """
    if type(t) is datetime.date:
        return False

    if type(t) is datetime.datetime:
        return True

    raise RuntimeError(f"想定外の型が渡された: = {type(t)}")

def hava_timezone(t) -> bool:
    """
        TimeZone情報があるかないか。
    """
    return hasattr(t, "tzinfo")
##

def ics_time_to_csv(ics_parts, rrule_start=None):
    """
    返り値:
    タプルで文字列４つ
      "開始日","開始時刻","終了日","終了時刻"

    例:
      "2025/09/05","14:40:00","2025/09/05","16:10:00"

    通日日程の場合は時間情報がないため、下記のようになる。
      "2025/09/05","","2025/09/05","" (1日間の通日日程)
      "2025/09/05","","2025/09/06","" (2日間の通日日程)

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

    flag_add_am12_time: Bool型
         True: 通日日程の場合は時間情報がないが開始時刻を「00:00:00」とする。
　　　　　　　　　終了時刻を翌日の「00:00:00」とする。これはOutlook/legacyが
               出力する形式。

         False(Default): 通日日程の場合は開始時刻,終了時刻を空文字列とする。
　　　　　　　　　これはGaroonが出力する形式。
               ただし、入力者が明示的に時間を入力した場合は、通日日程であっても
               時間情報があるため、Trueと同じ出力になる。

         例: 2025年9月11日の終日日程の場合。
            "開始日","開始時刻","終了日","終了時刻"
         Trueのとき
            "2025/09/11", "00:00:00", "2025/09/12", "00:00:00"

         Falseのとき
            "2025/09/11", "", "2025/09/11", ""

"""
    start = get_ics_val(ics_parts, 'dtstart')
    end = get_ics_val(ics_parts, 'dtend')

    if False:
        print(f"DEBUG: rrule_start = {rrule_start}")
        print(f"DEBUG: type(rrule_start) = {type(rrule_start)}")
        print(f"DEBUG: start = {start}")
        print(f"DEBUG: type(start) = {type(start)}")
        print(f"DEBUG: end = {end}")
        print(f"DEBUG: type(end) = {type(end)}")
        print(f"DEBUG: flag_csv_remove_timezone = {flag_csv_remove_timezone}")
        print(f"DEBUG: flag_add_am12_time = {flag_add_am12_time}")


    if hava_timezone(start) != hava_timezone(end):
        raise RuntimeError("想定外: timezoneありなし混在")

    if hava_time(start) != hava_time(end):
        raise RuntimeError("想定外: timeありなし混在")

    # 時間情報がない(日付のみ)の情報にTimeZoneがある。
    # 本来は許容されるべきだが、何がおきるかわからないので、未対応とする。
    if hava_timezone(start) and (not hava_time(start)):
        raise RuntimeError("想定外: タイムゾーン情報があるのに時間情報がない")

    if rrule_start:
        # rrule_startは常に時間情報ありのdatetime型
        # start/endは時間情報ありのdatetime型と時間情報なしのdate型のときがある。
        # 足し算でdate型がdetatime型に昇格するとまずいので、rrule_startの
        # 時刻が00:00:00のときはdete型に降格させる。

        if is_am12(rrule_start):
            # rrule_start時刻が00:00:00のときは時間情報が無いとみなしdate型に降格
            rrule_start = rrule_start.date()
            if False:
                print(f"DEBUG: rrule_start(conv) = {rrule_start}")
                print(f"DEBUG: type(rrule_start)(conv) = {type(rrule_start)}")

        end = rrule_start+(end-start)
        start = rrule_start

    # TimeZone情報あるなら無条件にJSTに変換。
    if hava_timezone(start):
        # https://engineering.mobalab.net/2022/05/27/python-datetime/
        tz_jst = datetime.timezone(datetime.timedelta(hours=9))
        start = start.astimezone(tz_jst)
        end = end.astimezone(tz_jst)

    s_d = "N/A"
    s_t = "N/A"
    e_d = "N/A"
    e_t = "N/A"
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
        #以下時間情報がない場合の処理
        # 終日日程
        # Outlook(legacy) style CSV
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

##########################################################################

# ICSのsummaryの分割を試みる。
flag_split_summary = True

#松元用の予定の選択肢を追加
flag_matumoto_modify = False
#
def split_garoon_style_summary(summary:str):
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
    if flag_matumoto_modify:
        head += ['TODO','MEMO','授業']
    # コロンの半角。
    splitter = [':']
    # コロンの全角。
    splitter += ['：']

    for s in splitter:
        ret = summary.split(s,1)
        if len(ret) == 1:
            continue
        for h in head:
            if re.fullmatch(h, ret[0]):
                ret[1] = re.sub('^\s+','', ret[1])
                return ret[0], ret[1]
    #分割失敗
    return "", summary
#
##########################################################################
#メモ欄(description)の4行目以降を消す。
flag_description_delete_4th_line_onwards=False
# Teamsの会議インフォメーションを消す。パスワードが入ってる。
flag_remove_teams_infomation = True
#各要素の最後の改行と空白をすべて取り除く。
flag_remove_tail_cr = False
#
def modify_description(description:str) -> str:
    """
    メモ欄(description)の加工を行う。長いと見にくいのと、
    Teamsのパスワードが入ってることがあるので。

    バグ: Teamsの会議インフォメーションの削除は、フォーマットが変わったら無効です。
    2025/9現在のフォーマットをもとに削除を行います。
"""
    if flag_remove_teams_infomation:
        lines = description.splitlines()
        new_line = []
        for i in lines:
            if re.search("Microsoft Teams ヘルプが必要ですか", i):
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
        if len(lines)>r:
            description = "\n".join(lines[:r])
        else:
            description = "\n".join(lines)
        description += "\n"

    return description

def csv_oneline_write(csv_stream, ics_parts, rrule_start=None):
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
    summary_t = get_ics_val(ics_parts, 'summary', 'N/A')
    #print(f"DEBUG: summary = {summary_t}")
    #ics_parts.prettyPrint()

    if flag_split_summary:
        #ICS形式の場合は予定(選択肢のところ)が無いので生成試みる
        summary_h, summary_t =  split_garoon_style_summary(summary_t)

    row.append(summary_h) # 予定(選択肢のところ)
    row.append(summary_t) # 予定本文

    description = modify_description(get_ics_val(ics_parts, 'description', ''))
    row.append(description) # メモ

    #各要素の最後の改行と空白をすべて取り除く。
    if flag_remove_tail_cr:
        for i in range(len(row)):
            row[i] = re.sub('\s+$','',row[i])

    csv_stream.writerow(row)

##########################################################################
def guess_timerange(FILENAME):
    """
ファイル名をもとに期間指定の推測を行います。
"""
    if re.search("all", FILENAME):
        return 0

    t = re.search("\d{6}", FILENAME)
    if not t:
        return None
    #print (f"DEBUG: {t}")
    ret = int(t.group())
    if not format_check_timerange(ret):
        return None
    return ret

##########################################################################

# GaroonのCSVの最初のヘッダ
flag_print_csv_header = False
#

def ics2csv(ics_file_path:str, csv_file_path:str, timerange:int=0):
    """
    ICS(iCalendar)ファイルをCSVファイルに変換する。

    引数:
        ics_file_path (str): 変換元のICS(iCalendar)ファイル。"stdin"を指定すると標準入力。
        csv_file_path (str): 変換先のCSVファイル。"stdout"を指定すると標準出力。
        timerange (int): CSVに変換する日時を限定する場合は、指定する。
                         2025年8月分がほしい場合は「202508」と指定する。
                         未指定だと全部変換する。
    返り値:
        void。失敗したら停止する。
    """
    ######################

    if vobject.VERSION != "0.9.9":
        print("ERROR: 依存ライブラリvobjectのバージョンが開発環境と異なります。",file=sys.stderr)
        print(f"ERROR: githubよりプルリクエストください。 vobject.VERSION={vobject.VERSION}",file=sys.stderr)
        sys.exit()
    

    if ics_file_path == "stdout" or ics_file_path[0] == "-":
        raise RuntimeError(f"入力ファイル名指定エラー: {ics_file_path}")

    if ics_file_path == "stdin":
        ics_data = sys.stdin.read()
    else:
        with open(ics_file_path, 'r', encoding='utf-8') as f:
            ics_data = f.read()

    if len(ics_data) < 10:
        raise RuntimeError(f"ファイル読み込みエラー: len(ics_data) = {len(ics_data)}")
    ######################
    # vobjectが例外を履く記述の修正/vobjectのRRULEの計算ミスを修正
    if flag_rrule_bugfix:
        ics_data = bug_fix_rrule(ics_data)
    ######################
    # 読み込んだデータをvobjectに変換。
    calendar = vobject.readOne(ics_data)
    ######################

    if csv_file_path == "stdin"  or csv_file_path[0] == "-":
        raise RuntimeError(f"ファイル名指定エラー: {csv_file_path}")

    # errors='xmlcharrefreplace' utfからsjisに変換時にsjis未定義コードが出たときに "&#xxxx;"に変換する。
    # Ref: https://docs.python.org/ja/3/howto/unicode.html
    # Ref: https://zenn.dev/hassaku63/articles/f7ca587b86398c
    #
    #escale_type='backslashreplace'
    escale_type='xmlcharrefreplace'
    if csv_file_path == "stdout":
        #https://geroforce.hatenablog.com/entry/2018/12/05/114633
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding=G_CSV_ENCODING, errors=escale_type, newline="")
        csv_out = sys.stdout
    else:
        csv_out = open(csv_file_path, 'w', encoding=G_CSV_ENCODING, errors=escale_type, newline="")

    # Pythonライブラリの仕様でCSVの最後の改行はCR+LF。
    # Ref: https://docs.python.org/ja/3/library/csv.html
    # ->「Dialect.lineterminator」
    # -> 「writer が作り出す各行を終端する際に用いられる文字列です。デフォルトでは '\r\n' です。」
    csv_writer = csv.writer(csv_out,quoting=csv.QUOTE_ALL)

    #Garoonの出力形式。
    if flag_print_csv_header:
        header = ["開始日","開始時刻","終了日","終了時刻","予定","予定詳細","メモ"]
        csv_writer.writerow(header)

    for component in calendar.components():
        #print(f"DEBUG: component.name = {component.name}")
        if component.name == 'VEVENT':  # Process only VEVENT components
            #print("DEBUG")
            dtstart = get_ics_val(component, 'dtstart')

            # ICSのRDATE命令はGaroonやOutlookで出現例を作れなかったため、
            # 未対応。
            if "N/A" != get_ics_val(component, 'rdate', "N/A"):
                raise RuntimeError("ICSデータに未対応の繰り返し命令RDATEが使われています。")

            # ICSのRRULE命令が未使用ならそのまま出力する。
            if "N/A" == get_ics_val(component, 'rrule', "N/A"):
                if is_collect_timerange(dtstart, timerange):
                    csv_oneline_write(csv_writer, component)
                continue

            # ICSのRRULE命令の処理。
            rrule_set = component.getrruleset(addRDate=True)
            for s in rrule_set:
                if is_collect_timerange(s, timerange):
                    csv_oneline_write(csv_writer, component, s)

    # end for()
    if csv_file_path != "stdout":
        csv_out.close()
        print(f"変換に成功しました: '{ics_file_path}' to '{csv_file_path}'")
#end func

############################################
if __name__ == '__main__':
    help("libics2csv")

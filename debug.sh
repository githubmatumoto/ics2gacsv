#!/bin/bash
echo "引数のデバック用バッチファイル。Linuxのデバック用。"
echo "カレントディレクトリにICSファイルを calendar.ics という名前で置いてください。"
#sjis
python3 ics2gacsv.py       all calendar.ics  gen-debug-sjis.csv
python3 ics2gacsv.py -m    all calendar.ics  gen-debug-sjis-m.csv
python3 ics2gacsv.py       all calendar.ics  stdout > gen-debug-stdout.csv
#utf-8
python3 ics2gacsv.py -u    all calendar.ics  gen-debug-u.csv
python3 ics2gacsv.py -u  guess calendar.ics  gen-debug-202510.csv
python3 ics2gacsv.py -u  202511 calendar.ics  gen-debug-202511.csv
python3 ics2gacsv.py -u -m all calendar.ics  gen-debug-m.csv
#
python3 ics2gacsv.py -u -s all calendar.ics  gen-debug-s.csv
python3 ics2gacsv.py -u -t all calendar.ics  gen-debug-t.csv
python3 ics2gacsv.py -u -o all calendar.ics  gen-debug-o.csv
python3 ics2gacsv.py -u -g all calendar.ics  gen-debug-g.csv
python3 ics2gacsv.py -u -d all calendar.ics  gen-debug-d.csv
python3 ics2gacsv.py -u -p all calendar.ics  gen-debug-p.csv
python3 ics2gacsv.py -u -r all calendar.ics  gen-debug-r.csv
python3 ics2gacsv.py -u -w all calendar.ics  gen-debug-w.csv

#
echo "TimeZoneとしてJSTとESTを明示指定。"
echo "明示指定のTimeZoneはTimeZoneが不明な文脈でのみ利用される"
python3 ics2gacsv.py -u  -T"Asia/Tokyo"  all calendar.ics  gen-debug-jst.csv
echo "以下は(おそらく)ICSが想定しているTimeZoneと異なる値を指定"
python3 ics2gacsv.py -u  -T"US/Eastern"  all calendar.ics  gen-debug-est.csv

echo "以下はRECURRENCE-IDの処理をしないため多数警告でる場合あるが、それで正常"
python3 ics2gacsv.py -u -x all calendar.ics  gen-debug-x.csv

echo "以下はUNTILの処理をしないため例外をはく場合があるが、それで正常"
python3 ics2gacsv.py -u -b all calendar.ics  gen-debug-b.csv


echo "以下のコマンドで比較できる。"
echo "ファイルに出力した時とstdoutに出した時の違い"
echo diff gen-debug-sjis.csv gen-debug-stdout.csv
echo "SUMMARY分割をするしない。"
echo diff gen-debug-u.csv gen-debug-s.csv
echo "SUMMARY分割の作者用の修正"
echo diff gen-debug-u.csv gen-debug-m.csv
echo "日付出力の違い(Garoon形式とWindowsの旧版Outlookが出力する形式)"
echo diff gen-debug-u.csv gen-debug-o.csv
echo "日付出力の違い(0時開始翌日0時終了の日程を終日日程とみなす)"
echo diff gen-debug-u.csv gen-debug-g.csv
echo "Teamsの会議インフォメーションを消す消さない"
echo diff gen-debug-u.csv gen-debug-p.csv
echo "メモ欄(description)の4行目以降を消す消さない"
echo diff gen-debug-u.csv gen-debug-d.csv
echo "ダブルクオート内の最後の改行や空白の除去をするしない"
echo diff gen-debug-u.csv gen-debug-r.csv
echo "RECURRENCE_ID処理で上書きされたイベントを(Hidden: )を表示するしない。"
echo diff gen-debug-u.csv gen-debug-w.csv
echo "RECURRENCE_ID処理を一切しない。"
echo diff gen-debug-u.csv gen-debug-x.csv
echo "TimeZoneを無理やり指定(JST)/指定が適切なら差分なし(もしくは不明な文脈なし)"
echo diff gen-debug-u.csv gen-debug-jst.csv
echo "TimeZoneを無理やり指定(EST)/指定が適切なら差分なし(もしくは不明な文脈なし)"
echo diff gen-debug-u.csv gen-debug-est.csv
echo "日付出力の違い(TimeZoneの有り無し)"
echo head gen-debug-u.csv gen-debug-t.csv
echo
#EOF

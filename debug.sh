#!/bin/bash
echo "引数のデバック用バッチファイル。Linuxのデバック用。"
echo "カレントディレクトリにICSファイルを calendar.ics という名前で置いてください。"
python3 ics2gacsv.py       all calendar.ics  gen-ou-debug-sjis.csv
python3 ics2gacsv.py -m    all calendar.ics  gen-ou-debug-sjis-m.csv
python3 ics2gacsv.py       all calendar.ics  stdout > gen-ou-debug-stdout.csv
python3 ics2gacsv.py -u    all calendar.ics  gen-ou-debug-u.csv
python3 ics2gacsv.py -u  guess calendar.ics  gen-ou-debug-202508.csv
python3 ics2gacsv.py -u  202509 calendar.ics  gen-ou-debug-202509.csv
python3 ics2gacsv.py -u -m all calendar.ics  gen-ou-debug-m.csv
python3 ics2gacsv.py -u -s all calendar.ics  gen-ou-debug-s.csv
python3 ics2gacsv.py -u -t all calendar.ics  gen-ou-debug-t.csv
python3 ics2gacsv.py -u -o all calendar.ics  gen-ou-debug-o.csv
python3 ics2gacsv.py -u -d all calendar.ics  gen-ou-debug-d.csv
python3 ics2gacsv.py -u -p all calendar.ics  gen-ou-debug-p.csv
python3 ics2gacsv.py -u -r all calendar.ics  gen-ou-debug-r.csv

echo "以下はRRULEのバグフィックスを行わない場合。"
echo "RRULEの処理で例外を吐いて問題ない。"
python3 ics2gacsv.py -u -b all   calendar.ics  gen-ou-debug-b.csv

echo "以下のコマンドで比較できる。"
echo "ファイルに出力した時とstdoutに出した時の違い"
echo diff gen-ou-debug-sjis.csv gen-ou-debug-stdout.csv
echo "SUMMARY分割をするしない。"
echo diff gen-ou-debug-u.csv gen-ou-debug-s.csv
echo "SUMMARY分割の作者用の修正"
echo diff gen-ou-debug-u.csv gen-ou-debug-m.csv
echo "日付出力の違い(Garoon形式とWindowsの旧版Outlookが出力する形式)"
echo diff gen-ou-debug-u.csv gen-ou-debug-o.csv
echo "Teamsの会議インフォメーションを消す消さない"
echo diff gen-ou-debug-u.csv gen-ou-debug-p.csv
echo "メモ欄(description)の4行目以降を消す消さない"
echo diff gen-ou-debug-u.csv gen-ou-debug-d.csv
echo "ダブルクオート内の最後の改行や空白の除去をするしない"
echo diff gen-ou-debug-u.csv gen-ou-debug-r.csv
echo "日付出力の違い(TimeZoneの有り無し)"
echo head gen-ou-debug-u.csv gen-ou-debug-t.csv
echo 
#EOF

#!/bin/bash
echo Linuxのテスト用バッチファイル
echo CSVはUTF-8で生成しています。ShiftJISにする場合はオプション -u を消してください。
#
python3 ics2gacsv.py -u -m guess calendar.ics  gen-ou202507.csv
python3 ics2gacsv.py -u -m guess calendar.ics  gen-ou202508.csv
python3 ics2gacsv.py -u -m guess calendar.ics  gen-ou202509.csv
python3 ics2gacsv.py -u -m guess calendar.ics  gen-ou202510.csv
python3 ics2gacsv.py -u -m guess calendar.ics  gen-ou202511.csv
python3 ics2gacsv.py -u -m guess calendar.ics  gen-ou202512.csv
python3 ics2gacsv.py -u -m guess calendar.ics  gen-ou202601.csv
python3 ics2gacsv.py -u -m guess calendar.ics  gen-ou202602.csv
python3 ics2gacsv.py -u -m all   calendar.ics  gen-ou-all.csv
#EOF

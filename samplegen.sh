#!/bin/bash
echo "Linuxのテスト用バッチファイル"
echo "カレントディレクトリにICSファイルを calendar.ics という名前で置いてください。"
echo "Excelで読む時は文字コードShiftJISのファイルを使ってください"
#
echo "文字コードUTF-8で生成"
python3 ics2gacsv.py -u -m 202511 calendar.ics  utf8-gen-202511.csv
python3 ics2gacsv.py -u -m guess calendar.ics  utf8-gen-202604.csv
python3 ics2gacsv.py -u -m guess calendar.ics  utf8-gen-202605.csv
python3 ics2gacsv.py -u -m guess calendar.ics  utf8-gen-202606.csv
python3 ics2gacsv.py -u -m all   calendar.ics  utf8-gen-all.csv
echo "文字コードShiftJISで生成"
python3 ics2gacsv.py -m 202511 calendar.ics  sjis-gen-202511.csv
python3 ics2gacsv.py -m guess calendar.ics  sjis-gen-202604.csv
python3 ics2gacsv.py -m guess calendar.ics  sjis-gen-202605.csv
python3 ics2gacsv.py -m guess calendar.ics  sjis-gen-202606.csv
python3 ics2gacsv.py -m all   calendar.ics  sjis-gen-all.csv
#EOF

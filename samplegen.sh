#!/bin/bash
echo "Linuxのテスト用バッチファイル"
echo "カレントディレクトリにICSファイルを calendar.ics という名前で置いてください。"
echo "Excelで読む時は文字コードShiftJISのファイルを使ってください"
#
echo "文字コードUTF-8で生成"
python3 ics2gacsv.py -u  202509 calendar.ics  utf8-gen-202509.csv
python3 ics2gacsv.py -u  guess  calendar.ics  utf8-gen-202510.csv
python3 ics2gacsv.py -u  all    calendar.ics  utf8-gen-all.csv
echo "文字コードShiftJISで生成"
python3 ics2gacsv.py  202509 calendar.ics  sjis-gen-202509.csv
python3 ics2gacsv.py  guess  calendar.ics  sjis-gen-202510.csv
python3 ics2gacsv.py  all    calendar.ics  sjis-gen-all.csv
#EOF

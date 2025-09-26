#!/bin/bash
echo "Linuxのテスト用バッチファイル"
echo "カレントディレクトリにICSファイルを calendar.ics という名前で置いてください。"
echo "Excelで読む時は文字コードShiftJISのファイルを使ってください"
#
echo "文字コードUTF-8で生成"
python3 ics2gacsv.py -u -m guess calendar.ics  utf8-gen-ou202507.csv
python3 ics2gacsv.py -u -m 202508 calendar.ics  utf8-gen-ou202508.csv
python3 ics2gacsv.py -u -m guess calendar.ics  utf8-gen-ou202509.csv
python3 ics2gacsv.py -u -m guess calendar.ics  utf8-gen-ou202510.csv
python3 ics2gacsv.py -u -m guess calendar.ics  utf8-gen-ou202511.csv
python3 ics2gacsv.py -u -m guess calendar.ics  utf8-gen-ou202512.csv
python3 ics2gacsv.py -u -m guess calendar.ics  utf8-gen-ou202601.csv
python3 ics2gacsv.py -u -m guess calendar.ics  utf8-gen-ou202602.csv
python3 ics2gacsv.py -u -m all   calendar.ics  utf8-gen-ou-all.csv
echo "文字コードShiftJISで生成"
python3 ics2gacsv.py -m guess calendar.ics  sjis-gen-ou202507.csv
python3 ics2gacsv.py -m guess calendar.ics  sjis-gen-ou202508.csv
python3 ics2gacsv.py -m 202509 calendar.ics  sjis-gen-ou202509.csv
python3 ics2gacsv.py -m guess calendar.ics  sjis-gen-ou202510.csv
python3 ics2gacsv.py -m guess calendar.ics  sjis-gen-ou202511.csv
python3 ics2gacsv.py -m guess calendar.ics  sjis-gen-ou202512.csv
python3 ics2gacsv.py -m guess calendar.ics  sjis-gen-ou202601.csv
python3 ics2gacsv.py -m guess calendar.ics  sjis-gen-ou202602.csv
python3 ics2gacsv.py -m all   calendar.ics  sjis-gen-ou-all.csv
#EOF

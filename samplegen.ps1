echo "Windowsのテスト用PowerShell"
echo "カレントディレクトリにICSファイルを calendar.ics という名前で置いてください。"
echo "CSVはShiftJISで生成しています。UTF-8にする場合は引数 -u を追加してください。"
python.exe ics2gacsv.py -m guess calendar.ics  sjis-gen-ou202507.csv
python.exe ics2gacsv.py -m 202508 calendar.ics  sjis-gen-ou202508.csv
python.exe ics2gacsv.py -m guess calendar.ics  sjis-gen-ou202509.csv
python.exe ics2gacsv.py -m guess calendar.ics  sjis-gen-ou202510.csv
python.exe ics2gacsv.py -m guess calendar.ics  sjis-gen-ou202511.csv
python.exe ics2gacsv.py -m guess calendar.ics  sjis-gen-ou202512.csv
python.exe ics2gacsv.py -m guess calendar.ics  sjis-gen-ou202601.csv
python.exe ics2gacsv.py -m guess calendar.ics  sjis-gen-ou202602.csv
python.exe ics2gacsv.py -m all   calendar.ics  sjis-gen-ou-all.csv

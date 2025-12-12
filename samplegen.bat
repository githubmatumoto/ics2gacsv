REM Windowsのテスト用バッチファイル
REM カレントディレクトリにICSファイルを calendar.ics という名前で置いてください。
REM CSVはShiftJISで生成しています。UTF-8にする場合は引数 -u を追加してください。
python.exe ics2gacsv.py -m 202511 calendar.ics  sjis-gen-202511.csv
python.exe ics2gacsv.py -m guess calendar.ics  sjis-gen-202604.csv
python.exe ics2gacsv.py -m guess calendar.ics  sjis-gen-202605.csv
python.exe ics2gacsv.py -m guess calendar.ics  sjis-gen-202606.csv
python.exe ics2gacsv.py -m all   calendar.ics  sjis-gen-all.csv
REM EOF

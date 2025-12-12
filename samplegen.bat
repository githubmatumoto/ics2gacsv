REM Windowsのテスト用バッチファイル
REM カレントディレクトリにICSファイルを calendar.ics という名前で置いてください。
REM CSVはShiftJISで生成しています。UTF-8にする場合は引数 -u を追加してください。
python3 ics2gacsv.py  202509 calendar.ics  sjis-gen-202509.csv
python3 ics2gacsv.py  guess  calendar.ics  sjis-gen-202510.csv
python3 ics2gacsv.py  all    calendar.ics  sjis-gen-all.csv
REM EOF

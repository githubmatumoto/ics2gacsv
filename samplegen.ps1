echo "Windows�̃e�X�g�pPowerShell"
echo "�J�����g�f�B���N�g����ICS�t�@�C���� calendar.ics �Ƃ������O�Œu���Ă��������B"
echo "CSV��ShiftJIS�Ő������Ă��܂��BUTF-8�ɂ���ꍇ�͈��� -u ��ǉ����Ă��������B"
python.exe ics2gacsv.py -m guess calendar.ics  sjis-gen-ou202507.csv
python.exe ics2gacsv.py -m 202508 calendar.ics  sjis-gen-ou202508.csv
python.exe ics2gacsv.py -m guess calendar.ics  sjis-gen-ou202509.csv
python.exe ics2gacsv.py -m guess calendar.ics  sjis-gen-ou202510.csv
python.exe ics2gacsv.py -m guess calendar.ics  sjis-gen-ou202511.csv
python.exe ics2gacsv.py -m guess calendar.ics  sjis-gen-ou202512.csv
python.exe ics2gacsv.py -m guess calendar.ics  sjis-gen-ou202601.csv
python.exe ics2gacsv.py -m guess calendar.ics  sjis-gen-ou202602.csv
python.exe ics2gacsv.py -m all   calendar.ics  sjis-gen-ou-all.csv

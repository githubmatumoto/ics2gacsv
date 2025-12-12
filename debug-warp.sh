echo "GaroonのICSテスト"


rm -rf *.csv calendar.ics debug-ga debug-ou debug-ouc
cp -p SAMPLE/ga.ics calendar.ics
sh debug.sh
mkdir debug-ga
mv *.csv  debug-ga

echo "Outlook(classic)のICSテスト"

rm -f calendar.ics
cp -p SAMPLE/outlook-classic.ics calendar.ics
sh debug.sh
mkdir debug-ouc
mv *.csv  debug-ouc

echo "Outlook(new)のICSテスト"

rm -f calendar.ics
cp -p SAMPLE/outlook-new.ics calendar.ics
sh debug.sh
mkdir debug-ou
mv *.csv  debug-ou

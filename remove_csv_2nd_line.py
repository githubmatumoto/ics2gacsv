#!/usr/bin/env python3
# -*- python -*-
# -*- coding: utf-8 -*-
__doc__="""
STDINから読み込んだCSVの各要素の2行目を削除する。
1行目のみにする。1行目の最後の改行を除去する。
出力はSTDOUT。文字コードはUTF-8を想定している。
UTF-8以外ならnkfなどで適切に変換する。
"""

import sys
import csv

reader = csv.reader(sys.stdin)
csv_writer = csv.writer(sys.stdout,quoting=csv.QUOTE_ALL)

for row in reader:
    new_row = []
    for i in row:
        lines = i.splitlines()
        if len(lines) >= 1:
            new_row.append(lines[0])
        else:
            new_row.append("")
    csv_writer.writerow(new_row)

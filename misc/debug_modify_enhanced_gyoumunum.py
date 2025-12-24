#!/usr/bin/env python3
# -*- python -*-
# -*- coding: utf-8 -*-
#
# Copyright (c) 2025 MATSUMOTO Ryuji.
# License: Apache License 2.0
#
import re
import sys
import libics2gacsv

__doc__="""ICS(iCalendar)のmodify_enhanced_gyoumunum()のデバグ用
"""

########################################

def debug(description: str, summary: str):
    print("===")
    print(f"--(Pre)--\n{description}----")
    ret = libics2gacsv.modify_enhanced_gyoumunum(description, summary)
    print(f"--(Aft)--\n{ret}----\n")
    print("===")

    if not re.search("^[0-9０-９]", description):
        return

    #頭に全角空白いれる
    description = "　"+ description
    print(f"--(Pre)--\n{description}----")
    ret = libics2gacsv.modify_enhanced_gyoumunum(description, summary)
    print(f"--(Aft)--\n{ret}----\n")

if __name__ == '__main__':
    if libics2gacsv.G_VERSION != "2.1":
        print("ERROR: ファイルが古いです。最新のics2gacsv.pyとlibics2gacsv.pyをダウンロードしてください。",file=sys.stderr)
        sys.exit()

    s = "出張: 東京特許許可局 %333"
    s = "出張: 東京特許許可局 g333"
    s = "出張: 東京特許許可局 g３３３" # 全角数字
    s = "出張: 東京特許許可局 g３３３　" # 全角数字+最後に全角空白
    #s = "出張: 東京特許許可局"

    if True:
        """Description-Type1-1:
        -(Pre)-------
        数字B\\s*
        -------------
        -(Aft)-------
        数字A
        -------------
        """
        d = "9999"
        debug(d,s)
        d = "9999 "
        debug(d,s)
        d = "9999 \n"
        debug(d,s)
        d = "9999　\n" # 全角SPACE
        debug(d,s)

        d = "9999"
        debug(d,s)
        d = "9999 "
        debug(d,s)
        d = "9999 \n"
        debug(d,s)
        d = "9999　\n" # 全角SPACE
        debug(d,s)

    if True:
        """Description-Type1-2:
        -(Pre)-------
        (N/A)\\s*
        -------------
        -(Aft)-------
        数字A
        -------------
        """
        d = "(N/A)"
        debug(d,s)
        d = "(N/A) "
        debug(d,s)
        d = "(N/A)\n"
        debug(d,s)
        d = "(N/A) \n"
        debug(d,s)
        d = "(N/A)　\n" #  全角SPACE
        debug(d,s)
        d = "(N/A)\nABCD"
        debug(d,s)
        d = "(N/A)ABCD"
        debug(d,s)

    if True:
        """Description-Type1-3:
        -(Pre)-------
        \\s*
        -------------
        -(Aft)-------
        数字A
        -------------
        """
        d = ""
        debug(d,s)
        d = " "
        debug(d,s)
        d = "\n"
        debug(d,s)
        d = "　" # 全角SPACE
        debug(d,s)
        d = "　 \t"
        debug(d,s)

    if True:
        """
        Description-Type2:
        行数は変化しない。
        -(Pre)-------
        数字B(空白文字)*(改行文字).*
        -------------
        -(Aft)-------
        数字A(改行文字).*
        -------------

        """
        print("**行数は変化しない**")
        d = "9999\nABC"
        debug(d,s)
        d = "9999　\nABC"# 全角SPACE
        debug(d,s)
        d = "9999\n\nABC"
        debug(d,s)

    if True:

        """
        Description-Type3:
        行数は変化しない。
        -(Pre)-------
        (空白文字)*(改行文字).*
        -------------
        -(Aft)-------
        数字A*(改行文字).*
        -------------
        """
        print("**行数は変化しない**")
        d = "\nABC"
        debug(d,s)
        d = "　\nABC"# 全角SPACE
        debug(d,s)

    if True:
        """
        Description-Type4:
        行数が1行増える
        -(Pre)-------
        [可|急](空白文字)*(改行文字).*
        -------------
        -(Aft)-------
        数字A(改行文字)[可|急](空白文字)*(改行文字).*
        -------------
        """
        print("**行数が1行増える**")
        d = "可"
        debug(d,s)
        d = "急\nABC"
        debug(d,s)
        d = "急\nABC\n"
        debug(d,s)

        d = "急 "
        debug(d,s)
        d = "可　\nABC"
        debug(d,s)
        d = "可 \nABC\n"

        debug(d,s)
        d = " 可"
        debug(d,s)
        d = " 急\nABC"
        debug(d,s)
        d = " 急\nABC\n"
        debug(d,s)

        d = "　急 "
        debug(d,s)
        d = "　可　\nABC"
        debug(d,s)
        d = "　可 \nABC\n"
        debug(d,s)

    if True:
        """
        Description-Type5:
        行数が2行増える
        -(Pre)-------
        .*
        -------------
        -(Aft)-------
        数字A(改行文字).*
        -------------
        """
        print("**行数が1行増える**")
        d = "ABC"
        debug(d,s)
        d = "ABC\n"
        debug(d,s)


#End of main()

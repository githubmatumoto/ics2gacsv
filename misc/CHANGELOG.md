# Cangelog:
## 2025/9/24: Version:1.0 (公開終了)

初版公開

## 2025/9/25: Version:1.1 (公開終了)

- 用語統一
  変数名等でicalとicsが混ざってたので、icsに統一。

## 2025/9/26: Version:1.2 (公開終了)
(内部メモ:subversion revision 2070, フォルダv1.2)

- CSVの改行コードの扱いが悪く無駄な空行が入るのを改善。生成される改行
  コードについては misc/TECH-MEMO.txt を参照ください。

- CSVでダブルクオートで囲まれた文字列の最後に改行が空白があった場合除
  去するコードを追加。

- 出力文字コードがShiftJISの時にUTF-8からShiftJISに変換する作業に失敗
  した時のエスケープ手段がstdoutに出力した時とファイルに出力した時で異なっ
  たので、統一。

- macOS 13.7.8/Intel/日本語環境動作確認しました。

- ライブラリのバージョン管理用の変数を追加しています。

- その他用語の統一など行ってます。

## 2025/9/30: Version:1.3 (非公開)
(内部メモ:subversion revision 2079, フォルダv1.3)

- 上書スケジュール(RECURRENCE-ID)へ限定対応。細かいRECURRENCE-ID命令に
  は対応していない。挙動がおかしい場合はオプション「-w」で無効化された
  命令をCSVに出力するようにして欲しい。

- 上書スケジュール(RECURRENCE-ID命令)の処理を無効にするオプション「-x」
  を追加。

- 上書スケジュール対応のため、CSVに出力する前に一度バッファリングを行
うように変更。

- CSVに出力時の特殊な値を定義
  ここで述べる値をカレンダーアプリに直接記載すると誤動作する。

-- 「(N/A)」: 指定の要素がVEVENTに無かった。SUMMARY/DESCRIPTONのみ。
   ICSアプリにより、上記要素が無い場合、下記の挙動がある。

```
   要素を未定義とする場合: 「(N/A)」が入る。
   改行をいれる場合: 空行が入る。
```

-- 「(REFERENCE DATA DOES NOT EXIST)」: 滅多にないはずだが、
 RECURRENCE-IDで上書をするVEVENTにはSUMMARY/DESCRIPTONが無い場合がある。
 基のVEVENTからコピーする処理を行っているが、基のVEVENTが見つからなかっ
 た。

-- 「Hidden: 」オプション「-w」を指定した時のみ。

  RECURRENCE-IDで上書をした基のVEVENTのSUMMARYに「Hidden: 」を付けて出
  力する。

- 暗にTimeZoneとしてJSTを想定しているコードに 「*DEPEND ON JST*」とい
  う印をつけた。(把握している分のみ)

## 2025/12/10: Version:1.4(非公開)
(内部メモ:subversion revision 2100, フォルダv1.4。間違えてv1.4版の
「libics2gacsv.py」を更新してしまったので2105で差し戻しを行っている。
なので、「libics2gacsv.py」は現在2105だが、2100と同じファイルのはず)

- 用語統一
   Windowsアプリ版の「Outlook(classic)」を「Outlook(legacy)」と誤記。
    →正しい「Outlook(classic)」に修正。

```
   「日程」→「スケジュール」に統一。
   日本語の「時間」と「時刻」の誤用→正しい日本語の意味に修正。
```

- 概ね下記のICSファイルに対応。

```
    Cybozu Garoon(Version 5.0.2)
    Web版 Outlook
    Windowsアプリ版 Outlook(classic)
```

いずれも2025年12月ごろに生成されたICSファイルの出力で確認。

- 警告やメッセージはv1.3まではSTDOUTに出力してたが、STDERRに出力するよ
  うに修正。

- 上書スケジュール(RECURRENCE-ID)の対応強化。概ね確認できた範囲では正
  常に変換できる。

- 開始時刻が0:00、終了時刻が翌日0:00のスケジュールを終日スケジュールと
  みなすオプション「-g」を追加。

- (不完全)旧版では常にTimeZoneとしてJST(Asia/Tokyo)を想定していたが、
  他のTimeZoneに対応。真面目に調査していない。

  v1.3で記載した「*DEPEND ON JST*」削除。

- TimeZone情報はICSファイルのVTIMEZONEというエリアに記載されているが、
  VTIMEZONEが無いICSファイルは「Flating Time」と呼ばれる。
  その場合、以下の情報をSTDERRに出力する。

```
  "INFO: ICSデータにTimeZoneデータがありません。"
  "INFO: Floating Timeのデータです。(Ref: RFC5545, 3.3.12. TIME)"
```

- 一部ICSファイルはVTIMEZONEがにTimeZoneが2個以上定義されている。
  その場合、以下の情報をSTDERRに出力する。

```
   "INFO: ICSファイルにTimzeZoneが複数定義されています。"
   "INFO: 現在定義されているTimeZone一覧: {一覧表示}"
   "INFO: TimeZoneとして1番目に定義されている[初出のTimeZone]を採用します。"
   "WARNING: 採用したTimeZoneが不適切な場合、繰返しスケジュールの最終日(UNTIL)の計算に失敗し、"
   "WARNING: スケジュールが欠落する可能性あります。不適切な場合は引数で指定してください。"
```

   UNTILの時刻はGMTで記載されていることがあります。そのため、文脈によっ
   てはTimeZone情報が必須となります。

   複数定義されている場合、TimeZoneの指定を行うオプション「-T」を追加。

- マニュアル「DOWNLOAD-Outlook-ICS.pdf」加筆。

- macOS(Intel)の動作確認環境が無くなったため、サポートOSから削除
- macOS 26.0.1/ARM/日本語環境をサポートOSに追加。

- Pythonのサポートバージョンを3.9以降に修正。
  Python3.9はすでにEOLのため本来なら3.10以上にしたいが
  macOSの標準Pythonが3.9なので3.9以上としています。。

- サポートOSからAlmalinux8.10を削除しました。
  OS標準のPythonが3.6のため。ただしOS標準パッケージにPython3.11が含ま
  れるため、インストールすれば、動作可能。

## 2025/12/18: Version:2.0 (公開版)
(内部メモ:subversion revision 2124, フォルダv2.0)

- 公開用にversion打ち直し。

- マニュアル「DOWNLOAD-Outlook-ICS.pdf」を削除して、Qiitaに移動。
  https://qiita.com/qiitamatumoto/items/24343d860ccc065b4cc8

- Ubuntuのサポートバージョンを22.04LTSから24.04LTSに変更。
  上記は誤記修正になります。開発用に使っていた環境を22.04と誤認してい
  たため。リリース版公開に合わせて再確認を行ったところ24.04でした。

- コード内の各種コメントの精査。

	- 変更履歴をソースコード内から misc/CHANGELOG.md に移動。

## 2025/12/20: Version:2.1 (beta版)
(内部メモ:subversion revision 未定, フォルダv2.1)

- 職場の業務記録提出用のスクリプト kiroku.py 追加。ファイル名を決
  め打ちで生成する。

- SUMMARYに業務番号を記載する、拡張追加。
```
  libics2gacsv.py:
  関数 modify_enhanced_gyoumunum()　追加
  制御変数 flag_enhanced_gyoumunum 追加

  ics2gacsv.py:
  引数 -z 追加(flag_enhanced_gyoumunum=True)
```

- 上書スケジュール(RECURRENCE-ID)のバグ対応。

Outlookが生成したICSファイルでもとの繰返し(RRULE)スケジュールのDESCRIPTIONが未定義"(N/A)"であり、
上書きスケジュールのDESCRIPTIONが未定義"(N/A)"であった場合は、
誤ってDESCRIPTIONに"(REFERENCE DATA DOES NOT EXIST)"が入いるようになった。

```
libics2gacsv.py:
def modify_reference_id_data()
   誤:     if csv_buffer[i][k] is None:
   正:     if (csv_buffer[i][k] is None) or (outlook_bugfix and csv_buffer[i][k] == UNREF):
```

- 修正split_garoon_style_summary
```
  変数名修正
  旧:flag_matumoto_modify

  新: flag_split_summary_enhance
  具体的な追加項目を今まで関数内にハードコーディングしてたが、isc2gacsv.pyで
  下記変数に代入する形に修正
  G_SPLIT_SUMMARY_ENHANCE
```

# Known bugs:

- 西暦を判断する基準の正規表現が「[^\\d]20[\\d]{6}」などになってる
	ので、
西暦2000年から2099年までしか動作しない。

- 上書スケジュール(RECURRENCE-ID)で上書きされる元のスケジュールが1999年
以前もしくは2100年以降だと動作しない。

- 繰り返し命令でRDATEというのがあるらしいが、手元のICSのアプリでは生成さ
れないので、未対応。ICSに含まれる場合は例外を送出します。

- (未調査)サマータイムの切り替えがあるTimeZoneの場合、サマータイムの前後
で1日の長さが23時間もしくは25時間の場合はおそらく正常に動作しない。

- Teamsの会議インフォーメーションの削除はフォーマットが変わったら無効。
2025年9月のフォーマットを元に削除を行います。

- ICSの各要素になにも入ってないというのを示すのに"(N/A)"という文字列を使っ
ています。SUMMARYやDESCRIPTIONに最初から"(N/A)"と入っていた場合は誤動作
する。

- ICSの上書スケジュール(RECURRENCE-ID)繰返し要素の処理に
     "(REFERENCE DATA DOES NOT EXIST)"
という文字列を使っています。SUMMARYやDESCRIPTIONに最初から上記が入って
場合は誤動作する。

- 業務番号を暗に4桁と想定している。5桁以上なら下記関数の正規表現を修正する。
	modify_enhanced_gyoumunum()

- (2025/12/25追記)Windowsアプリ版 Outlook(classic)で生成したICSファイルで一部条件
でICSのRDATE命令が使われるようです。RDATEが生成されたICSファイルは
変換できません。把握した範囲では、ICSファイル出力時の詳細情報の設定で
「詳細情報の一部」を選んだ時に生成される事があります。

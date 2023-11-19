# DupFileWithDate : Windows のエクスプローラーで、日付を付けてファイルを複製

## 使い方

Windows のエクスプローラーでファイルを選択して、右クリックし[ファイル複製(日付付き)]を選ぶと、日付付きでファイルを複製します。

ファイル名に日付と思われる数値がある場合は、今日の日付に置き換えます。(日付の書式は YYYY-MM-DD と、YYYYMMDD に対応しています。)

## インストール方法

1. フォルダー C:\Dev\DupFileWithDate を作成して、DupFileWithDate.vbs を配置してください。
2. DupFileWithDate.reg を実行して、レジストリーに登録してください。

## カスタマイズ方法

インストール先を変更したい場合は、DupFileWithDate.reg の 7 行目を修正してください。

```Windows Registry Entries:DupFileWithDate.reg 
@="wscript.exe \"C:\\Dev\\DupFileWithDate\\DupFileWithDate.vbs\" \"%1\""
```

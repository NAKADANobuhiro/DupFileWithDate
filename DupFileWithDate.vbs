'オプション指定
Option Explicit 

'変数宣言
Dim fsObj
Dim strFileName
Dim strNewFileName
Dim strParentFolderName
Dim strBaseName
Dim ext
dim objShell
Set objShell = WScript.CreateObject("WScript.Shell")

Dim regEx, strDate

'引数が1でなければ終了
If WScript.Arguments.Count <> 1 Then WScript.Quit

strFileName = WScript.Arguments(0) 'ファイル名取得
Set fsObj = CreateObject("Scripting.FileSystemObject") 'ファイルシステムオブジェクトを取得
strBaseName = fsObj.GetBaseName(strFileName) 'ファイルのベース名取得

Set regEx = New RegExp
regEx.Global		= True	
regEx.IgnoreCase	= True	
regEx.Global = False        

regEx.pattern = "[0-9]{4}(0[1-9]|1[0-2])(0[1-9]|[12][0-9]|3[01])" '日付らしき yyyymmdd 数値にマッチ
strDate = Replace(Left(Now(),10), "/", "") '今日の日付を yyyymmdd 型で取得

If regEx.Test(strBaseName) Then 'ファイル名に日付(yyyymmdd)が含まれていれば
    strBaseName = regEx.Replace(strBaseName, strDate) 'その日付を今日の日付に置き換える
else
    regEx.pattern = "[0-9]{4}-(0[1-9]|1[0-2])-(0[1-9]|[12][0-9]|3[01])" '日付らしき数値 yyyy-mm-dd にマッチ
    strDate = Replace(Left(Now(),10), "/", "-") '今日の日付を yyyy-mm-dd 型で取得
    If regEx.Test(strBaseName) Then 'ファイル名に日付(yyyy-mm-dd)が含まれていれば
        strBaseName = regEx.Replace(strBaseName, strDate) 'その日付を今日の日付に置き換える
    else    
        strBaseName = strBaseName & "_" & strDate '含まれていなければ今日の日付を付加する
    end if
end if

ext = LCase(fsObj.GetExtensionName(strFileName))'ファイルの拡張子取得
strParentFolderName = fsObj.GetParentFolderName(strFileName)'ファイルの親フォルダ名取得
strNewFileName = strParentFolderName & "\" & strBaseName & "." & ext'日付日時を付与したファイル名を作成

'コピー処理
if strFileName <> strNewFileName then
    fsObj.CopyFile strFileName, strNewFileName, true
else
    objShell.Popup "既に今日の日付がついています。", 5,, vbInformation
end if

'終了処理
set fsObj = Nothing
set strFileName = Nothing
set strNewFileName = Nothing
set strParentFolderName = Nothing
set strBaseName = Nothing
set ext = Nothing
WScript.Quit
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
## 余談:なぜ作成したか

日頃の業務で、議事録-2023-11-19.docx といった名前のファイルを作成することがあり、体裁は同じなので前回のファイルを開いて、修正して別名で保存するといった方法をとっている方も多いと思います。

ところが最近のアプリでは自動での上書き保存の機能がついていたり、クセで保存してしまうことで、前回のファイルを壊してしまう危険性があります。

そこで、この VBScript を使用することで、まず最初にファイルを複製し、その複製したファイルを修正することで、上記のような事故を避けることを目的としています。

## 余談:VBScript のサポート終了

自分用に作成して使っていたものを、気まぐれで公開したのですが、 
[Windows クライアントの非推奨の機能 \- What's new in Windows \| Microsoft Learn](https://learn.microsoft.com/ja-jp/windows/whats-new/deprecated-features) によると、VBScript は非推奨で、しばらくはオンデマンドで提供されるとのことです。("Windows の機能の有効化または無効化"で提供されるものと思われる。) 

もっと早く公開しておけば良かった... JScript ならまだいけるのか? PowerShell で移植する必要があるのか?

## 余談:Windows 11 エクスプローラーの仕様変更

Windows 11 では、エクスプローラーの右クリックの仕様が Windows 10 までとは異なり、右クリック後[その他のプションを表示]で、さらにメニューを開く必要があります。

Windows 10 までのすべてが表示される仕様に戻すためには、レジストリーを変更する必要があるとことですが、自分で個別に変更する方法はないのだろうか?

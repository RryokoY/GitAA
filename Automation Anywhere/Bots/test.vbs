Option Explicit

'エラーが発生しても中断しない
On Error Resume Next

Dim objStdIn    '標準入力用オブジェクト
Dim objStdOut   '標準出力用オブジェクト
Dim intExitCode '終了コード

'標準入力用オブジェクトのインスタンス作成
Set objStdIn = Wscript.StdIn

'標準出力用オブジェクトのインスタンス作成
Set objStdOut = Wscript.StdOut

'標準入力を1行ずつ読み込み標準出力へ書き込む

Do While objStdIn.AtEndOfStream = false

  objStdOut.WriteLine  objStdIn.ReadLine

Loop

'インスタンスの破棄
Set objStdIn   = Nothing
Set objStdOut  = Nothing

If Err.Number <> 0 Then
    intExitCode = 1 'エラー
Else
    intExitCode = 0 '正常
End If

'Quitメソッドの引数値は、バッチファイルでerrorlevelになる。
Wscript.Quit(intExitCode)
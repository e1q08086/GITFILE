Option Explicit
 
Dim intArgsCount
 
intArgsCount = Wscript.Arguments.Count
 
'引数の数を取得し数が合わなければ処理を中止する。
If intArgsCount <> 2 Then
  WScript.Echo "引数が" & intArgsCount & "個渡されました。" & vbCrLf _
             & "指定可能な引数の数は 2 個なので処理を中止します。"
  
  '処理を中断
  WScript.Quit
End If
 
WScript.Echo "処理開始"
WScript.Echo "1 個目の引数の値：" & Wscript.Arguments(0)
WScript.Echo "2 個目の引数の値：" & Wscript.Arguments(1) 
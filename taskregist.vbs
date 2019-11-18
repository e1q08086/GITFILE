call test()

Sub test()
    Dim execDate '処理実行日付を格納
    
    ' 起動日を取得
    execDate = GetLastDay(Now)
    WScript.Echo "実行日時は" & execDate
    
    ' タスクスケジューラタスク削除
    Call DeleteSchtask("NDATA\MyTask")
    WScript.Echo "実行日時は" & execDate

    ' タスクスケジューラタスク作成
    Call CreateSchtask("NDATA\MyTask","once","2019/12/31", "07:30")
    WScript.Echo "実行日時は" & execDate
    
'schtasks /create /tn "Backup App" /tr c:\windows\backup.cmd /sc daily /st '04:00:00
    
    
End Sub

Function GetLastDay(MyDate)
    Dim strDate
    Dim str
    Dim youbi
    
    '与えられた日付の翌月を求めます。
    strDate = DateAdd("m", 1, MyDate)
    
    '日にちを10日にします。
    strDate = Year(strDate) & "/" & Right("0" & Month(strDate), 2) & "/10"
    
    '曜日を求める
    'WScript.Echo "短い形式では " & WeekdayName(Weekday(strDate), True)
    youbi = WeekdayName(Weekday(strDate), True)
    
    Select Case youbi
    Case "月"
    Case "火"
    Case "水"
    Case "木"
    Case "金"
    Case "土"
        '月曜日起動に変更（+2日）
        strDate = DateAdd("d", 2, strDate)
    Case "日"
        '月曜日起動に変更（+1日）
        strDate = DateAdd("d", 1, strDate)
    End Select
    
    GetLastDay = strDate
    
End Function
             
Sub CreateSchtask(strTn, strSc, strSd, strSt) 
    Set objShell = CreateObject("WScript.Shell")
    objShell.Run "cmd /c schtasks /delete /tn ""NDATA\MyTask"" /F",0,True    
End Sub
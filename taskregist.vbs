call test()

Sub test()
    Dim execDate '�������s���t���i�[
    
    ' �N�������擾
    execDate = GetLastDay(Now)
    WScript.Echo "���s������" & execDate
    
    ' �^�X�N�X�P�W���[���^�X�N�폜
    Call DeleteSchtask("NDATA\MyTask")
    WScript.Echo "���s������" & execDate

    ' �^�X�N�X�P�W���[���^�X�N�쐬
    Call CreateSchtask("NDATA\MyTask","once","2019/12/31", "07:30")
    WScript.Echo "���s������" & execDate
    
'schtasks /create /tn "Backup App" /tr c:\windows\backup.cmd /sc daily /st '04:00:00
    
    
End Sub

Function GetLastDay(MyDate)
    Dim strDate
    Dim str
    Dim youbi
    
    '�^����ꂽ���t�̗��������߂܂��B
    strDate = DateAdd("m", 1, MyDate)
    
    '���ɂ���10���ɂ��܂��B
    strDate = Year(strDate) & "/" & Right("0" & Month(strDate), 2) & "/10"
    
    '�j�������߂�
    'WScript.Echo "�Z���`���ł� " & WeekdayName(Weekday(strDate), True)
    youbi = WeekdayName(Weekday(strDate), True)
    
    Select Case youbi
    Case "��"
    Case "��"
    Case "��"
    Case "��"
    Case "��"
    Case "�y"
        '���j���N���ɕύX�i+2���j
        strDate = DateAdd("d", 2, strDate)
    Case "��"
        '���j���N���ɕύX�i+1���j
        strDate = DateAdd("d", 1, strDate)
    End Select
    
    GetLastDay = strDate
    
End Function
             
Sub CreateSchtask(strTn, strSc, strSd, strSt) 
    Set objShell = CreateObject("WScript.Shell")
    objShell.Run "cmd /c schtasks /delete /tn ""NDATA\MyTask"" /F",0,True    
End Sub
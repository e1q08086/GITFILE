Option Explicit
 
Dim intArgsCount
 
intArgsCount = Wscript.Arguments.Count
 
'�����̐����擾����������Ȃ���Ώ����𒆎~����B
If intArgsCount <> 2 Then
  WScript.Echo "������" & intArgsCount & "�n����܂����B" & vbCrLf _
             & "�w��\�Ȉ����̐��� 2 �Ȃ̂ŏ����𒆎~���܂��B"
  
  '�����𒆒f
  WScript.Quit
End If
 
WScript.Echo "�����J�n"
WScript.Echo "1 �ڂ̈����̒l�F" & Wscript.Arguments(0)
WScript.Echo "2 �ڂ̈����̒l�F" & Wscript.Arguments(1) 
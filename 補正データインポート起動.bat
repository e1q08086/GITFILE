@echo off

REM フォーマットチェック
cscript C:\script\wday.vbs
if errorlevel 1 cscript C:\script\mailSend.vbs "異常" "桁数エラー" & goto ERROR
if errorlevel 2 cscript C:\script\mailSend.vbs "異常" "桁数エラー" & goto ERROR
REM インポート処理
cscript C:\script\wday.vbs
if errorlevel 1 call mailSend.bat "異常" "桁数エラー" & goto ERROR
if errorlevel 2 call mailSend.bat "異常" "年月エラー" & goto ERROR
REM 削除処理
cscript C:\script\wday.vbs
if errorlevel 1 call mailSend.bat "異常" "桁数エラー" & goto ERROR
if errorlevel 2 call mailSend.bat "異常" "年月エラー" & goto ERROR

Echo "正常終了"
cscript C:\script\mailSend.vbs "桁数エラー"
goto END

:ERROR
Echo "異常終了"

:END


pause
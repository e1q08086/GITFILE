@echo off

REM �t�H�[�}�b�g�`�F�b�N
cscript C:\script\wday.vbs
if errorlevel 1 cscript C:\script\mailSend.vbs "�ُ�" "�����G���[" & goto ERROR
if errorlevel 2 cscript C:\script\mailSend.vbs "�ُ�" "�����G���[" & goto ERROR
REM �C���|�[�g����
cscript C:\script\wday.vbs
if errorlevel 1 call mailSend.bat "�ُ�" "�����G���[" & goto ERROR
if errorlevel 2 call mailSend.bat "�ُ�" "�N���G���[" & goto ERROR
REM �폜����
cscript C:\script\wday.vbs
if errorlevel 1 call mailSend.bat "�ُ�" "�����G���[" & goto ERROR
if errorlevel 2 call mailSend.bat "�ُ�" "�N���G���[" & goto ERROR

Echo "����I��"
cscript C:\script\mailSend.vbs "�����G���["
goto END

:ERROR
Echo "�ُ�I��"

:END


pause
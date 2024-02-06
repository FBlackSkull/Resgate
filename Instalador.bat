::Automação do Resgate de Arquivos::

@echo off

TITLE Automacao do Resgate de Arquivos
CLS
echo AGUARDE INSTALANDO REQUESITOS BASICOS...
echo.
echo.
pip install paramiko
pip install o openpyxl
python.exe -m pip install --upgrade pip

icacls c:\BKP\* /GRANT Administrators:(F)
icacls c:\BKP\* /GRANT System:(F)
icacls c:\BKP\* /GRANT Guest:(F)
icacls c:\BKP\* /GRANT DefaultAccount:(F)

icacls D:\BKP\* /GRANT Administrators:(F)
icacls D:\BKP\* /GRANT System:(F)
icacls D:\BKP\* /GRANT Guest:(F)
icacls D:\BKP\* /GRANT DefaultAccount:(F)

md C:\RESGATE
net use J: \\172.18.10.2\Assistencia$
xcopy J:\#ARQUIVOS#\RESGATE C:\RESGATE /y /c
net use J: /delete

set TARGET='C:\RESGATE\RESGATE.bat'
set SHORTCUT='%USERPROFILE%\Desktop\RESGATE PMB.lnk'
set Work='C:\RESGATE'
set PWS=powershell.exe -ExecutionPolicy Bypass -NoLogo -NonInteractive -NoProfile

%PWS% -Command "$ws = New-Object -ComObject WScript.Shell; $s = $ws.CreateShortcut(%SHORTCUT%); $S.TargetPath = %TARGET%; $S.WorkingDirectory = %Work%; $S.Save()"
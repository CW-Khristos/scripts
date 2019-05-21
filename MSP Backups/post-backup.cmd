@echo off
set errorlevel=
cscript.exe //nologo "C:\Scripts\msp_postbackup.vbs" > "C:\Scripts\postbackup_vbsoutput.txt"
if %errorlevel% == 0 goto end

:fail
set /A interr=%errorlevel%+2147221504
set errorlevel=
echo %date% : %time% : POST-BACKUP.CMD FAIL : %interr% >> "C:\Scripts\postbackup_cmdout.txt"
echo %interr%
exit %interr%

:end
set /A interr=%errorlevel%+2147221504
set errorlevel=
echo %date% : %time% : POST-BACKUP.CMD SUCCESS : %interr% >> "C:\Scripts\postbackup_cmdout.txt"
exit 0
set errorlevel=
cscript.exe //nologo "C:\Scripts\msp_prebackup.vbs" > "C:\Scripts\prebackup_vbsoutput.txt"
if %errorlevel% == 0 goto end

:fail
set /A interr=%errorlevel%+2147221504
set errorlevel=
echo %date% : %time% : PRE-BACKUP.CMD FAIL : %interr% >> "C:\Scripts\prebackup_cmdout.txt"
echo %interr%
exit %interr%

:end
set /A interr=%errorlevel%+2147221504
set errorlevel=
echo %date% : %time% : PRE-BACKUP.CMD SUCCESS : %interr% >> "C:\Scripts\prebackup_cmdout.txt"
exit 0
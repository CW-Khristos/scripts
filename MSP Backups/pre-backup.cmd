@echo off
cscript.exe //nologo "C:\Scripts\msp_prebackup.vbs" > "C:\Scripts\prebackup_output.txt" 
if errorlevel == 0 goto end

:fail
set /A interr=%errorlevel%+2147221504
echo %interr%
exit /B %interr%

:end
exit /B 0
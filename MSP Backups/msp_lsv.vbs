''MSP_LSV.VBS
''SCRIPT IS DESIGNED TO SIMPLY EXPORT MSP BACKUP SETTINGS USING CLIENTTOOL.EXE UTILITY
''EXPORTS MSP BACKUP SETTINGS TO C:\TEMP\LSV.TXT
''MUST BE USED IN CONJUNCTION WITH MSP BACKUP SYNCHRONIZATION - LSV SYNCHRONIZATION.AMP CUSTOM SERVICE
''WRITTEN BY : CJ BLEDSOE / CJ<@>THECOMPUTERWARRIORS.COM
''NO REQUIRED PARAMETERS / DOES NOT ACCEPT PARAMETERS
on error resume next
''DEFINE VARIABLES
dim errRET, retDEL, strDLM, intDIFF
dim objIN, objOUT, objARG, objWSH, objFSO, objLOG, objEXEC
''DEFAULT SUCCESS
errRET = 0
''STDIN / STDOUT
set objIN = wscript.stdin
set objOUT = wscript.stdout
set objARG = wscript.arguments
''CREATE OBJECTS
set objWSH = createobject("wscript.shell")
set objFSO = createobject("scripting.filesystemobject")
''PREPARE LOGFILE
if (objFSO.fileexists("C:\temp\lsv.txt")) then  ''PREVIOUS LOGFILE EXISTS
  objFSO.deletefile "C:\temp\lsv.txt", true
end if
if (objFSO.fileexists("C:\temp\lsv.txt")) then  ''LOGFILE EXISTS
  objFSO.deletefile "C:\temp\lsv.txt", true
  set objLOG = objFSO.createtextfile("C:\temp\lsv.txt")
  objLOG.close
  set objLOG = objFSO.opentextfile("C:\temp\lsv.txt", 8)
else                                            ''LOGFILE NEEDS TO BE CREATED
  set objLOG = objFSO.createtextfile("C:\temp\lsv.txt")
  objLOG.close
  set objLOG = objFSO.opentextfile("C:\temp\lsv.txt", 8)
end if
''CHECK EXECUTION METHOD OF SCRIPT
strIN = lcase(mid(wscript.fullname, instrrev(wscript.fullname, "\") + 1))
if (strIN <> "cscript.exe") Then                ''LAUNCHED VIA WSCRIPT, RE-LAUNCH WITH CSCRIPT
  objOUT.write vbnewline & "SCRIPT LAUNCHED VIA EXPLORER, EXECUTING SCRIPT VIA CSCRIPT..."
  objWSH.run "cscript.exe //nologo " & chr(34) & Wscript.ScriptFullName & chr(34)
  wscript.quit
end if
''NO ARGUMENTS REQUIRED
''READ PASSED COMMANDLINE ARGUMENTS
'if (wscript.arguments.count > 0) then          ''ARGUMENTS WERE PASSED
'  for x = 0 to (wscript.arguments.count - 1)
'    objOUT.write vbnewline & now & vbtab & "ARGUMENT " & (x + 1) & " (ITEM " & x & ") " & " PASSED : " & ucase(objARG.item(x))
'  next 
'end if

''------------
''BEGIN SCRIPT
''EXPORT MSP BACKUP SETTINGS USING CLIENTTOOL UTILITY
'call HOOK("C:\Program Files\Backup Manager\clienttool.exe control.setting.list > " & chr(34) & "C:\temp\lsv.txt" & chr(34))
set objHOOK = objWSH.exec("C:\Program Files\Backup Manager\clienttool.exe control.setting.list")
strIN = objHOOK.stdout.readall
arrIN = split(strIN, vbnewline)
for intIN = 0 to ubound(arrIN)                                  ''CHECK SETTINGS LINE BY LINE, EXCLUDE THE 'C:\WINDOWS\TEMP' AND 'C:\TEMP' DIRECTORIES TO AVOID FALSE MONITOR ALERTS
  if ((instr(1, ucase(arrIN(intIN)), "c:\") = 0) and _
    (instr(1,ucase(arrIN(intIN)), "\temp") = 0)) then
      objOUT.write vbnewline & now & vbtab & arrIN(intIN) 
      objLOG.write vbnewline & now & vbtab & arrIN(intIN)
  end if
next
set objHOOK = nothing
''CLEAN OBJECTS
call CLEANUP
''END SCRIPT
''------------

''SUB-ROUTINES
sub HOOK(strCMD)                                ''CALL HOOK TO MONITOR OUTPUT OF CALLED COMMAND
  on error resume next
  'comspec = objWSH.ExpandEnvironmentStrings("%comspec%")
  ''CREATE OBJECT HOOK TO CALLED COMMAND
  set objHOOK = objWSH.exec(strCMD)
  'while (objHOOK.status = 0)
    ''MONITOR STDOUT OF OBJECT HOOK
    while (not objHOOK.stdout.atendofstream)
      strIN = objHOOK.stdout.readline
      if (strIN <> vbnullstring) then
        objOUT.write vbnewline & now & vbtab & strIN 
        objLOG.write vbnewline & now & vbtab & strIN 
      end if
    wend
    wscript.sleep 10
  'wend
  ''WRITE STDOUT TO LOG
  strIN = objHOOK.stdout.readall
  if (strIN <> vbnullstring) then
    objOUT.write vbnewline & now & vbtab & strIN 
    objLOG.write vbnewline & now & vbtab & strIN 
  end if
  ''WRITE ERRORS TO LOG
  'retSTOP = objHOOK.exitcode
  set objHOOK = nothing
  if (err.number <> 0) then
    objOUT.write vbnewline & now & vbtab & err.number & vbtab & err.description
  end if
end sub

sub CLEANUP()                                   ''SCRIPT CLEANUP
  objLOG.close
  ''EMPTY OBJECTS
  set objLOG = nothing
  set objFSO = nothing
  set objWSH = nothing
  set objARG = nothing
  set objOUT = nothing
  set objIN = nothing
  ''END SCRIPT, RETURN DEFAULT NO ERROR
  wscript.quit 0
end sub
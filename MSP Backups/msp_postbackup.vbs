'on error resume next
''DECLARE VARIABLES
dim retSTART
dim objDB, objSRV
dim objWSH, objFSO, objOUT
''CREATE SCRIPTING SHELL & FILE SYSTEM OBJECTS
set objOUT = wscript.stdout
set objWSH = createobject("wscript.shell")
set objFSO = createobject("scripting.filesystemobject")
''DEFAULT FAIL
retSTART = 5
call STARTDB()
''START EAGLESOFT DATABASE
sub STARTDB()
  objOUT.write vbnewline & "STARTING EAGLESOFT DATABASE : " & now
  ''CALL PATTERSONSERVERSTATUS.EXE WITH 'START' SWITCH, DO NOT MONITOR, PROCESS DOES NOT EXIT
  retSTART = objWSH.run(chr(34) & "C:\EagleSoft\Shared Files\PattersonServerStatus.exe" & chr(34) & " -start", 0, false)
  'while (objDB.status = 0)
  '  while(not objDB.stdout.atendofstream)
  '    objOUT.write vbnewline & (objDB.stdout.readline())
  '  wend
  '  wscript.sleep 10
  'wend
  'objOUT.write vbnewline & objDB.stdout.readall()
  'retSTART = objDB.exitcode
  'set objDB = nothing
  ''ERROR RETURNED
  if (retSTART = 0) then
    objOUT.write vbnewline & vbnewline & retSTART & vbtab & "EAGLESOFT DATABASE STARTED : " & now
  elseif (retSTART <> 0) then
    objOUT.write vbnewline & vbnewline & retSTART & vbtab & "ERROR STARTING EAGLESOFT DATABASE : " & now
    retSTART = 1
    call CLEANUP()
  end if
  call STARTEAGLE()
end sub

''START EAGLESOFT SERVICES
sub STARTEAGLE()
  objOUT.write vbnewline & vbnewline & "STARTING EAGLESOFT SERVICES : " & now
  ''START PATTERSON APP SERVICE
  set objSRV = objWSH.exec("net start " & chr(34) & "PattersonAppService" & chr(34))
  while (objSRV.status = 0)
    while (not objSRV.stdout.atendofstream)
      objOUT.write vbnewline & vbtab & objSRV.stdout.readline()
    wend
    wscript.sleep 10
  wend
  objOUT.write vbnewline & vbtab & objSRV.stdout.readall()
  retSTART = objSRV.exitcode
  set objSRV = nothing
  ''ERROR RETURNED
  if (retSTART <> 0) then
    ''SERVICE ALREADY STARTED
    if (retSTART = 2) then
      objOUT.write vbnewline & retSTART & vbtab & "SERVICE ALREADY STARTED : PattersonAppService : " & now
      retSTART = 0
    ''ANY OTHER ERROR
    elseif (retSTART <> 2) then
      objOUT.write vbnewline & retSTART & vbtab & "ERROR STARTING : PattersonAppService : " & now
      retSTART = 2
    end if
  end if
  objOUT.write vbnewline & vbnewline & "STARTING EAGLESOFT UPDATE SERVICES : " & now
  ''START PATTERSON APP SERVICE
  set objSRV = objWSH.exec("net start " & chr(34) & "PattersonUpdateService" & chr(34))
  while (objSRV.status = 0)
    while (not objSRV.stdout.atendofstream)
      objOUT.write vbnewline & vbtab & objSRV.stdout.readline()
    wend
    wscript.sleep 10
  wend
  objOUT.write vbnewline & vbtab & objSRV.stdout.readall()
  retSTART = objSRV.exitcode
  set objSRV = nothing
  ''ERROR RETURNED
  if (retSTART <> 0) then
    ''SERVICE ALREADY STARTED
    if (retSTART = 2) then
      objOUT.write vbnewline & retSTART & vbtab & "SERVICE ALREADY STARTED : PattersonUpdateService : " & now
      retSTART = 0
    ''ANY OTHER ERROR
    elseif (retSTART <> 2) then
      objOUT.write vbnewline & retSTART & vbtab & "ERROR STARTING : PattersonUpdateService : " & now
      retSTART = 3
    end if
  end if
  ''END SCRIPT, RETURN EXIT CODE
  call CLEANUP()
end sub

''SCRIPT CLEANUP
sub CLEANUP()
  if (retSTART = 0) then
    objOUT.write vbnewline & "POST-BACKUP COMPLETE : " & now
  elseif (retSTART <> 0) then
    objOUT.write vbnewline & "POST-BACKUP FAILURE : " & now
    Call Err.Raise(vbObjectError + retSTART, "post-backup", "fail")
  end if

  set objOUT = nothing
  set objFSO = nothing
  set objWSH = nothing
  wscript.quit retSTART
end sub
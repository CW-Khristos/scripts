on error resume next
''DECLARE VARIABLES
dim retSTOP
dim objSRV, objDB, objCOPY
dim objWSH, objFSO, objOUT
''CREATE SCRIPTING SHELL & FILE SYSTEM OBJECTS
set objOUT = wscript.stdout
set objWSH = createobject("wscript.shell")
set objFSO = createobject("scripting.filesystemobject")
''DEFAULT FAIL
retSTOP = 5
call STOPEAGLE()
''STOP EAGLESOFT SERVICES
sub STOPEAGLE()
  objOUT.write vbnewline & vbnewline & "STOPPING EAGLESOFT UPDATE SERVICES : " & now
  ''STOP PATTERSON APP SERVICE
  set objSRV = objWSH.exec("net stop " & chr(34) & "PattersonUpdateService" & chr(34))
  while (objSRV.status = 0)
    while (not objSRV.stdout.atendofstream)
      objOUT.write vbnewline & vbtab & objSRV.stdout.readline()
    wend
    wscript.sleep 10
  wend
  objOUT.write vbnewline & vbtab & objSRV.stdout.readall()
  retSTOP = objSRV.exitcode
  set objSRV = nothing
  ''ERROR RETURNED
  if (retSTOP <> 0) then
    ''SERVICE ALREADY STOPPED
    if (retSTOP = 2) then
      objOUT.write vbnewline & retSTOP & vbtab & "SERVICE ALREADY STOPPED : PattersonUpdateService : " & now
      retSTOP = 0
    ''ANY OTHER ERROR RETURNED
    elseif (retSTOP <> 2) then
      objOUT.write vbnewline & retSTOP & vbtab & "ERROR STOPPING : PattersonUpdateService : " & now
      retSTOP = 1
      call CLEANUP()
    end if
  end if
  objOUT.write vbnewline & vbnewline & "STOPPING EAGLESOFT SERVICES : " & now
  ''STOP PATTERSON APP SERVICE
  set objSRV = objWSH.exec("net stop " & chr(34) & "PattersonAppService" & chr(34))
  while (objSRV.status = 0)
    while (not objSRV.stdout.atendofstream)
      objOUT.write vbnewline & vbtab & objSRV.stdout.readline()
    wend
    wscript.sleep 10
  wend
  objOUT.write vbnewline & vbtab & objSRV.stdout.readall()
  retSTOP = objSRV.exitcode
  set objSRV = nothing
  ''ERROR RETURNED
  if (retSTOP <> 0) then
    ''SERVICE ALREADY STOPPED
    if (retSTOP = 2) then
      objOUT.write vbnewline & retSTOP & vbtab & "SERVICE ALREADY STOPPED : PattersonAppService : " & now
      retSTOP = 0
    ''ANY OTHER ERROR RETURNED
    elseif (retSTOP <> 2) then
      objOUT.write vbnewline & retSTOP & vbtab & "ERROR STOPPING : PattersonAppService : " & now
      retSTOP = 2
      call CLEANUP()
    end if
  end if
  call STOPDB()
end sub

''STOP EAGLESOFT DATABASE
sub STOPDB()
  objOUT.write vbnewline & vbnewline & "STOPPING EAGLESOFT DATABASE : " & now
  ''CALL PATTERSONSERVERSTATUS.EXE UTILITY WITH 'STOP' SWITCH
  set objDB = objWSH.exec(chr(34) & "C:\EagleSoft\Shared Files\PattersonServerStatus.exe" & chr(34) & " -stop")
  while (objDB.status = 0)
    while (not objDB.stdout.atendofstream)
      objOUT.write vbnewline & (objDB.stdout.readline())
    wend
    wscript.sleep 10
  wend
  objOUT.write vbnewline & objDB.stdout.readall()
  retSTOP = objDB.exitcode
  set objDB = nothing
  ''ERROR RETURNED
  if (retSTOP <> 0) then
    objOUT.write vbnewline & retSTOP & vbtab & "EAGLESOFT DATABASE : ERROR STOPPING: " & now
    retSTOP = 3
    call CLEANUP()
  end if
  objOUT.write vbnewline & vbtab & "EAGLESOFT DATABASE : STOPPED : " & now
  call DBCOPY()
end sub

sub DBCOPY()                                                              ''COPY EAGLESOFT DATA FOLDER
  objOUT.write vbnewline & vbnewline & "COPYING EAGLESOFT DATA : " & now
  ''USE ROBOCOPY TO COPY C:\EAGLESOFT\DATA FOLDER, OLVERWRITE ALL FILES IN DESTINATION
  set objCOPY = objWSH.exec("robocopy " & chr(34) & "C:\EagleSoft\Data" & chr(34) & " " & chr(34) & "B:\EaglesoftBackup" & chr(34) & " /MIR /z /w:1 /r:1 /mt /v")
  while (objCOPY.status = 0)
    while (not objCOPY.stdout.atendofstream)
      objOUT.write vbnewline & (objCOPY.stdout.readline())
    wend
    wscript.sleep 10
  wend
  objOUT.write vbnewline & objCOPY.stdout.readall()
  retSTOP = objCOPY.exitcode
  set objCOPY = nothing
  ''SUCCESSFULLY COPIED DATA
  if (retSTOP = 0) then
    objOUT.write vbnewline & "COPY EAGLESOFT DATA COMPLETE : " & now
  ''ERROR RETURNED
  elseif (retSTOP <> 0) then
    objOUT.write vbnewline & retSTOP & vbtab & "ERROR : XCOPY C:\EAGLESOFT\DATA B:\EAGLESOFTBACKUP : " & now
    retSTOP = 4
  end if
  ''END SCRIPT, RETURN EXIT CODE
  call CLEANUP()
end sub

''SCRIPT CLEANUP
sub CLEANUP()
  if (retSTOP = 0) then
    objOUT.write vbnewline & "PRE-BACKUP COMPLETE : " & retSTOP & " : " & now
  elseif (retSTOP <> 0) then
    objOUT.write vbnewline & "PRE-BACKUP FAILURE : " & retSTOP & " : " & now
    call err.raise(vbObjectError + retSTOP, "pre-backup", "fail")
  end if

  set objOUT = nothing
  set objFSO = nothing
  set objWSH = nothing
  wscript.quit retSTOP
end sub
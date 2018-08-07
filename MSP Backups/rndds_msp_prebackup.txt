on error resume next
''DECLARE VARIABLES
dim retSTOP
dim objWSH, objFSO, objOUT, objHOOK
''CREATE SCRIPTING SHELL & FILE SYSTEM OBJECTS
set objOUT = wscript.stdout
set objWSH = createobject("wscript.shell")
set objFSO = createobject("scripting.filesystemobject")
''DEFAULT FAIL
retSTOP = 5
''INITIATE STOP SERVICES
call STOPEAGLE

''SUB-ROUTINES
sub STOPEAGLE()                 ''STOP EAGLESOFT SERVICES
  objOUT.write vbnewline & vbnewline & "STOPPING EAGLESOFT SERVICES : " & NOW
  ''STOP PATTERSON APP SERVICE
  call HOOK("net stop " & chr(34) & "PattersonAppService" & chr(34))
  if (retSTOP <> 0) then        ''ERROR RETURNED
    if (retSTOP = 2) then       ''SERVICE ALREADY STOPPED
      objOUT.write vbnewline & retSTOP & vbtab & "SERVICE ALREADY STOPPED : PattersonAppService : " & NOW
      retSTOP = 0
    elseif (retSTOP <> 2) then  ''ANY OTHER ERROR RETURNED
      objOUT.write vbnewline & retSTOP & vbtab & "ERROR STOPPING : PattersonAppService : " & NOW
      retSTOP = 1
      ''END SCRIPT, RETURN EXIT CODE
      call CLEANUP
    end if
  end if
  ''STOP EAGLESOFT DATABASE
  call STOPDB
end sub

sub STOPDB()                    ''STOP EAGLESOFT DATABASE
  objOUT.write vbnewline & vbnewline & "STOPPING EAGLESOFT DATABASE : " & NOW
  ''CALL PATTERSONSERVERSTATUS.EXE UTILITY WITH 'STOP' SWITCH
  call HOOK(chr(34) & "C:\EagleSoft\Shared Files\PattersonServerStatus.exe" & chr(34) & " -stop")
  if (retSTOP <> 0) then        ''ERROR RETURNED
    objOUT.write vbnewline & retSTOP & vbtab & "EAGLESOFT DATABASE : ERROR STOPPING: " & NOW
    retSTOP = 2
    ''END SCRIPT, RETURN EXIT CODE
    call CLEANUP
  end if
  objOUT.write vbnewline & vbtab & "EAGLESOFT DATABASE : STOPPED : " & NOW
  ''COPY EAGLESOFT DATA
  call DBCOPY
end sub

sub DBCOPY()                    ''COPY EAGLESOFT DATA FOLDER
  objOUT.write vbnewline & vbnewline & "COPYING EAGLESOFT DATA : " & NOW
  ''USE XCOPY TO COPY C:\EAGLESOFT\DATA FOLDER, OLVERWRITE ALL FILES IN DESTINATION
  call HOOK("xcopy " & chr(34) & "C:\EagleSoft\Data" & chr(34) & " " & chr(34) & "E:\Backup" & chr(34) & " /E /F /H /I /K /R /Y")
  if (retSTOP = 0) then       ''SUCCESSFULLY COPIED DATA
    objOUT.write vbnewline & "COPY EAGLESOFT DATA COMPLETE : " & NOW
  elseif (retSTOP <> 0) then  ''ERROR RETURNED
    objOUT.write vbnewline & retSTOP & vbtab & "ERROR : XCOPY C:\EAGLESOFT\DATA E:\BACKUP : " & NOW
    retSTOP = 3
  end if
  ''END SCRIPT, RETURN EXIT CODE
  call CLEANUP
end sub

sub CLEANUP()                   ''SCRIPT CLEANUP
  if (retSTOP = 0) then         ''PRE-BACKUP COMPLETED SUCCESSFULLY
    objOUT.write vbnewline & "PRE-BACKUP COMPLETE : " & NOW
  elseif (retSTOP <> 0) then    ''PRE-BACKUP FAILED
    objOUT.write vbnewline & "PRE-BACKUP FAILURE : " & NOW
    ''RAISE CUSTOMIZED ERROR CODE, ERROR CODE WILL BE DEFINE RESTOP NUMBER INDICATING WHICH SECTION FAILED
    call err.raise(vbObjectError + retSTOP, "pre-backup", "fail")
  end if
  ''EMPTY OBJECTS
  set objOUT = nothing
  set objFSO = nothing
  set objWSH = nothing
  ''END SCRIPT, RETURN ERROR
  wscript.quit err.number
end sub

sub HOOK(strCMD)                ''CALL HOOK TO MONITOR OUTPUT OF CALLED COMMAND
  set objHOOK = objWSH.exec(strCMD)
  while (objHOOK.status = 0)
    while (not objHOOK.stdout.atendofstream)
      objOUT.write vbnewline & (objHOOK.stdout.readline())
    wend
    wscript.sleep 10
  wend
  objOUT.write vbnewline & objHOOK.stdout.readall()
  retSTOP = objHOOK.exitcode
  set objHOOK = nothing
end sub
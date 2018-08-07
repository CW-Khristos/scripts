on error resume next
''DECLARE VARIABLES
dim retSTART
dim objWSH, objFSO, objOUT, objHOOK
''CREATE SCRIPTING SHELL & FILE SYSTEM OBJECTS
set objOUT = wscript.stdout
set objWSH = createobject("wscript.shell")
set objFSO = createobject("scripting.filesystemobject")
''DEFAULT FAIL
retSTART = 5
''INITIATE START SERVICES
call STARTDB

''SUB-ROUTINES
sub STARTDB()                    ''START EAGLESOFT DATABASE
  objOUT.write vbnewline & "STARTING EAGLESOFT DATABASE : " & NOW
  ''CALL PATTERSONSERVERSTATUS.EXE WITH 'START' SWITCH, DO NOT MONITOR, PROCESS DOES NOT EXIT
  retSTART = objWSH.run(chr(34) & "C:\EagleSoft\Shared Files\PattersonServerStatus.exe" & chr(34) & " -start", 0, false)
  if (retSTART = 0) then         ''DATABASE SUCCESSFULLY STARTED
    objOUT.write vbnewline & vbnewline & retSTART & vbtab & "EAGLESOFT DATABASE STARTED : " & NOW
  elseif (retSTART <> 0) then    ''ERROR RETURNED
    objOUT.write vbnewline & vbnewline & retSTART & vbtab & "ERROR STARTING EAGLESOFT DATABASE : " & NOW
    retSTART = 1
    ''END SCRIPT, RETURN EXIT CODE
    call CLEANUP
  end if
  call STARTEAGLE
end sub

sub STARTEAGLE()                 ''START EAGLESOFT SERVICES
  objOUT.write vbnewline & vbnewline & "STARTING EAGLESOFT SERVICES : " & NOW
  ''START PATTERSON APP SERVICE
  call HOOK("net start " & chr(34) & "PattersonAppService" & chr(34))
  if (retSTART <> 0) then        ''ERROR RETURNED
    if (retSTART = 2) then       ''SERVICE ALREADY STARTED
      objOUT.write vbnewline & retSTOP & vbtab & "SERVICE ALREADY STARTED : PattersonAppService : " & NOW
      retSTART = 0
    elseif (retSTART <> 2) then  ''ANY OTHER ERROR
      objOUT.write vbnewline & retSTART & vbtab & "ERROR STARTING : PattersonAppService : " & NOW
      retSTART = 2
    end if
  end if
  ''END SCRIPT, RETURN EXIT CODE
  call CLEANUP
end sub

sub CLEANUP()                    ''SCRIPT CLEANUP
  if (retSTART = 0) then         ''POST-BACKUP COMPLETED SUCCESSFULLY
    objOUT.write vbnewline & "POST-BACKUP COMPLETE : " & NOW
  elseif (retSTART <> 0) then    ''POST-BACKUP FAILED
    objOUT.write vbnewline & "POST-BACKUP FAILURE : " & NOW
    ''RAISE CUSTOMIZED ERROR CODE, ERROR CODE WILL BE DEFINE RESTOP NUMBER INDICATING WHICH SECTION FAILED
    call err.raise(vbObjectError + retSTART, "post-backup", "fail")
  end if
  ''EMPTY OBJECTS
  set objOUT = nothing
  set objFSO = nothing
  set objWSH = nothing
  ''END SCRIPT, RETURN ERROR
  wscript.quit err.number
end sub

sub HOOK(strCMD)                 ''CALL HOOK TO MONITOR OUTPUT OF CALLED COMMAND
  set objHOOK = objWSH.exec(strCMD)
  while (objHOOK.status = 0)
    while (not objHOOK.stdout.atendofstream)
      objOUT.write vbnewline & (objHOOK.stdout.readline())
    wend
    wscript.sleep 10
  wend
  objOUT.write vbnewline & objHOOK.stdout.readall()
  retSTART = objHOOK.exitcode
  set objHOOK = nothing
end sub
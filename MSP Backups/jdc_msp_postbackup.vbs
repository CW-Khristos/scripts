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
''INITIATE SERVICE STARTS
call STARTPSQL

''SUB-ROUTINES
sub STARTPSQL()                  ''START PERVASIVE SQL SERVICE
  objOUT.write vbnewline & vbnewline & "STARTING PERVASIVE SQL SERVICE : " & NOW
  ''START PERVASIVE SQL SERVICE
  call HOOK("net start " & chr(34) & "psqlWGE" & chr(34))
  if (retSTART <> 0) then        ''ERROR RETURNED
    if (retSTART = 2) then       ''SERVICE ALREADY STARTED
      objOUT.write vbnewline & retSTOP & vbtab & "SERVICE ALREADY STARTED : psqlWGE : " & NOW
      retSTART = 0
    elseif (retSTART <> 2) then  ''ANY OTHER ERROR
      objOUT.write vbnewline & retSTART & vbtab & "ERROR STARTING : psqlWGE : " & NOW
      retSTART = 1
    end if
  end if
  ''START SAGE SERVICES
  call STARTSAGE
end sub

sub STARTSAGE()                  ''START SAGE SERVICES
  objOUT.write vbnewline & "STARTING SAGE SERVICES : " & NOW
  ''START SAGE 50 SMARTPOSTING SERVICE
  call HOOK("net start " & chr(34) & "Sage 50 SmartPosting 2017" & chr(34))
  if (retSTART <> 0) then        ''ERROR RETURNED
    if (retSTART = 2) then       ''SERVICE ALREADY STARTED
      objOUT.write vbnewline & retSTOP & vbtab & "SERVICE ALREADY STARTED : Sage SmartPosting 2017 : " & NOW
      retSTART = 0
    elseif (retSTART <> 2) then  ''ANY OTHER ERROR
      objOUT.write vbnewline & retSTART & vbtab & "ERROR STARTING : Sage 50 SmartPosting 2017 : " & NOW
      retSTART = 2
    end if
  end if
  ''START SAGE AUTOUPDATE MANAGER SERVICE
  call HOOK("net start " & chr(34) & "Sage AutoUpdate Manager Service" & chr(34))
  if (retSTART <> 0) then        ''ERROR RETURNED
    if (retSTART = 2) then       ''SERVICE ALREADY STARTED
      objOUT.write vbnewline & retSTOP & vbtab & "SERVICE ALREADY STARTED : Sage AutoUpdate Manager Service : " & NOW
      retSTART = 0
    elseif (retSTART <> 2) then  ''ANY OTHER ERROR
      objOUT.write vbnewline & retSTART & vbtab & "ERROR STARTING : Sage AutoUpdate Manager Service : " & NOW
      retSTART = 3
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
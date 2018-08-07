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
''INITIATE SERVICE STOPS
call STOPSAGE

''SUB-ROUTINES
sub STOPSAGE()                  ''STOP SAGE SERVICES
  objOUT.write vbnewline & vbnewline & "STOPPING SAGE SERVICES : " & NOW
  ''STOP SAGE AUTOUPDATE MANAGER SERVICE
  call HOOK("net stop " & chr(34) & "Sage AutoUpdate Manager Service" & chr(34))
  if (retSTOP <> 0) then        ''ERROR RETURNED
    if (retSTOP = 2) then       ''SERVICE ALREADY STOPPED
      objOUT.write vbnewline & retSTOP & vbtab & "SERVICE ALREADY STOPPED : Sage AutoUpdate Manager Service : " & NOW
      retSTOP = 0
    elseif (retSTOP <> 2) then  ''ANY OTHER ERROR RETURNED
      objOUT.write vbnewline & retSTOP & vbtab & "ERROR STOPPING : Sage AutoUpdate Manager Service : " & NOW
      retSTOP = 1
      ''END SCRIPT, RETURN EXIT CODE
      call CLEANUP
    end if
  end if
  ''STOP SAGE 50 SMARTPOSTING SERVICE
  call HOOK("net stop " & chr(34) & "Sage 50 SmartPosting 2017" & chr(34))
  if (retSTOP <> 0) then        ''ERROR RETURNED
    if (retSTOP = 2) then       ''SERVICE ALREADY STOPPED
      objOUT.write vbnewline & retSTOP & vbtab & "SERVICE ALREADY STOPPED : Sage 50 SmartPosting 2017 : " & NOW
      retSTOP = 0
    elseif (retSTOP <> 2) then  ''ANY OTHER ERROR
      objOUT.write vbnewline & retSTOP & vbtab & "ERROR STOPPING : Sage 50 SmartPosting 2017 : " & NOW
      retSTOP = 2
      ''END SCRIPT, RETURN EXIT CODE
      call CLEANUP
    end if
  end if
  ''STOP PERVASIVE SQL SERVICE
  call STOPPSQL
end sub

sub STOPPSQL()                  ''STOP PERVASIVE SQL SERVICE
  objOUT.write vbnewline & "STOPPING PERVASIVE SQL SERVICE : " & NOW
  ''STOP PERVASIVE SQL SERVICE
  call HOOK("net stop " & chr(34) & "psqlWGE" & chr(34))
  if (retSTOP <> 0) then        ''ERROR RETURNED
    if (retSTOP = 2) then       ''SERVICE ALREADY STOPPED
      objOUT.write vbnewline & retSTOP & vbtab & "SERVICE ALREADY STOPPED : psqlWGE : " & NOW
      retSTOP = 0
    elseif (retSTOP <> 2) then  ''ANY OTHER ERROR
      objOUT.write vbnewline & retSTOP & vbtab & "ERROR STOPPING : psqlWGE : " & NOW
      retSTOP = 3
      ''END SCRIPT, RETURN EXIT CODE
      call CLEANUP
    end if
  end if
  ''COPY SAGE DATA
  call SAGECOPY
end sub

sub SAGECOPY()                  ''COPY SAGE FOLDER
  objOUT.write vbnewline & "COPYING SAGE DATA : " & NOW
  ''USE XCOPY TO COPY D:\SAGE FOLDER, OLVERWRITE ALL FILES IN DESTINATION
  call HOOK("xcopy " & chr(34) & "D:\Sage" & chr(34) & " " & chr(34) & "D:\CW MSP Sage\Sage" & chr(34) & " /E /F /H /I /K /R /Y")
  if (retSTOP = 0) then         ''SUCCESSFULLY COPIED DATA
    objOUT.write vbnewline & "COPY SAGE DATA COMPLETE : " & NOW
  elseif (retSTOP <> 0) then    ''ERROR RETURNED
    objOUT.write vbnewline & retSTOP & vbtab & "ERROR : XCOPY D:\SAGE D:\CW MSP SAGE\SAGE : " & NOW
    retSTOP = 4
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
on error resume next
''SCRIPT VARIABLES
DIM objSIN, objSOUT, strORG, strREP, strSID
dim objIN, objOUT, objARG, objWSH, objFSO, objLOG, objEXEC, objHOOK
dim strIN, strOUT, strLSV, strUSR, strPWD, colUSR(), colSID(), retSTOP
''DEFAULT SUCCESS
retSTOP = 0
''STDIN / STDOUT
set objIN = wscript.stdin
set objOUT = wscript.stdout
set objARG = wscript.arguments
''CREATE SCRIPTING OBJECTS
set objWSH = createobject("wscript.shell")
set objFSO = createobject("scripting.filesystemobject")
''PREPARE LOGFILE
if (objFSO.fileexists("C:\temp\LSVperm")) then              ''LOGFILE EXISTS
  objFSO.deletefile "C:\temp\LSVperm", true
  set objLOG = objFSO.createtextfile("C:\temp\LSVperm")
  objLOG.close
  set objLOG = objFSO.opentextfile("C:\temp\LSVperm", 8)
else                                                        ''LOGFILE NEEDS TO BE CREATED
  set objLOG = objFSO.createtextfile("C:\temp\LSVperm")
  objLOG.close
  set objLOG = objFSO.opentextfile("C:\temp\LSVperm", 8)
end if
''READ PASSED COMMANDLINE ARGUMENTS
if (wscript.arguments.count > 0) then                       ''ARGUMENTS WERE PASSED
  for x = 0 to (wscript.arguments.count - 1)
    objOUT.write vbnewline & now & vbtab & " - ARGUMENT " & (x + 1) & " (ITEM " & x & ") " & " PASSED : " & ucase(objARG.item(x))
    objLOG.write vbnewline & now & vbtab & " - ARGUMENT " & (x + 1) & " (ITEM " & x & ") " & " PASSED : " & ucase(objARG.item(x))
  next 
  ''ARGUMENT 0 - FOLDER PATH
  strLSV = objARG.item(0)
  if (wscript.arguments.count > 1) then                     ''SET RMMTECH LOGON ARGUMENTS FOR UPDATING 'BACKUP SERVICE CONTROLLER' LOGON
    strUSR = objARG.item(1)
    strPWD = objARG.item(2)
    ''PASSED USER ACCOUNT IS A LOCAL ACCOUNT
    if (instr(1, strUSR, "\") = 0) then
      strUSR = ".\" & strUSR
    end if
  else                                                      ''NOT ENOUGH ARGUMENTS PASSED, END SCRIPT
    retSTOP = 1
    call CLEANUP
  end if
else                                                        ''NO ARGUMENTS PASSED, END SCRIPT
  objOUT.write vbnewline & vbnewline & now & vbtab & " - SCRIPT REQUIRES PATH TO MSP LSV DESTINATION"
  objLOG.write vbnewline & vbnewline & now & vbtab & " - SCRIPT REQUIRES PATH TO MSP LSV DESTINATION"
  retSTOP = 1
  call CLEANUP
end if
objOUT.write vbnewline & vbnewline & now & vbtab & " - EXECUTING LSVPERM"
objLOG.write vbnewline & vbnewline & now & vbtab & " - EXECUTING LSVPERM"
''GET SIDS OF ALL USERS
intUSR = 0
intSID = 0
objOUT.write vbnewline & vbnewline & now & vbtab & vbtab & " - ENUMERATING USERNAMES AND SIDS"
objLOG.write vbnewline & vbnewline & now & vbtab & vbtab & " - ENUMERATING USERNAMES AND SIDS"
set objEXEC = objWSH.exec("wmic useraccount get name,sid /format:csv")
while (not objEXEC.stdout.atendofstream)
  strIN = objEXEC.stdout.readline
  'objOUT.write vbnewline & now & vbtab & vbtab & strIN
  'objLOG.write vbnewline & now & vbtab & vbtab & strIN
  if ((trim(strIN) <> vbnullstring) and (instr(1, strIN, ","))) then
    if ((trim(split(strIN, ",")(1)) <> vbnullstring) and (trim(split(strIN, ",")(1)) <> "Name")) then
      redim preserve colUSR(intUSR + 1)
      redim preserve colSID(intSID + 1)
      colUSR(intUSR) = trim(split(strIN, ",")(1))
      colSID(intSID) = trim(split(strIN, ",")(2))
      intUSR = (intUSR + 1)
      intSID = (intSID + 1)
    end if
  end if
  if (err.number <> 0) then
    retSTOP = 2
    objOUT.write vbnewline & now & vbtab & vbtab & vbtab & err.number & vbtab & err.description
    objLOG.write vbnewline & now & vbtab & vbtab & vbtab & err.number & vbtab & err.description
  end if
wend
err.clear
''VALIDATE COLLECTED USERNAMES AND SIDS
objOUT.write vbnewline & vbnewline & now & vbtab & vbtab & " - COLLECTED USERNAMES AND SIDS"
objLOG.write vbnewline & vbnewline & now & vbtab & vbtab & " - COLLECTED USERNAMES AND SIDS"
for intUSR = 0 to ubound(colUSR)
  intSID = intUSR
  if (instr(1, lcase(colUSR(intUSR)), "rmmtech")) then
    strSID = colSID(intUSR)
  end if
  objOUT.write vbnewline & now & vbtab & vbtab & vbtab & colUSR(intUSR) & " : " & colSID(intSID)
  objLOG.write vbnewline & now & vbtab & vbtab & vbtab & colUSR(intUSR) & " : " & colSID(intSID)
next
''STOP 'BACKUP SERVICE CONTROLLER' AND UPDATE ACCOUNT LOGON TO RMMTECH
objOUT.write vbnewline & vbnewline & now & vbtab & vbtab & " - UPDATING BACKUP SERVICE CONTROLLER"
objLOG.write vbnewline & vbnewline & now & vbtab & vbtab & " - UPDATING BACKUP SERVICE CONTROLLER"
call HOOK("sc.exe stop " & chr(34) & "Backup Service Controller" & chr(34))
call HOOK("sc.exe config " & chr(34) & "Backup Service Controller" & chr(34) & " obj= " & chr(34) & strUSR & chr(34) & " password= " & chr(34) & strPWD & chr(34) & " TYPE= own")
''GRANT 'LOGON AS A SERVICE' TO RMMTECH USER
objOUT.write vbnewline & now & vbtab & vbtab & vbtab & " - GRANT LONGON AS SERVICE : " & strUSR
objLOG.write vbnewline & now & vbtab & vbtab & vbtab & " - GRANT LONGON AS SERVICE : " & strUSR
strORG = "SeServiceLogonRight ="
strREP = "SeServiceLogonRight = " & "*" & strSID & ","
''EXPORT CURRENT SECURITY DATABASE CONFIGS
call HOOK("secedit /export /cfg c:\temp\config.inf")
''READ CURRENT EXPORTED SECURITY DATABASE CONFIGS
set objSIN = objFSO.opentextfile("c:\temp\config.inf", 1, 1, -1)
strIN = objSIN.readall
objSIN.close
set objSIN = nothing
''WRITE SECURITY DATABASE CONFIGS WITH 'SeServiceLogonRight' FOR RMMTECH
set objSOUT = objFSO.opentextfile("c:\temp\config.inf", 2, 1, -1)
objSOUT.write (replace(strIN,strORG,strREP))
objSOUT.close
set objSOUT = nothing
''APPLY NEW SECURITY DATABASE CONFIGS
call HOOK("secedit /import /db secedit.sdb /cfg c:\temp\config.inf")
call HOOK("secedit /configure /db secedit.sdb")
call HOOK("gpupdate")
''REMOVE TEMP FILES
'objFSO.deletefile("c:\temp\config.inf") 
objOUT.write vbnewline & now & vbtab & vbtab & vbtab & " - LOGON AS SERVICE GRANTED : " & strUSR
objLOG.write vbnewline & now & vbtab & vbtab & vbtab & " - LOGON AS SERVICE GRANTED : " & strUSR
objOUT.write vbnewline & now & vbtab & vbtab & " - BACKUP SERVICE CONTROLLER UPDATED"
objLOG.write vbnewline & now & vbtab & vbtab & " - BACKUP SERVICE CONTROLLER UPDATED"
''TAKEOWN USING CURRENT USERS, THIS SHOULD BE RMMTECH
objOUT.write vbnewline & vbnewline & now & vbtab & vbtab & " - ASSIGNING OWNERSHIP"
objLOG.write vbnewline & vbnewline & now & vbtab & vbtab & " - ASSIGNING OWNERSHIP"
call HOOK("takeown /F " & chr(34) & strLSV & chr(34) & " /R /D Y")
''ADD RMMTECH USER EXPLICIT FULL CONTROL
objOUT.write vbnewline & now & vbtab & vbtab & " - ASSIGNING RMMTECH FULL CONTROL"
objLOG.write vbnewline & now & vbtab & vbtab & " - ASSIGNING RMMTECH FULL CONTROL"
for intUSR = 0 to ubound(colUSR)
  intSID = intUSR
  if (instr(1, lcase(colUSR(intUSR)), "rmmtech")) then
    call HOOK("icacls " & chr(34) & strLSV & chr(34) & " /grant " & colUSR(intUSR) & ":(OI)(CI)F /T /C /Q")
    call HOOK("icacls " & chr(34) & strLSV & chr(34) & " /grant *" & colSID(intSID) & ":(OI)(CI)F /T /C /Q")
  end if
next
''DISABLE INHERITANCE ON LSV DESTINATION, AND ONLY THE ROOT MSP BACKUP LSV DESTINATION
''THIS MUST BE DONE LAST, BEFORE REMOVING ALL OTHER USER PERMISSIONS SO RMMTECH PERMISSIONS FULLY PROPAGATE TO ALL FILES / FOLDERS
objOUT.write vbnewline & vbnewline & now & vbtab & vbtab & " - REMOVING INHERITANCE ON LSV DESTINATION"
objLOG.write vbnewline & vbnewline & now & vbtab & vbtab & " - REMOVING INHERITANCE ON LSV DESTINATION"
call HOOK("icacls " & chr(34) & strLSV & chr(34) & " /inheritance:r /C")
''REMOVE USER PERMISSIONS, EXCLUDE RMMTECH FROM REMOVAL
objOUT.write vbnewline & vbnewline & now & vbtab & vbtab & " - REMOVING ALL OTHER ENUMERATED USERS' PERMISSIONS"
objLOG.write vbnewline & vbnewline & now & vbtab & vbtab & " - REMOVING ALL OTHER ENUMERATED USERS' PERMISSIONS"
for intUSR = 0 to ubound(colUSR)
  intSID = intUSR
  if ((colUSR(intUSR) <> vbnullstring) and (instr(1, lcase(colUSR(intUSR)), "rmmtech")=0)) then
    call HOOK("icacls " & chr(34) & strLSV & chr(34) & " /remove:g " & colUSR(intUSR) & " /T /C /Q")
    call HOOK("icacls " & chr(34) & strLSV & chr(34) & " /remove:g *" & colSID(intSID) & " /T /C /Q")
  end if
next
''RESTART 'BACKUP SERVICE CONTROLLER'
objOUT.write vbnewline & vbnewline & now & vbtab & vbtab & " - RESTARTING BACKUP SERVICE CONTROLLER"
objLOG.write vbnewline & vbnewline & now & vbtab & vbtab & " - RESTARTING BACKUP SERVICE CONTROLLER"
call HOOK("sc.exe start " & chr(34) & "Backup Service Controller" & chr(34))
''END SCRIPT
call CLEANUP

''SUB-ROUTINES
sub HOOK(strCMD)                                        ''CALL HOOK TO MONITOR OUTPUT OF CALLED COMMAND
  on error resume next
  'comspec = objWSH.ExpandEnvironmentStrings("%comspec%")
  objOUT.write vbnewline & now & vbtab & vbtab & "EXECUTING : " & strCMD
  objLOG.write vbnewline & now & vbtab & vbtab & "EXECUTING : " & strCMD
  set objHOOK = objWSH.exec(strCMD)
  if (instr(1, strCMD, "takeown /F ") = 0) then         ''SUPPRESS 'TAKEOWN' SUCCESS MESSAGES
    'while (objHOOK.status = 0)
    while (not objHOOK.stdout.atendofstream)
      strIN = objHOOK.stdout.readline
      if (strIN <> vbnullstring) then
        objOUT.write vbnewline & now & vbtab & vbtab & vbtab & strIN 
        objLOG.write vbnewline & now & vbtab & vbtab & vbtab & strIN 
      end if
    wend
    wscript.sleep 10
    'wend
    strIN = objHOOK.stdout.readall
    if (strIN <> vbnullstring) then
      objOUT.write vbnewline & now & vbtab & vbtab & vbtab & strIN 
      objLOG.write vbnewline & now & vbtab & vbtab & vbtab & strIN 
    end if
  end if
  'retSTOP = objHOOK.exitcode
  set objHOOK = nothing
  if (err.number <> 0) then
    retSTOP = 3
    objOUT.write vbnewline & now & vbtab & vbtab & vbtab & err.number & vbtab & err.description
    objLOG.write vbnewline & now & vbtab & vbtab & vbtab & err.number & vbtab & err.description
    err.clear
  end if
end sub

sub CLEANUP()                                           ''SCRIPT CLEANUP
  if (retSTOP = 0) then                                 ''SCRIPT COMPLETED SUCCESSFULLY
    objOUT.write vbnewline & vbnewline & now & vbtab & " - LSVPERM COMPLETE : " & now
    objLOG.write vbnewline & vbnewline & now & vbtab & " - LSVPERM COMPLETE : " & now
    err.clear
  elseif (retSTOP <> 0) then                            ''SCRIPT FAILED
    objOUT.write vbnewline & vbnewline & now & vbtab & " - LSVPERM FAILURE : " & now
    objLOG.write vbnewline & vbnewline & now & vbtab & " - LSVPERM FAILURE : " & now
    ''RAISE CUSTOMIZED ERROR CODE, ERROR CODE WILL BE DEFINED RESTOP NUMBER INDICATING WHICH SECTION FAILED
    call err.raise(vbObjectError + retSTOP, "LSVperm", "fail")
  end if
  ''EMPTY OBJECTS
  set objEXEC = nothing
  set objLOG = nothing
  set objFSO = nothing
  set objWSH = nothing
  set objARG = nothing
  set objOUT = nothing
  set objIN = nothing
  ''END SCRIPT, RETURN ERROR
  wscript.quit err.number
end sub
''SVCPERM.VBS
''DESIGNED TO GRANTING SERVICE LOGON PERMISSIONS
''REQUIRES 1 PARAMETER; 'STRUSR'
''WRITTEN BY : CJ BLEDSOE / CJ<@>THECOMPUTERWARRIORS.COM
on error resume next
''SCRIPT VARIABLES
dim errRET, strVER, strSEL
dim strIN, strOUT, colUSR(), colSID()
dim objSIN, objSOUT, strORG, strREP, strSID
''VARIABLES ACCEPTING PARAMETERS - TARGET USERNAME
dim strUSR
''SCRIPT OBJECTS
dim objLOG, objEXEC, objHOOK
dim objIN, objOUT, objARG, objWSH, objFSO
''VERSION FOR SCRIPT UPDATE, SVCPERM.VBS, REF #2 , FIXES #21
strVER = 2
''DEFAULT SUCCESS
errRET = 0
''STDIN / STDOUT
set objIN = wscript.stdin
set objOUT = wscript.stdout
set objARG = wscript.arguments
''OBJECTS FOR LOCATING FOLDERS
set objWSH = createobject("wscript.shell")
set objFSO = createobject("scripting.filesystemobject")
''PREPARE LOGFILE
if (objFSO.fileexists("C:\temp\SVCperm")) then              ''LOGFILE EXISTS
  objFSO.deletefile "C:\temp\SVCperm", true
  set objLOG = objFSO.createtextfile("C:\temp\SVCperm")
  objLOG.close
  set objLOG = objFSO.opentextfile("C:\temp\SVCperm", 8)
else                                                        ''LOGFILE NEEDS TO BE CREATED
  set objLOG = objFSO.createtextfile("C:\temp\SVCperm")
  objLOG.close
  set objLOG = objFSO.opentextfile("C:\temp\SVCperm", 8)
end if
''READ PASSED COMMANDLINE ARGUMENTS
if (wscript.arguments.count > 0) then                       ''ARGUMENTS WERE PASSED
  for x = 0 to (wscript.arguments.count - 1)
    objOUT.write vbnewline & now & vbtab & " - ARGUMENT " & (x + 1) & " (ITEM " & x & ") " & " PASSED : " & ucase(objARG.item(x))
    objLOG.write vbnewline & now & vbtab & " - ARGUMENT " & (x + 1) & " (ITEM " & x & ") " & " PASSED : " & ucase(objARG.item(x))
  next 
  if (wscript.arguments.count > 0) then                     ''SET RMMTECH LOGON ARGUMENTS FOR UPDATING 'BACKUP SERVICE CONTROLLER' LOGON
    strUSR = objARG.item(0)
    ''PASSED USER ACCOUNT IS A LOCAL ACCOUNT
    'if (instr(1, strUSR, "\") = 0) then
    '  strUSR = ".\" & strUSR
    'end if
  else                                                      ''NOT ENOUGH ARGUMENTS PASSED, END SCRIPT
    errRET = 1
    call CLEANUP
  end if
else                                                        ''NO ARGUMENTS PASSED, END SCRIPT
  objOUT.write vbnewline & vbnewline & now & vbtab & " - SCRIPT REQUIRES USER TO GRANT SERVICE LOGON"
  objLOG.write vbnewline & vbnewline & now & vbtab & " - SCRIPT REQUIRES USER TO GRANT SERVICE LOGON"
  errRET = 1
  call CLEANUP
end if
objOUT.write vbnewline & vbnewline & now & vbtab & " - EXECUTING SVCPERM"
objLOG.write vbnewline & vbnewline & now & vbtab & " - EXECUTING SVCPERM"
''AUTOMATIC UPDATE, SVCPERM.VBS, REF #2
call CHKAU()
''PRE-MATURE END SCRIPT, TESTING AUTOMATIC UPDATE SVCPERM.VBS, REF #2 , FIXES #21
'call CLEANUP()
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
    errRET = 2
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
  if (instr(1, lcase(colUSR(intUSR)), lcase(strUSR))) then
    strSID = colSID(intUSR)
  end if
  objOUT.write vbnewline & now & vbtab & vbtab & vbtab & colUSR(intUSR) & " : " & colSID(intSID)
  objLOG.write vbnewline & now & vbtab & vbtab & vbtab & colUSR(intUSR) & " : " & colSID(intSID)
next
''GRANT 'LOGON AS A SERVICE' TO RMMTECH USER
objOUT.write vbnewline & now & vbtab & vbtab & " - GRANT LONGON AS SERVICE : " & strUSR & " : " & strSID
objLOG.write vbnewline & now & vbtab & vbtab & " - GRANT LONGON AS SERVICE : " & strUSR & " : " & strSID
strORG = "SeServiceLogonRight = "
strREP = "SeServiceLogonRight = " & "*" & strSID & ","
''EXPORT CURRENT SECURITY DATABASE CONFIGS
call HOOK("secedit /export /cfg c:\temp\config.inf")
''READ CURRENT EXPORTED SECURITY DATABASE CONFIGS
set objSIN = objFSO.opentextfile("c:\temp\config.inf", 1, 1, -1)
strIN = objSIN.readall
objSIN.close
set objSIN = nothing
''WRITE SECURITY DATABASE CONFIGS WITH 'SetServiceLogonRight' FOR RMMTECH
set objSOUT = objFSO.opentextfile("c:\temp\config.inf", 2, 1, -1)
objSOUT.write (replace(strIN,strORG,strREP))
objSOUT.close
set objSOUT = nothing
wscript.sleep 1000
''APPLY NEW SECURITY DATABASE CONFIGS
call HOOK("secedit /import /db secedit.sdb /cfg c:\temp\config.inf")
call HOOK("secedit /configure /db secedit.sdb")
call HOOK("gpupdate /force")
''REMOVE TEMP FILES
'objFSO.deletefile("c:\temp\config.inf") 
objOUT.write vbnewline & now & vbtab & vbtab & " - LOGON AS SERVICE GRANTED : " & strUSR
objLOG.write vbnewline & now & vbtab & vbtab & " - LOGON AS SERVICE GRANTED : " & strUSR
''STOP 'BACKUP SERVICE CONTROLLER' AND UPDATE ACCOUNT LOGON TO RMMTECH
'objOUT.write vbnewline & vbnewline & now & vbtab & vbtab & " - UPDATING BACKUP SERVICE CONTROLLER"
'objLOG.write vbnewline & vbnewline & now & vbtab & vbtab & " - UPDATING BACKUP SERVICE CONTROLLER"
'call HOOK("sc.exe stop " & chr(34) & "Backup Service Controller" & chr(34))
'call HOOK("sc.exe config " & chr(34) & "Backup Service Controller" & chr(34) & " obj= " & chr(34) & strUSR & chr(34) & " password= " & chr(34) & strPWD & chr(34) & " TYPE= own")
'objOUT.write vbnewline & now & vbtab & vbtab & " - BACKUP SERVICE CONTROLLER UPDATED"
'objLOG.write vbnewline & now & vbtab & vbtab & " - BACKUP SERVICE CONTROLLER UPDATED"
''END SCRIPT
call CLEANUP()

''SUB-ROUTINES
sub CHKAU()																					''CHECK FOR SCRIPT UPDATE, SVCPERM.VBS, REF #2 , FIXES #21
  ''REMOVE WINDOWS AGENT CACHED VERSION OF SCRIPT
  if (objFSO.fileexists("C:\Program Files (x86)\N-Able Technologies\Windows Agent\cache\" & wscript.scriptname)) then
    objFSO.deletefile "C:\Program Files (x86)\N-Able Technologies\Windows Agent\cache\" & wscript.scriptname, true
  end if
	''ADD WINHTTP SECURE CHANNEL TLS REGISTRY KEYS
	call HOOK("reg add " & chr(34) & "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\WinHttp" & chr(34) & _
		" /f /v DefaultSecureProtocols /t REG_DWORD /d 0x00000A00 /reg:32")
	call HOOK("reg add " & chr(34) & "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\WinHttp" & chr(34) & _
		" /f /v DefaultSecureProtocols /t REG_DWORD /d 0x00000A00 /reg:64")
	''SCRIPT OBJECT FOR PARSING XML
	set objXML = createobject("Microsoft.XMLDOM")
	''FORCE SYNCHRONOUS
	objXML.async = false
	''LOAD SCRIPT VERSIONS DATABASE XML
	if objXML.load("https://github.com/CW-Khristos/scripts/raw/master/version.xml") then
		set colVER = objXML.documentelement
		for each objSCR in colVER.ChildNodes
			''LOCATE CURRENTLY RUNNING SCRIPT
			if (lcase(objSCR.nodename) = lcase(wscript.scriptname)) then
				''CHECK LATEST VERSION
				if (cint(objSCR.text) > cint(strVER)) then
					objOUT.write vbnewline & now & " - UPDATING " & objSCR.nodename & " : " & objSCR.text & vbnewline
					objLOG.write vbnewline & now & " - UPDATING " & objSCR.nodename & " : " & objSCR.text & vbnewline
					''DOWNLOAD LATEST VERSION OF SCRIPT
					call FILEDL("https://github.com/CW-Khristos/scripts/raw/master/SVCperm.vbs", wscript.scriptname)
					''RUN LATEST VERSION
					if (wscript.arguments.count > 0) then             ''ARGUMENTS WERE PASSED
						for x = 0 to (wscript.arguments.count - 1)
							strTMP = strTMP & " " & objARG.item(x)
						next
            objOUT.write vbnewline & now & vbtab & " - RE-EXECUTING  " & objSCR.nodename & " : " & objSCR.text & vbnewline
            objLOG.write vbnewline & now & vbtab & " - RE-EXECUTING  " & objSCR.nodename & " : " & objSCR.text & vbnewline
						objWSH.run "cscript.exe //nologo " & chr(34) & "c:\temp\" & wscript.scriptname & chr(34) & strTMP, 0, false
					elseif (wscript.arguments.count = 0) then         ''NO ARGUMENTS WERE PASSED
            objOUT.write vbnewline & now & vbtab & " - RE-EXECUTING  " & objSCR.nodename & " : " & objSCR.text & vbnewline
            objLOG.write vbnewline & now & vbtab & " - RE-EXECUTING  " & objSCR.nodename & " : " & objSCR.text & vbnewline
						objWSH.run "cscript.exe //nologo " & chr(34) & "c:\temp\" & wscript.scriptname & chr(34), 0, false
					end if
					''END SCRIPT
					call CLEANUP()
				end if
			end if
		next
	end if
	set colVER = nothing
	set objXML = nothing
end sub

sub HOOK(strCMD)                                        ''CALL HOOK TO MONITOR OUTPUT OF CALLED COMMAND
  on error resume next
  'comspec = objWSH.ExpandEnvironmentStrings("%comspec%")
  objOUT.write vbnewline & now & vbtab & vbtab & "EXECUTING : " & strCMD
  objLOG.write vbnewline & now & vbtab & vbtab & "EXECUTING : " & strCMD
  set objHOOK = objWSH.exec(strCMD)
  if (instr(1, strCMD, "takeown /F ") = 0) then         ''SUPPRESS 'TAKEOWN' SUCCESS MESSAGES
    while (not objHOOK.stdout.atendofstream)
      strIN = objHOOK.stdout.readline
      if (strIN <> vbnullstring) then
        objOUT.write vbnewline & now & vbtab & vbtab & vbtab & strIN 
        objLOG.write vbnewline & now & vbtab & vbtab & vbtab & strIN 
      end if
    wend
    wscript.sleep 10
    strIN = objHOOK.stdout.readall
    if (strIN <> vbnullstring) then
      objOUT.write vbnewline & now & vbtab & vbtab & vbtab & strIN 
      objLOG.write vbnewline & now & vbtab & vbtab & vbtab & strIN 
    end if
  end if
  set objHOOK = nothing
  if (err.number <> 0) then
    errRET = 3
    objOUT.write vbnewline & now & vbtab & vbtab & vbtab & err.number & vbtab & err.description
    objLOG.write vbnewline & now & vbtab & vbtab & vbtab & err.number & vbtab & err.description
    err.clear
  end if
end sub

sub CLEANUP()                                           ''SCRIPT CLEANUP
  if (errRET = 0) then                                 ''SCRIPT COMPLETED SUCCESSFULLY
    objOUT.write vbnewline & vbnewline & now & vbtab & " - SVCPERM COMPLETE : " & now
    objLOG.write vbnewline & vbnewline & now & vbtab & " - SVCPERM COMPLETE : " & now
    err.clear
  elseif (errRET <> 0) then                            ''SCRIPT FAILED
    objOUT.write vbnewline & vbnewline & now & vbtab & " - SVCPERM FAILURE : " & errRET & " : " & now
    objLOG.write vbnewline & vbnewline & now & vbtab & " - SVCPERM FAILURE : " & errRET & " : " & now
    ''RAISE CUSTOMIZED ERROR CODE, ERROR CODE WILL BE DEFINED RESTOP NUMBER INDICATING WHICH SECTION FAILED
    call err.raise(vbObjectError + errRET, "SVCPERM", "fail")
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
''JDC_MSP_POSTBACKUP.VBS
''DESIGNED TO RESTART SAGE DATABASE AND SERVICES
''WRITTEN BY : CJ BLEDSOE / CJ<@>THECOMPUTERWARRIORS.COM
on error resume next
''SCRIPT VARIABLES
dim errRET, strVER, strIN
''SCRIPT OBJECTS
dim objIN, objOUT, objARG, objWSH, objFSO
dim objLOG, objHOOK, objEXEC, objHTTP, objXML
''CREATE SCRIPTING SHELL & FILE SYSTEM OBJECTS
set objOUT = wscript.stdout
set objWSH = createobject("wscript.shell")
set objFSO = createobject("scripting.filesystemobject")
''VERSION FOR SCRIPT UPDATE , JDC_MSP_POSTBACKUP.VBS , REF #2 , REF #50
strVER = 3
''DEFAULT FAIL
errRET = 5
''PREPARE LOGFILE
if (objFSO.fileexists("C:\temp\MSP_POSTBACKUP")) then       ''LOGFILE EXISTS
  objFSO.deletefile "C:\temp\MSP_POSTBACKUP", true
  set objLOG = objFSO.createtextfile("C:\temp\MSP_POSTBACKUP")
  objLOG.close
  set objLOG = objFSO.opentextfile("C:\temp\MSP_POSTBACKUP", 8)
else                                                        ''LOGFILE NEEDS TO BE CREATED
  set objLOG = objFSO.createtextfile("C:\temp\MSP_POSTBACKUP")
  objLOG.close
  set objLOG = objFSO.opentextfile("C:\temp\MSP_POSTBACKUP", 8)
end if
''READ PASSED COMMANDLINE ARGUMENTS
if (wscript.arguments.count > 0) then                       ''ARGUMENTS WERE PASSED
  for x = 0 to (wscript.arguments.count - 1)
    objOUT.write vbnewline & now & vbtab & " - ARGUMENT " & (x + 1) & " (ITEM " & x & ") " & " PASSED : " & ucase(objARG.item(x))
    objLOG.write vbnewline & now & vbtab & " - ARGUMENT " & (x + 1) & " (ITEM " & x & ") " & " PASSED : " & ucase(objARG.item(x))
  next 
  if (wscript.arguments.count > 0) then                     ''SET PASSED ARG7UMENTS
  else                                                      ''NOT ENOUGH ARGUMENTS PASSED, END SCRIPT
    'errRET = 1
    'call CLEANUP()
  end if
else                                                        ''NO ARGUMENTS PASSED, END SCRIPT
  'errRET = 1
  'call CLEANUP()
end if

''------------
''BEGIN SCRIPT
objOUT.write vbnewline & vbnewline & now & vbtab & " - EXECUTING JDC_MSP_POSTBACKUP" & vbnewline
objLOG.write vbnewline & vbnewline & now & vbtab & " - EXECUTING JDC_MSP_POSTBACKUP" & vbnewline
''AUTOMATIC UPDATE , 'ERRRET'=10 , JDC_MSP_POSTBACKUP.VBS , REF #2 , REF #50
call CHKAU()
''INITIATE SERVICE STARTS
call STARTPSQL()
''END SCRIPT, RETURN EXIT CODE
call CLEANUP()
''END SCRIPT
''------------

''SUB-ROUTINES
sub STARTPSQL()                                             ''START PERVASIVE SQL SERVICE
  objOUT.write vbnewline & vbnewline & "STARTING PERVASIVE SQL SERVICE : " & now
  objLOG.write vbnewline & vbnewline & "STARTING PERVASIVE SQL SERVICE : " & now
  ''START PERVASIVE SQL SERVICE
  call HOOK("net start " & chr(34) & "psqlWGE" & chr(34))
  if (errRET <> 0) then                                     ''ERROR RETURNED
    if (errRET = 2) then                                    ''SERVICE ALREADY STARTED
      objOUT.write vbnewline & retSTOP & vbtab & "SERVICE ALREADY STARTED : psqlWGE : " & now
      objLOG.write vbnewline & retSTOP & vbtab & "SERVICE ALREADY STARTED : psqlWGE : " & now
      errRET = 0
      err.clear
    elseif (errRET <> 2) then                               ''ANY OTHER ERROR
      objOUT.write vbnewline & errRET & vbtab & "ERROR STARTING : psqlWGE : " & now
      objLOG.write vbnewline & errRET & vbtab & "ERROR STARTING : psqlWGE : " & now
      call LOGERR(4)
    end if
  end if
  ''START SAGE SERVICES
  call STARTSAGE()
end sub

sub STARTSAGE()                                             ''START SAGE SERVICES
  objOUT.write vbnewline & "STARTING SAGE SERVICES : " & now
  objLOG.write vbnewline & "STARTING SAGE SERVICES : " & now
  ''START SAGE 50 SMARTPOSTING SERVICE
  call HOOK("net start " & chr(34) & "Sage 50 SmartPosting 2017" & chr(34))
  if (errRET <> 0) then                                     ''ERROR RETURNED
    if (errRET = 2) then                                    ''SERVICE ALREADY STARTED
      objOUT.write vbnewline & retSTOP & vbtab & "SERVICE ALREADY STARTED : Sage SmartPosting 2017 : " & now
      objLOG.write vbnewline & retSTOP & vbtab & "SERVICE ALREADY STARTED : Sage SmartPosting 2017 : " & now
      errRET = 0
      err.clear
    elseif (errRET <> 2) then                               ''ANY OTHER ERROR
      objOUT.write vbnewline & errRET & vbtab & "ERROR STARTING : Sage 50 SmartPosting 2017 : " & now
      objLOG.write vbnewline & errRET & vbtab & "ERROR STARTING : Sage 50 SmartPosting 2017 : " & now
      call LOGERR(5)
    end if
  end if
  ''START SAGE AUTOUPDATE MANAGER SERVICE
  call HOOK("net start " & chr(34) & "Sage AutoUpdate Manager Service" & chr(34))
  if (errRET <> 0) then                                     ''ERROR RETURNED
    if (errRET = 2) then                                    ''SERVICE ALREADY STARTED
      objOUT.write vbnewline & retSTOP & vbtab & "SERVICE ALREADY STARTED : Sage AutoUpdate Manager Service : " & now
      errRET = 0
      err.clear
    elseif (errRET <> 2) then                               ''ANY OTHER ERROR
      objOUT.write vbnewline & errRET & vbtab & "ERROR STARTING : Sage AutoUpdate Manager Service : " & now
      call LOGERR(6)
    end if
  end if
end sub

sub CHKAU()																					        ''CHECK FOR SCRIPT UPDATE , JDC_MSP_POSTBACKUP.VBS , REF #2 , REF #50
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
	if objXML.load("https://github.com/CW-Khristos/scripts/raw/dev/version.xml") then
		set colVER = objXML.documentelement
		for each objSCR in colVER.ChildNodes
			''LOCATE CURRENTLY RUNNING SCRIPT
			if (lcase(objSCR.nodename) = lcase(wscript.scriptname)) then
				''CHECK LATEST VERSION
				if (cint(objSCR.text) > cint(strVER)) then
					objOUT.write vbnewline & now & " - UPDATING " & objSCR.nodename & " : " & objSCR.text & vbnewline
					objLOG.write vbnewline & now & " - UPDATING " & objSCR.nodename & " : " & objSCR.text & vbnewline
					''DOWNLOAD LATEST VERSION OF SCRIPT
					call FILEDL("https://github.com/CW-Khristos/scripts/raw/dev/MSP%20Backups/jdc_msp_psotbackup.vbs", wscript.scriptname)
					''RUN LATEST VERSION
					if (wscript.arguments.count > 0) then             ''ARGUMENTS WERE PASSED
						for x = 0 to (wscript.arguments.count - 1)
							strTMP = strTMP & " " & chr(34) & objARG.item(x) & chr(34)
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

sub FILEDL(strURL, strFILE)                                 ''CALL HOOK TO DOWNLOAD FILE FROM URL
  strSAV = vbnullstring
  ''SET DOWNLOAD PATH
  strSAV = "C:\temp\" & strFILE
  objOUT.write vbnewline & now & vbtab & vbtab & vbtab & "HTTPDOWNLOAD-------------DOWNLOAD : " & strURL & " : SAVE AS :  " & strSAV
  objLOG.write vbnewline & now & vbtab & vbtab & vbtab & "HTTPDOWNLOAD-------------DOWNLOAD : " & strURL & " : SAVE AS :  " & strSAV
  ''CREATE HTTP OBJECT
  set objHTTP = createobject( "WinHttp.WinHttpRequest.5.1" )
  ''DOWNLOAD FROM URL
  objHTTP.open "GET", strURL, false
  objHTTP.send
  if objFSO.fileexists(strSAV) then
    objFSO.deletefile(strSAV)
  end if
  if (objHTTP.status = 200) then
    dim objStream
    set objStream = createobject("ADODB.Stream")
    with objStream
      .Type = 1 'adTypeBinary
      .Open
      .Write objHTTP.ResponseBody
      .SaveToFile strSAV
      .Close
    end with
    set objStream = nothing
  end if
  ''CHECK THAT FILE EXISTS
  if objFSO.fileexists(strSAV) then
    objOUT.write vbnewline & now & vbtab & vbtab & " - DOWNLOAD : " & strSAV & " : SUCCESSFUL"
    objLOG.write vbnewline & now & vbtab & vbtab & " - DOWNLOAD : " & strSAV & " : SUCCESSFUL"
  end if
	set objHTTP = nothing
  if (err.number <> 0) then
    objOUT.write vbnewline & now & vbtab & vbtab & err.number & vbtab & err.description
    objLOG.write vbnewline & now & vbtab & vbtab & err.number & vbtab & err.description
    call LOGERR(2)
    err.clear
  end if
end sub

sub HOOK(strCMD)                                            ''CALL HOOK TO MONITOR OUTPUT OF CALLED COMMAND
  on error resume next
  objOUT.write vbnewline & now & vbtab & vbtab & "EXECUTING : " & strCMD
  objLOG.write vbnewline & now & vbtab & vbtab & "EXECUTING : " & strCMD
  set objHOOK = objWSH.exec(strCMD)
  if (instr(1, strCMD, "takeown /F ") = 0) then             ''SUPPRESS 'TAKEOWN' SUCCESS MESSAGES
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
  ''CHECK FOR ERRORS
  errRET = objHOOK.exitcode
  set objHOOK = nothing
  if (err.number <> 0) then
    objOUT.write vbnewline & now & vbtab & vbtab & vbtab & err.number & vbtab & err.description
    objLOG.write vbnewline & now & vbtab & vbtab & vbtab & err.number & vbtab & err.description
    call LOGERR(3)
    err.clear
  end if
end sub

sub LOGERR(intSTG)                                          ''CALL HOOK TO MONITOR OUTPUT OF CALLED COMMAND
  if (err.number <> 0) then
    objOUT.write vbnewline & now & vbtab & vbtab & vbtab & err.number & vbtab & err.description & vbnewline
    objLOG.write vbnewline & now & vbtab & vbtab & vbtab & err.number & vbtab & err.description & vbnewline
		errRET = intSTG
		err.clear
  end if
end sub

sub CLEANUP()                                               ''SCRIPT CLEANUP
  if (errRET = 0) then                                      ''POST-BACKUP COMPLETED SUCCESSFULLY
    objOUT.write vbnewline & vbnewline & now & vbtab & " - POST-BACKUP COMPLETE : " & now
    objLOG.write vbnewline & vbnewline & now & vbtab & " - POST-BACKUP COMPLETE : " & now
    err.clear
  elseif (errRET <> 0) then                                 ''POST-BACKUP FAILED
    objOUT.write vbnewline & vbnewline & now & vbtab & " - POST-BACKUP FAILURE : " & errRET & " : " & now
    objLOG.write vbnewline & vbnewline & now & vbtab & " - POST-BACKUP FAILURE : " & errRET & " : " & now
    ''RAISE CUSTOMIZED ERROR CODE, ERROR CODE WILL BE DEFINE RESTOP NUMBER INDICATING WHICH SECTION FAILED
    call err.raise(vbObjectError + errRET, "PRE-BACKUP", "FAIL")
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
''SS_MSP_PREBACKUP.VBS
''DESIGNED TO STOP EAGLESOFT SERVICES AND DATABASE TO ALLOW FOR FILE COPY TO 'OFFLINE' DIRECTORY
''SCRIPT UTILIZES ROBOCOPY TO 'MIRROR' SOURCE TO DESTINATION EXACTLY
''MSP BACKUPS EXCLUDE 'ONLINE' EAGLESOFT DIRECTORY AND INCLUDE 'OFFLINE' DIRECTORY
''WRITTEN BY : CJ BLEDSOE / CJ<@>THECOMPUTERWARRIORS.COM
'on error resume next
''DECLARE VARIABLES
dim errRET, strVER
dim objWSH, objFSO, objOUT, objHOOK
''CREATE SCRIPTING SHELL & FILE SYSTEM OBJECTS
set objOUT = wscript.stdout
set objWSH = createobject("wscript.shell")
set objFSO = createobject("scripting.filesystemobject")
''SCRIPT VARIABLES
dim errRET, strVER, strIN
''SCRIPT OBJECTS
dim objIN, objOUT, objARG, objWSH, objFSO
dim objLOG, objHOOK, objEXEC, objHTTP, objXML
''VERSION FOR SCRIPT UPDATE , SS_MSP_PREBACKUP.VBS , REF #2 , REF #50
strVER = 2
''DEFAULT FAIL
errRET = 5
''PREPARE LOGFILE
if (objFSO.fileexists("C:\temp\MSP_PREBACKUP")) then        ''LOGFILE EXISTS
  objFSO.deletefile "C:\temp\MSP_PREBACKUP", true
  set objLOG = objFSO.createtextfile("C:\temp\MSP_PREBACKUP")
  objLOG.close
  set objLOG = objFSO.opentextfile("C:\temp\MSP_PREBACKUP", 8)
else                                                        ''LOGFILE NEEDS TO BE CREATED
  set objLOG = objFSO.createtextfile("C:\temp\MSP_PREBACKUP")
  objLOG.close
  set objLOG = objFSO.opentextfile("C:\temp\MSP_PREBACKUP", 8)
end if
''READ PASSED COMMANDLINE ARGUMENTS
if (wscript.arguments.count > 0) then                       ''ARGUMENTS WERE PASSED
  for x = 0 to (wscript.arguments.count - 1)
    objOUT.write vbnewline & now & vbtab & " - ARGUMENT " & (x + 1) & " (ITEM " & x & ") " & " PASSED : " & ucase(objARG.item(x))
    objLOG.write vbnewline & now & vbtab & " - ARGUMENT " & (x + 1) & " (ITEM " & x & ") " & " PASSED : " & ucase(objARG.item(x))
  next 
  if (wscript.arguments.count > 0) then                     ''SET USER , PASSWORD , AND OPERATION LEVEL VARIABLES
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
objOUT.write vbnewline & vbnewline & now & vbtab & " - EXECUTING SS_MSP_PREBACKUP" & vbnewline
objLOG.write vbnewline & vbnewline & now & vbtab & " - EXECUTING SS_MSP_PREBACKUP" & vbnewline
''AUTOMATIC UPDATE , 'ERRRET'=10 , SS_MSP_PREBACKUP.VBS , REF #2 , REF #50
call CHKAU()
''INITIATE STOP SERVICES
call STOPEAGLE()
''END SCRIPT
call CLEANUP()
''END SCRIPT
''------------

''SUB-ROUTINES
sub STOPEAGLE()                                             ''STOP EAGLESOFT SERVICES
  objOUT.write vbnewline & vbnewline & "STOPPING EAGLESOFT SERVICES : " & now
  ''STOP PATTERSON APP SERVICE
  call HOOK("net stop " & chr(34) & "PattersonAppService" & chr(34))
  if (errRET <> 0) then                                     ''ERROR RETURNED
    if (errRET = 2) then                                    ''SERVICE ALREADY STOPPED
      objOUT.write vbnewline & errRET & vbtab & "SERVICE ALREADY STOPPED : PattersonAppService : " & now
      errRET = 0
      err.clear
    elseif (errRET <> 2) then                               ''ANY OTHER ERROR RETURNED
      objOUT.write vbnewline & errRET & vbtab & "ERROR STOPPING : PattersonAppService : " & now
      call LOGERR(4)
    end if
  end if
  ''STOP EAGLESOFT DATABASE
  call STOPDB()
end sub

sub STOPDB()                                                ''STOP EAGLESOFT DATABASE
  objOUT.write vbnewline & vbnewline & "STOPPING EAGLESOFT DATABASE : " & now
  ''CALL PATTERSONSERVERSTATUS.EXE UTILITY WITH 'STOP' SWITCH
  call HOOK(chr(34) & "C:\EagleSoft\Shared Files\PattersonServerStatus.exe" & chr(34) & " -stop")
  if (errRET <> 0) then                                     ''ERROR RETURNED
    objOUT.write vbnewline & errRET & vbtab & "EAGLESOFT DATABASE : ERROR STOPPING: " & now
    call LOGERR(5)
    ''END SCRIPT, RETURN EXIT CODE
    call CLEANUP()
  end if
  objOUT.write vbnewline & vbtab & "EAGLESOFT DATABASE : STOPPED : " & now
  ''COPY EAGLESOFT DATA
  call DBCOPY()
end sub

sub DBCOPY()                                                ''COPY EAGLESOFT DATA FOLDER
  objOUT.write vbnewline & vbnewline & "COPYING EAGLESOFT DATA : " & now
  ''USE ROBOCOPY TO COPY C:\EAGLESOFT\DATA FOLDER, OLVERWRITE ALL FILES IN DESTINATION , SS_MSP_PREBACKUP.VBS , REF #2 , REF #49
  call HOOK("robocopy " & chr(34) & "C:\EagleSoft\Data" & chr(34) & " " & chr(34) & "B:\EaglesoftBackup" & chr(34) & " /MIR /z /w:1 /r:1 /mt /v")
  if (errRET > 4) then                                      ''SUCCESSFULLY COPIED DATA
    objOUT.write vbnewline & "COPY EAGLESOFT DATA COMPLETE : " & now
  elseif (errRET < 5) then                                 ''ERROR RETURNED
    objOUT.write vbnewline & errRET & vbtab & "ERROR : ROBOCOPY C:\EAGLESOFT\DATA E:\BACKUP : " & now
    call LOGERR(6)
  end if
end sub

sub CHKAU()																					        ''CHECK FOR SCRIPT UPDATE , SS_MSP_PREBACKUP.VBS , REF #2 , REF #50
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
					call FILEDL("https://github.com/CW-Khristos/scripts/raw/dev/MSP%20Backups/ss_msp_prebackup.vbs", wscript.scriptname)
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
  if (errRET = 0) then                                      ''PRE-BACKUP COMPLETED SUCCESSFULLY
    objOUT.write vbnewline & vbnewline & now & vbtab & " - PRE-BACKUP COMPLETE : " & now
    objLOG.write vbnewline & vbnewline & now & vbtab & " - PRE-BACKUP COMPLETE : " & now
    err.clear
  elseif (errRET <> 0) then                                 ''PRE-BACKUP FAILED
    objOUT.write vbnewline & vbnewline & now & vbtab & " - PRE-BACKUP FAILURE : " & errRET & " : " & now
    objLOG.write vbnewline & vbnewline & now & vbtab & " - PRE-BACKUP FAILURE : " & errRET & " : " & now
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
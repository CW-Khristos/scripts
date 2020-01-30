on error resume next
''ALWAYS RIGHT-CLICK SCRIPT, CHOOSE "PROPERTIES", CLICK "UNBLOCK"
''SCRIPT VARIABLES
dim strNUL, strSEL, strRUN
dim strIDL, strTMP, arrTMP, strIN
dim strNAM, strACT, strDAT, strDAY, strMON, strTIM, strARC
dim errRET, strVER, blnRUN, blnSUP
''SCRIPT OBJECTS
dim objIN, objOUT, objARG, objWSH
dim objFSO, objLOG, objHOOK, objHTTP, objXML
''SET 'ERRRET' CODE
errRET = 0
''VERSION FOR SCRIPT UPDATE, SET_MSPARCHIVE.VBS, REF #2
strVER = 2
''DEFAULT 'BLNRUN' FLAG - RESTART BACKUPS IF WRITERS ARE STABLE
blnRUN = false
''STDIN / STDOUT
set objIN = wscript.stdin
set objOUT = wscript.stdout
set objARG = wscript.arguments
''OBJECTS FOR LOCATING FOLDERS
set objWSH = createobject("wscript.shell")
set objFSO = createobject("scripting.filesystemobject")
''SET EXECUTION FLAG
strRUN = "false"
''PREPARE LOGFILE
if (objFSO.fileexists("C:\temp\set_msparchive")) then      ''LOGFILE EXISTS
  objFSO.deletefile "C:\temp\set_msparchive", true
  set objLOG = objFSO.createtextfile("C:\temp\set_msparchive")
  objLOG.close
  set objLOG = objFSO.opentextfile("C:\temp\set_msparchive", 8)
else                                                  ''LOGFILE NEEDS TO BE CREATED
  set objLOG = objFSO.createtextfile("C:\temp\set_msparchive")
  objLOG.close
  set objLOG = objFSO.opentextfile("C:\temp\set_msparchive", 8)
end if
''READ PASSED COMMANDLINE ARGUMENTS
if (wscript.arguments.count > 0) then                 ''ARGUMENTS WERE PASSED
  for x = 0 to (wscript.arguments.count - 1)
    objOUT.write vbnewline & now & vbtab & "ARGUMENT " & (x + 1) & " (ITEM " & x & ") " & " PASSED : " & objARG.item(x)
    objLOG.write vbnewline & now & vbtab & "ARGUMENT " & (x + 1) & " (ITEM " & x & ") " & " PASSED : " & objARG.item(x)
  next
  ''SET ARCHIVE SCHEDULE NAME
  strNAM = objARG.item(0)
  ''SET SERVICE / CONFIGURATION OPTIONS
  if (wscript.arguments.count = 7) then               ''ALL ARGUMENTS PASSED
    ''SET ARCHIVE SCHEDULE ACTIVE
    strACT = objARG.item(1)
    ''SET ARCHIVE DATASOURCES
    strDAT = objARG.item(2)
    ''SET ARCHIVE DAYS OF MONTH
    strDAY = objARG.item(3)
    ''SET ARCHIVE MONTHS
    strMON = objARG.item(4)
    ''SET ARCHIVE TIME
    strTIM = objARG.item(5)
    ''SET SCRIPT RUN LEVEL
    strRUN = objARG.item(6)
    if (strRUN = "true") then
      strARC = "Y"
    end if
  end if
else                                                  ''NO ARGUMENTS PASSED
  objOUT.write vbnewline & now & vbtab & " - NO ARGUMENTS PASSED. SCRIPT WILL REQUEST SETTINGS DURING EXECUTION"
  objLOG.write vbnewline & now & vbtab & " - NO ARGUMENTS PASSED. SCRIPT WILL REQUEST SETTINGS DURING EXECUTION"
end if

''------------
''BEGIN SCRIPT
objOUT.write vbnewline & now & " - STARTING SET_MSPARCHIVE" & vbnewline
objLOG.write vbnewline & now & " - STARTING SET_MSPARCHIVE" & vbnewline
''AUTOMATIC UPDATE, SET_MSPARCHIVE.VBS, REF #2
call CHKAU()
''ENTER CALL VERIFY LOOP
call VERIFY()
''END SCRIPT
call CLEANUP()
''END SCRIPT
''------------

''SUB-ROUTINES
sub VERIFY()                                          ''CALL HOOK TO VERIFY SCRIPT CONFIGURATIONS, RUNS IN A LOOP UNTIL EXECUTION / TERMINATION
  if (wscript.arguments.count = 0) then               ''SCRIPT NOT PRE-CONFIGURED
    strSEL = vbnullstring
    ''ENABLE MSP BACKUP ARCHIVES
    if (strARC = vbnullstring) then
      objOUT.write vbnewline & vbtab & "ENABLE MSP BACKUP ARCHIVES (Y / N)?"
      objLOG.write vbnewline & vbtab & "ENABLE MSP BACKUP ARCHIVES (Y / N)?"
      strSEL = objIN.readline
      strARC = strSEL
      if (ucase(strSEL) = "Y") then
        objOUT.write vbnewline & vbtab & "ENABLE MSP BACKUP ARCHIVES : " & strARC
        objLOG.write vbnewline & vbtab & "ENABLE MSP BACKUP ARCHIVES : " & strARC
      end if
    elseif (strARC <> vbnullstring) then
      objOUT.write vbnewline & vbtab & "ENABLE MSP BACKUP ARCHIVES : " & strARC
      objLOG.write vbnewline & vbtab & "ENABLE MSP BACKUP ARCHIVES : " & strARC
    end if
    strSEL = vbnullstring
  elseif (wscript.arguments.count > 0) then           ''SCRIPT PRE-CONFIGURED
    ''PLACEHOLDER
  end if
  ''EXECUTION CHECK
  if (lcase(strRUN) <> "true") then
    objOUT.write vbnewline & vbnewline & vbtab & "EXECUTE SCRIPT WITH CONFIGURED SETTINGS (Y / N)"
    objLOG.write vbnewline & vbnewline & vbtab & "EXECUTE SCRIPT WITH CONFIGURED SETTINGS (Y / N)"
    strSEL = objIN.readline
    if (ucase(strSEL) = "Y") then
      strRUN = "true"
      objOUT.write vbnewline & vbnewline & now & " - EXECUTING SET_MSPARCHIVE SCRIPT" & vbnewline
      objLOG.write vbnewline & vbnewline & now & " - EXECUTING SET_MSPARCHIVE SCRIPT" & vbnewline
      ''EXIT VERIFY LOOP, RUN SCRIPT EXECUTION
      call EXECUTE()
    elseif (ucase(strSEL) = "N") then
      strRUN = "false"
      objOUT.write vbnewline & vbnewline & now & " - SKIPPING SCRIPT EXECUTION" & vbnewline
      objLOG.write vbnewline & vbnewline & now & " - SKIPPING SCRIPT EXECUTION" & vbnewline
      ''RETURN CALL TO VERIFY LOOP
      call VERIFY()
    end if
  elseif (lcase(strRUN) = "true") then
    ''EXIT VERIFY LOOP, RUN SCRIPT EXECUTION
    call EXECUTE()
  end if
end sub

sub EXECUTE()                                         ''CALL HOOK TO EXECUTE SCRIPT CHANGES
  if (strRUN = "true") then
    ''PAUSE TO ENABLE MSP BACKUP ARCHIVES
    if (lcase(strARC) = "y") then
      ''CHECK MSP BACKUP STATUS VIA MSP BACKUP CLIENTTOOL UTILITY AFTER RESTART
      for intLOOP = 0 to 10
        objOUT.write vbnewline & now & vbtab & " - CHECKING MSP BACKUP STATUS"
        objLOG.write vbnewline & now & vbtab & " - CHECKING MSP BACKUP STATUS"
        set objHOOK = objWSH.exec(chr(34) & "c:\Program Files\Backup Manager\ClientTool.exe" & chr(34) & " control.status.get")
        strIDL = objHOOK.stdout.readall
        objOUT.write vbnewline & now & vbtab & vbtab & vbtab & strIDL
        objLOG.write vbnewline & now & vbtab & vbtab & vbtab & strIDL
        set objHOOK = nothing
        ''SERVICE NOT STARTED
        if (strIDL = vbnullstring) then
          objOUT.write vbnewline & now & vbtab & " - CLIENTTOOL NOT READY, RESTARTING BACKUP SERVICE"
          objLOG.write vbnewline & now & vbtab & " - CLIENTTOOL NOT READY, RESTARTING BACKUP SERVICE"
          call HOOK("net start " & chr(34) & "Backup Service Controller" & chr(34))
          set objHOOK = objWSH.exec(chr(34) & "c:\Program Files\Backup Manager\ClientTool.exe" & chr(34) & " control.status.get")
          strIDL = objHOOK.stdout.readall
          objOUT.write vbnewline & now & vbtab & vbtab & vbtab & strIDL
          objLOG.write vbnewline & now & vbtab & vbtab & vbtab & strIDL
          set objHOOK = nothing
        end if
        if ((instr(1, strIDL, "Idle")) or (instr(1, strIDL, "RegSync"))) then     			      ''BACKUPS NOT IN PROGRESS
            ''FORCE RUN OF SYSTEM STATE
            blnRUN = true
            if (blnRUN) then														      ''ENABLE ARCHIVING
              ''ADDITIONAL DELAY TO GIVE SERVICE A BIT EXTRA Time
              wscript.sleep (60000)
              objOUT.write vbnewline & now & vbtab & vbtab & vbtab & "CLIENTTOOL READY, ENABLING ARCHIVING"
              objLOG.write vbnewline & now & vbtab & vbtab & vbtab & "CLIENTTOOL READY, ENABLING ARCHIVING"
              blnRUN = true
            end if
            exit for
        elseif ((strIDL = vbnullstring) or (instr(1, strIDL, "Idle") = 0) or _
          (instr(1, strIDL, "RegSync") = 0) or (instr(1, strIDL, "Suspended"))) then					''BACKUPS IN PROGRESS, SERVICE NOT READY
            objOUT.write vbnewline & now & vbtab & vbtab & vbtab & "BACKUPS IN PROGRESS, SERVICE NOT READY" 
            objLOG.write vbnewline & now & vbtab & vbtab & vbtab & "BACKUPS IN PROGRESS, SERVICE NOT READY"
            blnRUN = false
        end if
        wscript.sleep 12000
      next
      if (not blnRUN) then                                        ''SERVICE DID NOT INITIALIZE , 'ERRRET'=1
        objOUT.write vbnewline & now & vbtab & vbtab & vbtab & "SERVICE NOT READY, TERMINATING" 
        objLOG.write vbnewline & now & vbtab & vbtab & vbtab & "SERVICE NOT READY, TERMINATING"
        call LOGERR(1)
      elseif (blnRUN) then
        ''REQUEST LOCAL SPEEDVAULT PATH
        'objOUT.write vbnewline & vbnewline & now & vbtab & " - ENTER LOCAL SPEEDVAULT PATH : "
        'objLOG.write vbnewline & vbnewline & now & vbtab & " - ENTER LOCAL SPEEDVAULT PATH : "
        'strLSVL = objIN.readline
        ''REQUEST RMMTECH CREDENTIALS FOR LOCAL SPEEDVAULT
        'objOUT.write vbnewline & vbnewline & now & vbtab & " - ENTER RMMTECH USERNAME FOR LSV ACCESS : "
        'objLOG.write vbnewline & vbnewline & now & vbtab & " - ENTER RMMTECH USERNAME FOR LSV ACCESS : "
        'strLSVU = objIN.readline
        'objOUT.write vbnewline & vbnewline & now & vbtab & " - ENTER RMMTECH PASSWORD FOR LSV ACCESS : "
        'objLOG.write vbnewline & vbnewline & now & vbtab & " - ENTER RMMTECH PASSWORD FOR LSV ACCESS : "
        'strLSVP = objIN.readline
        ''SET DEFAULT 'CW_DEFAULT_MSPARCHIVE" ARCHIVING SCHEDULE
        objOUT.write vbnewline & vbnewline & now & vbtab & " - APPLYING MSP BACKUP ARCHIVE SCHEDULE '" & strNAM & "'"
        objLOG.write vbnewline & vbnewline & now & vbtab & " - APPLYING MSP BACKUP ARCHIVE SCHEDULE '" & strNAM & "'"
        call HOOK("C:\Program Files\Backup Manager\ClientTool.exe control.archiving.add -name " & chr(34) & strNAM & chr(34) & _
          " -active " & chr(34) & strACT & chr(34) & " -datasources " & chr(34) & strDAT & chr(34) & " -days-of-month " & chr(34) & strDAY & chr(34) & _
          " -months " & chr(34) & strMON & chr(34) & " -time " & chr(34) & strTIM & chr(34))
        ''STEP TO VERIFY MSP BACKUP ARCHIVING SCHEDULE
        objOUT.write vbnewline & vbnewline & now & vbtab & " - PLEASE VERIFY MSP BACKUP ARCHIVE SCHEDULE 'CW_DEFAULT_MSPARCHIVE'"
        objOUT.write vbnewline & now & vbtab & " - VIA N-CENTRAL>CONFIGURATION>BACKUP MANAGER>MSP BACKUPS>DASHBOARD>LAUNCH BACKUP CLIENT>PREFERENCES>ARCHIVING"
        objOUT.write vbnewline & now & vbtab & " - PRESS 'ENTER' WHEN READY"
        objLOG.write vbnewline & vbnewline & now & vbtab & " - PLEASE VERIFY MSP BACKUP ARCHIVE SCHEDULE 'CW_DEFAULT_MSPARCHIVE'"
        objLOG.write vbnewline & now & vbtab & " - VIA N-CENTRAL>CONFIGURATION>BACKUP MANAGER>MSP BACKUPS>DASHBOARD>LAUNCH BACKUP CLIENT>PREFERENCES>ARCHIVING"
        if (wscript.arguments.count <> 7) then               ''NOT ALL ARGUMENTS PASSED
          objLOG.write vbnewline & now & vbtab & " - PRESS 'ENTER' WHEN READY"
          strNUL = objIN.readline
        end if
      end if
    end if
  end if
end sub

sub CHKAU()																					        ''CHECK FOR SCRIPT UPDATE, SET_MSPARCHIVE.VBS, REF #2 , FIXES #4
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
        objOUT.write vbnewline & now & vbtab & " - Set_MSPArchive :  " & strVER & " : GitHub : " & objSCR.text & vbnewline
        objLOG.write vbnewline & now & vbtab & " - Set_MSPArchive :  " & strVER & " : GitHub : " & objSCR.text & vbnewline
				if (cint(objSCR.text) > cint(strVER)) then
					objOUT.write vbnewline & now & " - UPDATING " & objSCR.nodename & " : " & objSCR.text & vbnewline
					objLOG.write vbnewline & now & " - UPDATING " & objSCR.nodename & " : " & objSCR.text & vbnewline
					''DOWNLOAD LATEST VERSION OF SCRIPT
					call FILEDL("https://github.com/CW-Khristos/scripts/raw/dev/MSP%20Backups/set_msparchive.vbs", wscript.scriptname)
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

sub FILEDL(strURL, strFILE)                                 ''CALL HOOK TO DOWNLOAD FILE FROM URL , 'ERRRET'=2
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
  ''CHECK IF FILE ALREADY EXISTS
  if objFSO.fileexists(strSAV) then
    ''DELETE FILE FOR OVERWRITE
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
  ''ERROR RETURNED
  if (err.number <> 0) then
    objOUT.write vbnewline & now & vbtab & vbtab & err.number & vbtab & err.description
    objLOG.write vbnewline & now & vbtab & vbtab & err.number & vbtab & err.description
    call LOGERR(2)
    err.clear
  end if
end sub

sub HOOK(strCMD)                                            ''CALL HOOK TO MONITOR OUTPUT OF CALLED COMMAND , 'ERRRET'=3
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
  if ((not blnSUP) and (err.number <> 0)) then
    objOUT.write vbnewline & now & vbtab & vbtab & vbtab & err.number & vbtab & err.description
    objLOG.write vbnewline & now & vbtab & vbtab & vbtab & err.number & vbtab & err.description
    call LOGERR(3)
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

sub CLEANUP()                                 			        ''SCRIPT CLEANUP
  if (errRET = 0) then         											        ''SET_MSPARCHIVE COMPLETED SUCCESSFULLY
    objOUT.write vbnewline & "SET_MSPARCHIVE SUCCESSFUL : " & NOW
  elseif (errRET <> 0) then    											        ''SET_MSPARCHIVE FAILED
    objOUT.write vbnewline & "SET_MSPARCHIVE FAILURE : " & NOW & " : " & errRET
    ''RAISE CUSTOMIZED ERROR CODE, ERROR CODE WILL BE DEFINE RESTOP NUMBER INDICATING WHICH SECTION FAILED
    call err.raise(vbObjectError + errRET, "SET_MSPARCHIVE", "FAILURE")
  end if
  objOUT.write vbnewline & vbnewline & now & " - SET_MSPARCHIVE COMPLETE" & vbnewline
  objLOG.write vbnewline & vbnewline & now & " - SET_MSPARCHIVE COMPLETE" & vbnewline
  objLOG.close
  ''EMPTY OBJECTS
  set objLOG = nothing
  set objFSO = nothing
  set objWSH = nothing
  set objARG = nothing
  set objOUT = nothing
  set objIN = nothing
  ''END SCRIPT, RETURN ERROR NUMBER
  wscript.quit err.number
end sub
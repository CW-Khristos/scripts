''CODEDROP_FIX.VBS
''SCRIPT IS DESIGNED TO DOWNLOAD AND AUTOMATE 'CODEDROP' FIX FROM SOLARWINDS FOR SELF-HEAL AND COPY/PASTE ISSUES, REF #2 , REF #1
''WRITTEN BY : CJ BLEDSOE / CJ<@>THECOMPUTERWARRIORS.COM
'on error resume next
dim strFIX
''SCRIPT VARIABLES
dim errRET, strVER, strIN, strCDD
''SCRIPT OBJECTS
dim objIN, objOUT, objARG, objWSH
dim objFSO, objLOG, objHOOK, objHTTP, objXML
''SET 'ERRRET' CODE
errRET = 0
''VERSION FOR SCRIPT UPDATE, CODEDROP_FIX.VBS, REF #2 , REF #1
strVER = 5
''STDIN / STDOUT
set objIN = wscript.stdin
set objOUT = wscript.stdout
set objARG = wscript.arguments
''OBJECTS FOR LOCATING FOLDERS
set objWSH = createobject("wscript.shell")
set objFSO = createobject("scripting.filesystemobject")
''PREPARE LOGFILE
if (objFSO.fileexists("C:\temp\CODEDROP_FIX")) then		               ''LOGFILE EXISTS
  objFSO.deletefile "C:\temp\CODEDROP_FIX", true
  set objLOG = objFSO.createtextfile("C:\temp\CODEDROP_FIX")
  objLOG.close
  set objLOG = objFSO.opentextfile("C:\temp\CODEDROP_FIX", 8)
else                                                                 ''LOGFILE NEEDS TO BE CREATED
  set objLOG = objFSO.createtextfile("C:\temp\CODEDROP_FIX")
  objLOG.close
  set objLOG = objFSO.opentextfile("C:\temp\CODEDROP_FIX", 8)
end if
''CHECK EXECUTION METHOD OF SCRIPT
strIN = lcase(mid(wscript.fullname, instrrev(wscript.fullname, "\") + 1))
if (strIN <> "cscript.exe") Then
	''needs to save and pass arguments
  objOUT.write vbnewline & "SCRIPT LAUNCHED VIA EXPLORER, EXECUTING SCRIPT VIA CSCRIPT..."
  objWSH.run "cscript.exe //nologo " & chr(34) & Wscript.ScriptFullName & chr(34)
  wscript.quit
end if
''READ PASSED COMMANDLINE ARGUMENTS
if (wscript.arguments.count < 1) then                                ''NO ARGUMENTS PASSED, END SCRIPT, 'ERRRET'=1
  objOUT.write vbnewline & vbnewline & now & vbtab & " - SCRIPT REQUIRES CODE DROP FIX SELECTION : SELFHEAL / COPYPASTE"
  objLOG.write vbnewline & vbnewline & now & vbtab & " - SCRIPT REQUIRES CODE DROP FIX SELECTION : SELFHEAL / COPYPASTE"
  call LOGERR(1)
  call CLEANUP()
elseif (wscript.arguments.count = 1) then                             ''ARGUMENTS WERE PASSED
  for x = 0 to (wscript.arguments.count - 1)
    objOUT.write vbnewline & now & vbtab & " - ARGUMENT " & (x + 1) & " (ITEM " & x & ") " & " PASSED : " & ucase(objARG.item(x))
    objLOG.write vbnewline & now & vbtab & " - ARGUMENT " & (x + 1) & " (ITEM " & x & ") " & " PASSED : " & ucase(objARG.item(x))
  next 
  strFIX = objARG.item(0)                                             ''SET STRING 'STRFIX', CODE DROP FIX SELECTION
end if

''------------
''BEGIN SCRIPT
objOUT.write vbnewline & now & " - STARTING CODEDROP_FIX" & vbnewline
objLOG.write vbnewline & now & " - STARTING CODEDROP_FIX" & vbnewline
''AUTOMATIC UPDATE, CODEDROP_FIX.VBS, REF #2 , REF #1
call CHKAU()
''STOP WINDOWS AGENT SERVICES
objOUT.write vbnewline & now & vbtab & " - STOPPING WINDOWS AGENT SERVICES"
objLOG.write vbnewline & now & vbtab & " - STOPPING WINDOWS AGENT SERVICES"
call HOOK("net stop " & chr(34) & "Windows Agent Maintenance Service" & chr(34))
call HOOK("net stop " & chr(34) & "Windows Agent Service" & chr(34))
wscript.sleep 5000
''DOWNLOAD CODEDROP 'FIX' FILES
objOUT.write vbnewline & now & vbtab & " - DOWNLOADING CODEDROP 'FIX' FILES"
objLOG.write vbnewline & now & vbtab & " - DOWNLOADING CODEDROP 'FIX' FILES"
if (ucase(strFIX) = "SELFHEAL") then
  ''WINDOWS AGENT CODEDROP Directory
  strCDD = "C:\Program Files (x86)\N-able Technologies\Windows Agent\bin"
  objOUT.write vbnewline & now & vbtab & " - DOWNLOADING CODEDROP 'SELF-HEAL' FILES"
  objLOG.write vbnewline & now & vbtab & " - DOWNLOADING CODEDROP 'SELF-HEAL' FILES"
  call FILEDL("https://github.com/CW-Khristos/scripts/raw/dev/CodeDrop/selfheal/codedrop_MAR30_NCI-15758/agent.exe", "agent.exe")
  call FILEDL("https://github.com/CW-Khristos/scripts/raw/dev/CodeDrop/selfheal/codedrop_MAR30_NCI-15758/CodeDropMeta.xml", "CodeDropMeta.xml")
  'call FILEDL("https://github.com/CW-Khristos/scripts/raw/dev/CodeDrop/selfheal/codedrop_MAR17_NCI-15758/agent.exe", "agent.exe")
  'call FILEDL("https://github.com/CW-Khristos/scripts/raw/dev/CodeDrop/selfheal/codedrop_MAR17_NCI-15758/CodeDropMeta.xml", "CodeDropMeta.xml")
  ''RENAME 'OLD' CODEDROP FILES
  objOUT.write vbnewline & now & vbtab & " - RENAMING 'OLD' CODEDROP FILES"
  objLOG.write vbnewline & now & vbtab & " - RENAMING 'OLD' CODEDROP FILES"
  if objFSO.fileexists(strCDD & "\agent.exe") then
    call HOOK("cmd.exe /C move /y " & chr(34) & strCDD & "\agent.exe" & chr(34) & " " & chr(34) & strCDD & "\agent.old" & chr(34))
  end if
  'if objFSO.fileexists(strSAV) then
  '  call HOOK("cmd.exe /C move /y " & chr(34) & strCDD & "\CodeDropMeta.xml" & chr(34) & " " & chr(34) & strCDD & "\CodeDropMeta.old" & chr(34))
  'end if
  ''MOVE CODEDROP 'FIX' FILES TO APPROPRIATE LOCATION
  objOUT.write vbnewline & now & vbtab & " - MOVING CODEDROP 'FIX' FILES TO APPROPRIATE LOCATION"
  objLOG.write vbnewline & now & vbtab & " - MOVING CODEDROP 'FIX' FILES TO APPROPRIATE LOCATION"
  ''CHECK THAT FILE EXISTS
  if objFSO.fileexists("C:\Temp\agent.exe") then
    call HOOK("cmd.exe /C move /y " & chr(34) & "c:\temp\agent.exe" & chr(34) & " " & chr(34) & strCDD & chr(34))
    'objFSO.copyfile "C:\Temp\agent.exe", strCDD & "\agent.exe", true
  end if
  ''CHECK THAT FILE EXISTS
  if objFSO.fileexists("C:\Temp\CodeDropMeta.xml") then
    call HOOK("cmd.exe /C move /y " & chr(34) & "c:\temp\CodeDropMeta.xml" & chr(34) & " " & chr(34) & strCDD & chr(34))
    'objFSO.copyfile "C:\Temp\CodeDropMeta.xml", strCDD & "\CodeDropMeta.xml", true
  end if
elseif (ucase(strFIX) = "COPYPASTE") then
  strCDD = "C:\Program Files (x86)\N-Able Technologies\Reactive\bin"
  objOUT.write vbnewline & now & vbtab & " - DOWNLOADING CODEDROP 'COPY/PASTE' FILES"
  objLOG.write vbnewline & now & vbtab & " - DOWNLOADING CODEDROP 'COPY/PASTE' FILES"
  call FILEDL("https://github.com/CW-Khristos/scripts/blob/dev/CodeDrop/copypaste/ConsoleAPIWrapper32_64/ConsoleAPIWrapper32.dll", "ConsoleAPIWrapper32.dll")
  call FILEDL("https://github.com/CW-Khristos/scripts/blob/dev/CodeDrop/copypaste/ConsoleAPIWrapper32_64/ConsoleAPIWrapper64.dll", "ConsoleAPIWrapper64.dll")
  ''RENAME 'OLD' CODEDROP FILES
  objOUT.write vbnewline & now & vbtab & " - RENAMING 'OLD' CODEDROP FILES"
  objLOG.write vbnewline & now & vbtab & " - RENAMING 'OLD' CODEDROP FILES"
  if objFSO.fileexists(strCDD & "\ConsoleAPIWrapper32.dll") then
    call HOOK("cmd.exe /C move /y " & chr(34) & strCDD & "\ConsoleAPIWrapper32.dll" & chr(34) & " " & chr(34) & strCDD & "\ConsoleAPIWrapper32.old" & chr(34))
  end if
  if objFSO.fileexists(strCDD & "\ConsoleAPIWrapper64.dll") then
    call HOOK("cmd.exe /C move /y " & chr(34) & strCDD & "\ConsoleAPIWrapper64.dll" & chr(34) & " " & chr(34) & strCDD & "\ConsoleAPIWrapper64.old" & chr(34))
  end if
  ''MOVE CODEDROP 'FIX' FILES TO APPROPRIATE LOCATION
  objOUT.write vbnewline & now & vbtab & " - MOVING CODEDROP 'FIX' FILES TO APPROPRIATE LOCATION"
  objLOG.write vbnewline & now & vbtab & " - MOVING CODEDROP 'FIX' FILES TO APPROPRIATE LOCATION"
  ''CHECK THAT FILE EXISTS
  if objFSO.fileexists("C:\Temp\ConsoleAPIWrapper32.dll") then
    call HOOK("cmd.exe /C move /y " & chr(34) & "c:\temp\ConsoleAPIWrapper32.dll" & chr(34) & " " & chr(34) & strCDD & chr(34))
    'objFSO.copyfile "C:\Temp\ConsoleAPIWrapper32.dll", strCDD & "\ConsoleAPIWrapper32.dll", true
  end if
  ''CHECK THAT FILE EXISTS
  if objFSO.fileexists("C:\Temp\ConsoleAPIWrapper64.dll") then
    call HOOK("cmd.exe /C move /y " & chr(34) & "c:\temp\ConsoleAPIWrapper64.dll" & chr(34) & " " & chr(34) & strCDD & chr(34))
    'objFSO.copyfile "C:\Temp\ConsoleAPIWrapper64.dll", strCDD & "\ConsoleAPIWrapper64.dll", true
  end if
end if
''RESTART WINDOWS AGENT SERVICES
objOUT.write vbnewline & now & vbtab & " - RESTARTING WINDOWS AGENT SERVICES"
objLOG.write vbnewline & now & vbtab & " - RESTARTING WINDOWS AGENT SERVICES"
call HOOK("net start " & chr(34) & "Windows Agent Maintenance Service" & chr(34))
call HOOK("net start " & chr(34) & "Windows Agent Service" & chr(34))
''END SCRIPT
call CLEANUP()
''END SCRIPT
''------------

sub CHKAU()																					        ''CHECK FOR SCRIPT UPDATE, CODEDROP_FIX.VBS, REF #2 , REF #1
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
        objOUT.write vbnewline & now & vbtab & " - CODEDROP_FIX :  " & strVER & " : GitHub : " & objSCR.text & vbnewline
        objLOG.write vbnewline & now & vbtab & " - CODEDROP_FIX :  " & strVER & " : GitHub : " & objSCR.text & vbnewline
				if (cint(objSCR.text) > cint(strVER)) then
					objOUT.write vbnewline & now & " - UPDATING " & objSCR.nodename & " : " & objSCR.text & vbnewline
					objLOG.write vbnewline & now & " - UPDATING " & objSCR.nodename & " : " & objSCR.text & vbnewline
					''DOWNLOAD LATEST VERSION OF SCRIPT
					call FILEDL("https://github.com/CW-Khristos/scripts/raw/dev/CodeDrop/CodeDrop_Fix.vbs", wscript.scriptname)
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
  errRET = intSTG
  if (err.number <> 0) then
    objOUT.write vbnewline & now & vbtab & vbtab & vbtab & err.number & vbtab & err.description & vbnewline
    objLOG.write vbnewline & now & vbtab & vbtab & vbtab & err.number & vbtab & err.description & vbnewline
		err.clear
  end if
end sub

sub CLEANUP()                                 			        ''SCRIPT CLEANUP
  if (errRET = 0) then         											        ''CODEDROP_FIX COMPLETED SUCCESSFULLY
    objOUT.write vbnewline & "CODEDROP_FIX SUCCESSFUL : " & now
  elseif (errRET <> 0) then    											        ''CODEDROP_FIX FAILED
    objOUT.write vbnewline & "CODEDROP_FIX FAILURE : " & now & " : " & errRET
    ''RAISE CUSTOMIZED ERROR CODE, ERROR CODE WILL BE DEFINE RESTOP NUMBER INDICATING WHICH SECTION FAILED
    call err.raise(vbObjectError + errRET, "CODEDROP_FIX", "FAILURE")
  end if
  objOUT.write vbnewline & vbnewline & now & " - CODEDROP_FIX COMPLETE" & vbnewline
  objLOG.write vbnewline & vbnewline & now & " - CODEDROP_FIX COMPLETE" & vbnewline
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
''PMESERVICE_INSTALL.VBS
''SCRIPT IS DESIGNED TO DOWNLOAD AND EXECUTE PME SERVICE UPDATE
''WRITTEN BY : CJ BLEDSOE / CJ<@>THECOMPUTERWARRIORS.COM
on error resume next
''SCRIPT VARIABLES
dim errRET, strVER, strIN
''REGISTRY CONSTANTS
const HKCR = &H80000000
const HKLM = &H80000002
''SCRIPT OBJECTS
dim objAPP, objSRC, objTGT
dim objIN, objOUT, objARG, objWSH
dim objFSO, objLOG, objHOOK, objHTTP, objXML
''SET 'ERRRET' CODE
errRET = 0
''VERSION FOR SCRIPT UPDATE, PMESERVICE_INSTALL.VBS, REF #2
strVER = 1
''STDIN / STDOUT
set objIN = wscript.stdin
set objOUT = wscript.stdout
set objARG = wscript.arguments
''OBJECTS FOR LOCATING FOLDERS
set objWSH = createobject("wscript.shell")
set objFSO = createobject("scripting.filesystemobject")
''PREPARE LOGFILE
if (objFSO.fileexists("C:\temp\PMESERVICE_INSTALL")) then		        ''LOGFILE EXISTS
  objFSO.deletefile "C:\temp\PMESERVICE_INSTALL", true
  set objLOG = objFSO.createtextfile("C:\temp\PMESERVICE_INSTALL")
  objLOG.close
  set objLOG = objFSO.opentextfile("C:\temp\PMESERVICE_INSTALL", 8)
else                                                            ''LOGFILE NEEDS TO BE CREATED
  set objLOG = objFSO.createtextfile("C:\temp\PMESERVICE_INSTALL")
  objLOG.close
  set objLOG = objFSO.opentextfile("C:\temp\PMESERVICE_INSTALL", 8)
end if
''CHECK EXECUTION METHOD OF SCRIPT
strIN = lcase(mid(wscript.fullname, instrrev(wscript.fullname, "\") + 1))
if (strIN <> "cscript.exe") Then
	''needs to save and pass arguments
  objOUT.write vbnewline & "SCRIPT LAUNCHED VIA EXPLORER, EXECUTING SCRIPT VIA CSCRIPT..."
  objWSH.run "cscript.exe //nologo " & chr(34) & Wscript.ScriptFullName & chr(34)
  wscript.quit
end if

''------------
''BEGIN SCRIPT
objOUT.write vbnewline & now & " - STARTING PMESERVICE_INSTALL" & vbnewline
objLOG.write vbnewline & now & " - STARTING PMESERVICE_INSTALL" & vbnewline
''AUTOMATIC UPDATE, PMESERVICE_INSTALL.VBS, REF #2 , FIXES #4
call CHKAU()
''STOP WINDOWS PROBE SERVICES
objOUT.write vbnewline & vbnewline & now & vbtab & " - STOPPING WINDOWS PROBE SERVICES" & vbnewline
objLOG.write vbnewline & vbnewline & now & vbtab & " - STOPPING WINDOWS PROBE SERVICES" & vbnewline
intRET = objWSH.run("sc query " & chr(34) & "Windows Software Probe Service" & chr(34), 0, true)
if (intRET = 0) then
  call HOOK("net stop " & chr(34) & "N-able Patch Repository Service" & chr(34))
  call HOOK("net stop " & chr(34) & "Windows Software Probe Maintenance Service" & chr(34))
  call HOOK("net stop " & chr(34) & "Windows Software Probe Service" & chr(34))
end if
wscript.sleep 5000
''DOWNLOAD AND RUN 'CCLUTTERV2.VBS' WHICH INCLUDES NABLEPATCHCACHE AND NABLEUPDATECACHE DIRECTORIES
'call FILEDL("https://github.com/CW-Khristos/scripts/raw/dev/CClutterV2.vbs", "CClutterV2.vbs")
'call HOOK("cscript.exe " & chr(34) & "c:\temp\CClutterV2.vbs" & chr(34) & " " & chr(34) & "true" & chr(34))
''MAKE NECESSARY REGISTRY CHANGES TO ALLOW POWERSHELL 'INVOKE-WEBREQUEST' CMDLET USED BY PME SERVICE TO DOWNLOAD FILES
objOUT.write vbnewline & vbnewline & now & vbtab & " - CHANGING IE FIRST-RUN TO ALLOW POWERSHELL INVOKE-WEBREQUEST"
objLOG.write vbnewline & vbnewline & now & vbtab & " - CHANGING IE FIRST-RUN TO ALLOW POWERSHELL INVOKE-WEBREQUEST"
call HOOK("reg add " & chr(34) & "HKLM\SOFTWARE\Policies\Microsoft\Internet Explorer\Main" & chr(34) & _
  " /f /v DisableFirstRunCustomize /t REG_DWORD /d 0x00000001 /reg:32")
call HOOK("reg add " & chr(34) & "HKLM\SOFTWARE\Policies\Microsoft\Internet Explorer\Main" & chr(34) & _
  " /f /v DisableFirstRunCustomize /t REG_DWORD /d 0x00000001 /reg:64")
call HOOK("reg add " & chr(34) & "HKCU\SOFTWARE\Policies\Microsoft\Internet Explorer\Main" & chr(34) & _
  " /f /v DisableFirstRunCustomize /t REG_DWORD /d 0x00000001 /reg:32")
call HOOK("reg add " & chr(34) & "HKCU\SOFTWARE\Policies\Microsoft\Internet Explorer\Main" & chr(34) & _
  " /f /v DisableFirstRunCustomize /t REG_DWORD /d 0x00000001 /reg:64")
call HOOK("reg add " & chr(34) & "HKLM\SOFTWARE\Microsoft\Internet Explorer\Main" & chr(34) & _
  " /f /v DisableFirstRunCustomize /t REG_DWORD /d 0x00000001 /reg:32")
call HOOK("reg add " & chr(34) & "HKLM\SOFTWARE\Microsoft\Internet Explorer\Main" & chr(34) & _
  " /f /v DisableFirstRunCustomize /t REG_DWORD /d 0x00000001 /reg:64")
call HOOK("reg add " & chr(34) & "HKCU\SOFTWARE\Microsoft\Internet Explorer\Main" & chr(34) & _
  " /f /v DisableFirstRunCustomize /t REG_DWORD /d 0x00000001 /reg:32")
call HOOK("reg add " & chr(34) & "HKCU\SOFTWARE\Microsoft\Internet Explorer\Main" & chr(34) & _
  " /f /v DisableFirstRunCustomize /t REG_DWORD /d 0x00000001 /reg:64")
''DOWNLOAD PME SERVICE SUPPORTING FILES
objOUT.write vbnewline & vbnewline & now & vbtab & " - DOWNLOADING PME SERVICE SUPPORTING FILES" & vbnewline
objLOG.write vbnewline & vbnewline & now & vbtab & " - DOWNLOADING PME SERVICE SUPPORTING FILES" & vbnewline
call FILEDL("https://github.com/CW-Khristos/scripts/raw/dev/PMEService.zip", "PMEService.zip")
wscript.sleep 5000
''DOWNLOAD SUPPORTING FILES
if (not objFSO.fileexists("c:\temp\PMEService.zip")) then
  call FILEDL("https://github.com/CW-Khristos/scripts/raw/dev/PMEService.zip", "PMEService.zip")
end if
if (objFSO.fileexists("C:\temp\PMEService.zip")) then
  ''EXTRACT PME SERVICE SUPPORTING FILES
  set objSRC = objAPP.namespace("C:\temp\PMEService.zip").items()
  set objTGT = objAPP.namespace("C:\temp")
  objTGT.copyhere objSRC, intOPT
end if
''CHECK FOR EXTRACTED X.ROBOT
if (objFSO.fileexists("c:\temp\PMEService\AnniversaryUpdates_details.xml")) then
  ''MOVE PME SERVICE SUPPORTING FILES TO 'C:\PROGRAMDATA\SOLARWINDS MSP\PME\ARCHIVES'
  objOUT.write vbnewline & vbnewline & now & vbtab & " - MOVING ANNIVERSARYUPDATES_DETAILS.XML" & vbnewline
  objLOG.write vbnewline & vbnewline & now & vbtab & " - MOVING ANNIVERSARYUPDATES_DETAILS.XML" & vbnewline
  call HOOK("cmd.exe /C move /y " & chr(34) & "c:\temp\PMEService\AnniversaryUpdates_details.xml" & chr(34) & " " & chr(34) & "C:\ProgramData\SolarWinds MSP\PME\Archives" & chr(34))
end if
''CHECK FOR EXTRACTED X.ROBOT
if (objFSO.fileexists("c:\temp\PMEService\AnniversaryUpdates_details.xml")) then
  ''MOVE PME SERVICE SUPPORTING FILES TO 'C:\PROGRAMDATA\SOLARWINDS MSP\PME\ARCHIVES'
  objOUT.write vbnewline & vbnewline & now & vbtab & " - MOVING ANNIVERSARYUPDATES.ZIP" & vbnewline
  objLOG.write vbnewline & vbnewline & now & vbtab & " - MOVING ANNIVERSARYUPDATES.ZIP" & vbnewline
  call HOOK("cmd.exe /C move /y " & chr(34) & "c:\temp\PMEService\AnniversaryUpdates.zip" & chr(34) & " " & chr(34) & "C:\ProgramData\SolarWinds MSP\PME\Archives" & chr(34))
end if
''CHECK FOR EXTRACTED X.ROBOT
if (objFSO.fileexists("c:\temp\PMEService\AnniversaryUpdates_details.xml")) then
  ''MOVE PME SERVICE SUPPORTING FILES TO 'C:\PROGRAMDATA\SOLARWINDS MSP\PME\ARCHIVES'
  objOUT.write vbnewline & vbnewline & now & vbtab & " - MOVING SECURITYUPDATES_DETAILS.XML" & vbnewline
  objLOG.write vbnewline & vbnewline & now & vbtab & " - MOVING SECURITYUPDATES_DETAILS.XML" & vbnewline
  call HOOK("cmd.exe /C move /y " & chr(34) & "c:\temp\PMEService\SecurityUpdates_details.xml" & chr(34) & " " & chr(34) & "C:\ProgramData\SolarWinds MSP\PME\Archives" & chr(34))
end if
''CHECK FOR EXTRACTED X.ROBOT
if (objFSO.fileexists("c:\temp\PMEService\AnniversaryUpdates_details.xml")) then
  ''MOVE PME SERVICE SUPPORTING FILES TO 'C:\PROGRAMDATA\SOLARWINDS MSP\PME\ARCHIVES'
  objOUT.write vbnewline & vbnewline & now & vbtab & " - MOVING SECURITYUPDATES.ZIP" & vbnewline
  objLOG.write vbnewline & vbnewline & now & vbtab & " - MOVING SECURITYUPDATES.ZIP" & vbnewline
  call HOOK("cmd.exe /C move /y " & chr(34) & "c:\temp\PMEService\SecurityUpdates.zip" & chr(34) & " " & chr(34) & "C:\ProgramData\SolarWinds MSP\PME\Archives" & chr(34))
end if
''RUN PME SERVICE UPDATE WITH /VERYSILENT SWITCH
if (objFSO.fileexists("c:\temp\PMEService\PMESetup.exe")) then
  objOUT.write vbnewline & vbnewline & now & vbtab & " - EXECUTING PME SERVICE UPDATE" & vbnewline
  objLOG.write vbnewline & vbnewline & now & vbtab & " - EXECUTING PME SERVICE UPDATE" & vbnewline
  call HOOK("cmd.exe /C " & chr(34) & "c:\temp\PMEService\PMESetup.exe" & chr(34) & " /verysilent")
end if
''RESTART WINDOWS PROBE SERVICES
objOUT.write vbnewline & vbnewline & now & vbtab & " - RESTARTING WINDOWS PROBE SERVICES" & vbnewline
objLOG.write vbnewline & vbnewline & now & vbtab & " - RESTARTING WINDOWS PROBE SERVICES" & vbnewline
intRET = objWSH.run("sc query " & chr(34) & "Windows Software Probe Service" & chr(34), 0, true)
if (intRET = 0) then
  call HOOK("net start " & chr(34) & "N-able Patch Repository Service" & chr(34))
  call HOOK("net start " & chr(34) & "Windows Software Probe Maintenance Service" & chr(34))
  call HOOK("net start " & chr(34) & "Windows Software Probe Service" & chr(34))
end if
wscript.sleep 10000
'call HOOK("wuauclt /resetauthorization /detectnow")
'call HOOK("cmd.exe /C " & chr(34) & "PowerShell.exe (New-Object -ComObject Microsoft.Update.AutoUpdate).DetectNow()" & chr(34))
''END SCRIPT
call CLEANUP()
''END SCRIPT
''------------

sub CHKAU()																					        ''CHECK FOR SCRIPT UPDATE, PMESERVICE_INSTALL.VBS, REF #2
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
        objOUT.write vbnewline & now & vbtab & " - PMESERVICE_INSTALL :  " & strVER & " : GitHub : " & objSCR.text & vbnewline
        objLOG.write vbnewline & now & vbtab & " - PMESERVICE_INSTALL :  " & strVER & " : GitHub : " & objSCR.text & vbnewline
				if (cint(objSCR.text) > cint(strVER)) then
					objOUT.write vbnewline & now & " - UPDATING " & objSCR.nodename & " : " & objSCR.text & vbnewline
					objLOG.write vbnewline & now & " - UPDATING " & objSCR.nodename & " : " & objSCR.text & vbnewline
					''DOWNLOAD LATEST VERSION OF SCRIPT
					call FILEDL("https://github.com/CW-Khristos/scripts/raw/dev/PMEService_Install.vbs", wscript.scriptname)
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
  if (errRET = 0) then         											        ''PMESERVICE_INSTALL COMPLETED SUCCESSFULLY
    objOUT.write vbnewline & "PMESERVICE_INSTALL SUCCESSFUL : " & now
  elseif (errRET <> 0) then    											        ''PMESERVICE_INSTALL FAILED
    objOUT.write vbnewline & "PMESERVICE_INSTALL FAILURE : " & now & " : " & errRET
    ''RAISE CUSTOMIZED ERROR CODE, ERROR CODE WILL BE DEFINE RESTOP NUMBER INDICATING WHICH SECTION FAILED
    call err.raise(vbObjectError + errRET, "PMESERVICE_INSTALL", "FAILURE")
  end if
  objOUT.write vbnewline & vbnewline & now & " - PMESERVICE_INSTALL COMPLETE" & vbnewline
  objLOG.write vbnewline & vbnewline & now & " - PMESERVICE_INSTALL COMPLETE" & vbnewline
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
''PMESERVICE_INSTALL.VBS
''SCRIPT IS DESIGNED TO DOWNLOAD AND EXECUTE PME SERVICE UPDATE
''THIS VERSION ONLY FOCUSES ON INSTALLING THE PME SERVICE AND SUPPORTING FILES
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
''ZIP ARCHIVE OPTIONS
intOPT = 16
''VERSION FOR SCRIPT UPDATE, PMESERVICE_INSTALL.VBS, REF #2 , REF #68 , REF #69
strVER = 3
strREPO = "scripts"
strBRCH = "master"
strDIR = vbnullstring
''STDIN / STDOUT
blnSUP = false
set objIN = wscript.stdin
set objOUT = wscript.stdout
set objARG = wscript.arguments
''OBJECTS FOR LOCATING FOLDERS
set objWSH = createobject("wscript.shell")
Set objAPP = createobject("shell.application")
set objFSO = createobject("scripting.filesystemobject")
''CHECK 'PERSISTENT' FOLDERS , REF #2 , REF #73
if (not (objFSO.folderexists("C:\IT\"))) then
  objFSO.createfolder("C:\IT\")
end if
if (not (objFSO.folderexists("C:\IT\Scripts\"))) then
  objFSO.createfolder("C:\IT\Scripts\")
end if
''PREPARE LOGFILE
if (objFSO.fileexists("C:\temp\PMESERVICE_INSTALL")) then       ''LOGFILE EXISTS
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
if (errRET = 0) then                                        ''NO ERRORS DURING INITIAL START
  objOUT.write vbnewline & vbnewline & now & " - STARTING PMESERVICE_INSTALL" & vbnewline
  objLOG.write vbnewline & vbnewline & now & " - STARTING PMESERVICE_INSTALL" & vbnewline
  ''AUTOMATIC UPDATE, PMESERVICE_INSTALL.VBS, REF #2 , REF #69 , REF #68 , FIXES #4
  ''DOWNLOAD CHKAU.VBS SCRIPT, REF #2 , REF #69 , REF #68
  call FILEDL("https://raw.githubusercontent.com/CW-Khristos/scripts/master/chkAU.vbs", "C:\IT\Scripts", "chkAU.vbs")
  ''EXECUTE CHKAU.VBS SCRIPT, REF #69
  objOUT.write vbnewline & now & vbtab & vbtab & " - CHECKING FOR UPDATE : PMESERVICE_INSTALL : " & strVER
  objLOG.write vbnewline & now & vbtab & vbtab & " - CHECKING FOR UPDATE : PMESERVICE_INSTALL : " & strVER
  intRET = objWSH.run ("cmd.exe /C " & chr(34) & "cscript.exe " & chr(34) & "C:\IT\Scripts\chkAU.vbs" & chr(34) & " " & _
    chr(34) & strREPO & chr(34) & " " & chr(34) & strBRCH & chr(34) & " " & chr(34) & strDIR & chr(34) & " " & _
    chr(34) & wscript.scriptname & chr(34) & " " & chr(34) & strVER & chr(34) & chr(34), 0, true)
  ''CHKAU RETURNED - NO UPDATE FOUND , REF #2 , REF #69 , REF #68
  objOUT.write vbnewline & "errRET='" & intRET & "'"
  objLOG.write vbnewline & "errRET='" & intRET & "'"
  intRET = (intRET - vbObjectError)
  objOUT.write vbnewline & "errRET='" & intRET & "'"
  objLOG.write vbnewline & "errRET='" & intRET & "'"
  if ((intRET = 4) or (intRET = 10) or (intRET = 11) or (intRET = 1) or (intRET = 2147221505) or (intRET = 2147221517)) then
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
    'call FILEDL("https://raw.githubusercontent.com/CW-Khristos/scripts/dev/CClutterV2.vbs", "C:\IT\Scripts", "CClutterV2.vbs")
    'call HOOK("cscript.exe " & chr(34) & "C:\IT\Scripts\CClutterV2.vbs" & chr(34) & " " & chr(34) & "true" & chr(34))
    ''REMOVE POSSIBLE TRASHED 'ARCHIVES'
    if (objFSO.fileexists("%ProgramData%\SolarWinds MSP\PME\Archives")) then
      objFSO.deletefile "%ProgramData%\SolarWinds MSP\PME\Archives", true
    end if
    if (not (objFSO.folderexists("%ProgramData%\SolarWinds MSP\SolarWinds.MSP.CacheService"))) then
      call HOOK("cmd.exe /C rd /s /q " & chr(34) & "%ProgramData%\SolarWinds MSP\SolarWinds.MSP.CacheService" & chr(34))
    end if
    if (not (objFSO.folderexists("%ProgramData%\SolarWinds MSP\SolarWinds.MSP.PME.Agent.PmeService"))) then
      call HOOK("cmd.exe /C rd /s /q " & chr(34) & "%ProgramData%\SolarWinds MSP\SolarWinds.MSP.PME.Agent.PmeService" & chr(34))
    end if
    if (not (objFSO.folderexists("%ProgramData%\SolarWinds MSP\SolarWinds.MSP.RPCServerService"))) then
      call HOOK("cmd.exe /C rd /s /q " & chr(34) & "%ProgramData%\SolarWinds MSP\SolarWinds.MSP.RPCServerService" & chr(34))
    end if
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
    call FILEDL("https://raw.githubusercontent.com/CW-Khristos/scripts/master/PMEService.zip", "C:\IT", "PMEService.zip")
    wscript.sleep 5000
    ''DOWNLOAD SUPPORTING FILES
    if (not objFSO.fileexists("C:\IT\PMEService.zip")) then
      call FILEDL("https://raw.githubusercontent.com/CW-Khristos/scripts/master/PMEService.zip", "C:\IT", "PMEService.zip")
      wscript.sleep 10000
    end if
    if (objFSO.fileexists("C:\IT\PMEService.zip")) then
      ''EXTRACT PME SERVICE SUPPORTING FILES
      objOUT.write vbnewline & vbnewline & now & vbtab & " - EXTRACTING PME SERVICE SUPPORTING FILES" & vbnewline
      objLOG.write vbnewline & vbnewline & now & vbtab & " - EXTRACTING PME SERVICE SUPPORTING FILES" & vbnewline
      set objSRC = objAPP.namespace("C:\IT\PMEService.zip").items()
      set objTGT = objAPP.namespace("C:\IT")
      objTGT.copyhere objSRC, intOPT
    end if
    ''CHECK FOR EXTRACTED X.ROBOT
    if (objFSO.fileexists("C:\IT\AnniversaryUpdates_details.xml")) then
      ''COPY PME SERVICE SUPPORTING FILES TO 'C:\PROGRAMDATA\SOLARWINDS MSP\PME\ARCHIVES'
      objOUT.write vbnewline & vbnewline & now & vbtab & " - MOVING ANNIVERSARYUPDATES_DETAILS.XML" & vbnewline
      objLOG.write vbnewline & vbnewline & now & vbtab & " - MOVING ANNIVERSARYUPDATES_DETAILS.XML" & vbnewline
      call HOOK("cmd.exe /C copy /y " & chr(34) & "C:\IT\AnniversaryUpdates_details.xml" & chr(34) & " " & chr(34) & "C:\ProgramData\SolarWinds MSP\PME\Archives" & chr(34))
    end if
    ''CHECK FOR EXTRACTED X.ROBOT
    if (objFSO.fileexists("C:\IT\AnniversaryUpdates.zip")) then
      ''COPY PME SERVICE SUPPORTING FILES TO 'C:\PROGRAMDATA\SOLARWINDS MSP\PME\ARCHIVES'
      objOUT.write vbnewline & vbnewline & now & vbtab & " - MOVING ANNIVERSARYUPDATES.ZIP" & vbnewline
      objLOG.write vbnewline & vbnewline & now & vbtab & " - MOVING ANNIVERSARYUPDATES.ZIP" & vbnewline
      call HOOK("cmd.exe /C copy /y " & chr(34) & "C:\IT\AnniversaryUpdates.zip" & chr(34) & " " & chr(34) & "C:\ProgramData\SolarWinds MSP\PME\Archives" & chr(34))
    end if
    ''CHECK FOR EXTRACTED X.ROBOT
    if (objFSO.fileexists("C:\IT\SecurityUpdates_details.xml")) then
      ''COPY PME SERVICE SUPPORTING FILES TO 'C:\PROGRAMDATA\SOLARWINDS MSP\PME\ARCHIVES'
      objOUT.write vbnewline & vbnewline & now & vbtab & " - MOVING SECURITYUPDATES_DETAILS.XML" & vbnewline
      objLOG.write vbnewline & vbnewline & now & vbtab & " - MOVING SECURITYUPDATES_DETAILS.XML" & vbnewline
      call HOOK("cmd.exe /C copy /y " & chr(34) & "C:\IT\SecurityUpdates_details.xml" & chr(34) & " " & chr(34) & "C:\ProgramData\SolarWinds MSP\PME\Archives" & chr(34))
    end if
    ''CHECK FOR EXTRACTED X.ROBOT
    if (objFSO.fileexists("C:\IT\SecurityUpdates-2020.2.11.20.zip")) then
      ''COPY PME SERVICE SUPPORTING FILES TO 'C:\PROGRAMDATA\SOLARWINDS MSP\PME\ARCHIVES'
      objOUT.write vbnewline & vbnewline & now & vbtab & " - MOVING SECURITYUPDATES.ZIP" & vbnewline
      objLOG.write vbnewline & vbnewline & now & vbtab & " - MOVING SECURITYUPDATES.ZIP" & vbnewline
      call HOOK("cmd.exe /C copy /y " & chr(34) & "C:\IT\SecurityUpdates-2020.2.11.20.zip" & chr(34) & " " & chr(34) & "C:\ProgramData\SolarWinds MSP\PME\Archives" & chr(34))
    end if
    if (objFSO.fileexists("C:\IT\SecurityUpdates-2020.4.24.4.zip")) then
      ''COPY PME SERVICE SUPPORTING FILES TO 'C:\PROGRAMDATA\SOLARWINDS MSP\PME\ARCHIVES'
      objOUT.write vbnewline & vbnewline & now & vbtab & " - MOVING SECURITYUPDATES.ZIP" & vbnewline
      objLOG.write vbnewline & vbnewline & now & vbtab & " - MOVING SECURITYUPDATES.ZIP" & vbnewline
      call HOOK("cmd.exe /C copy /y " & chr(34) & "C:\IT\SecurityUpdates-2020.4.24.4.zip" & chr(34) & " " & chr(34) & "C:\ProgramData\SolarWinds MSP\PME\Archives" & chr(34))
    end if
    ''RUN PME SERVICE UPDATE WITH /VERYSILENT SWITCH
    if (objFSO.fileexists("C:\IT\PMESetup.exe")) then
      objOUT.write vbnewline & vbnewline & now & vbtab & " - EXECUTING PME SERVICE UPDATE" & vbnewline
      objLOG.write vbnewline & vbnewline & now & vbtab & " - EXECUTING PME SERVICE UPDATE" & vbnewline
      call HOOK("cmd.exe /C " & chr(34) & "C:\IT\PMESetup.exe" & chr(34) & " /verysilent /log=" & chr(34) & "C:\temp\PMESetup.log" & chr(34))
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
  end if
elseif (errRET <> 0) then                                   ''ERRORS ENCOUNTERED DURING INITIAL START
  call LOGERR(errRET)
end if
''END SCRIPT
call CLEANUP()
''END SCRIPT
''------------

''SUB-ROUTINES
sub FILEDL(strURL, strDL, strFILE)                          ''CALL HOOK TO DOWNLOAD FILE FROM URL , 'ERRRET'=11
  strSAV = vbnullstring
  ''SET DOWNLOAD PATH
  strSAV = strDL & "\" & strFILE
  objOUT.write vbnewline & now & vbtab & vbtab & vbtab & "HTTPDOWNLOAD-------------DOWNLOAD : " & strURL & " : SAVE AS :  " & strSAV
  objLOG.write vbnewline & now & vbtab & vbtab & vbtab & "HTTPDOWNLOAD-------------DOWNLOAD : " & strURL & " : SAVE AS :  " & strSAV
  ''CHECK IF FILE ALREADY EXISTS
  if objFSO.fileexists(strSAV) then
    ''DELETE FILE FOR OVERWRITE
    objFSO.deletefile(strSAV)
  end if
  ''CREATE HTTP OBJECT
  set objHTTP = createobject( "WinHttp.WinHttpRequest.5.1" )
  ''DOWNLOAD FROM URL
  objHTTP.open "GET", strURL, false
  objHTTP.send
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
    call LOGERR(11)
    err.clear
  end if
end sub

sub HOOK(strCMD)                                            ''CALL HOOK TO MONITOR OUTPUT OF CALLED COMMAND , 'ERRRET'=12
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
    call LOGERR(12)
  end if
end sub

sub LOGERR(intSTG)                                          ''CALL HOOK TO MONITOR OUTPUT OF CALLED COMMAND
  errRET = intSTG
  if (err.number <> 0) then
    objOUT.write vbnewline & now & vbtab & vbtab & vbtab & err.number & vbtab & err.description & vbnewline
    objLOG.write vbnewline & now & vbtab & vbtab & vbtab & err.number & vbtab & err.description & vbnewline
		err.clear
  end if
  ''CUSTOM ERROR CODES
  select case intSTG
    case 1                                                  '' 'ERRRET'=1 - NOT ENOUGH ARGUMENTS
      objOUT.write vbnewline & vbnewline & now & vbtab & " - NOT ENOUGH ARGUMENTS"
      objLOG.write vbnewline & vbnewline & now & vbtab & " - NOT ENOUGH ARGUMENTS"
  end select
end sub

sub CLEANUP()                                 			        ''SCRIPT CLEANUP
  on error resume next
  if (errRET = 0) then         											        ''PMESERVICE_INSTALL COMPLETED SUCCESSFULLY
    objOUT.write vbnewline & vbnewline & now & vbtab & " - PMESERVICE_INSTALL SUCCESSFUL : " & now
    objLOG.write vbnewline & vbnewline & now & vbtab & " - PMESERVICE_INSTALL SUCCESSFUL : " & now
    err.clear
  elseif (errRET <> 0) then    											        ''PMESERVICE_INSTALL FAILED
    objOUT.write vbnewline & vbnewline & now & vbtab & " - PMESERVICE_INSTALL FAILURE : " & errRET & " : " & now
    objLOG.write vbnewline & vbnewline & now & vbtab & " - PMESERVICE_INSTALL FAILURE : " & errRET & " : " & now
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
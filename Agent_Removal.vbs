''AGENT_REMOVAL.VBS
''DESIGNED TO AUTOMATICALLY UNINSTALL WINDOWS AGENT SILENTLY
''WRITTEN BY : CJ BLEDSOE / CBLEDSOE<@>IPMCOMPUTERS.COM
on error resume next
''SCRIPT VARIABLES
dim errRET, strVER
''SCRIPT OBJECTS
dim objIN, objOUT, objARG, objWSH, objFSO
dim objLOG, objEXEC, objHOOK, objHTTP, objXML
''VERSION FOR SCRIPT UPDATE, AGENT_REMOVAL.VBS, REF #2 , REF #68 , REF #69 , FIXES #21 , FIXES #31
strVER = 1
strREPO = "scripts"
strBRCH = "dev"
strDIR = vbnullstring
''DEFAULT SUCCESS
errRET = 0
''STDIN / STDOUT
set objIN = wscript.stdin
set objOUT = wscript.stdout
set objARG = wscript.arguments
''OBJECTS FOR LOCATING FOLDERS
set objWSH = createobject("wscript.shell")
set objFSO = createobject("scripting.filesystemobject")
''ENVIRONMENT VARIABLES
strPD = objWSH.expandenvironmentstrings("%PROGRAMDATA%")
strPF = objWSH.expandenvironmentstrings("%PROGRAMFILES%")
strPF86 = objWSH.expandenvironmentstrings("%PROGRAMFILES(X86)%")
''CHECK 'PERSISTENT' FOLDERS , REF #2 , REF #73
if (not (objFSO.folderexists("c:\temp"))) then
  objFSO.createfolder("c:\temp")
end if
if (not (objFSO.folderexists("C:\IT\"))) then
  objFSO.createfolder("C:\IT\")
end if
if (not (objFSO.folderexists("C:\IT\Scripts\"))) then
  objFSO.createfolder("C:\IT\Scripts\")
end if
''PREPARE LOGFILE
if (objFSO.fileexists("C:\temp\AGENT_REMOVAL")) then          ''LOGFILE EXISTS
  objFSO.deletefile "C:\temp\AGENT_REMOVAL", true
  set objLOG = objFSO.createtextfile("C:\temp\AGENT_REMOVAL")
  objLOG.close
  set objLOG = objFSO.opentextfile("C:\temp\AGENT_REMOVAL", 8)
else                                                          ''LOGFILE NEEDS TO BE CREATED
  set objLOG = objFSO.createtextfile("C:\temp\AGENT_REMOVAL")
  objLOG.close
  set objLOG = objFSO.opentextfile("C:\temp\AGENT_REMOVAL", 8)
end if
''READ PASSED COMMANDLINE ARGUMENTS
if (wscript.arguments.count > 0) then                         ''ARGUMENTS WERE PASSED
  ''ARGUMENT OUTPUT DISABLED TO SANITIZE
  'for x = 0 to (wscript.arguments.count - 1)
  '  objOUT.write vbnewline & now & vbtab & " - ARGUMENT " & (x + 1) & " (ITEM " & x & ") " & " PASSED : " & ucase(objARG.item(x))
  '  objLOG.write vbnewline & now & vbtab & " - ARGUMENT " & (x + 1) & " (ITEM " & x & ") " & " PASSED : " & ucase(objARG.item(x))
  'next 
  if (wscript.arguments.count >= 1) then                      ''SET VARIABLES ACCEPTING ARGUMENTS
  end if
elseif (wscript.arguments.count < 1) then                     ''NOT ENOUGH ARGUMENTS PASSED , END SCRIPT , 'ERRRET'=1
  'call LOGERR(1)
  'call CLEANUP()
end if

''------------
''BEGIN SCRIPT
if (errRET = 0) then
  ''DETERMINE OS ARCHITECTURE
  if (GetOSbits = 64) then
    strPF = strPF86
  elseif (GetOSbits = 32) then
    strPF = strPF
  end if
  objOUT.write vbnewline & vbnewline & now & vbtab & " - EXECUTING AGENT_REMOVAL"
  objLOG.write vbnewline & vbnewline & now & vbtab & " - EXECUTING AGENT_REMOVAL"
  ''STOP SERVICES
  objOUT.write vbnewline & now & vbtab & vbtab & " - STOPPING WINDOWS AGENT SERVICES"
  objLOG.write vbnewline & now & vbtab & vbtab & " - STOPPING WINDOWS AGENT SERVICES"
  call HOOK("sc stop " & chr(34) & "Windows Agent Service" & chr(34))
  call HOOK("sc stop " & chr(34) & "Windows Agent Maintenance Service" & chr(34))
  ''KILL SERVICE PROCESSES
  objOUT.write vbnewline & now & vbtab & vbtab & " - STOPPING WINDOWS AGENT PROCESSES"
  objLOG.write vbnewline & now & vbtab & vbtab & " - STOPPING WINDOWS AGENT PROCESSES"
  call HOOK("taskkill /F /IM agent.exe /T")
  call HOOK("taskkill /F /IM AgentMaint.exe /T")
  ''UNINSTALL WINDOWS AGENT 2021.1.2.391
  objOUT.write vbnewline & now & vbtab & vbtab & " - UNINSTALLING WINDOWS AGENT"
  objLOG.write vbnewline & now & vbtab & vbtab & " - UNINSTALLING WINDOWS AGENT"
  call HOOK("MsiExec.exe /X {1D35A03E-E581-4838-9EE3-244DBBF51415} /qn")
  ''CLEANUP REGISTRY KEYS
  objOUT.write vbnewline & now & vbtab & vbtab & " - REMOVING REGISTRY KEYS"
  objLOG.write vbnewline & now & vbtab & vbtab & " - REMOVING REGISTRY KEYS"
  call HOOK("reg delete " & chr(34) & "HKLM\SOFTWARE\N-able Technologies\NcentralAsset" & chr(34) & " /f")
  call HOOK("reg delete " & chr(34) & "HKLM\SOFTWARE\N-able Technologies" & chr(34) & " /f")
  ''REMOVE SERVICES
  objOUT.write vbnewline & now & vbtab & vbtab & " - REMOVING AGENT SERVICES"
  objLOG.write vbnewline & now & vbtab & vbtab & " - REMOVING AGENT SERVICES"
  call HOOK("sc delete " & chr(34) & "Windows Agent Service" & chr(34))
  call HOOK("sc delete " & chr(34) & "Windows Agent Maintenance Service" & chr(34))
  ''CLEAR PROGRAM FILES / PROGRAM FILES (X86) FOLDER
  if (objFSO.fileexists(strPF & "\N-Able Technologies")) then
    objOUT.write vbnewline & now & vbtab & vbtab & " - REMOVING PROGRAM FILES\N-ABLE TECHNOLOGIES DRIECTORY"
    objLOG.write vbnewline & now & vbtab & vbtab & " - REMOVING PROGRAM FILES\N-ABLE TECHNOLOGIES DRIECTORY"
    objFSO.deletefolder chr(34) & strPF & "\N-Able Technologies" & chr(34), true
  end if
  ''CLEAR PROGRAMDATA FOLDER
  if (objFSO.fileexists(strPD & "\N-Able Technologies")) then
    objOUT.write vbnewline & now & vbtab & vbtab & " - REMOVING PROGRAMDATA\N-ABLE TECHNOLOGIES DRIECTORY"
    objLOG.write vbnewline & now & vbtab & vbtab & " - REMOVING PROGRAMDATA\N-ABLE TECHNOLOGIES DRIECTORY"
    objFSO.deletefolder chr(34) & strPD & "\N-Able Technologies" & chr(34), true
  end if
elseif (errRET <> 0) then
  call LOGERR(errRET)
end if
''END SCRIPT
call CLEANUP()
''END SCRIPT
''------------

''FUNCTIONS
function GetOSbits()
   if (objWSH.ExpandEnvironmentStrings("%PROCESSOR_ARCHITECTURE%") = "AMD64") then
      GetOSbits = 64
   else
      GetOSbits = 32
   end if
end function

''SUB-ROUTINES
sub FILEDL(strURL, strDL, strFILE)                            ''CALL HOOK TO DOWNLOAD FILE FROM URL , 'ERRRET'=11
  strSAV = vbnullstring
  ''SET DOWNLOAD PATH
  strSAV = strDL & "\" & strFILE
  objOUT.write vbnewline & now & vbtab & vbtab & "HTTPDOWNLOAD-------------DOWNLOAD : " & strURL & " : SAVE AS :  " & strSAV
  objLOG.write vbnewline & now & vbtab & vbtab & "HTTPDOWNLOAD-------------DOWNLOAD : " & strURL & " : SAVE AS :  " & strSAV
  ''ADD WINHTTP SECURE CHANNEL TLS REGISTRY KEYS
  call HOOK("reg add " & chr(34) & "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\WinHttp" & chr(34) & _
    " /f /v DefaultSecureProtocols /t REG_DWORD /d 0x00000A00 /reg:32")
  call HOOK("reg add " & chr(34) & "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\WinHttp" & chr(34) & _
    " /f /v DefaultSecureProtocols /t REG_DWORD /d 0x00000A00 /reg:64")
  ''CHECK IF FILE ALREADY EXISTS
  if objFSO.fileexists(strSAV) then
    ''DELETE FILE FOR OVERWRITE
    objFSO.deletefile(strSAV)
  end if
  ''CREATE HTTP OBJECT
  set objHTTP = createobject("WinHttp.WinHttpRequest.5.1")
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
    objOUT.write vbnewline & vbnewline & now & vbtab & " - DOWNLOAD : " & strSAV & " : SUCCESSFUL"
    objLOG.write vbnewline & vbnewline & now & vbtab & " - DOWNLOAD : " & strSAV & " : SUCCESSFUL"
  end if
	set objHTTP = nothing
  if ((err.number <> 0) and (err.number <> 58)) then          ''ERROR RETURNED DURING UPDATE CHECK , 'ERRRET'=11
    call LOGERR(11)
  end if
end sub

sub HOOK(strCMD)                                              ''CALL HOOK TO MONITOR OUTPUT OF CALLED COMMAND , 'ERRRET'=12
  on error resume next
  objOUT.write vbnewline & now & vbtab & vbtab & " - EXECUTING : HOOK(" & strCMD & ")"
  objLOG.write vbnewline & now & vbtab & vbtab & " - EXECUTING : HOOK(" & strCMD & ")"
  set objHOOK = objWSH.exec(strCMD)
  while (not objHOOK.stdout.atendofstream)
    if (instr(1, strCMD, "takeown /F ") = 0) then             ''SUPPRESS 'TAKEOWN' SUCCESS MESSAGES
      strIN = objHOOK.stdout.readline
      if (strIN <> vbnullstring) then
        objOUT.write vbnewline & now & vbtab & vbtab & vbtab & strIN 
        objLOG.write vbnewline & now & vbtab & vbtab & vbtab & strIN 
      end if
    end if
  wend
  wscript.sleep 10
  if (instr(1, strCMD, "takeown /F ") = 0) then               ''SUPPRESS 'TAKEOWN' SUCCESS MESSAGES
    strIN = objHOOK.stdout.readall
    if (strIN <> vbnullstring) then
      objOUT.write vbnewline & now & vbtab & vbtab & vbtab & strIN 
      objLOG.write vbnewline & now & vbtab & vbtab & vbtab & strIN 
    end if
  end if
  set objHOOK = nothing
  if (err.number <> 0) then                                   ''ERROR RETURNED , 'ERRRET'=12
    call LOGERR(12)
  end if
end sub

sub LOGERR(intSTG)                                            ''CALL HOOK TO MONITOR OUTPUT OF CALLED COMMAND
  errRET = intSTG
  if (err.number <> 0) then
    objOUT.write vbnewline & now & vbtab & vbtab & vbtab & err.number & vbtab & err.description & vbnewline
    objLOG.write vbnewline & now & vbtab & vbtab & vbtab & err.number & vbtab & err.description & vbnewline
		err.clear
  end if
  select case intSTG
    case 1                                                    ''NOT ENOUGH ARGUMENTS , 'ERRRET'=1
      objOUT.write vbnewline & vbnewline & now & vbtab & " - NOT ENOUGH ARGUMENTS PASSED"
      objLOG.write vbnewline & vbnewline & now & vbtab & " - NOT ENOUGH ARGUMENTS PASSED"
  end select
end sub

sub CLEANUP()                                                 ''SCRIPT CLEANUP
  on error resume next
  if (errRET = 0) then         															  ''AGENT_REMOVAL COMPLETED SUCCESSFULLY
    objOUT.write vbnewline & vbnewline & now & vbtab & " - AGENT_REMOVAL SUCCESSFUL : " & errRET & " : " & now
    objLOG.write vbnewline & vbnewline & now & vbtab & " - AGENT_REMOVAL SUCCESSFUL : " & errRET & " : " & now
    err.clear
  elseif (errRET <> 0) then    															  ''AGENT_REMOVAL FAILED
    objOUT.write vbnewline & vbnewline & now & vbtab & " - AGENT_REMOVAL FAILURE : " & errRET & " : " & now
    objLOG.write vbnewline & vbnewline & now & vbtab & " - AGENT_REMOVAL FAILURE : " & errRET & " : " & now
    ''RAISE CUSTOMIZED ERROR CODE, ERROR CODE WILL BE DEFINE RESTOP NUMBER INDICATING WHICH SECTION FAILED
    call err.raise(vbObjectError + errRET, "AGENT_REMOVAL", "FAILURE")
  end if
  objOUT.write vbnewline & vbnewline & now & " - AGENT_REMOVAL COMPLETE" & vbnewline
  objLOG.write vbnewline & vbnewline & now & " - AGENT_REMOVAL COMPLETE" & vbnewline
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
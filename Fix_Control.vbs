''FIX_CONTROL.VBS
''DESIGNED TO FIX TAKE CONTROL
''WRITTEN BY : CJ BLEDSOE / CBLEDSOE<@>IPMCOMPUTERS.COM
on error resume next
''SCRIPT VARIABLES
dim strREPO, strBRCH, strDIR
dim errRET, strVER, strSEL, strIN
''VARIABLES ACCEPTING PARAMETERS
''SCRIPT OBJECTS
dim objIN, objOUT, objARG, objWSH, objFSO
dim objLOG, objEXEC, objHOOK, objSIN, objSOUT
''VERSION FOR SCRIPT UPDATE, FIX_CONTROL.VBS, REF #2 , REF #68 , REF #69 , FIXES #21 , FIXES #31
strVER = 1
strREPO = "scripts"
strBRCH = "master"
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
if (objFSO.fileexists("C:\temp\FIX_CONTROL")) then              ''LOGFILE EXISTS
  objFSO.deletefile "C:\temp\FIX_CONTROL", true
  set objLOG = objFSO.createtextfile("C:\temp\FIX_CONTROL")
  objLOG.close
  set objLOG = objFSO.opentextfile("C:\temp\FIX_CONTROL", 8)
else                                                        ''LOGFILE NEEDS TO BE CREATED
  set objLOG = objFSO.createtextfile("C:\temp\FIX_CONTROL")
  objLOG.close
  set objLOG = objFSO.opentextfile("C:\temp\FIX_CONTROL", 8)
end if
''READ PASSED COMMANDLINE ARGUMENTS
if (wscript.arguments.count > 0) then                       ''ARGUMENTS WERE PASSED
  ''ARGUMENT OUTPUT DISABLED TO SANITIZE
  'for x = 0 to (wscript.arguments.count - 1)
  '  objOUT.write vbnewline & now & vbtab & " - ARGUMENT " & (x + 1) & " (ITEM " & x & ") " & " PASSED : " & ucase(objARG.item(x))
  '  objLOG.write vbnewline & now & vbtab & " - ARGUMENT " & (x + 1) & " (ITEM " & x & ") " & " PASSED : " & ucase(objARG.item(x))
  'next 
  if (wscript.arguments.count > 1) then                     ''REQUIRED ARGUMENTS PASSED
  else                                                      ''NOT ENOUGH ARGUMENTS PASSED , END SCRIPT , 'ERRRET'=1
    'call LOGERR(1)
  end if
elseif (wscript.arguments.count = 0) then                   ''NO ARGUMENTS PASSED , END SCRIPT , 'ERRRET'=1
  'call LOGERR(1)
end if

''------------
''BEGIN SCRIPT
if (errRET = 0) then                                        ''NO ERRORS DURING INITIAL START
  ''DETERMINE OS ARCHITECTURE
  if (GetOSbits = 64) then
    strPF = strPF86
  elseif (GetOSbits = 32) then
    strPF = strPF
  end if
  objOUT.write vbnewline & vbnewline & now & vbtab & " - EXECUTING FIX_CONTROL"
  objLOG.write vbnewline & vbnewline & now & vbtab & " - EXECUTING FIX_CONTROL"
	''AUTOMATIC UPDATE, FIX_CONTROL.VBS, REF #2 , REF #69 , REF #68
  ''DOWNLOAD CHKAU.VBS SCRIPT, REF #2 , REF #69 , REF #68
  'call FILEDL("https://raw.githubusercontent.com/CW-Khristos/scripts/master/chkAU.vbs", "C:\IT\Scripts", "chkAU.vbs")
  ''EXECUTE CHKAU.VBS SCRIPT, REF #69
  objOUT.write vbnewline & now & vbtab & vbtab & " - CHECKING FOR UPDATE : FIX_CONTROL : " & strVER
  objLOG.write vbnewline & now & vbtab & vbtab & " - CHECKING FOR UPDATE : FIX_CONTROL : " & strVER
  'intRET = objWSH.run ("cmd.exe /C " & chr(34) & "cscript.exe " & chr(34) & "C:\IT\Scripts\chkAU.vbs" & chr(34) & " " & _
  '  chr(34) & strREPO & chr(34) & " " & chr(34) & strBRCH & chr(34) & " " & chr(34) & strDIR & chr(34) & " " & _
  '  chr(34) & wscript.scriptname & chr(34) & " " & chr(34) & strVER & chr(34) & chr(34), 0, true)
  '''CHKAU RETURNED - NO UPDATE FOUND , REF #2 , REF #69 , REF #68
  'objOUT.write vbnewline & "errRET='" & intRET & "'"
  'objLOG.write vbnewline & "errRET='" & intRET & "'"
  'intRET = (intRET - vbObjectError)
  'objOUT.write vbnewline & "errRET='" & intRET & "'"
  'objLOG.write vbnewline & "errRET='" & intRET & "'"
  intRET = 4
  if ((intRET = 4) or (intRET = 10) or (intRET = 11) or (intRET = 1) or (intRET = 2147221505) or (intRET = 2147221517)) then
    ''STOP TAKE CONTROL SERVICES
    objOUT.write vbnewline & now & vbtab & vbtab & " - STOPPING TAKE CONTROL SERVICES"
    objLOG.write vbnewline & now & vbtab & vbtab & " - STOPPING TAKE CONTROL SERVICES"
    call HOOK("sc stop " & chr(34) & "BASupportExpressSrvcUpdater_N_Central" & chr(34))
    call HOOK("sc stop " & chr(34) & "BASupportExpressStandaloneService_N_Central" & chr(34))
    ''KILL SERVICE PROCESSES
    objOUT.write vbnewline & now & vbtab & vbtab & " - STOPPING TAKE CONTROL PROCESSES"
    objLOG.write vbnewline & now & vbtab & vbtab & " - STOPPING TAKE CONTROL PROCESSES"
    call HOOK("taskkill /F /IM BASupSrvc.exe /T")
    call HOOK("taskkill /F /IM BASupSysInf.exe /T")
    call HOOK("taskkill /F /IM BASupSrvcCnfg.exe /T")
    call HOOK("taskkill /F /IM BASupSrvcUpdater.exe /T")
    ''GetSupportService
    if (objFSO.folderexists(strPF & "\BeAnywhere Support Express\GetSupportService")) then
      ''RUN UNINSTALL
      objOUT.write vbnewline & now & vbtab & vbtab & " - UNINSTALLING TAKE CONTROL"
      objLOG.write vbnewline & now & vbtab & vbtab & " - UNINSTALLING TAKE CONTROL"
      if (objFSO.fileexists(strPF & "\BeAnywhere Support Express\GetSupportService\uninstall.exe")) then
        'call HOOK(strPF & "\BeAnywhere Support Express\GetSupportService\uninstall.exe" & " /S")
        objWSH.run chr(34) & strPF & "\BeAnywhere Support Express\GetSupportService\uninstall.exe" & chr(34) & " /S", 0, true
      end if
      ''REMOVE DIRECTORY
      wscript.sleep 10000
      objOUT.write vbnewline & now & vbtab & vbtab & " - REMOVING TAKE CONTROL DIRECTORY"
      objLOG.write vbnewline & now & vbtab & vbtab & " - REMOVING TAKE CONTROL DIRECTORY"
      ''CALL HOOK(RMDIR) GETTING STUCK DUE TO PROCESSES IN DIRECTORY
      'call HOOK("rmdir /s /q " & chr(34) & strPF & "\BeAnywhere Support Express\GetSupportService" & chr(34))
      intRET = objWSH.run ("rmdir /s /q " & chr(34) & strPF & "\BeAnywhere Support Express\GetSupportService" & chr(34), 0, false)
      if (intRET <> 0) then
        for intLOOP = 0 to 10
          wscript.sleep 5000
          intRET = objWSH.run ("rmdir /s /q " & chr(34) & strPF & "\BeAnywhere Support Express\GetSupportService" & chr(34), 0, false)
          if (intRET = 0) then
            exit for
          end if
        next
      end if
      if (err.number <> 0) then
        call LOGERR(1)
      end if
    ''GetSupportService_N-Central
    elseif (objFSO.folderexists(strPF & "\BeAnywhere Support Express\GetSupportService_N-Central")) then
      ''RUN UNINSTALL
      objOUT.write vbnewline & now & vbtab & vbtab & " - UNINSTALLING TAKE CONTROL"
      objLOG.write vbnewline & now & vbtab & vbtab & " - UNINSTALLING TAKE CONTROL"
      if (objFSO.fileexists(strPF & "\BeAnywhere Support Express\GetSupportService_N-Central\uninstall.exe")) then
        'call HOOK(strPF86 & "\BeAnywhere Support Express\GetSupportService_N-Central\uninstall.exe" & " /S")
        objWSH.run chr(34) & strPF & "\BeAnywhere Support Express\GetSupportService_N-Central\uninstall.exe" & chr(34) & " /S", 0, true
      end if
      ''REMOVE DIRECTORY
      wscript.sleep 10000
      objOUT.write vbnewline & now & vbtab & vbtab & " - REMOVING TAKE CONTROL DIRECTORY"
      objLOG.write vbnewline & now & vbtab & vbtab & " - REMOVING TAKE CONTROL DIRECTORY"
      ''CALL HOOK(RMDIR) GETTING STUCK DUE TO PROCESSES IN DIRECTORY
      'call HOOK("rmdir /s /q " & strPF86 & "\BeAnywhere Support Express\GetSupportService_N-Central")
      intRET = objWSH.run ("rmdir /s /q " & chr(34) & strPF86 & "\BeAnywhere Support Express\GetSupportService_N-Central" & chr(34), 0, false)
      if (intRET <> 0) then
        for intLOOP = 0 to 10
          wscript.sleep 5000
          intRET = objWSH.run ("rmdir /s /q " & chr(34) & strPF86 & "\BeAnywhere Support Express\GetSupportService_N-Central" & chr(34), 0, false)
          if (intRET = 0) then
            exit for
          end if
        next
      end if
      if (err.number <> 0) then
        call LOGERR(2)
      end if
    end if
    ''CALL HOOK(RMDIR) GETTING STUCK DUE TO PROCESSES IN DIRECTORY
    'call HOOK("rmdir /s /q " & chr(34) & strPF & "\BeAnywhere Support Express" & chr(34))
    intRET = objWSH.run ("rmdir /s /q " & chr(34) & strPF & "\BeAnywhere Support Express" & chr(34), 0, false)
    if (intRET <> 0) then
      for intLOOP = 0 to 10
        wscript.sleep 5000
        intRET = objWSH.run ("rmdir /s /q " & chr(34) & strPF & "\BeAnywhere Support Express" & chr(34), 0, false)
        if (intRET = 0) then
          exit for
        end if
      next
    end if
    if (err.number <> 0) then
      call LOGERR(3)
    end if
    ''PROGRAMDATA DIRECTORY
    if (objFSO.folderexists(chr(34) & strPD & "\GetSupportService_N-Central" & chr(34))) then
      objOUT.write vbnewline & now & vbtab & vbtab & " - REMOVING " & chr(34) & ucase(strPD & "\GetSupportService_N-Central") & chr(34)
      objLOG.write vbnewline & now & vbtab & vbtab & " - REMOVING " & chr(34) & ucase(strPD & "\GetSupportService_N-Central") & chr(34)
      call HOOK("rmdir /s /q " & chr(34) & strPD & "\GetSupportService_N-Central" & chr(34))
      if (err.number <> 0) then
        call LOGERR(4)
      end if
    end if
    if (objFSO.folderexists(chr(34) & strPD & "\GetSupportService" & chr(34))) then
      objOUT.write vbnewline & now & vbtab & vbtab & " - REMOVING " & chr(34) & ucase(strPD & "\GetSupportService") & chr(34)
      objLOG.write vbnewline & now & vbtab & vbtab & " - REMOVING " & chr(34) & ucase(strPD & "\GetSupportService") & chr(34)
      call HOOK("rmdir /s /q " & chr(34) & strPD & "\GetSupportService_N-Central" & chr(34))
      if (err.number <> 0) then
        call LOGERR(5)
      end if
    end if
    ''REMOVE TAKE CONTROL SERVICES
    objOUT.write vbnewline & now & vbtab & vbtab & " - REMOVING TAKE CONTROL SERVICES"
    objLOG.write vbnewline & now & vbtab & vbtab & " - REMOVING TAKE CONTROL SERVICES"
    call HOOK("sc delete " & chr(34) & "BASupportExpressSrvcUpdater_N_Central" & chr(34))
    call HOOK("sc delete " & chr(34) & "BASupportExpressStandaloneService_N_Central" & chr(34))
  end if
elseif (errRET <> 0) then                                   ''ERRORS ENCOUNTERED DURING INITIAL START
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
sub FILEDL(strURL, strDL, strFILE)                          ''CALL HOOK TO DOWNLOAD FILE FROM URL , 'ERRRET'=11
  strSAV = vbnullstring
  ''SET DOWNLOAD PATH
  strSAV = strDL & "\" & strFILE
  objOUT.write vbnewline & now & vbtab & vbtab & vbtab & "HTTPDOWNLOAD-------------DOWNLOAD : " & strURL & " : SAVE AS :  " & strSAV
  objLOG.write vbnewline & now & vbtab & vbtab & vbtab & "HTTPDOWNLOAD-------------DOWNLOAD : " & strURL & " : SAVE AS :  " & strSAV
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
    objOUT.write vbnewline & now & vbtab & vbtab & " - DOWNLOAD : " & strSAV & " : SUCCESSFUL"
    objLOG.write vbnewline & now & vbtab & vbtab & " - DOWNLOAD : " & strSAV & " : SUCCESSFUL"
  end if
	set objHTTP = nothing
  if ((err.number <> 0) and (err.number <> 58)) then        ''ERROR RETURNED DURING DOWNLOAD , 'ERRRET'=11
    call LOGERR(11)
  end if
end sub

sub HOOK(strCMD)                                            ''CALL HOOK TO MONITOR OUTPUT OF CALLED COMMAND , 'ERRRET'=12
  on error resume next
  objOUT.write vbnewline & now & vbtab & vbtab & " - EXECUTING : HOOK(" & strCMD & ")"
  objLOG.write vbnewline & now & vbtab & vbtab & " - EXECUTING : HOOK(" & strCMD & ")"
  set objHOOK = objWSH.exec(strCMD)
  if (instr(1, strCMD, "takeown /F ") = 0) then             ''SUPPRESS 'TAKEOWN' SUCCESS MESSAGES
    while (not objHOOK.stdout.atendofstream)
      strIN = objHOOK.stdout.readline
      if (strIN <> vbnullstring) then
        objOUT.write vbnewline & now & vbtab & vbtab & vbtab & strIN 
        objLOG.write vbnewline & now & vbtab & vbtab & vbtab & strIN 
      end if
      wscript.sleep 10
    wend
    wscript.sleep 10
    strIN = objHOOK.stdout.readall
    if (strIN <> vbnullstring) then
      objOUT.write vbnewline & now & vbtab & vbtab & vbtab & strIN 
      objLOG.write vbnewline & now & vbtab & vbtab & vbtab & strIN 
    end if
  end if
  set objHOOK = nothing
  if (err.number <> 0) then                                 ''ERROR RETURNED DURING UPDATE CHECK , 'ERRRET'=12
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
  end select
end sub

sub CLEANUP()                                               ''SCRIPT CLEANUP
  on error resume next
  if (errRET = 0) then                                      ''SCRIPT COMPLETED SUCCESSFULLY
    objOUT.write vbnewline & vbnewline & now & vbtab & " - FIX_CONTROL SUCCESSFUL : " & now
    objLOG.write vbnewline & vbnewline & now & vbtab & " - FIX_CONTROL SUCCESSFUL : " & now
    err.clear
  elseif (errRET <> 0) then                                 ''SCRIPT FAILED
    objOUT.write vbnewline & vbnewline & now & vbtab & " - FIX_CONTROL FAILURE : " & errRET & " : " & now
    objLOG.write vbnewline & vbnewline & now & vbtab & " - FIX_CONTROL FAILURE : " & errRET & " : " & now
    ''RAISE CUSTOMIZED ERROR CODE , ERROR CODE WILL BE DEFINED RESTOP NUMBER INDICATING WHICH SECTION FAILED
    call err.raise(vbObjectError + errRET, "FIX_CONTROL", "fail")
  end if
  objOUT.write vbnewline & vbnewline & now & " - FIX_CONTROL COMPLETE" & vbnewline
  objLOG.write vbnewline & vbnewline & now & " - FIX_CONTROL COMPLETE" & vbnewline
  objLOG.close
  ''EMPTY OBJECTS
  set objEXEC = nothing
  set objLOG = nothing
  set objFSO = nothing
  set objWSH = nothing
  set objARG = nothing
  set objOUT = nothing
  set objIN = nothing
  ''END SCRIPT , RETURN ERROR
  wscript.quit err.number
end sub
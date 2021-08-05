''FIX_CONTROL.VBS
''DESIGNED TO FIX TAKE CONTROL
''WRITTEN BY : CJ BLEDSOE / CBLEDSOE<@>IPMCOMPUTERS.COM
on error resume next
''SCRIPT VARIABLES
dim errRET, strVER, strSEL
dim strREPO, strBRCH, strDIR
dim strIN, strOUT, strORG, strREP
dim colUSR(), colSID(), arrUSR(), arrSID()
''VARIABLES ACCEPTING PARAMETERS
dim strUSR, strOPT, strPWD, strSVC
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
  strPD = objWSH.expandenvironmentstrings("%PROGRAMDATA%")
  strPF = objWSH.expandenvironmentstrings("%PROGRAMFILES%")
  strPF86 = objWSH.expandenvironmentstrings("%PROGRAMFILES(X86)%")
  objOUT.write vbnewline & vbnewline & now & vbtab & " - EXECUTING FIX_CONTROL"
  objLOG.write vbnewline & vbnewline & now & vbtab & " - EXECUTING FIX_CONTROL"
  ''32BIT INSTALL
  if (objFSO.folderexists(strPF & "\BeAnywhere Support Express\GetSupportService_N-Central")) then
    ''RUN UNINSTALL
    objOUT.write vbnewline & now & vbtab & vbtab & " - UNINSTALLING 32BIT TAKE CONTROL"
    objLOG.write vbnewline & now & vbtab & vbtab & " - UNINSTALLING 32BIT TAKE CONTROL"
    if (objFSO.fileexists(strPF & "\BeAnywhere Support Express\GetSupportService_N-Central\uninstall.exe")) then
      'call HOOK(strPF & "\BeAnywhere Support Express\GetSupportService_N-Central\uninstall.exe" & " /S")
      objWSH.run strPF & "\BeAnywhere Support Express\GetSupportService_N-Central\uninstall.exe" & " /S", 0, true
    end if
    ''KILL SERVICE PROCESSES
    objOUT.write vbnewline & now & vbtab & vbtab & " - STOPPING TAKE CONTROL PROCESSES"
    objLOG.write vbnewline & now & vbtab & vbtab & " - STOPPING TAKE CONTROL PROCESSES"
    call HOOK("taskkill /F /IM BASupSrvc.exe /T")
    call HOOK("taskkill /F /IM BASupSrvcCnfg.exe /T")
    call HOOK("taskkill /F /IM BASupSrvcUpdater.exe /T")
    ''REMOVE DIRECTORY
    objOUT.write vbnewline & now & vbtab & vbtab & " - REMOVING TAKE CONTROL DIRECTORY"
    objLOG.write vbnewline & now & vbtab & vbtab & " - REMOVING TAKE CONTROL DIRECTORY"
    call HOOK("cmd.exe /C rmdir /s /q " & chr(34) & strPF & "\BeAnywhere Support Express\GetSupportService_N-Central" & chr(34))
    if (err.number <> 0) then
      call LOGERR(1)
    end if
  ''64BIT INSTALL
  elseif (objFSO.folderexists(strPF86 & "\BeAnywhere Support Express\GetSupportService_N-Central")) then
    ''RUN UNINSTALL
    objOUT.write vbnewline & now & vbtab & vbtab & " - UNINSTALLING 64BIT TAKE CONTROL"
    objLOG.write vbnewline & now & vbtab & vbtab & " - UNINSTALLING 64BIT TAKE CONTROL"
    if (objFSO.fileexists(strPF86 & "\BeAnywhere Support Express\GetSupportService_N-Central\uninstall.exe")) then
      'call HOOK(strPF86 & "\BeAnywhere Support Express\GetSupportService_N-Central\uninstall.exe" & " /S")
      objWSH.run strPF86 & "\BeAnywhere Support Express\GetSupportService_N-Central\uninstall.exe" & " /S", 0, true
    end if
    ''KILL SERVICE PROCESSES
    objOUT.write vbnewline & now & vbtab & vbtab & " - STOPPING TAKE CONTROL PROCESSES"
    objLOG.write vbnewline & now & vbtab & vbtab & " - STOPPING TAKE CONTROL PROCESSES"
    call HOOK("taskkill /F /IM BASupSrvc.exe /T")
    call HOOK("taskkill /F /IM BASupSrvcCnfg.exe /T")
    call HOOK("taskkill /F /IM BASupSrvcUpdater.exe /T")
    ''REMOVE DIRECTORY
    objOUT.write vbnewline & now & vbtab & vbtab & " - REMOVING TAKE CONTROL DIRECTORY"
    objLOG.write vbnewline & now & vbtab & vbtab & " - REMOVING TAKE CONTROL DIRECTORY"
    call HOOK("cmd.exe /c rmdir /s /q " & chr(34) & strPF86 & "\BeAnywhere Support Express\GetSupportService_N-Central" & chr(34))
    if (err.number <> 0) then
      call LOGERR(2)
    end if
  end if
  ''PROGRAMDATA DIRECTORY
  if (objFSO.folderexists(chr(34) & strPD & "\GetSupportService_N-Central" & chr(34))) then
    objOUT.write vbnewline & now & vbtab & vbtab & " - REMOVING " & chr(34) & ucase(strPD & "\GetSupportService_N-Central") & chr(34)
    objLOG.write vbnewline & now & vbtab & vbtab & " - REMOVING " & chr(34) & ucase(strPD & "\GetSupportService_N-Central") & chr(34)
    call HOOK("cmd.exe /C rmdir /s /q " & chr(34) & strPD & "\GetSupportService_N-Central" & chr(34))
    if (err.number <> 0) then
      call LOGERR(3)
    end if
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
  objOUT.write vbnewline & now & vbtab & vbtab & "EXECUTING : HOOK(" & strCMD & ")"
  objLOG.write vbnewline & now & vbtab & vbtab & "EXECUTING : HOOK(" & strCMD & ")"
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
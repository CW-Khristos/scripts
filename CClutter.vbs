on error resume next
''SCRIPT VARIABLES
dim errRET, strVER, strNEW
dim colFOL(31), blnLOG, lngSIZ, strFOL
''SCRIPT OBJECTS
dim objLOG, objHOOK, objHTTP, objXML
dim objIN, objOUT, objARG, objWSH, objFSO, objFOL
''VERSION FOR SCRIPT UPDATE, CCLUTTER.VBS, REF #2
strVER = 3
''DEFAULT SUCCESS
errRET = 0
''STDIN / STDOUT
set objIN = wscript.stdin
set objOUT = wscript.stdout
set objARG = wscript.arguments
''OBJECTS FOR LOCATING FOLDERS
set objWSH = createobject("wscript.shell")
set objFSO = createobject("scripting.filesystemobject")
''FILESIZE COUNTER
lngSIZ = 0
''SET BLNLOG TO 'TRUE' TO ENABLE A TEXT LOG
''THIS SHOULD ONLY BE USED IF NEEDED
''PREFERRABLY USING CSCRIPT.EXE CCLUTTER.VBS FROM COMMAND-LINE
blnLOG = true
''CHECK EXECUTION METHOD OF SCRIPT
strIN = lcase(mid(wscript.fullname, instrrev(wscript.fullname, "\") + 1))
if (strIN <> "cscript.exe") Then
  objOUT.write vbnewline & "SCRIPT LAUNCHED VIA EXPLORER, EXECUTING SCRIPT VIA CSCRIPT..."
  objWSH.run "cscript.exe //nologo " & chr(34) & Wscript.ScriptFullName & chr(34)
  wscript.quit
end if
''COLLECTION OF FOLDERS TO CHECK
''THESE FOLDERS REQUIRE RETRIEVAL FROM ENVIRONMENTAL VARIABLES
colFOL(0) = objWSH.expandenvironmentstrings("%windir%") & "\Temp"
colFOL(1) = objWSH.expandenvironmentstrings("%windir%") & "\Logs\CBS"
colFOL(2) = objWSH.expandenvironmentstrings("%windir%") & "\SoftwareDistribution"
''THESE FOLDERS ARE NORMAL FOLDER PATHS
colFOL(3) =  "C:\temp"
colFOL(4) = "C:\Program Files\N-able Technologies\NablePatchCache"
colFOL(5) = "C:\Program Files\N-able Technologies\UpdateServerCache"
colFOL(6) = "C:\Program Files (x86)\N-able Technologies\NablePatchCache"
colFOL(7) = "C:\Program Files (x86)\N-able Technologies\UpdateServerCache"
colFOL(8) = "C:\ProgramData\N-able Technologies\AutomationManager\Logs"
colFOL(9) = "C:\ProgramData\N-able Technologies\AutomationManager\temp"
colFOL(10) = "C:\ProgramData\N-able Technologies\AutomationManager\ScriptResults"
colFOL(11) = "C:\ProgramData\GetSupportService\logs"
colFOL(12) = "C:\ProgramData\GetSupportService_N-Central\logs"
colFOL(13) = "C:\ProgramData\GetSupportService_N-Central\Updates"
colFOL(14) = "C:\inetpub\logs\LogFiles\W3SVC2"
colfol(15) = "C:\inetpub\logs\LogFiles\W3SVC1"
''EXCHANGE LOGGING FOLDERS
if (objFSO.folderexists("C:\Program Files\Microsoft\Exchange Server")) then
  colFOL(16) = "C:\Program Files\Microsoft\Exchange Server\V15\Logging\Diagnostics\AnalyzerLogs"
  colFOL(17) = "C:\Program Files\Microsoft\Exchange Server\V15\Logging\Diagnostics\CertificateLogs"
  colFOL(18) = "C:\Program Files\Microsoft\Exchange Server\V15\Logging\Diagnostics\CosmosLog"
  colFOL(19) = "C:\Program Files\Microsoft\Exchange Server\V15\Logging\Diagnostics\DailyPerformanceLogs"
  colFOL(20) = "C:\Program Files\Microsoft\Exchange Server\V15\Logging\Diagnostics\Dumps"
  colFOL(21) = "C:\Program Files\Microsoft\Exchange Server\V15\Logging\Diagnostics\EtwTraces"
  colFOL(22) = "C:\Program Files\Microsoft\Exchange Server\V15\Logging\Diagnostics\Poison"
  colFOL(23) = "C:\Program Files\Microsoft\Exchange Server\V15\Logging\Diagnostics\ServiceLogs"
  colFOL(24) = "C:\Program Files\Microsoft\Exchange Server\V15\Logging\Diagnostics\Watermarks"
  colFOL(25) = "C:\Program Files\Microsoft\Exchange Server\V15\Logging\MailboxAssistantsLog"
  colFOL(26) = "C:\Program Files\Microsoft\Exchange Server\V15\Logging\MailboxAssociationLog"
  colFOL(27) = "C:\Program Files\Microsoft\Exchange Server\V15\Logging\MigrationMonitorLogs"
  colFOL(28) = "C:\Program Files\Microsoft\Exchange Server\V15\Logging\RpcHttp\W3SVC1"
  colFOL(29) = "C:\Program Files\Microsoft\Exchange Server\V15\Logging\RpcHttp\W3SVC2"
  colFOL(30) = "C:\Program Files\Microsoft\Exchange Server\V15\Logging\HttpProxy\RpcHttp"
end if
''C:\ProgramData\MXB\Backup Manager\logs
''READ PASSED COMMANDLINE ARGUMENTS
if (wscript.arguments.count > 0) then                       ''ARGUMENTS WERE PASSED
  for x = 0 to (wscript.arguments.count - 1)
    objOUT.write vbnewline & "ARGUMENT " & (x + 1) & " (ITEM " & x & ") " & " PASSED : " & ucase(objARG.item(x))
    objLOG.write vbnewline & "ARGUMENT " & (x + 1) & " (ITEM " & x & ") " & " PASSED : " & ucase(objARG.item(x))
  next 
  blnLOG = bool(objARG.item(0))
  strFOL = objARG.item(1)
else                                                        ''NO ARGUMENTS PASSED, DO NOT CREATE LOGFILE, NO CUSTOM DESTINATION
  blnLOG = false
end if

''------------
''BEGIN SCRIPT
objOUT.write vbnewline & vbnewline & now & vbtab & " - EXECUTING CCLUTTER"
objLOG.write vbnewline & vbnewline & now & vbtab & " - EXECUTING CCLUTTER"
''AUTOMATIC UPDATE, CCLUTTER.VBS, REF #2
call CHKAU()
''CREATE LOGFILE, IF ENABLED
if (blnLOG) then
  if (objFSO.fileexists("C:\temp\cclutter.txt")) then       ''LOGFILE EXISTS
    objFSO.deletefile "C:\temp\cclutter.txt", true
    set objLOG = objFSO.createtextfile("C:\temp\cclutter.txt")
    objLOG.close
    set objLOG = objFSO.opentextfile("C:\temp\cclutter.txt", 8)
  else                                                      ''LOGFILE NEEDS TO BE CREATED
    set objLOG = objFSO.createtextfile("C:\temp\cclutter.txt")
    objLOG.close
    set objLOG = objFSO.opentextfile("C:\temp\cclutter.txt", 8)
  end if
end if
''ENUMERATE THROUGH FOLDER COLLECTION
for x = 0 to ubound(colFOL)
  if (colFOL(x) <> vbnullstring) then                       ''ENSURE COLFOL(X) IS NOT EMPTY
    if (objFSO.folderexists(colFOL(x))) then                ''ENSURE FOLDER EXISTS BEFORE CLEARING
      set objFOL = objFSO.getfolder(colFOL(x))
      strNEW = vbnewline & "CLEARING : " & objFOL.path
      objOUT.write strNEW
      if (blnLOG) then                                      ''WRITE TO LOGFILE, IF ENABLED
        objLOG.write strNEW
      end if
      ''CLEAR NORMAL FOLDERS
      if (objFOL.path <> objWSH.expandenvironmentstrings("%windir%") & "\SoftwareDistribution") then
        strNEW = vbnewline & "CLEARING : " & objFOL.path
        objOUT.write strNEW
        if (blnLOG) then                                    ''WRITE TO LOGFILE, IF ENABLED
          objLOG.write strNEW
        end if
        call cFolder(objFOL)
      ''CLEARING WINDOWS UPDATES
      elseif (objFOL.path = objWSH.expandenvironmentstrings("%windir%") & "\SoftwareDistribution") then
        ''CHECK FOR 'PENDING.XML IF CLEARING SOFTWAREDISTRIBUTION
        if (objFSO.fileexists(objWSH.expandenvironmentstrings("%windir%") & "\WinSxS\pending.xml")) then
          strNEW = vbnewline & "'PENDING.XML' FOUND : SKIPPING : " & objFOL.path
          objOUT.write strNEW
          if (blnLOG) then                                  ''WRITE TO LOGFILE, IF ENABLED
            objLOG.write strNEW
          end if
        elseif (not (objFSO.fileexists(objWSH.expandenvironmentstrings("%windir%") & "\WinSxS\pending.xml"))) then
          strNEW = vbnewline & "'PENDING.XML' NOT FOUND : CLEARING : " & objFOL.path
          objOUT.write strNEW
          if (blnLOG) then                                  ''WRITE TO LOGFILE, IF ENABLED
            objLOG.write strNEW
          end if
          ''STOP WINDOWS UPDATE SERVICE TO CLEAR WINDOWS UPDATE FOLDER
          call HOOK("net stop wuauserv /y")
          call cFolder(objFOL)
          ''RESTART WINDOWS UPDATE SERVICE
          call HOOK("net start wuauserv")
        end if
      end if
    else                                                    ''NON-EXISTENT FOLDER
      strNEW = vbnewline & "NON-EXISTENT : " & colFOL(x)
      objOUT.write strNEW
      if (blnLOG) then                                      ''WRITE TO LOGFILE, IF ENABLED
        objLOG.write strNEW
      end if
    end if
  end if
next
''ENUMERATE THROUGH PASSED FOLDER PATH
'if (strFOL <> vbnullstring) then
'  if (objFSO.folderexists(strFOL)) then                    ''ENSURE FOLDER EXISTS BEFORE CLEARING
'    set objFOL = objFSO.getfolder(strFOL)
'    strNEW = vbnewline & "CLEARING : " & objFOL.path
'    objOUT.write strNEW
'    ''WRITE TO LOGFILE, IF ENABLED
'    if (blnLOG) then
'      objLOG.write strNEW
'    end if
'    ''CLEAR CONTENTS OF FOLDER
'    call cFolder(objFOL)
'  else                                                     ''NON-EXISTENT FOLDER
'    strNEW = vbnewline & "NON-EXISTENT : " & colFOL(x)
'    objOUT.write strNEW
'    ''WRITE TO LOGFILE, IF ENABLED
'    if (blnLOG) then
'      objLOG.write strNEW
'    end if
'  end if
'end if
''END SCRIPT
call CLEANUP()
''END SCRIPT
''------------

''SUB-ROUTINES
sub cFolder (ByVal Folder)                                  ''SUB-ROUTINE TO CLEAR CONTENTS OF FOLDER
  ''SUB-ROUTINE IS RECURSIVE, WILL CLEAR FOLDER AND SUB-FOLDERS!
  on error resume next
  dim colFIL, colFOL
  ''DELETE FILES
  set colFIL = Folder.files
  for each objFIL in colFIL                                 ''ENUMERATE EACH FILE
    filSIZ = round((objFIL.size / 1024), 2)
    strFIL = objFIL.path
    objFIL.delete(True)
    if (err.number = 0) then                                ''SUCCESSFULLY DELETED FILE
      lngSIZ = (lngSIZ + filSIZ)
      strNEW = vbnewline & "DELETED FILE: " & strFIL
      'objOUT.write strNEW
      if (blnLOG) then                                      ''WRITE TO LOGFILE, IF ENABLED
        objLOG.write strNEW
      end if
    elseif (err.number <> 0) then                           ''ERROR ENCOUNTERED DELETING FILE
      strNEW = vbnewline & "ERROR : " &  err.number & vbtab & err.description & vbtab & strFIL
      objOUT.write strNEW
      if (blnLOG) then                                      ''WRITE TO LOGFILE, IF ENABLED
        objLOG.write strNEW
      end if
    end if
  next
  ''EMPTY AND DELETE SUB-FOLDERS
  set colFOL = Folder.SubFolders
  for each subFOL in colFOL                                 ''ENUMERATE EACH SUB-FOLDER
    strFOL = subFOL.path
    call cFolder(subFOL)
    subFOL.delete(True)
    if (err.number = 0 ) then                               ''SUCCESSFULLY DELETED FOLDER
      strNEW = vbnewline & "REMOVED FOLDER : " & strFOL
      objOUT.write strNEW
      if (blnLOG) then                                      ''WRITE TO LOGFILE, IF ENABLED
        objLOG.write strNEW
      end if
    elseif (err.number <> 0) then                           ''ENCOUNTERED ERROR TRYING TO DELETE FOLDER
      strNEW = vbnewline & "ERROR : " &  err.number & vbtab & err.description & vbtab & strFOL
      objOUT.write strNEW
      if (blnLOG) then                                      ''WRITE TO LOGFILE, IF ENABLED
        objLOG.write strNEW
      end if
    end if
  next
end sub

sub CHKAU()																					        ''CHECK FOR SCRIPT UPDATE, PWDCHG.VBS, REF #2 , FIXES #21
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
        objOUT.write vbnewline & now & vbtab & " - CClutter :  " & strVER & " : GitHub : " & objSCR.text & vbnewline
        objLOG.write vbnewline & now & vbtab & " - CClutter :  " & strVER & " : GitHub : " & objSCR.text & vbnewline
        if (cint(objSCR.text) > cint(strVER)) then
          objOUT.write vbnewline & now & " - UPDATING " & objSCR.nodename & " : " & objSCR.text & vbnewline
          objLOG.write vbnewline & now & " - UPDATING " & objSCR.nodename & " : " & objSCR.text & vbnewline
          ''DOWNLOAD LATEST VERSION OF SCRIPT
          call FILEDL("https://github.com/CW-Khristos/scripts/raw/master/cclutter.vbs", wscript.scriptname)
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
    errRET = 2
		err.clear
  end if
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

sub LOGERR(intSTG)                                          ''CALL HOOK TO MONITOR OUTPUT OF CALLED COMMAND
  if (err.number <> 0) then
    objOUT.write vbnewline & now & vbtab & vbtab & vbtab & err.number & vbtab & err.description & vbnewline
    objLOG.write vbnewline & now & vbtab & vbtab & vbtab & err.number & vbtab & err.description & vbnewline
		errRET = intSTG
		err.clear
  end if
end sub

sub CLEANUP()                                     ''SCRIPT CLEANUP
  strNEW = vbnewline & "CCLUTTER COMPLETE : CLEARED " & round((lngSIZ / 1024),2) & " MB"
  objOUT.write strNEW
  if (blnLOG) then                                ''LOGFILE CLEANUP, IF ENABLED
    objLOG.write strNEW
    objLOG.close
    set objLOG = nothing
    ''UNCOMMENT LINES BELOW TO CAUSE LOGFILE TO OPEN AUTOMATICALLY
    'objWSH.run "C:\cclutter.txt", 1
    'wscript.sleep 1000
    ''UNCOMMENT THE FOLLOWING LINE TO DELETE LOGFILE AFTER EXECUTION
    'objFSO.deletefile "C:\cclutter.txt", true
  end if
  ''SCRIPT / OBJECT CLEANUP
  set objFOL = nothing
  set objFSO = nothing
  set objOUT = nothing
  set objWSH = nothing
  ''END SCRIPT, RETURN DEFAULT NO ERROR
  wscript.quit 0
end sub
''MSP_LSV.VBS
''NO REQUIRED PARAMETERS / DOES NOT ACCEPT PARAMETERS
''SCRIPT IS DESIGNED TO SIMPLY EXPORT MSP BACKUP SETTINGS USING CLIENTTOOL.EXE UTILITY
''EXPORTS MSP BACKUP SETTINGS TO C:\TEMP\LSV.TXT
''MUST BE USED IN CONJUNCTION WITH MSP BACKUP SYNCHRONIZATION - LSV SYNCHRONIZATION.AMP CUSTOM SERVICE
''WRITTEN BY : CJ BLEDSOE / CJ<@>THECOMPUTERWARRIORS.COM
'on error resume next
''DEFINE VARIABLES
dim strDLM, intDIFF
dim errRET, strVER, retDEL
''SCRIPT OBJECTS
dim objIN, objOUT, objARG, objWSH
dim objFSO, objLOG, objLSV, objHOOK, objHTTP, objXML
''DEFAULT SUCCESS
errRET = 0
''VERSION FOR SCRIPT UPDATE, MSP_LSV.VBS, REF #2
strVER = 1
''STDIN / STDOUT
set objIN = wscript.stdin
set objOUT = wscript.stdout
set objARG = wscript.arguments
''CREATE OBJECTS
set objWSH = createobject("wscript.shell")
set objFSO = createobject("scripting.filesystemobject")
''PREPARE LOGFILE
if (objFSO.fileexists("C:\temp\msp_lsv.txt")) then  ''PREVIOUS LOGFILE EXISTS
  objFSO.deletefile "C:\temp\msp_lsv.txt", true
end if
if (objFSO.fileexists("C:\temp\msp_lsv.txt")) then  ''LOGFILE EXISTS
  objFSO.deletefile "C:\temp\msp_lsv.txt", true
  set objLOG = objFSO.createtextfile("C:\temp\msp_lsv.txt")
  objLOG.close
  set objLOG = objFSO.opentextfile("C:\temp\msp_lsv.txt", 8)
else                                            ''LOGFILE NEEDS TO BE CREATED
  set objLOG = objFSO.createtextfile("C:\temp\msp_lsv.txt")
  objLOG.close
  set objLOG = objFSO.opentextfile("C:\temp\msp_lsv.txt", 8)
end if
''PREPARE MONITOR FILE
if (objFSO.fileexists("C:\temp\lsv.txt")) then  ''PREVIOUS LOGFILE EXISTS
  objFSO.deletefile "C:\temp\lsv.txt", true
end if
if (objFSO.fileexists("C:\temp\lsv.txt")) then  ''LOGFILE EXISTS
  objFSO.deletefile "C:\temp\lsv.txt", true
  set objLSV = objFSO.createtextfile("C:\temp\lsv.txt")
  objLSV.close
  set objLSV = objFSO.opentextfile("C:\temp\lsv.txt", 8)
else                                            ''LOGFILE NEEDS TO BE CREATED
  set objLSV = objFSO.createtextfile("C:\temp\lsv.txt")
  objLSV.close
  set objLSV = objFSO.opentextfile("C:\temp\lsv.txt", 8)
end if
''CHECK EXECUTION METHOD OF SCRIPT
strIN = lcase(mid(wscript.fullname, instrrev(wscript.fullname, "\") + 1))
if (strIN <> "cscript.exe") Then                ''LAUNCHED VIA WSCRIPT, RE-LAUNCH WITH CSCRIPT
  objOUT.write vbnewline & "SCRIPT LAUNCHED VIA EXPLORER, EXECUTING SCRIPT VIA CSCRIPT..."
  objWSH.run "cscript.exe //nologo " & chr(34) & Wscript.ScriptFullName & chr(34)
  wscript.quit
end if
''NO ARGUMENTS REQUIRED
''READ PASSED COMMANDLINE ARGUMENTS
'if (wscript.arguments.count > 0) then          ''ARGUMENTS WERE PASSED
'  for x = 0 to (wscript.arguments.count - 1)
'    objOUT.write vbnewline & now & vbtab & "ARGUMENT " & (x + 1) & " (ITEM " & x & ") " & " PASSED : " & ucase(objARG.item(x))
'  next 
'end if

''------------
''BEGIN SCRIPT
objOUT.write vbnewline & now & " - STARTING MSP_LSV" & vbnewline
objLOG.write vbnewline & now & " - STARTING MSP_LSV" & vbnewline
''AUTOMATIC UPDATE, MSP_LSV.VBS, REF #2
call CHKAU()
''EXPORT MSP BACKUP SETTINGS USING CLIENTTOOL UTILITY
'call HOOK("C:\Program Files\Backup Manager\clienttool.exe control.setting.list > " & chr(34) & "C:\temp\lsv.txt" & chr(34))
set objHOOK = objWSH.exec("C:\Program Files\Backup Manager\clienttool.exe control.setting.list")
strIN = objHOOK.stdout.readall
arrIN = split(strIN, vbnewline)
''WRITE SCRIPT LOGFILE
for intIN = 0 to ubound(arrIN)                                  ''CHECK SETTINGS LINE BY LINE, EXCLUDE THE 'C:\WINDOWS\TEMP' AND 'C:\TEMP' DIRECTORIES TO AVOID FALSE MONITOR ALERTS
  if ((instr(1, lcase(arrIN(intIN)), "c:\") = 0) and _
    (instr(1,lcase(arrIN(intIN)), "\temp") = 0)) then
      objOUT.write vbnewline & now & vbtab & arrIN(intIN)
      objLOG.write vbnewline & now & vbtab & arrIN(intIN)
  end if
next
intIN = 0
''WRITE MONITOR FILE
for intIN = 0 to ubound(arrIN)                                  ''CHECK SETTINGS LINE BY LINE, EXCLUDE THE 'C:\WINDOWS\TEMP' AND 'C:\TEMP' DIRECTORIES TO AVOID FALSE MONITOR ALERTS
  if ((instr(1, lcase(arrIN(intIN)), "c:\") = 0) and _
    (instr(1,lcase(arrIN(intIN)), "\temp") = 0)) then
      ''EXCLUDE ALL OUTPUT EXCEPT FOR LSV LOCATION
      if (instr(1, lcase(arrIN(intIN)),"localspeedvaultlocation")) then
        ''REMOVE LOCALSPEEDVAULTLOCATION 'LABEL', OUTPUT ONLY THE ACTUAL LSV DIRECTORY
        strTMP = split(lcase(arrIN(intIN)), "localspeedvaultlocation ")(1)
        objOUT.write vbnewline & now & vbtab & arrIN(intIN) & " - WRITTEN TO LSV.TXT"
        objLSV.write strTMP
        exit for
      end if
  end if
next
set objHOOK = nothing
''CLEAN OBJECTS
call CLEANUP()
''END SCRIPT
''------------

''SUB-ROUTINES
sub CHKAU()																					        ''CHECK FOR SCRIPT UPDATE, MSP_LSV.VBS, REF #2 , FIXES #4
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
        objOUT.write vbnewline & now & vbtab & " - MSP_LSV :  " & strVER & " : GitHub : " & objSCR.text & vbnewline
        objLOG.write vbnewline & now & vbtab & " - MSP_LSV :  " & strVER & " : GitHub : " & objSCR.text & vbnewline
        if (cint(objSCR.text) > cint(strVER)) then
          objOUT.write vbnewline & now & " - UPDATING " & objSCR.nodename & " : " & objSCR.text & vbnewline
          objLOG.write vbnewline & now & " - UPDATING " & objSCR.nodename & " : " & objSCR.text & vbnewline
          ''DOWNLOAD LATEST VERSION OF SCRIPT
          call FILEDL("https://github.com/CW-Khristos/scripts/raw/master/MSP%20Backups/msp_lsv.vbs", wscript.scriptname)
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
  if (errRET = 0) then         											        ''MSP_LSV COMPLETED SUCCESSFULLY
    objOUT.write vbnewline & "MSP_LSV SUCCESSFUL : " & now
    objLOG.write vbnewline & "MSP_LSV SUCCESSFUL : " & now
  elseif (errRET <> 0) then    											        ''MSP_LSV FAILED
    objOUT.write vbnewline & "MSP_LSV FAILURE : " & now & " : " & errRET
    objLOG.write vbnewline & "MSP_LSV FAILURE : " & now & " : " & errRET
    ''RAISE CUSTOMIZED ERROR CODE, ERROR CODE WILL BE DEFINE RESTOP NUMBER INDICATING WHICH SECTION FAILED
    call err.raise(vbObjectError + errRET, "MSP_LSV", "FAILURE")
  end if
  objOUT.write vbnewline & vbnewline & now & " - MSP_LSV COMPLETE" & vbnewline
  objLOG.write vbnewline & vbnewline & now & " - MSP_LSV COMPLETE" & vbnewline
  objLOG.close
  objLSV.close
  ''EMPTY OBJECTS
  set objLSV = nothing
  set objLOG = nothing
  set objFSO = nothing
  set objWSH = nothing
  set objARG = nothing
  set objOUT = nothing
  set objIN = nothing
  ''END SCRIPT, RETURN ERROR NUMBER
  wscript.quit err.number
end sub
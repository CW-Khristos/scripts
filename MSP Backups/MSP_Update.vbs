''https://cdn.cloudbackup.management/maxdownloads/mxb-windows-x86_x64.exe
''MSP_UPDATE.VBS
''SCRIPT IS DESIGNED TO AUTOMATICALLY DOWNLOAD AND INSTALL MSP BACKUPS FROM DIRECT LINK
''WRITTEN BY : CJ BLEDSOE / CJ<@>THECOMPUTERWARRIORS.COM
'on error resume next
''SCRIPT VARIABLES
dim objFSO, objLOG, objHOOK
dim objIN, objOUT, objARG, objWSH
dim errRET, strIDL, strTMP, arrTMP, strIN
''VSS WRITER FLAGS
dim blnSQL, blnTSK, blnVSS, blnWMI
dim blnAHS, blnBIT, blnCSVC, blnRDP, blnRUN
''SET 'ERRRET' CODE
errRET = 0
''DEFAULT 'BLNRUN' FLAG
blnRUN = false
''STDIN / STDOUT
set objIN = wscript.stdin
set objOUT = wscript.stdout
set objARG = wscript.arguments
''OBJECTS FOR LOCATING FOLDERS
set objWSH = createobject("wscript.shell")
set objFSO = createobject("scripting.filesystemobject")
''CHECK EXECUTION METHOD OF SCRIPT
strIN = lcase(mid(wscript.fullname, instrrev(wscript.fullname, "\") + 1))
if (strIN <> "cscript.exe") Then
  objOUT.write vbnewline & "SCRIPT LAUNCHED VIA EXPLORER, EXECUTING SCRIPT VIA CSCRIPT..."
  objWSH.run "cscript.exe //nologo " & chr(34) & Wscript.ScriptFullName & chr(34)
  wscript.quit
end if
''PREPARE LOGFILE
if (objFSO.fileexists("C:\temp\MSP_UPDATE")) then      ''LOGFILE EXISTS
  objFSO.deletefile "C:\temp\MSP_UPDATE", true
  set objLOG = objFSO.createtextfile("C:\temp\MSP_UPDATE")
  objLOG.close
  set objLOG = objFSO.opentextfile("C:\temp\MSP_UPDATE", 8)
else                                                  ''LOGFILE NEEDS TO BE CREATED
  set objLOG = objFSO.createtextfile("C:\temp\MSP_UPDATE")
  objLOG.close
  set objLOG = objFSO.opentextfile("C:\temp\MSP_UPDATE", 8)
end if
''READ PASSED COMMANDLINE ARGUMENTS
if (wscript.arguments.count > 0) then                       ''ARGUMENTS WERE PASSED
  for x = 0 to (wscript.arguments.count - 1)
    objOUT.write vbnewline & now & vbtab & " - ARGUMENT " & (x + 1) & " (ITEM " & x & ") " & " PASSED : " & ucase(objARG.item(x))
    objLOG.write vbnewline & now & vbtab & " - ARGUMENT " & (x + 1) & " (ITEM " & x & ") " & " PASSED : " & ucase(objARG.item(x))
  next 
else                                                        ''NO ARGUMENTS PASSED, END SCRIPT
  objOUT.write vbnewline & vbnewline & now & vbtab & " - NO ARGUMENTS PASSED"
  objLOG.write vbnewline & vbnewline & now & vbtab & " - NO ARGUMENTS PASSED"
  errRET = 1
  'call CLEANUP
end if
objOUT.write vbnewline & now & " - STARTING MSP_UPDATE" & vbnewline
objLOG.write vbnewline & now & " - STARTING MSP_UPDATE" & vbnewline
''DOWNLOAD MSP BACKUP CLIENT
objOUT.write vbnewline & now & vbtab & " - DOWNLOADING LATEST MSP BACKUP CLIENT"
objLOG.write vbnewline & now & vbtab & " - DOWNLOADING LATEST MSP BACKUP CLIENT"
call FILEDL("https://cdn.cloudbackup.management/maxdownloads/mxb-windows-x86_x64.exe", "mxb-windows-x86_x64.exe")
''INSTALL MSP BACKUP MANAGER
objOUT.write vbnewline & now & vbtab & " - INSTALLING LATEST MSP BACKUP CLIENT"
objLOG.write vbnewline & now & vbtab & " - INSTALLING LATEST MSP BACKUP CLIENT"
call HOOK("C:\temp\mxb-windows-x86_x64.exe")
''END SCRIPT
call CLEANUP()

''SUB-ROUTINES
sub FILEDL(strURL, strFILE)                                 ''CALL HOOK TO DOWNLOAD FILE FROM URL
  strSAV = vbnullstring
  ''SET DOWNLOAD PATH
  strSAV = "C:\temp\" & strFILE
  objOUT.write vbnewline & now & vbtab & vbtab & "HTTPDOWNLOAD-------------DOWNLOAD : " & strURL & " : SAVE AS :  " & strSAV
  objLOG.write vbnewline & now & vbtab & vbtab & "HTTPDOWNLOAD-------------DOWNLOAD : " & strURL & " : SAVE AS :  " & strSAV
  ''CREATE HTTP OBJECT
  set objHTTP = createobject( "WinHttp.WinHttpRequest.5.1" )
  ''DOWNLOAD FROM URL
  objHTTP.open "GET", strURL, false
  objHTTP.send
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
    objOUT.write vbnewline & vbnewline & now & vbtab & " - DOWNLOAD : " & strSAV & " : SUCCESSFUL"
    objLOG.write vbnewline & vbnewline & now & vbtab & " - DOWNLOAD : " & strSAV & " : SUCCESSFUL"
    set objHTTP = nothing
  end if
  if (err.number <> 0) then
    errRET = 2
  end if
end sub

sub HOOK(strCMD)                                            ''CALL HOOK TO MONITOR OUTPUT OF CALLED COMMAND
  on error resume next
  'comspec = objWSH.ExpandEnvironmentStrings("%comspec%")
  set objHOOK = objWSH.exec(strCMD)
  'while (objHOOK.status = 0)
    while (not objHOOK.stdout.atendofstream)
      strIN = objHOOK.stdout.readline
      if (strIN <> vbnullstring) then
        objOUT.write vbnewline & now & vbtab & vbtab & strIN 
        objLOG.write vbnewline & now & vbtab & vbtab & strIN 
      end if
    wend
    wscript.sleep 10
  'wend
  strIN = objHOOK.stdout.readall
  if (strIN <> vbnullstring) then
    objOUT.write vbnewline & now & vbtab & vbtab & strIN 
    objLOG.write vbnewline & now & vbtab & vbtab & strIN 
  end if
  'errRET = objHOOK.exitcode
  set objHOOK = nothing
  if (err.number <> 0) then
    errRET = 3
    objOUT.write vbnewline & now & vbtab & vbtab & err.number & vbtab & err.description
    objLOG.write vbnewline & now & vbtab & vbtab & err.number & vbtab & err.description
  end if
end sub

sub CLEANUP()                                               ''SCRIPT CLEANUP
  objOUT.write vbnewline & vbnewline & now & " - MSP_UPDATE COMPLETE" & vbnewline
  objLOG.write vbnewline & vbnewline & now & " - MSP_UPDATE COMPLETE" & vbnewline
  objLOG.close
  ''EMPTY OBJECTS
  set objLOG = nothing
  set objFSO = nothing
  set objWSH = nothing
  set objARG = nothing
  set objOUT = nothing
  set objIN = nothing
  ''END SCRIPT, RETURN ERROR NUMBER
  wscript.quit errRET
end sub
on error resume next
''DEFINE VARIABLES
dim strNEW
dim colFOL(21), blnLOG, lngSIZ, strFOL
dim objOUT, objARG, objWSH, objFSO, objFOL, objLOG
''FILESIZE COUNTER
lngSIZ = 0
''SET BLNLOG TO 'TRUE' TO ENABLE A TEXT LOG
''THIS SHOULD ONLY BE USED IF NEEDED
''PREFERRABLY USING CSCRIPT.EXE CCLUTTER.VBS FROM COMMAND-LINE
blnLOG = true
''CREATE SCRIPTING OBJECTS
set objOUT = wscript.stdout
set objARG = wscript.arguments
set objWSH = createobject("wscript.shell")
set objFSO = createobject("scripting.filesystemobject")
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
''colFOL(1) =  objWSH.expandenvironmentstrings("%temp%")
colFOL(1) = objWSH.expandenvironmentstrings("%windir%") & "\Logs\CBS"
colFOL(2) = objWSH.expandenvironmentstrings("%windir%") & "\SoftwareDistribution"
''THESE FOLDERS ARE NORMAL FOLDER PATHS
colFOL(3) = "C:\Program Files\N-able Technologies\NablePatchCache"
colFOL(4) = "C:\Program Files (x86)\N-able Technologies\NablePatchCache"
colFOL(5) = "C:\inetpub\logs\LogFiles\W3SVC2"
colfol(6) = "C:\inetpub\logs\LogFiles\W3SVC1"
''EXCHANGE LOGGING FOLDERS
if (objFSO.folderexists("C:\Program Files\Microsoft\Exchange Server")) then
  colFOL(7) = "C:\Program Files\Microsoft\Exchange Server\V15\Logging\Diagnostics\AnalyzerLogs"
  colFOL(8) = "C:\Program Files\Microsoft\Exchange Server\V15\Logging\Diagnostics\CertificateLogs"
  colFOL(9) = "C:\Program Files\Microsoft\Exchange Server\V15\Logging\Diagnostics\CosmosLog"
  colFOL(10) = "C:\Program Files\Microsoft\Exchange Server\V15\Logging\Diagnostics\DailyPerformanceLogs"
  colFOL(11) = "C:\Program Files\Microsoft\Exchange Server\V15\Logging\Diagnostics\Dumps"
  colFOL(12) = "C:\Program Files\Microsoft\Exchange Server\V15\Logging\Diagnostics\EtwTraces"
  colFOL(13) = "C:\Program Files\Microsoft\Exchange Server\V15\Logging\Diagnostics\Poison"
  colFOL(14) = "C:\Program Files\Microsoft\Exchange Server\V15\Logging\Diagnostics\ServiceLogs"
  colFOL(15) = "C:\Program Files\Microsoft\Exchange Server\V15\Logging\Diagnostics\Watermarks"
  colFOL(16) = "C:\Program Files\Microsoft\Exchange Server\V15\Logging\MailboxAssistantsLog"
  colFOL(17) = "C:\Program Files\Microsoft\Exchange Server\V15\Logging\MailboxAssociationLog"
  colFOL(18) = "C:\Program Files\Microsoft\Exchange Server\V15\Logging\MigrationMonitorLogs"
  colFOL(19) = "C:\Program Files\Microsoft\Exchange Server\V15\Logging\RpcHttp\W3SVC1"
  colFOL(20) = "C:\Program Files\Microsoft\Exchange Server\V15\Logging\RpcHttp\W3SVC2"
  colFOL(21) = "C:\Program Files\Microsoft\Exchange Server\V15\Logging\HttpProxy\RpcHttp"
end if
''C:\ProgramData\MXB\Backup Manager\logs
''READ PASSED COMMANDLINE ARGUMENTS
if (wscript.arguments.count > 0) then             ''ARGUMENTS WERE PASSED
  for x = 0 to (wscript.arguments.count - 1)
    objOUT.write vbnewline & "ARGUMENT " & (x + 1) & " (ITEM " & x & ") " & " PASSED : " & ucase(objARG.item(x))
    objLOG.write vbnewline & "ARGUMENT " & (x + 1) & " (ITEM " & x & ") " & " PASSED : " & ucase(objARG.item(x))
  next 
  blnLOG = bool(objARG.item(0))
  strFOL = objARG.item(1)
else                                              ''NO ARGUMENTS PASSED, DO NOT CREATE LOGFILE, NO CUSTOM DESTINATION
  blnLOG = false
end if
''CREATE LOGFILE, IF ENABLED
if (blnLOG) then
  if (objFSO.fileexists("C:\cclutter.txt")) then  ''LOGFILE EXISTS
    objFSO.deletefile "C:\cclutter.txt", true
    set objLOG = objFSO.createtextfile("C:\cclutter.txt")
    objLOG.close
    set objLOG = objFSO.opentextfile("C:\cclutter.txt", 8)
  else                                            ''LOGFILE NEEDS TO BE CREATED
    set objLOG = objFSO.createtextfile("C:\cclutter.txt")
    objLOG.close
    set objLOG = objFSO.opentextfile("C:\cclutter.txt", 8)
  end if
end if
''ENUMERATE THROUGH FOLDER COLLECTION
for x = 0 to ubound(colFOL)
  if (colFOL(x) <> vbnullstring) then             ''ENSURE COLFOL(X) IS NOT EMPTY
    if (objFSO.folderexists(colFOL(x))) then      ''ENSURE FOLDER EXISTS BEFORE CLEARING
	  set objFOL = objFSO.getfolder(colFOL(x))
	  strNEW = vbnewline & "CLEARING : " & objFOL.path
	  objOUT.write strNEW
      if (blnLOG) then                            ''WRITE TO LOGFILE, IF ENABLED
        objLOG.write strNEW
      end if
      ''CLEAR CONTENTS OF FOLDER
      call cFolder(objFOL)
    else                                          ''NON-EXISTENT FOLDER
	  strNEW = vbnewline & "NON-EXISTENT : " & colFOL(x)
      objOUT.write strNEW
      if (blnLOG) then                            ''WRITE TO LOGFILE, IF ENABLED
        objLOG.write strNEW
      end if
    end if
  end if
next
''ENUMERATE THROUGH PASSED FOLDER PATH
'if (strFOL <> vbnullstring) then
'  if (objFSO.folderexists(strFOL)) then           ''ENSURE FOLDER EXISTS BEFORE CLEARING
'    set objFOL = objFSO.getfolder(strFOL)
'    strNEW = vbnewline & "CLEARING : " & objFOL.path
'    objOUT.write strNEW
'    ''WRITE TO LOGFILE, IF ENABLED
'    if (blnLOG) then
'      objLOG.write strNEW
'    end if
'    ''CLEAR CONTENTS OF FOLDER
'    call cFolder(objFOL)
'  else                                            ''NON-EXISTENT FOLDER
'    strNEW = vbnewline & "NON-EXISTENT : " & colFOL(x)
'    objOUT.write strNEW
'    ''WRITE TO LOGFILE, IF ENABLED
'    if (blnLOG) then
'      objLOG.write strNEW
'    end if
'  end if
'end if
''END SCRIPT
call CLEANUP

''SUB-ROUTINES
sub cFolder (ByVal Folder)                        ''SUB-ROUTINE TO CLEAR CONTENTS OF FOLDER
  ''SUB-ROUTINE IS RECURSIVE, WILL CLEAR FOLDER AND SUB-FOLDERS!
  on error resume next
  dim colFIL, colFOL
  ''DELETE FILES
  set colFIL = Folder.files
  for each objFIL in colFIL                       ''ENUMERATE EACH FILE
    filSIZ = round((objFIL.size / 1024), 2)
    strFIL = objFIL.path
    objFIL.delete(True)
    if (err.number = 0) then                      ''SUCCESSFULLY DELETED FILE
      lngSIZ = (lngSIZ + filSIZ)
      strNEW = vbnewline & "DELETED FILE: " & strFIL
      'objOUT.write strNEW
      if (blnLOG) then                            ''WRITE TO LOGFILE, IF ENABLED
        objLOG.write strNEW
      end if
    elseif (err.number <> 0) then                 ''ERROR ENCOUNTERED DELETING FILE
      strNEW = vbnewline & "ERROR : " &  err.number & vbtab & err.description & vbtab & strFIL
      objOUT.write strNEW
      if (blnLOG) then                            ''WRITE TO LOGFILE, IF ENABLED
	    objLOG.write strNEW
      end if
    end if
  next
  ''EMPTY AND DELETE SUB-FOLDERS
  set colFOL = Folder.SubFolders
  for each subFOL in colFOL                       ''ENUMERATE EACH SUB-FOLDER
    strFOL = subFOL.path
    call cFolder(subFOL)
    subFOL.delete(True)
    if (err.number = 0 ) then                     ''SUCCESSFULLY DELETED FOLDER
      strNEW = vbnewline & "REMOVED FOLDER : " & strFOL
      objOUT.write strNEW
      if (blnLOG) then                            ''WRITE TO LOGFILE, IF ENABLED
        objLOG.write strNEW
      end if
    elseif (err.number <> 0) then                 ''ENCOUNTERED ERROR TRYING TO DELETE FOLDER
      strNEW = vbnewline & "ERROR : " &  err.number & vbtab & err.description & vbtab & strFOL
      objOUT.write strNEW
      if (blnLOG) then                            ''WRITE TO LOGFILE, IF ENABLED
	    objLOG.write strNEW
      end if
    end if
  next
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
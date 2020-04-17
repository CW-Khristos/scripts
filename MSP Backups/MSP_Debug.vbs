''MSP_DEBUG.VBS
''DESIGNED TO UPDATE THE MSP BACKUP 'CONFIG.INI' FILE IN AN AUTOMATED FASHION TO ENABLE DEBUG LOGGING
''REQUIRED PARAMETER : 'STRHDR' TO IDENTIFY SECTION OF 'CONFIG.INI' FILE TO MODIFY
''REQUIRED PARAMETER : 'STRCHG', SCRIPT VARIABLE TO CONTAIN STRING TO INJECT INTO 'CONFIG.INI' FILE
''WRITTEN BY : CJ BLEDSOE / CJ<@>THECOMPUTERWARRIORS.COM
on error resume next
''SCRIPT VARIABLES
dim strVER, errRET
dim blnHDR, blnINJ, blnMOD
dim strREPO, strBRCH, strDIR
''VARIABLES ACCEPTING PARAMETERS
dim strIN, arrIN, strHDR, strCHG
''SCRIPT OBJECTS
dim objIN, objOUT, objARG
dim objWSH, objFSO, objLOG, objCFG
''VERSION FOR SCRIPT UPDATE, MSP_DEBUG.VBS, REF #2 , REF #68 , REF #69 , FIXES #24
strVER = 4
strREPO = "scripts"
strBRCH = "dev"
strDIR = "MSP Backups"
''SET 'errRET' CODE
errRET = 0
''SET 'BLNHDR' FLAG
blnHDR = false
''SET 'BLNINJ' FLAG
blnINJ = false
''SET 'BLNMOD' FLAG
blnMOD = true
''SET HEADER TO INSERT INTO CONFIG.INI
strHDR = "[Logging]"
''SET STRING TO INSERT INTO CONFIG.INI
strCHG = "LoggingLevel=Debug"
''STDIN / STDOUT
set objIN = wscript.stdin
set objOUT = wscript.stdout
set objARG = wscript.arguments
''OBJECTS FOR LOCATING FOLDERS
strTMP = "C:\temp\"
set objWSH = createobject("wscript.shell")
set objFSO = createobject("scripting.filesystemobject")
''MSP BACKUP MANAGER CONFIG.INI FILE
set objCFG = objFSO.opentextfile("C:\Program Files\Backup Manager\config.ini")
''CHECK EXECUTION METHOD OF SCRIPT
strIN = lcase(mid(wscript.fullname, instrrev(wscript.fullname, "\") + 1))
if (strIN <> "cscript.exe") Then
  objOUT.write vbnewline & "SCRIPT LAUNCHED VIA EXPLORER, EXECUTING SCRIPT VIA CSCRIPT..."
  objWSH.run "cscript.exe //nologo " & chr(34) & Wscript.ScriptFullName & chr(34)
  wscript.quit
end if
''PREPARE LOGFILE
if (objFSO.fileexists("C:\temp\MSP_DEBUG")) then                      ''LOGFILE EXISTS
  objFSO.deletefile "C:\temp\MSP_DEBUG", true
  set objLOG = objFSO.createtextfile("C:\temp\MSP_DEBUG")
  objLOG.close
  set objLOG = objFSO.opentextfile("C:\temp\MSP_DEBUG", 8)
else                                                                  ''LOGFILE NEEDS TO BE CREATED
  set objLOG = objFSO.createtextfile("C:\temp\MSP_DEBUG")
  objLOG.close
  set objLOG = objFSO.opentextfile("C:\temp\MSP_DEBUG", 8)
end if
''READ PASSED COMMANDLINE ARGUMENTS
if (wscript.arguments.count > 0) then                                 ''ARGUMENTS WERE PASSED
  for x = 0 to (wscript.arguments.count - 1)
    objOUT.write vbnewline & now & vbtab & " - ARGUMENT " & (x + 1) & " (ITEM " & x & ") " & " PASSED : " & ucase(objARG.item(x))
    objLOG.write vbnewline & now & vbtab & " - ARGUMENT " & (x + 1) & " (ITEM " & x & ") " & " PASSED : " & ucase(objARG.item(x))
  next
  ''SET DEFAULT ARGUMENTS FOR DEBUG LOGGING , REF #17
  strHDR = "[Logging]"
  strCHG = "LoggingLevel=Debug"
elseif (wscript.arguments.count = 0) then                             ''SET DEFAULT ARGUMENTS FOR DEBUG LOGGING , REF #17
  strHDR = "[Logging]"
  strCHG = "LoggingLevel=Debug"
end if

''------------
''BEGIN SCRIPT
objOUT.write vbnewline & now & " - EXECUTING MSP_DEBUG" & vbnewline
objLOG.write vbnewline & now & " - EXECUTING MSP_DEBUG" & vbnewline
''AUTOMATIC UPDATE, MSP_DEBUG.VBS, REF #2 , REF #69 , REF #68 , FIXES #24
''DOWNLOAD CHKAU.VBS SCRIPT, REF #2 , REF #69 , REF #68
call FILEDL("https://github.com/CW-Khristos/scripts/raw/dev/chkAU.vbs", "chkAU.vbs")
''EXECUTE CHKAU.VBS SCRIPT, REF #69
objOUT.write vbnewline & now & vbtab & vbtab & " - CHECKING FOR UPDATE : MSP_DEBUG : " & strVER
objLOG.write vbnewline & now & vbtab & vbtab & " - CHECKING FOR UPDATE : MSP_DEBUG : " & strVER
intRET = objWSH.run ("cmd.exe /C " & chr(34) & "cscript.exe " & chr(34) & "C:\temp\chkAU.vbs" & chr(34) & " " & _
  chr(34) & strREPO & chr(34) & " " & chr(34) & strBRCH & chr(34) & " " & chr(34) & strDIR & chr(34) & " " & _
  chr(34) & wscript.scriptname & chr(34) & " " & chr(34) & strVER & chr(34) & chr(34), 0, true)
''CHKAU RETURNED - NO UPDATE FOUND , REF #2 , REF #69 , REF #68
intRET = (intRET - vbObjectError)
if ((intRET = 4) or (intRET = 10) or (intRET = 11) or (intRET = 1)) then
  ''PARSE CONFIG.INI FILE , REF #17
  objOUT.write vbnewline & now & vbtab & " - CURRENT CONFIG.INI"
  objLOG.write vbnewline & now & vbtab & " - CURRENT CONFIG.INI"
  strIN = objCFG.readall
  arrIN = split(strIN, vbnewline)
  for intIN = 0 to ubound(arrIN)                                        ''CHECK CONFIG.INI LINE BY LINE
    objOUT.write vbnewline & vbtab & vbtab & arrIN(intIN)
    objLOG.write vbnewline & vbtab & vbtab & arrIN(intIN)
    if (arrIN(intIN) = "[Logging]") then                                ''FOUND SPECIFIED 'HEADER' IN CONFIG.INI
      blnHDR = true
    end if
    if (arrIN(intIN) = "LoggingLevel=Debug") then                       ''STRING TO INJECT ALREADY IN CONFIG.INI
      blnMOD = false
    end if
    if ((blnHDR) and (blnMOD) and (arrIN(intIN) = vbnullstring)) then   ''STRING TO INJECT NOT FOUND, INJECT UNDER CURRENT 'HEADER'
      blnINJ = true
      blnHDR = false
      arrIN(intIN) = "LoggingLevel=Debug" & vbCrlf
    end if
  next
  if ((not blnHDR) and (blnMOD)) then                                   '' '[LOGGING]' HEADER NOT FOUND , REF #17
    blnINJ = true
    redim preserve arrIN(intIN)
    arrIN(intIN) = "[Logging]" & vbCrlf & "LoggingLevel=Debug" & vbCrlf
  end if
  objCFG.close
  set objCFG = nothing
  ''REPLACE CONFIG.INI FILE , REF #17
  if (blnINJ) then
    objOUT.write vbnewline & vbnewline & now & vbtab & " - NEW CONFIG.INI"
    objLOG.write vbnewline & vbnewline & now & vbtab & " - NEW CONFIG.INI"
    strIN = vbnullstring
    set objCFG = objFSO.opentextfile("C:\Program Files\Backup Manager\config.ini", 2)
    for intIN = 0 to ubound(arrIN)
      strIN = strIN & arrIN(intIN) & vbCrlf
      objOUT.write vbnewline & vbtab & vbtab & arrIN(intIN)
      objLOG.write vbnewline & vbtab & vbtab & arrIN(intIN)
    next
    objCFG.write strIN
    objCFG.close
    set objCFG = nothing
  end if
  ''CREATE MSP BACKUP DEBUG FOLDERS , REF #17
  objOUT.write vbnewline & now & vbtab & " - CHECKING 'BACKUPFP.PROTOCOL' LOGGING DIRECTORY"
  objLOG.write vbnewline & now & vbtab & " - CHECKING 'BACKUPFP.PROTOCOL' LOGGING DIRECTORY"
  ''WIN 7/8/10/2K8
  strFOL = "C:\ProgramData\MXB\Backup Manager\logs"
  if (objFSO.folderexists(strFOL)) then
    strFOL = "C:\ProgramData\MXB\Backup Manager\logs\BackupFP.Protocol"
    if (not objFSO.folderexists(strFOL)) then                           ''NEED TO CREATE 'BACKUPFP.PROTOCOL' LOGGING DIRECTORY
      objOUT.write vbnewline & now & vbtab & vbtab & "CREATING LOGGING DIRECTORY : " & strFOL
      objLOG.write vbnewline & now & vbtab & vbtab & "CREATING LOGGING DIRECTORY : " & strFOL
      objFSO.createfolder "C:\ProgramData\MXB\Backup Manager\logs\BackupFP.Protocol"
    end if
  end if
  ''WIN XP/2K3
  strFOL = "C:\Documents and Settings\All Users\Application Data\MXB\Backup Manager\logs"
  if (objFSO.folderexists(strFOL)) then
    strFOL = "C:\Documents and Settings\All Users\Application Data\MXB\Backup Manager\logs\BackupFP.Protocol"
    if (not objFSO.folderexists(strFOL)) then                           ''NEED TO CREATE 'BACKUPFP.PROTOCOL' LOGGING DIRECTORY
      objOUT.write vbnewline & now & vbtab & vbtab & "CREATING LOGGING DIRECTORY : " & strFOL
      objLOG.write vbnewline & now & vbtab & vbtab & "CREATING LOGGING DIRECTORY : " & strFOL
      objFSO.createfolder strFOL
    end if
  end if
end if
''END SCRIPT
call CLEANUP()
''END SCRIPT
''------------

''SUB-ROUTINES
sub FILEDL(strURL, strFILE)                   			                    ''CALL HOOK TO DOWNLOAD FILE FROM URL , 'ERRRET'=11
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
  if (objFSO.fileexists(strSAV)) then
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
  if (objFSO.fileexists(strSAV)) then
    objOUT.write vbnewline & vbnewline & now & vbtab & " - DOWNLOAD : " & strSAV & " : SUCCESSFUL"
    objLOG.write vbnewline & vbnewline & now & vbtab & " - DOWNLOAD : " & strSAV & " : SUCCESSFUL"
    set objHTTP = nothing
  end if
  if (err.number <> 0) then                                             ''ERROR RETURNED , 'ERRRET'=11
    call LOGERR(11)
  end if
end sub

sub HOOK(strCMD)                                                        ''CALL HOOK TO MONITOR OUTPUT OF CALLED COMMAND , 'ERRRET'=12
  on error resume next
  set objHOOK = objWSH.exec(strCMD)
  while (not objHOOK.stdout.atendofstream)
    if (instr(1, strCMD, "takeown /F ") = 0) then                       ''SUPPRESS 'TAKEOWN' SUCCESS MESSAGES
      strIN = objHOOK.stdout.readline
      if (strIN <> vbnullstring) then
        objOUT.write vbnewline & now & vbtab & vbtab & vbtab & strIN 
        objLOG.write vbnewline & now & vbtab & vbtab & vbtab & strIN 
      end if
    end if
  wend
  wscript.sleep 10
  if (instr(1, strCMD, "takeown /F ") = 0) then                         ''SUPPRESS 'TAKEOWN' SUCCESS MESSAGES
    strIN = objHOOK.stdout.readall
    if (strIN <> vbnullstring) then
      objOUT.write vbnewline & now & vbtab & vbtab & vbtab & strIN 
      objLOG.write vbnewline & now & vbtab & vbtab & vbtab & strIN 
    end if
  end if
  set objHOOK = nothing
  if (err.number <> 0) then                                             ''ERROR RETURNED , 'ERRRET'=12
    call LOGERR(12)
  end if
end sub

sub CLEANUP()                                                           ''SCRIPT CLEANUP
  on error resume next
  if (errRET = 0) then                                                  ''SCRIPT COMPLETED SUCCESSFULLY
    objOUT.write vbnewline & vbnewline & now & vbtab & " - MSP_DEBUG COMPLETE : " & errRET & " : " & now
    objLOG.write vbnewline & vbnewline & now & vbtab & " - MSP_DEBUG COMPLETE : " & errRET & " : " & now
    err.clear
  elseif (errRET <> 0) then                                             ''SCRIPT FAILED
    objOUT.write vbnewline & vbnewline & now & vbtab & " - MSP_DEBUG FAILURE : " & errRET & " : " & now
    objLOG.write vbnewline & vbnewline & now & vbtab & " - MSP_DEBUG FAILURE : " & errRET & " : " & now
    ''RAISE CUSTOMIZED ERROR CODE, ERROR CODE WILL BE DEFINED RESTOP NUMBER INDICATING WHICH SECTION FAILED
    call err.raise(vbObjectError + errRET, "MSP_DEBUG", "FAILURE")
  end if
  objLOG.close
  ''EMPTY OBJECTS
  set objCFG = nothing
  set objLOG = nothing
  set objFSO = nothing
  set objWSH = nothing
  set objARG = nothing
  set objOUT = nothing
  set objIN = nothing
  ''END SCRIPT, RETURN ERROR NUMBER
  wscript.quit errRET
end sub
''MSP_LSV.VBS
''NO REQUIRED PARAMETERS / DOES NOT ACCEPT PARAMETERS
''SCRIPT IS DESIGNED TO SIMPLY EXPORT MSP BACKUP SETTINGS USING CLIENTTOOL.EXE UTILITY
''EXPORTS MSP BACKUP SETTINGS TO C:\TEMP\LSV.TXT
''MUST BE USED IN CONJUNCTION WITH MSP BACKUP SYNCHRONIZATION - LSV SYNCHRONIZATION.AMP CUSTOM SERVICE
''WRITTEN BY : CJ BLEDSOE / CJ<@>THECOMPUTERWARRIORS.COM
on error resume next
''SCRIPT VARIABLES
dim strVER, errRET
dim strREPO, strBRCH, strDIR
dim strDLM, intDIFF, retDEL
''SCRIPT OBJECTS
dim objIN, objOUT, objARG
dim objHOOK, objHTTP, objXML
dim objWSH, objFSO, objLOG, objLSV
''DEFAULT SUCCESS
errRET = 0
''VERSION FOR SCRIPT UPDATE, MSP_LSV.VBS, REF #2 , REF #68 , REF #69
strVER = 2
strREPO = "scripts"
strBRCH = "dev"
strDIR = "MSP Backups"
''STDIN / STDOUT
set objIN = wscript.stdin
set objOUT = wscript.stdout
set objARG = wscript.arguments
''CREATE SCRIPTING OBJECTS
set objWSH = createobject("wscript.shell")
set objFSO = createobject("scripting.filesystemobject")
''PREPARE LOGFILE
if (objFSO.fileexists("C:\temp\msp_lsv")) then              ''PREVIOUS LOGFILE EXISTS
  objFSO.deletefile "C:\temp\msp_lsv", true
end if
if (objFSO.fileexists("C:\temp\msp_lsv")) then              ''LOGFILE EXISTS
  objFSO.deletefile "C:\temp\msp_lsv", true
  set objLOG = objFSO.createtextfile("C:\temp\msp_lsv")
  objLOG.close
  set objLOG = objFSO.opentextfile("C:\temp\msp_lsv", 8)
else                                                        ''LOGFILE NEEDS TO BE CREATED
  set objLOG = objFSO.createtextfile("C:\temp\msp_lsv")
  objLOG.close
  set objLOG = objFSO.opentextfile("C:\temp\msp_lsv", 8)
end if
''PREPARE MONITOR FILE
if (objFSO.fileexists("C:\temp\lsv.txt")) then              ''PREVIOUS LOGFILE EXISTS
  objFSO.deletefile "C:\temp\lsv.txt", true
end if
if (objFSO.fileexists("C:\temp\lsv.txt")) then              ''LOGFILE EXISTS
  objFSO.deletefile "C:\temp\lsv.txt", true
  set objLSV = objFSO.createtextfile("C:\temp\lsv.txt")
  objLSV.close
  set objLSV = objFSO.opentextfile("C:\temp\lsv.txt", 8)
else                                                        ''LOGFILE NEEDS TO BE CREATED
  set objLSV = objFSO.createtextfile("C:\temp\lsv.txt")
  objLSV.close
  set objLSV = objFSO.opentextfile("C:\temp\lsv.txt", 8)
end if
''CHECK EXECUTION METHOD OF SCRIPT
strIN = lcase(mid(wscript.fullname, instrrev(wscript.fullname, "\") + 1))
if (strIN <> "cscript.exe") Then                            ''LAUNCHED VIA WSCRIPT, RE-LAUNCH WITH CSCRIPT
  objOUT.write vbnewline & "SCRIPT LAUNCHED VIA EXPLORER, EXECUTING SCRIPT VIA CSCRIPT..."
  objWSH.run "cscript.exe //nologo " & chr(34) & Wscript.ScriptFullName & chr(34)
  wscript.quit
end if
''NO ARGUMENTS REQUIRED
''READ PASSED COMMANDLINE ARGUMENTS
'if (wscript.arguments.count > 0) then                      ''ARGUMENTS WERE PASSED
'  for x = 0 to (wscript.arguments.count - 1)
'    objOUT.write vbnewline & now & vbtab & "ARGUMENT " & (x + 1) & " (ITEM " & x & ") " & " PASSED : " & ucase(objARG.item(x))
'  next 
'end if

''------------
''BEGIN SCRIPT
if (errRET = 0) then                                        ''ARGUMENTS PASSED , CONTINUE SCRIPT
  objOUT.write vbnewline & vbnewline & now & vbtab & " - EXECUTING MSP_LSV"
  objLOG.write vbnewline & vbnewline & now & vbtab & " - EXECUTING MSP_LSV"
	''AUTOMATIC UPDATE, MSP_LSV.VBS, REF #2 , REF #69 , REF #68 , FIXES #32 , REF #71
  ''DOWNLOAD CHKAU.VBS SCRIPT, REF #2 , REF #69 , REF #68
  call FILEDL("https://github.com/CW-Khristos/scripts/raw/dev/chkAU.vbs", "chkAU.vbs")
  ''EXECUTE CHKAU.VBS SCRIPT, REF #69
  objOUT.write vbnewline & now & vbtab & vbtab & " - CHECKING FOR UPDATE : MSP_LSV : " & strVER
  objLOG.write vbnewline & now & vbtab & vbtab & " - CHECKING FOR UPDATE : MSP_LSV : " & strVER
  intRET = objWSH.run ("cmd.exe /C " & chr(34) & "cscript.exe " & chr(34) & "C:\temp\chkAU.vbs" & chr(34) & " " & _
    chr(34) & strREPO & chr(34) & " " & chr(34) & strBRCH & chr(34) & " " & chr(34) & strDIR & chr(34) & " " & _
    chr(34) & wscript.scriptname & chr(34) & " " & chr(34) & strVER & chr(34) & chr(34), 0, true)
  ''CHKAU RETURNED - NO UPDATE FOUND , REF #2 , REF #69 , REF #68
  intRET = (intRET - vbObjectError)
  if ((intRET = 4) or (intRET = 10) or (intRET = 11) or (intRET = 1)) then
    ''EXPORT MSP BACKUP SETTINGS USING CLIENTTOOL UTILITY
    'call HOOK("C:\Program Files\Backup Manager\clienttool.exe control.setting.list > " & chr(34) & "C:\temp\lsv.txt" & chr(34))
    set objHOOK = objWSH.exec("C:\Program Files\Backup Manager\clienttool.exe control.setting.list")
    strIN = objHOOK.stdout.readall
    arrIN = split(strIN, vbnewline)
    ''WRITE SCRIPT LOGFILE
    for intIN = 0 to ubound(arrIN)                              ''CHECK SETTINGS LINE BY LINE, EXCLUDE THE 'C:\WINDOWS\TEMP' AND 'C:\TEMP' DIRECTORIES TO AVOID FALSE MONITOR ALERTS
      if ((instr(1, lcase(arrIN(intIN)), "c:\") = 0) and _
        (instr(1,lcase(arrIN(intIN)), "\temp") = 0)) then
          objOUT.write vbnewline & now & vbtab & arrIN(intIN)
          objLOG.write vbnewline & now & vbtab & arrIN(intIN)
      end if
    next
    intIN = 0
    ''WRITE MONITOR FILE
    for intIN = 0 to ubound(arrIN)                              ''CHECK SETTINGS LINE BY LINE, EXCLUDE THE 'C:\WINDOWS\TEMP' AND 'C:\TEMP' DIRECTORIES TO AVOID FALSE MONITOR ALERTS
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
  end if
elseif (errRET <> 0) then                                   ''NO ARGUMENTS PASSED , END SCRIPT , 'ERRRET'=1
  call LOGERR(errRET)
end if
''END SCRIPT
call CLEANUP()
''END SCRIPT
''------------

''SUB-ROUTINES
sub FILEDL(strURL, strFILE)                                 ''CALL HOOK TO DOWNLOAD FILE FROM URL , 'ERRRET'=11
  strSAV = vbnullstring
  ''SET DOWNLOAD PATH
  strSAV = "C:\temp\" & strFILE
  objOUT.write vbnewline & now & vbtab & vbtab & vbtab & "HTTPDOWNLOAD-------------DOWNLOAD : " & strURL & " : SAVE AS :  " & strSAV
  objLOG.write vbnewline & now & vbtab & vbtab & vbtab & "HTTPDOWNLOAD-------------DOWNLOAD : " & strURL & " : SAVE AS :  " & strSAV
  if objFSO.fileexists(strSAV) then
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
  if (err.number <> 0) then                                 ''ERROR RETURNED DURING UPDATE CHECK , 'ERRRET'=11
    call LOGERR(11)
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
      objOUT.write vbnewline & vbnewline & now & vbtab & " - SCRIPT REQUIRES USER TO GRANT SERVICE LOGON"
      objLOG.write vbnewline & vbnewline & now & vbtab & " - SCRIPT REQUIRES USER TO GRANT SERVICE LOGON"
  end select
end sub

sub CLEANUP()                                 			        ''SCRIPT CLEANUP
  on error resume next
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
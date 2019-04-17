''QSMART.VBS
''DESIGNED TO QUERY AND REPORT SMART STATUS FOR ALL CONNECTED DRIVES
''WRITTEN BY : CJ BLEDSOE / CJ<@>THECOMPUTERWARRIORS.COM
on error resume next
''SCRIPT VARIABLES
dim errRET, strVER
dim strIN, strCOMP
dim arrDRV(), arrSMART()
dim intDRV, intTOT, intSMART
''SCRIPT OBJECTS
dim objWMI, objFPD, arrFPD
dim objLOG, objHOOK, objEXEC, objHTTP, objXML
dim objIN, objOUT, objARG, objWSH, objFSO
''VERSION FOR SCRIPT UPDATE, SMART.VBS, REF #2
strVER = 1
''DEFAULT SUCCESS
errRET = 0
''INITIALIZE ENUMERATED DRIVE ARRAY
redim arrDRV(0)
''INITIALIZE SMART ATTRIBUTE ARRAY
redim arrSMART(0,0)
arrSMART(0,0) = vbnullstring
''STDIN / STDOUT
strCOMP = "."
set objIN = wscript.stdin
set objOUT = wscript.stdout
set objARG = wscript.arguments
''OBJECTS FOR LOCATING FOLDERS
set objWSH = createobject("wscript.shell")
set objFSO = createobject("scripting.filesystemobject")
''WMI OBJECTS
Set objWMI = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strCOMP & "\root\wmi")
set objFPD = objWMI.instancesof("MSStorageDriver_FailurePredictData", 1) ''=" & chr(34) & strDRV & chr(34))
''PREPARE LOGFILE
if (objFSO.fileexists("C:\temp\QSMART")) then              ''LOGFILE EXISTS
  objFSO.deletefile "C:\temp\QSMART", true
  set objLOG = objFSO.createtextfile("C:\temp\QSMART")
  objLOG.close
  set objLOG = objFSO.opentextfile("C:\temp\QSMART", 8)
else                                                        ''LOGFILE NEEDS TO BE CREATED
  set objLOG = objFSO.createtextfile("C:\temp\QSMART")
  objLOG.close
  set objLOG = objFSO.opentextfile("C:\temp\QSMART", 8)
end if
''READ PASSED COMMANDLINE ARGUMENTS
if (wscript.arguments.count > 0) then                       ''ARGUMENTS WERE PASSED
  for x = 0 to (wscript.arguments.count - 1)
    objOUT.write vbnewline & now & vbtab & " - ARGUMENT " & (x + 1) & " (ITEM " & x & ") " & " PASSED : " & ucase(objARG.item(x))
    objLOG.write vbnewline & now & vbtab & " - ARGUMENT " & (x + 1) & " (ITEM " & x & ") " & " PASSED : " & ucase(objARG.item(x))
  next 
  if (wscript.arguments.count > 0) then                     ''SET USER , PASSWORD , AND OPERATION LEVEL VARIABLES
  else                                                      ''NOT ENOUGH ARGUMENTS PASSED, END SCRIPT
    'errRET = 1
    'call CLEANUP()
  end if
else                                                        ''NO ARGUMENTS PASSED, END SCRIPT
  'errRET = 1
  'call CLEANUP()
end if

''------------
''BEGIN SCRIPT
objOUT.write vbnewline & vbnewline & now & vbtab & " - EXECUTING QSMART" & vbnewline
objLOG.write vbnewline & vbnewline & now & vbtab & " - EXECUTING QSMART" & vbnewline
''AUTOMATIC UPDATE, QSMART.VBS, REF #2 , FIXES #42
call CHKAU()
''CHECK FOR SMARTCTL.EXE IN C:\TEMPERATURE
if (not objFSO.fileexists("c:\temp\smartctl.exe")) then
  call FILEDL("https://github.com/CW-Khristos/scripts/raw/dev/SMART/smartctl.exe", "smartctl.exe")
end if
''GET LIST OF AVAILABLE DRIVES , 'ERRRET'=2
intDRV = 0
intSMART = 0
objOUT.write vbnewline & vbnewline & now & vbtab & vbtab & " - ENUMERATING CONNECTED DRIVES" & vbnewline
objLOG.write vbnewline & vbnewline & now & vbtab & vbtab & " - ENUMERATING CONNECTED DRIVES" & vbnewline
set objEXEC = objWSH.exec("c:\temp\smartctl.exe --scan-open")
while (not objEXEC.stdout.atendofstream)
  strIN = objEXEC.stdout.readline
  'objOUT.write vbnewline & now & vbtab & vbtab & strIN
  'objLOG.write vbnewline & now & vbtab & vbtab & strIN
  if (trim(strIN) <> vbnullstring) then
    redim arrSMART((intDRV + 1), 0)
    redim preserve arrDRV(intDRV + 1)
    arrDRV(intDRV) = trim(split(strIN, " ")(0))
    intDRV = (intDRV + 1)
    intTOT = intDRV
  end if
  if (err.number <> 0) then
    call LOGERR(2)
  end if
wend
set objEXEC = nothing
err.clear
for intDRV = 0 to ubound(arrDRV)
  objOUT.write vbnewline & now & vbtab & vbtab & vbtab & arrDRV(intDRV)
next
''QUERY SMART STATUS FOR ALL ENUMERATED DRIVES
if (intTOT > 0) then
  for intDRV = 0 to (intTOT - 1)
    if (arrDRV(intDRV) <> vbnullstring) then
      set objEXEC = objWSH.exec("c:\temp\smartctl.exe -A " & arrDRV(intDRV))
      while (not objEXEC.stdout.atendofstream)
        strIN = objEXEC.stdout.readline
        'objOUT.write vbnewline & now & vbtab & vbtab & strIN
        'objLOG.write vbnewline & now & vbtab & vbtab & strIN
        if (trim(strIN) <> vbnullstring) then
          if ((instr(1, strIN, "smartctl") = 0) and (instr(1, strIN, "Copyright (C)") = 0) and (instr(1, strIN, "=== START") = 0) _ 
            and (instr(1, strIN, "SMART Attributes Data") = 0) and (instr(1, strIN, "Vendor Specific SMART") = 0) and (instr(1, strIN, "ID#") = 0)) then
              redim preserve arrSMART(intTOT, intSMART + 1)
              'wscript.echo vbnewline & strIN & vbnewline
              arrSMART(intDRV, intSMART) = trim(split(strIN, " ")(1)) & "[" & trim(split(strIN, " ")(ubound(split(strIN, " ")))) & "]"
              intSMART = (intSMART + 1)
          end if
        end if
        if (err.number <> 0) then
          call LOGERR(2)
        end if
        wscript.sleep 1000
      wend
      set objEXEC = nothing
    end if
  next
end if
for intDRV = 0 to (intTOT - 1)
  for intSMART = 0 to ubound(arrSMART(intDRV))
    objOUT.write vbnewline & vbtab & now & vbtab & arrSMART(intDRV, intSMART)
  next
next

''END SCRIPT
call CLEANUP()
''END SCRIPT
''------------
 
for each objDRV in objFPD
	arrFPD = objDRV.VendorSpecific
	objOUT.writeline vbnewline & "OBTAINING SMART STATUS : " & objDRV.instancename & " : "
  for intPOS = 0 to ubound(arrFPD)
    intATT = (intPOS * 12)
    if ((intATT + 2) < ubound(arrFPD)) then
      select case (arrFPD(intATT + 2))
        ''ROTATIONAL
        case 1
          objOUT.write "RAW READ ERROR RATE" & vbtab & ":"
          call wrtVAL(intPOS)
        case 5
          objOUT.write "REALLOCATED SECTOR COUNT" & vbtab & ":"
          call wrtVAL(intPOS)
        case 7
          objOUT.write "SEEK ERROR COUNT" & vbtab & ":"
          call wrtVAL(intPOS)
        case 9
          objOUT.write "POWER ON HOURS" & vbtab & ":"
          call wrtVAL(intPOS)
        case 10
          objOUT.write "SPIN-UP RETRY COUNT" & vbtab & ":"
          call wrtVAL(intPOS)
        ''SSD
        case 170
          objOUT.write "AVAILABLE SPACE" & vbtab & ":"
          call wrtVAL(intPOS)
        case 171
          objOUT.write "PROGRAM FAIL" & vbtab & ":"
          call wrtVAL(intPOS)
        case 172
          objOUT.write "ERASE FAIL" & vbtab & ":"
          call wrtVAL(intPOS)
        case 173
          objOUT.write "WEAR LEVELING" & vbtab & ":"
          call wrtVAL(intPOS)
        case 176
          objOUT.write "ERASE FAIL" & vbtab & ":"
          call wrtVAL(intPOS)
        case 177
          objOUT.write "WEAR RANGE" & vbtab & ":"
          call wrtVAL(intPOS)
        case 179
          objOUT.write "USED RESERVED" & vbtab & ":"
          call wrtVAL(intPOS)
        case 180
          objOUT.write "UN-USED RESERVED" & vbtab & ":"
          call wrtVAL(intPOS)
        case 181
          objOUT.write "PROGRAM FAIL COUNT" & vbtab & ":"
          call wrtVAL(intPOS)
        case 182
          objOUT.write "ERASE FAIL COUNT" & vbtab & ":"
          call wrtVAL(intPOS)
        case 230
          objOUT.write "DRIVE LIFE PROTECTION" & vbtab & ":"
          call wrtVAL(intPOS)
        case 231
          objOUT.write "LIFE LEFT" & vbtab & ":"
          call wrtVAL(intPOS)
        case 232
          objOUT.write "ENDURANCE REMAINING" & vbtab & ":"
          call wrtVAL(intPOS)
        case 233
          objOUT.write "MEDIA WEAROUT" & vbtab & ":"
          call wrtVAL(intPOS)
        case 234
          objOUT.write "AVG / MAX ERASE" & vbtab & ":"
          call wrtVAL(intPOS)
        case 235
          objOUT.write "GOOD BLOCK / SYSTEM FREE COUNT" & vbtab & ":"
          call wrtVAL(intPOS)
        ''ROTATIONAL
        case 194
          objOUT.write "TEMPERATURE (C)" & vbtab & ":"
          call wrtVAL(intPOS)
        case 196
          objOUT.write "REALLOCATION EVENT COUNT" & vbtab & ":"
          call wrtVAL(intPOS)
        case 197
          objOUT.write "CURRENT PENDING SECTOR COUNT" & vbtab & ":"
          call wrtVAL(intPOS)
        case 198
          objOUT.write "OFFLINE / UNRECOVERABLE SECTOR COUNT" & vbtab & ":"
          call wrtVAL(intPOS)
        case else
          objOUT.write "UNKNOWN ATTRIBUTE" & vbtab & ":"
          call wrtVAL(intPOS)
      end select
    end if
	next
next

''SUB-ROUTINES
sub wrtVAL(posFPD)
    for y = 0 to 11
      if (((posFPD * 12) + y) <= ubound(arrFPD)) then
        objOUT.write " " & arrFPD((posFPD * 12) + y) & ","
      end if
    next
    objOUT.write vbnewline
end sub

sub CHKAU()																					''CHECK FOR SCRIPT UPDATE, QSNART.VBS, REF #2 , FIXES #42
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
	if objXML.load("https://github.com/CW-Khristos/scripts/raw/dev/version.xml") then
		set colVER = objXML.documentelement
		for each objSCR in colVER.ChildNodes
			''LOCATE CURRENTLY RUNNING SCRIPT
			if (lcase(objSCR.nodename) = lcase(wscript.scriptname)) then
				''CHECK LATEST VERSION
				if (cint(objSCR.text) > cint(strVER)) then
					objOUT.write vbnewline & now & " - UPDATING " & objSCR.nodename & " : " & objSCR.text & vbnewline
					objLOG.write vbnewline & now & " - UPDATING " & objSCR.nodename & " : " & objSCR.text & vbnewline
					''DOWNLOAD LATEST VERSION OF SCRIPT
					call FILEDL("https://github.com/CW-Khristos/scripts/raw/dev/qSMART.vbs", wscript.scriptname)
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

sub CLEANUP()                                           ''SCRIPT CLEANUP
  if (errRET = 0) then                                 ''SCRIPT COMPLETED SUCCESSFULLY
    objOUT.write vbnewline & vbnewline & now & vbtab & " - QSMART COMPLETE : " & now
    objLOG.write vbnewline & vbnewline & now & vbtab & " - QSMART COMPLETE : " & now
    err.clear
  elseif (errRET <> 0) then                            ''SCRIPT FAILED
    objOUT.write vbnewline & vbnewline & now & vbtab & " - QSMART FAILURE : " & errRET & " : " & now
    objLOG.write vbnewline & vbnewline & now & vbtab & " - QSMART FAILURE : " & errRET & " : " & now
    ''RAISE CUSTOMIZED ERROR CODE, ERROR CODE WILL BE DEFINED RESTOP NUMBER INDICATING WHICH SECTION FAILED
    call err.raise(vbObjectError + errRET, "QSMART", "fail")
  end if
  ''EMPTY OBJECTS
  set objFPD = nothing
  set objWMI = nothing
  set objEXEC = nothing
  set objLOG = nothing
  set objFSO = nothing
  set objWSH = nothing
  set objARG = nothing
  set objOUT = nothing
  set objIN = nothing
  ''END SCRIPT, RETURN ERROR
  wscript.quit err.number
end sub

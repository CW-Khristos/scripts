''QSMART.VBS
''DESIGNED TO QUERY AND REPORT SMART STATUS FOR ALL CONNECTED DRIVES
''WRITTEN BY : CJ BLEDSOE / CJ<@>THECOMPUTERWARRIORS.COM
'on error resume next
''SCRIPT VARIABLES
dim errRET, strVER
dim strIN, strCOMP
dim arrDRV(), arrSMART()
dim intDRV, intTOT, intSMART
''SCRIPT OBJECTS
dim objWMI, objFPD, arrFPD
dim objIN, objOUT, objARG, objWSH, objFSO
dim objLOG, objHOOK, objEXEC, objHTTP, objXML
''VERSION FOR SCRIPT UPDATE , QSMART.VBS, REF #2 , REF #42 , FIXES #44
strVER = 2
''DEFAULT SUCCESS
errRET = 0
''INITIALIZE ENUMERATED DRIVE ARRAY , QSMART.VBS, REF #2 , REF #42 , FIXES #44
redim arrDRV(0)
''INITIALIZE SMART ATTRIBUTE ARRAY , QSMART.VBS, REF #2 , REF #42 , FIXES #44
redim arrSMART(1,9)
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
if (objFSO.fileexists("C:\temp\QSMART")) then               ''LOGFILE EXISTS
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
    strDRV = objARG.item(0)
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
''AUTOMATIC UPDATE , 'ERRRET'=10 , QSMART.VBS , REF #2 , REF #42 , FIXES #42
call CHKAU()
''CHECK FOR SMARTCTL.EXE IN C:\TEMP , QSMART.VBS, REF #2 , REF #42 , FIXES #44
if (not objFSO.fileexists("c:\temp\smartctl.exe")) then
  call FILEDL("https://github.com/CW-Khristos/scripts/raw/dev/SMART/smartctl.exe", "smartctl.exe")
end if
''GET LIST OF AVAILABLE DRIVES , 'ERRRET'=2 , QSMART.VBS, REF #2 , REF #42 , FIXES #44
intDRV = 0
intSMART = 0
objOUT.write vbnewline & vbnewline & now & vbtab & vbtab & " - ENUMERATING CONNECTED DRIVES" & vbnewline
objLOG.write vbnewline & vbnewline & now & vbtab & vbtab & " - ENUMERATING CONNECTED DRIVES" & vbnewline
''ENUMERATE AND COLLECT DRIVE LIST , QSMART.VBS, REF #2 , REF #42 , FIXES #44
set objEXEC = objWSH.exec("c:\temp\smartctl.exe --scan-open")
while (not objEXEC.stdout.atendofstream)
  strIN = objEXEC.stdout.readline
  if (trim(strIN) <> vbnullstring) then
    ''RE-SIZE 'ARRSMART'('DRIVE INDEX', 'SMART INDEX') ARRAY , QSMART.VBS, REF #2 , REF #42 , REF #44
    redim arrSMART((intDRV + 2), 9)
    ''RE-SIZE 'ARRDRV'('DRIVE INDEX') ARRAY , QSMART.VBS, REF #2 , REF #42 , REF #44
    redim preserve arrDRV(intDRV + 1)
    ''COLLECT 'SMARTCTL' DRIVE PATH , QSMART.VBS, REF #2 , REF #42 , FIXES #44
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
''LIST COLLECTED 'SMARTCTL' DRIVE LIST , QSMART.VBS, REF #2 , REF #42 , FIXES #44
if (strDRV = vbnullstring) then
  for intDRV = 0 to ubound(arrDRV)
    if (arrDRV(intDRV) <> vbnullstring) then
      objOUT.write vbnewline & now & vbtab & vbtab & vbtab & arrDRV(intDRV)
    end if
  next
  ''QUERY 'SMART STATUS' FOR ALL ENUMERATED DRIVES , 'ERRRET'=3 , QSMART.VBS, REF #2 , REF #42 , FIXES #44
  objOUT.write vbnewline & vbnewline & now & vbtab & vbtab & " - QUERYING DRIVES' 'SMART' STATUS" & vbnewline
  objLOG.write vbnewline & vbnewline & now & vbtab & vbtab & " - QUERYING DRIVES' 'SMART' STATUS" & vbnewline
  if (intTOT > 0) then
    intATT = 0
    intCOL = 9
    ''ENUMERATE THROUGH EACH DRIVE
    for intDRV = 0 to (intTOT)
      ''RESET 'SMART' INDEX
      intATT = 0
      intSMART = 0
      if (arrDRV(intDRV) <> vbnullstring) then
        ''QUERY 'SMART' ATTRIBUTES USING 'SMARTCTL' , QSMART.VBS, REF #2 , REF #42 , FIXES #44
        objOUT.write vbnewline & vbnewline & now & vbtab & vbtab & " - QUERYING DRIVE : " & arrDRV(intDRV) & vbnewline
        objLOG.write vbnewline & vbnewline & now & vbtab & vbtab & " - QUERYING DRIVE : " & arrDRV(intDRV) & vbnewline
        set objEXEC = objWSH.exec("c:\temp\smartctl.exe -A " & arrDRV(intDRV))
        ''ENUMERATE THROUGH EACH 'SMART' ATTRIBUTE
        while (not objEXEC.stdout.atendofstream)
          on error resume next
          strIN = trim(objEXEC.stdout.readline)
          if (strIN <> vbnullstring) then
            ''EXCLUDE 'HEADERS'
            if ((instr(1, strIN, "smartctl") = 0) and (instr(1, strIN, "Copyright (C)") = 0) and (instr(1, strIN, "=== START") = 0) _ 
              and (instr(1, strIN, "SMART Attributes Data") = 0) and (instr(1, strIN, "Vendor Specific SMART") = 0) and (instr(1, strIN, "ID#") = 0)) then
                'objOUT.write vbnewline & now & vbtab & vbtab & split(strIN, " ")(1)
                'objLOG.write vbnewline & now & vbtab & vbtab & split(strIN, " ")(1)
                ''PARSE 'SMARTCTL' OUTPUT , QSMART.VBS, REF #2 , REF #42 , FIXES #44
                'for intTMP = 1 to ubound(split(strIN, " "))
                  if ((instr(1, strIN, "  ") and (split(strIN, " ")(1) <> vbnullstring))) then
                    ''VALIDATE 'SMART' ATTRIBUTE NAME , QSMART.VBS, REF #2 , REF #42 , FIXES #44
                    'objOUT.write vbnewline & blnSMART(split(strIN, " ")(intTMP))
                    'objLOG.write vbnewline & blnSMART(split(strIN, " ")(intTMP))
                    if (blnSMART(split(strIN, " ")(1))) then
                      ''RE-SIZE 'ARRSMART'('DRIVE INDEX', 'SMART INDEX') ARRAY , QSMART.VBS, REF #2 , REF #42 , REF #44
                      if (intATT >= intCOL) then
                        intCOL = intATT + 1
                        wscript.echo "TEST"
                        wscript.echo vbnewline & vbtab & "REDEFINE : " & ubound(arrSMART , intDRV + 1) & vbtab & " / " & vbtab & intATT
                        redim preserve arrSMART(intTOT, intCOL)
                      end if
                      ''COLLECT 'SMARTCTL' DRIVE SMART ATTRIBUTES , QSMART.VBS, REF #2 , REF #42 , FIXES #44
                      'objOUT.write vbnewline & "DRIVE : " & intDRV & " - SMART ATT : " & intSMART 
                      'objOUT.write vbnewline & intSMART & vbtab & trim(split(strIN, " ")(1)) & "[" & trim(split(strIN, "  ")(ubound(split(strIN, "  ")))) & "]" & vbnewline
                      'objOUT.write vbnewline & strIN
                      'objLOG.write vbnewline & trim(split(strIN, " ")(1)) & "[" & trim(split(strIN, "  ")(ubound(split(strIN, "  ")))) & "]" & vbnewline
                      arrSMART(intDRV, intSMART) = trim(split(strIN, " ")(1)) & "[" & trim(split(strIN, "  ")(ubound(split(strIN, "  ")))) & "]"
                      'wscript.echo vbnewline & vbtab & ubound(arrSMART , intDRV + 1) & vbtab & " / " & vbtab & intATT
                      intSMART = (intSMART + 1)
                      intATT = (intATT + 1)
                      'exit for
                    end if
                  end if
                'next
            end if
          end if
          if (err.number <> 0) then
            call LOGERR(3)
          end if
          wscript.sleep 200
        wend
        set objEXEC = nothing
      end if
    next
  end if
elseif (strDRV <> vbnullstring) then
  intDRV = 0
  ''RESET 'SMART' INDEX
  intATT = 0
  intCOL = 9
  intSMART = 0
  ''QUERY 'SMART STATUS' FOR ALL ENUMERATED DRIVES , 'ERRRET'=3 , QSMART.VBS, REF #2 , REF #42 , FIXES #44
  objOUT.write vbnewline & vbnewline & now & vbtab & vbtab & " - QUERYING DRIVES' 'SMART' STATUS" & vbnewline
  objLOG.write vbnewline & vbnewline & now & vbtab & vbtab & " - QUERYING DRIVES' 'SMART' STATUS" & vbnewline
  ''QUERY 'SMART' ATTRIBUTES USING 'SMARTCTL' , QSMART.VBS, REF #2 , REF #42 , FIXES #44
  objOUT.write vbnewline & vbnewline & now & vbtab & vbtab & " - QUERYING DRIVE : " & arrDRV(intDRV) & vbnewline
  objLOG.write vbnewline & vbnewline & now & vbtab & vbtab & " - QUERYING DRIVE : " & arrDRV(intDRV) & vbnewline
  set objEXEC = objWSH.exec("c:\temp\smartctl.exe -A " & strDRV)
  ''ENUMERATE THROUGH EACH 'SMART' ATTRIBUTE
  while (not objEXEC.stdout.atendofstream)
    on error resume next
    strIN = trim(objEXEC.stdout.readline)
    if (strIN <> vbnullstring) then
      ''EXCLUDE 'HEADERS'
      if ((instr(1, strIN, "smartctl") = 0) and (instr(1, strIN, "Copyright (C)") = 0) and (instr(1, strIN, "=== START") = 0) _ 
        and (instr(1, strIN, "SMART Attributes Data") = 0) and (instr(1, strIN, "Vendor Specific SMART") = 0) and (instr(1, strIN, "ID#") = 0)) then
          'objOUT.write vbnewline & now & vbtab & vbtab & split(strIN, " ")(1)
          'objLOG.write vbnewline & now & vbtab & vbtab & split(strIN, " ")(1)
          ''PARSE 'SMARTCTL' OUTPUT , QSMART.VBS, REF #2 , REF #42 , FIXES #44
          'for intTMP = 1 to ubound(split(strIN, " "))
            if ((instr(1, strIN, "  ") and (split(strIN, " ")(1) <> vbnullstring))) then
              ''VALIDATE 'SMART' ATTRIBUTE NAME , QSMART.VBS, REF #2 , REF #42 , FIXES #44
              'objOUT.write vbnewline & blnSMART(split(strIN, " ")(intTMP))
              'objLOG.write vbnewline & blnSMART(split(strIN, " ")(intTMP))
              if (blnSMART(split(strIN, " ")(1))) then
                ''RE-SIZE 'ARRSMART'('DRIVE INDEX', 'SMART INDEX') ARRAY , QSMART.VBS, REF #2 , REF #42 , REF #44
                if (intATT >= intCOL) then
                  intCOL = intATT + 1
                  wscript.echo "TEST"
                  wscript.echo vbnewline & vbtab & "REDEFINE : " & ubound(arrSMART , intDRV + 1) & vbtab & " / " & vbtab & intATT
                  redim preserve arrSMART(intTOT, intCOL)
                end if
                ''COLLECT 'SMARTCTL' DRIVE SMART ATTRIBUTES , QSMART.VBS, REF #2 , REF #42 , FIXES #44
                'objOUT.write vbnewline & "DRIVE : " & intDRV & " - SMART ATT : " & intSMART 
                'objOUT.write vbnewline & intSMART & vbtab & trim(split(strIN, " ")(1)) & "[" & trim(split(strIN, "  ")(ubound(split(strIN, "  ")))) & "]" & vbnewline
                'objOUT.write vbnewline & strIN
                'objLOG.write vbnewline & trim(split(strIN, " ")(1)) & "[" & trim(split(strIN, "  ")(ubound(split(strIN, "  ")))) & "]" & vbnewline
                arrSMART(intDRV, intSMART) = trim(split(strIN, " ")(1)) & "[" & trim(split(strIN, "  ")(ubound(split(strIN, "  ")))) & "]"
                'wscript.echo vbnewline & vbtab & ubound(arrSMART , intDRV + 1) & vbtab & " / " & vbtab & intATT
                intSMART = (intSMART + 1)
                intATT = (intATT + 1)
                'exit for
              end if
            end if
          'next
      end if
    end if
    if (err.number <> 0) then
      call LOGERR(3)
    end if
    wscript.sleep 200
  wend
  set objEXEC = nothing
end if
''LIST COLLECTED 'SMARTCTL' DRIVE SMART ATTRIBUTES , QSMART.VBS , REF #2 , REF #42 , FIXES #44
objOUT.write vbnewline & vbnewline & now & vbtab & vbtab & " - QSMART 'SMART' STATUS" & vbnewline
objLOG.write vbnewline & vbnewline & now & vbtab & vbtab & " - QSMART 'SMART' STATUS" & vbnewline
if (strDRV = vbnullstring) then
  for intDRV = 0 to ubound(arrDRV)
    if (arrDRV(intDRV) <> vbnullstring) then
      objOUT.write vbnewline & vbtab & now & vbtab & arrDRV(intDRV)
      objLOG.write vbnewline & vbtab & now & vbtab & arrDRV(intDRV)
      for intSMART = 0 to 9'ubound(arrSMART, intDRV + 1)
        'if (arrSMART(intDRV, intSMART) <> vbnullstring) then
          objOUT.write vbnewline & vbtab & now & vbtab & arrSMART(intDRV, intSMART)
          objLOG.write vbnewline & vbtab & now & vbtab & arrSMART(intDRV, intSMART)
        'end if
      next
    end if
  next
elseif (strDRV <> vbnullstring) then
  objOUT.write vbnewline & vbtab & now & vbtab & arrDRV(intDRV)
  objLOG.write vbnewline & vbtab & now & vbtab & arrDRV(intDRV)
  for intSMART = 0 to 9'ubound(arrSMART, intDRV + 1)
    'if (arrSMART(intDRV, intSMART) <> vbnullstring) then
      objOUT.write vbnewline & vbtab & now & vbtab & arrSMART(intDRV, intSMART)
      objLOG.write vbnewline & vbtab & now & vbtab & arrSMART(intDRV, intSMART)
    'end if
  next
end if
''END SCRIPT
call CLEANUP()
''END SCRIPT
''------------
 
 ''FUNCTIONS
function blnSMART(varVAL)												            ''VALIDATE 'SMART' ATTRIBUTE NAME , QSMART.VBS, REF #2 , REF #42 , FIXES #44
  blnSMART = false
  varVAL = trim(ucase(replace(varVAL, "_", " ")))
  select case varVAL
    ''ROTATIONAL
    ''SMART ID 1
    case "RAW READ ERROR RATE"
      blnSMART = true
    ''SMART ID 5
    case "REALLOCATED SECTOR CT"
      blnSMART = true
    ''SMART ID 7
    case "SEEK ERROR RATE"
      blnSMART = true
    ''SMART ID 9
    case "POWER ON HOURS"
      blnSMART = true
    ''SMART ID 10
    case "SPIN RETRY COUNT"
      blnSMART = true
    ''SMART ID 194
    case "TEMPERATURE CELSIUS"
      blnSMART = true
    ''SMART ID 196
    case "REALLOCATION EVENT COUNT"
      blnSMART = true
    ''SMART ID 197
    case "CURRENT PENDING SECTOR"
      blnSMART = true
    ''SMART ID 198
    case "OFFLINE UNCORRECTABLE"
      blnSMART = true
    ''SSD
    ''SMART ID 170
    case "AVAILABLE SPACE"
      blnSMART = true
    ''SMART ID 171
    case "PROGRAM FAIL"
      blnSMART = true
    ''SMART ID 172
    case "ERASE FAIL"
      blnSMART = true
    ''SMART ID 173
    case "WEAR LEVELING"
      blnSMART = true
    ''SMART ID 176
    case "ERASE FAIL"
      blnSMART = true
    ''SMART ID 177
    case "WEAR RANGE DELTA"
      blnSMART = true
    ''SMART ID 179
    case "USED RESERVED"
      blnSMART = true
    ''SMART ID 180
    case "UN-USED RESERVED"
      blnSMART = true
    ''SMART ID 181
    case "PROGRAM FAIL COUNT"
      blnSMART = true
    ''SMART ID 182
    case "ERASE FAIL COUNT"
      blnSMART = true
    ''SMART ID 196
    case "REALLOCATED EVENT COUNT"
      blnSMART = true
    ''SMART ID 230
    case "DRIVE LIFE PROTECTION"
      blnSMART = true
    ''SMART ID 231
    case "SSD LIFE LEFT"
      blnSMART = true
    ''SMART ID 232
    case "ENDURANCE REMAINING"
      blnSMART = true
    ''SMART ID 233
    case "MEDIA WEAROUT"
      blnSMART = true
    ''SMART ID 234
    case "AVG / MAX ERASE"
      blnSMART = true
    ''SMART ID 235
    case "GOOD BLOCK / SYSTEM FREE COUNT"
      blnSMART = true
    ''UNKNOWNS
    case else
  end select
end function

''SUB-ROUTINES
sub CHKAU()																					        ''CHECK FOR SCRIPT UPDATE , QSMART.VBS , REF #2 , FIXES #42
  ''NO LONGER REQUIRED WITH NCENTRAL 2021; SCRIPTS ARE PLACED IN INDIVIDUAL 'TASK' DIRECTORIES
  ''REMOVE WINDOWS AGENT CACHED VERSION OF SCRIPT
  'if (objFSO.fileexists("C:\Program Files (x86)\N-Able Technologies\Windows Agent\cache\" & wscript.scriptname)) then
  '  objFSO.deletefile "C:\Program Files (x86)\N-Able Technologies\Windows Agent\cache\" & wscript.scriptname, true
  'end if
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

sub HOOK(strCMD)                                            ''CALL HOOK TO MONITOR OUTPUT OF CALLED COMMAND
  on error resume next
  'comspec = objWSH.ExpandEnvironmentStrings("%comspec%")
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
  if (err.number <> 0) then
    errRET = 3
    objOUT.write vbnewline & now & vbtab & vbtab & vbtab & err.number & vbtab & err.description
    objLOG.write vbnewline & now & vbtab & vbtab & vbtab & err.number & vbtab & err.description
    err.clear
  end if
end sub

sub LOGERR(intSTG)                                          ''CALL HOOK TO MONITOR OUTPUT OF CALLED COMMAND
  errRET = intSTG
  if (err.number <> 0) then
    objOUT.write vbnewline & now & vbtab & vbtab & vbtab & err.number & vbtab & err.description & vbnewline
    objLOG.write vbnewline & now & vbtab & vbtab & vbtab & err.number & vbtab & err.description & vbnewline
		err.clear
  end if
end sub

sub CLEANUP()                                               ''SCRIPT CLEANUP
  if (errRET = 0) then                                      ''SCRIPT COMPLETED SUCCESSFULLY
    objOUT.write vbnewline & vbnewline & now & vbtab & " - QSMART COMPLETE : " & now
    objLOG.write vbnewline & vbnewline & now & vbtab & " - QSMART COMPLETE : " & now
    err.clear
  elseif (errRET <> 0) then                                 ''SCRIPT FAILED
    objOUT.write vbnewline & vbnewline & now & vbtab & " - QSMART FAILURE : " & errRET & " : " & now
    objLOG.write vbnewline & vbnewline & now & vbtab & " - QSMART FAILURE : " & errRET & " : " & now
    ''RAISE CUSTOMIZED ERROR CODE, ERROR CODE WILL BE DEFINED RESTOP NUMBER INDICATING WHICH SECTION FAILED
    call err.raise(vbObjectError + errRET, "QSMART", "FAIL")
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

''MSP_CONFIG.VBS
''DESIGNED TO UPDATE THE MSP BACKUP 'CONFIG.INI' FILE IN AN AUTOMATED FASHION
''REQUIRED PARAMETER : 'STRHDR' , STRING TO IDENTIFY SECTION OF 'CONFIG.INI' FILE TO MODIFY
''REQUIRED PARAMETER : 'STRCHG' , SCRIPT VARIABLE TO CONTAIN STRING TO INJECT INTO 'CONFIG.INI' FILE
''WRITTEN BY : CJ BLEDSOE / CJ<@>THECOMPUTERWARRIORS.COM
''SCRIPT VARIABLES
dim blnHDR, blnINJ, blnMOD
dim errRET, strVER, strIN, arrIN
''VARIABLES ACCEPTING PARAMETERS
dim strHDR, strCHG, strVAL, blnFORCE
''SCRIPT OBJECTS
dim objLOG, objCFG
dim objIN, objOUT, objARG, objWSH, objFSO
''SET 'errRET' CODE
errRET = 0
''VERSION FOR SCRIPT UPDATE , MSP_CONFIG.VBS , REF #2 , FIXES #25
strVER = 4
''SET 'BLNHDR' FLAG
blnHDR = false
''SET 'BLNINJ' FLAG
blnINJ = false
''SET 'BLNMOD' FLAG
blnMOD = true
''SET 'BLNFORCE' FLAG
blnFORCE = false
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
if (objFSO.fileexists("C:\temp\MSP_CONFIG")) then                     ''LOGFILE EXISTS
  objFSO.deletefile "C:\temp\MSP_CONFIG", true
  set objLOG = objFSO.createtextfile("C:\temp\MSP_CONFIG")
  objLOG.close
  set objLOG = objFSO.opentextfile("C:\temp\MSP_CONFIG", 8)
else                                                                  ''LOGFILE NEEDS TO BE CREATED
  set objLOG = objFSO.createtextfile("C:\temp\MSP_CONFIG")
  objLOG.close
  set objLOG = objFSO.opentextfile("C:\temp\MSP_CONFIG", 8)
end if
''READ PASSED COMMANDLINE ARGUMENTS
if (wscript.arguments.count <= 2) then                                ''NO ARGUMENTS PASSED, END SCRIPT, 'ERRRET'=1
  objOUT.write vbnewline & vbnewline & now & vbtab & " - SCRIPT REQUIRES HEADER SELECTION, STRING TO INJECT, AND VALUE TO SET"
  objLOG.write vbnewline & vbnewline & now & vbtab & " - SCRIPT REQUIRES HEADER SELECTION, STRING TO INJECT, AND VALUE TO SET"
  call LOGERR(1)
  call CLEANUP()
elseif (wscript.arguments.count > 0) then                             ''ARGUMENTS WERE PASSED
  for x = 0 to (wscript.arguments.count - 1)
    objOUT.write vbnewline & now & vbtab & " - ARGUMENT " & (x + 1) & " (ITEM " & x & ") " & " PASSED : " & ucase(objARG.item(x))
    objLOG.write vbnewline & now & vbtab & " - ARGUMENT " & (x + 1) & " (ITEM " & x & ") " & " PASSED : " & ucase(objARG.item(x))
  next 
  strHDR = objARG.item(0)                                             ''SET STRING 'STRHDR', TARGET 'HEADER'
  if (wscript.arguments.count > 2) then
    strCHG = objARG.item(1)                                           ''SET STRING 'STRCHG', TARGET STRING TO INSERT
    strVAL = objARG.item(2)                                           ''SET STRING 'STRVAL', TARGET VALUE TO INSERT
    if (wscript.arguments.count > 3) then
      blnFORCE = objARG.item(3)                                       ''SET BOOLEAN 'BLNFORCE', FLAG TO FORCE MODIFY VALUE
    end if
  elseif (wscript.arguments.count <= 2) then                          ''NO ARGUMENTS PASSED, END SCRIPT, 'ERRRET'=1
    objOUT.write vbnewline & vbnewline & now & vbtab & " - SCRIPT REQUIRES HEADER SELECTION, STRING TO INJECT, AND VALUE TO SET"
    objLOG.write vbnewline & vbnewline & now & vbtab & " - SCRIPT REQUIRES HEADER SELECTION, STRING TO INJECT, AND VALUE TO SET"
    call LOGERR(1)
    call CLEANUP()  
  end if
end if

''------------
''BEGIN SCRIPT
objOUT.write vbnewline & now & " - STARTING MSP_CONFIG" & vbnewline
objLOG.write vbnewline & now & " - STARTING MSP_CONFIG" & vbnewline
''AUTOMATIC UPDATE, MSP_CONFIG.VBS, REF #2 , FIXES #25
call CHKAU()
''PARSE CONFIG.INI FILE
objOUT.write vbnewline & now & vbtab & " - CURRENT CONFIG.INI"
objLOG.write vbnewline & now & vbtab & " - CURRENT CONFIG.INI"
strIN = objCFG.readall
arrIN = split(strIN, vbnewline)
for intIN = 0 to ubound(arrIN)                                        ''CHECK CONFIG.INI LINE BY LINE
  objOUT.write vbnewline & vbtab & vbtab & arrIN(intIN)
  objLOG.write vbnewline & vbtab & vbtab & arrIN(intIN)
  if (arrIN(intIN) = strHDR) then                                     ''FOUND SPECIFIED 'HEADER' IN CONFIG.INI
    blnHDR = true
  end if
  if (instr(1, arrIN(intIN), strCHG)) then                            ''STRING TO INJECT ALREADY IN CONFIG.INI
    blnINJ = false
    blnMOD = false
    if (strVAL = split(arrIN(intIN), "=")(1)) then                    ''PASSED VALUE 'STRVAL' MATCHES INTERNAL STRING VALUE
      blnINJ = false
      blnMOD = false
    elseif (strVAL <> split(arrIN(intIN), "=")(1)) then               ''PASSED VALUE 'STRVAL' DOES NOT MATCH INTERNAL STRING VALUE
      if (not blnFORCE) then
        blnINJ = false
        blnMOD = false
      elseif (blnFORCE) then
        blnINJ = true
        blnMOD = false
        arrIN(intIN) = strCHG & "=" & strVAL
        exit for
      end if  
    end if
    exit for
  end if
  if ((blnHDR) and (blnMOD) and (arrIN(intIN) = vbnullstring)) then   ''STRING TO INJECT NOT FOUND, INJECT UNDER CURRENT 'HEADER'
    blnINJ = true
    blnHDR = false
    arrIN(intIN) = strCHG & "=" & strVAL & vbCrlf
  end if
next
objCFG.close
set objCFG = nothing
''REPLACE CONFIG.INI FILE
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
''CLEANUP
call CLEANUP()
''END SCRIPT
''------------

''SUB-ROUTINES
sub CHKAU()																					                  ''CHECK FOR SCRIPT UPDATE, MSP_CONFIG.VBS, REF #2 , FIXES #25
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
        objOUT.write vbnewline & now & vbtab & " - MSP_Config :  " & strVER & " : GitHub : " & objSCR.text & vbnewline
        objLOG.write vbnewline & now & vbtab & " - MSP_Config :  " & strVER & " : GitHub : " & objSCR.text & vbnewline
				if (cint(objSCR.text) > cint(strVER)) then
					objOUT.write vbnewline & now & " - UPDATING " & objSCR.nodename & " : " & objSCR.text & vbnewline
					objLOG.write vbnewline & now & " - UPDATING " & objSCR.nodename & " : " & objSCR.text & vbnewline
					''DOWNLOAD LATEST VERSION OF SCRIPT
					call FILEDL("https://github.com/CW-Khristos/scripts/raw/master/MSP%20Backups/MSP_Config.vbs", wscript.scriptname)
					''RUN LATEST VERSION
					if (wscript.arguments.count > 0) then                       ''ARGUMENTS WERE PASSED
						for x = 0 to (wscript.arguments.count - 1)
							strTMP = strTMP & " " & chr(34) & objARG.item(x) & chr(34)
						next
            objOUT.write vbnewline & now & vbtab & " - RE-EXECUTING  " & objSCR.nodename & " : " & objSCR.text & vbnewline
            objLOG.write vbnewline & now & vbtab & " - RE-EXECUTING  " & objSCR.nodename & " : " & objSCR.text & vbnewline
						objWSH.run "cscript.exe //nologo " & chr(34) & "c:\temp\" & wscript.scriptname & chr(34) & strTMP, 0, false
					elseif (wscript.arguments.count = 0) then                   ''NO ARGUMENTS WERE PASSED
            objOUT.write vbnewline & now & vbtab & " - RE-EXECUTING  " & objSCR.nodename & " : " & objSCR.text & vbnewline
            objLOG.write vbnewline & now & vbtab & " - RE-EXECUTING  " & objSCR.nodename & " : " & objSCR.text & vbnewline
						objWSH.run "cscript.exe //nologo " & chr(34) & "c:\temp\" & wscript.scriptname & chr(34), 0, false
					end if
					''SET 'ERRRET'=13, END SCRIPT
          call LOGERR(13)
					call CLEANUP()
				end if
			end if
		next
	end if
	set colVER = nothing
	set objXML = nothing
  if (err.number <> 0) then                                           ''ERROR RETURNED DURING UPDATE CHECK , 'ERRRET'=10
    call LOGERR(10)
  end if
end sub

sub FILEDL(strURL, strFILE)                                           ''CALL HOOK TO DOWNLOAD FILE FROM URL , 'ERRRET'=11
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
  if (err.number <> 0) then                                           ''ERROR RETURNED , 'ERRRET'=11
    call LOGERR(11)
  end if
end sub

sub HOOK(strCMD)                                                      ''CALL HOOK TO MONITOR OUTPUT OF CALLED COMMAND , 'ERRRET'=12
  on error resume next
  objOUT.write vbnewline & now & vbtab & vbtab & "EXECUTING : " & strCMD
  objLOG.write vbnewline & now & vbtab & vbtab & "EXECUTING : " & strCMD
  set objHOOK = objWSH.exec(strCMD)
  if (instr(1, strCMD, "takeown /F ") = 0) then                       ''SUPPRESS 'TAKEOWN' SUCCESS MESSAGES
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
    call LOGERR(12)
  end if
end sub

sub LOGERR(intSTG)                                                    ''CALL HOOK TO MONITOR OUTPUT OF CALLED COMMAND
  errRET = intSTG
  if (err.number <> 0) then
    objOUT.write vbnewline & now & vbtab & vbtab & vbtab & err.number & vbtab & err.description & vbnewline
    objLOG.write vbnewline & now & vbtab & vbtab & vbtab & err.number & vbtab & err.description & vbnewline
		err.clear
  end if
end sub

sub CLEANUP()                                 			                  ''SCRIPT CLEANUP
  if (errRET = 0) then         											                  ''MSP_CONFIG COMPLETED SUCCESSFULLY
    objOUT.write vbnewline & "MSP_CONFIG SUCCESSFUL : " & now
  elseif (errRET <> 0) then    											                  ''MSP_CONFIG FAILED
    objOUT.write vbnewline & "MSP_CONFIG FAILURE : " & now & " : " & errRET
    ''RAISE CUSTOMIZED ERROR CODE, ERROR CODE WILL BE DEFINE RESTOP NUMBER INDICATING WHICH SECTION FAILED
    call err.raise(vbObjectError + errRET, "MSP_CONFIG", "FAILURE")
  end if
  objOUT.write vbnewline & vbnewline & now & " - MSP_CONFIG COMPLETE" & vbnewline
  objLOG.write vbnewline & vbnewline & now & " - MSP_CONFIG COMPLETE" & vbnewline
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
  wscript.quit err.number
end sub
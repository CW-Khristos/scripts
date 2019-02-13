''MSP_CONFIG.VBS
''DESIGNED TO UPDATE THE MSP BACKUP 'CONFIG.INI' FILE IN AN AUTOMATED FASHION
''REQUIRED PARAMETER : 'STRHDR' , STRING TO IDENTIFY SECTION OF 'CONFIG.INI' FILE TO MODIFY
''REQUIRED PARAMETER : 'STRCHG' , SCRIPT VARIABLE TO CONTAIN STRING TO INJECT INTO 'CONFIG.INI' FILE
''WRITTEN BY : CJ BLEDSOE / CJ<@>THECOMPUTERWARRIORS.COM
''SCRIPT VARIABLES
dim blnHDR, blnINJ, blnMOD
dim errRET, strVER, strIN, arrIN
''VARIABLES ACCEPTING PARAMETERS
dim strHDR, strCHG
''SCRIPT OBJECTS
dim objLOG, objCFG
dim objIN, objOUT, objARG, objWSH, objFSO
''SET 'errRET' CODE
errRET = 0
''VERSION FOR SCRIPT UPDATE , MSP_CONFIG.VBS , REF #2 , FIXES #25
strVER = 3
''SET 'BLNHDR' FLAG
blnHDR = false
''SET 'BLNINJ' FLAG
blnINJ = false
''SET 'BLNMOD' FLAG
blnMOD = true
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
if (objFSO.fileexists("C:\temp\MSP_CONFIG")) then               ''LOGFILE EXISTS
  objFSO.deletefile "C:\temp\MSP_CONFIG", true
  set objLOG = objFSO.createtextfile("C:\temp\MSP_CONFIG")
  objLOG.close
  set objLOG = objFSO.opentextfile("C:\temp\MSP_CONFIG", 8)
else                                                            ''LOGFILE NEEDS TO BE CREATED
  set objLOG = objFSO.createtextfile("C:\temp\MSP_CONFIG")
  objLOG.close
  set objLOG = objFSO.opentextfile("C:\temp\MSP_CONFIG", 8)
end if
''READ PASSED COMMANDLINE ARGUMENTS
if (wscript.arguments.count > 0) then                           ''ARGUMENTS WERE PASSED
  for x = 0 to (wscript.arguments.count - 1)
    objOUT.write vbnewline & now & vbtab & " - ARGUMENT " & (x + 1) & " (ITEM " & x & ") " & " PASSED : " & ucase(objARG.item(x))
    objLOG.write vbnewline & now & vbtab & " - ARGUMENT " & (x + 1) & " (ITEM " & x & ") " & " PASSED : " & ucase(objARG.item(x))
  next 
  ''ARGUMENT 0 - TARGET 'HEADER'
  strHDR = objARG.item(0)
  if (wscript.arguments.count > 1) then                         ''SET STRING TO INSERT INTO CONFIG.INI
    strCHG = objARG.item(1)
  else                                                          ''NOT ENOUGH ARGUMENTS PASSED, END SCRIPT
    errRET = 1
    call CLEANUP
  end if
else                                                            ''NO ARGUMENTS PASSED, END SCRIPT
  objOUT.write vbnewline & vbnewline & now & vbtab & " - SCRIPT REQUIRES HEADER SELECTION AND STRING TO INJECT"
  objLOG.write vbnewline & vbnewline & now & vbtab & " - SCRIPT REQUIRES HEADER SELECTION AND STRING TO INJECT"
  errRET = 1
  call CLEANUP
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
for intIN = 0 to ubound(arrIN)                                  ''CHECK CONFIG.INI LINE BY LINE
  objOUT.write vbnewline & vbtab & vbtab & arrIN(intIN)
  objLOG.write vbnewline & vbtab & vbtab & arrIN(intIN)
  if (arrIN(intIN) = strHDR) then                               ''FOUND SPECIFIED 'HEADER' IN CONFIG.INI
    blnHDR = true
  end if
  if (arrIN(intIN) = strCHG) then                               ''STRING TO INJECT ALREADY IN CONFIG.INI
    blnMOD = false
  end if
  if ((blnHDR) and (blnMOD) and (arrIN(intIN) = vbnullstring)) then   ''STRING TO INJECT NOT FOUND, INJECT UNDER CURRENT 'HEADER'
    blnINJ = true
    blnHDR = false
    arrIN(intIN) = strCHG & vbCrlf
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
sub CHKAU()																					''CHECK FOR SCRIPT UPDATE, MSP_CONFIG.VBS, REF #2 , FIXES #25
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
				if (cint(objSCR.text) > cint(strVER)) then
					objOUT.write vbnewline & now & " - UPDATING " & objSCR.nodename & " : " & objSCR.text & vbnewline
					objLOG.write vbnewline & now & " - UPDATING " & objSCR.nodename & " : " & objSCR.text & vbnewline
					''DOWNLOAD LATEST VERSION OF SCRIPT
					call FILEDL("https://github.com/CW-Khristos/scripts/raw/master/MSP%20Backups/MSP_Config.vbs", wscript.scriptname)
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

sub HOOK(strCMD)                              			''CALL HOOK TO MONITOR OUTPUT OF CALLED COMMAND
  on error resume next
  set objHOOK = objWSH.exec(strCMD)
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
  set objHOOK = nothing
  if ((not blnSUP) and (err.number <> 0)) then
    objOUT.write vbnewline & now & vbtab & vbtab & vbtab & err.number & vbtab & err.description
    objLOG.write vbnewline & now & vbtab & vbtab & vbtab & err.number & vbtab & err.description
    errRET = 3
    err.clear
  end if
end sub

sub FILEDL(strURL, strFILE)                   			''CALL HOOK TO DOWNLOAD FILE FROM URL
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
		err.clear
  end if
end sub

sub CLEANUP()                                 			''SCRIPT CLEANUP
  if (errRET = 0) then         											''MSP_CONFIG COMPLETED SUCCESSFULLY
    objOUT.write vbnewline & "MSP_CONFIG SUCCESSFUL : " & NOW
  elseif (errRET <> 0) then    											''MSP_CONFIG FAILED
    objOUT.write vbnewline & "MSP_CONFIG FAILURE : " & NOW & " : " & errRET
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
  wscript.quit errRET
end sub
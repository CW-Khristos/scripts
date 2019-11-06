''DISK_USAGE.VBS
''DESIGNED TO AUTOMATE ANALYSIS AND REPORTING OF DISK USAGE STATISTICS USING X.ROBOT CMD UTILITY
''WRITTEN BY : CJ BLEDSOE / CJ<@>THECOMPUTERWARRIORS.COM
'on error resume next
''SCRIPT VARIABLES
dim errRET, strVER, strIN, intOPT
''SCRIPT OBJECTS
dim objIN, objOUT, objARG, objWSH, objFSO
dim objLOG, objHOOK, objEXEC, objHTTP, objXML
''VERSION FOR SCRIPT UPDATE, DISK_USAGE.VBS, REF #2 , FIXES #45
strVER = 2
''DEFAULT SUCCESS
errRET = 0
''ZIP ARCHIVE OPTIONS
intOPT = 256
''STDIN / STDOUT
set objIN = wscript.stdin
set objOUT = wscript.stdout
set objARG = wscript.arguments
''OBJECTS FOR LOCATING FOLDERS
set objWSH = createobject("wscript.shell")
Set objAPP = createobject("shell.application")
set objFSO = createobject("scripting.filesystemobject")
''PREPARE LOGFILE
if (objFSO.fileexists("C:\temp\DISK_USAGE")) then               ''LOGFILE EXISTS
  objFSO.deletefile "C:\temp\DISK_USAGE", true
  set objLOG = objFSO.createtextfile("C:\temp\DISK_USAGE")
  objLOG.close
  set objLOG = objFSO.opentextfile("C:\temp\DISK_USAGE", 8)
else                                                          ''LOGFILE NEEDS TO BE CREATED
  set objLOG = objFSO.createtextfile("C:\temp\DISK_USAGE")
  objLOG.close
  set objLOG = objFSO.opentextfile("C:\temp\DISK_USAGE", 8)
end if
''READ PASSED COMMANDLINE ARGUMENTS
if (wscript.arguments.count > 0) then                         ''ARGUMENTS WERE PASSED
  for x = 0 to (wscript.arguments.count - 1)
    objOUT.write vbnewline & now & vbtab & " - ARGUMENT " & (x + 1) & " (ITEM " & x & ") " & " PASSED : " & ucase(objARG.item(x))
    objLOG.write vbnewline & now & vbtab & " - ARGUMENT " & (x + 1) & " (ITEM " & x & ") " & " PASSED : " & ucase(objARG.item(x))
  next 
  if (wscript.arguments.count > 0) then                       ''REQUIRED ARGUMENTS PASSED
    strPATH = objARG.item(0)
    strFORM = objARG.item(1)
  end if
else                                                          ''NO ARGUMENTS PASSED , END SCRIPT , 'ERRRET'=1
  call LOGERR(1)
  err.clear
end if

''------------
''BEGIN SCRIPT
if (errRET = 1) then
  objOUT.write vbnewline & vbnewline & now & vbtab & " - EXECUTING DISK_USAGE"
  objLOG.write vbnewline & vbnewline & now & vbtab & " - EXECUTING DISK_USAGE"
  ''AUTOMATIC UPDATE , 'ERRRET'=10 , DISK_USAGE.VBS , REF #2 , FIXES #45
  call CHKAU()
  ''CHECK FOR X.ROBOT.EXE IN C:\TEMP\X.ROBOT32
  if (not objFSO.fileexists("c:\temp\X.Robot32\x.robot.exe")) then
    call FILEDL("https://github.com/CW-Khristos/scripts/raw/dev/XRobot/X.Robot32.zip", "X.Robot32.zip")
    wscript.sleep 5000
    ''CHECK FOR X.ROBOT32.ZIP IN C:\TEMP, REF #46
    if (not objFSO.fileexists("c:\temp\X.Robot32.zip")) then
      call FILEDL("https://github.com/CW-Khristos/scripts/raw/dev/XRobot/X.Robot32.zip", "X.Robot32.zip")
    end if
    if (objFSO.fileexists("C:\temp\X.Robot32.zip")) then
      ''EXTRACT X.ROBOT32.ZIP TO C:\TEMP\XROBOT
      set objSRC = objAPP.namespace("C:\temp\X.Robot32.zip").items()
      set objTGT = objAPP.namespace("C:\temp")
      objTGT.copyhere objSRC, intOPT
    end if
  end if
  if (objFSO.fileexists("c:\temp\X.Robot32\x.robot.exe")) then
    strRCMD = "c:\temp\x.robot32\x.robot.exe " & chr(34) & strPATH & chr(34)
    if (ucase(strFORM) = "HTM") then
      strRCMD = strRCMD & " /HTM{111111111120};c:\temp\robot.htm"
    elseif (ucase(strFORM) = "CSV") then
      strRCMD = strRCMD & " /CSV{113};c:\temp\robot.csv"
    end if
    call HOOK("CMD /C " & chr(34) & strRCMD & chr(34))
    ''DISABLED ZIP ARCHIVE CALLS
    'wscript.sleep 5000
    if (ucase(strFORM) = "HTM") then
      call FILEDL("https://github.com/CW-Khristos/scripts/raw/dev/XRobot/SavePage.exe", "savepage.exe")
      strRCMD = "c:\temp\savepage.exe " & chr(34) & "XRobot - Report" & chr(34) & " " & chr(34) & "file://c:/temp/robot.htm" & chr(34) & " " & chr(34) & "C:\temp\" & chr(34)
      call HOOK("CMD /C " & chr(34) & strRCMD & chr(34))
    '  call makZIP("c:\temp\robot.htm", "c:\temp\robot.zip")
    '  wscript.sleep 1000
    '  call makZIP("c:\temp\data", "c:\temp\robot.zip")
    end if
  end if
elseif (errRET = 0) then
  objOUT.write vbnewline & vbnewline & now & vbtab & " - EXECUTING DISK_USAGE"
  objLOG.write vbnewline & vbnewline & now & vbtab & " - EXECUTING DISK_USAGE"
  ''AUTOMATIC UPDATE , 'ERRRET'=10 , DISK_USAGE.VBS , REF #2 , FIXES #45
  call CHKAU()
  ''CHECK FOR X.ROBOT.EXE IN C:\TEMP\X.ROBOT32
  if (not objFSO.fileexists("c:\temp\X.Robot32\x.robot.exe")) then
    call FILEDL("https://github.com/CW-Khristos/scripts/raw/dev/XRobot/X.Robot32.zip", "X.Robot32.zip")
    wscript.sleep 5000
    ''CHECK FOR X.ROBOT32.ZIP IN C:\TEMP, REF #46
    if (not objFSO.fileexists("c:\temp\X.Robot32.zip")) then
      call FILEDL("https://github.com/CW-Khristos/scripts/raw/dev/XRobot/X.Robot32.zip", "X.Robot32.zip")
    end if
    if (objFSO.fileexists("C:\temp\X.Robot32.zip")) then
      ''EXTRACT X.ROBOT32.ZIP TO C:\TEMP\XROBOT
      set objSRC = objAPP.namespace("C:\temp\X.Robot32.zip").items()
      set objTGT = objAPP.namespace("C:\temp")
      objTGT.copyhere objSRC, intOPT
    end if
  end if
  ''CHECK FOR EXTRACTED X.ROBOT
  if (objFSO.fileexists("c:\temp\X.Robot32\x.robot.exe")) then
    strRCMD = "c:\temp\x.robot32\x.robot.exe " & chr(34) & strPATH & chr(34)
    if (ucase(strFORM) = "HTM") then
      strRCMD = strRCMD & " /HTM{111111111120};c:\temp\robot.htm"
    elseif (ucase(strFORM) = "CSV") then
      strRCMD = strRCMD & " /CSV{113};c:\temp\robot.csv"
    end if
    ''EXECUTE X.ROBOT
    call HOOK("CMD /C " & chr(34) & strRCMD & chr(34))
    ''DISABLED ZIP ARCHIVE CALLS
    'wscript.sleep 5000
    ''CONVERT TO HTM FORMAT
    if (ucase(strFORM) = "HTM") then
      call FILEDL("https://github.com/CW-Khristos/scripts/raw/dev/XRobot/SavePage.exe", "savepage.exe")
      strRCMD = "c:\temp\savepage.exe " & chr(34) & "XRobot - Report" & chr(34) & " " & chr(34) & "file://c:/temp/robot.htm" & chr(34) & " " & chr(34) & "C:\temp\" & chr(34)
      call HOOK("CMD /C " & chr(34) & strRCMD & chr(34))
      ''ARCHIVE HTM REPORT DATA
    '  call makZIP("c:\temp\data", "c:\temp\robot.zip")
    '  wscript.sleep 1000
    '  call makZIP("c:\temp\robot.htm", "c:\temp\robot.zip")
    end if
  end if
  if (objFSO.folderexists("C:\Windows\CSC")) then
    set objFOL = objFSO.getfolder("C:\Windows\CSC")
    intSIZ = (objFOL.size / 1048576)  ''CONVERT TO MB
    objOUT.write vbnewline & now & vbtab & vbtab & "CSC CACHE SIZE (MB) : " & intSIZ
    objLOG.write vbnewline & now & vbtab & vbtab & "CSC CACHE SIZE (MB) : " & intSIZ
  end if
end if
''END SCRIPT
call CLEANUP()
''END SCRIPT
''------------

''FUNCTIONS
function makZIP(strSRC, strZIP)                                                       ''MAKE ZIP ARCHIVE , 'ERRRET'=60
  strSRC = objFSO.getabsolutepathname(strSRC)
  strZIP = objFSO.getabsolutepathname(strZIP)
  if (not objFSO.fileexists(strZIP)) then                                             ''MAKE NEW ZIP ARCHIVE IF ARCHIVE DOES NOT EXIST , 'ERRRET'=61
    call newZIP(strZIP)
  end if
  ''ENUMERATE FILES
  sDupe = false
  aFileName = split(strSRC, "\")
  sFileName = (aFileName(ubound(aFileName)))
  sZipFileCount = objAPP.namespace(strZIP).items.count
  if (sZipFileCount > 0) then                                                         ''CHECK FOR DUPLICATES
    for each strZIPFILE in objAPP.namespace(strZIP).items
      if lcase(sFileName) = lcase(strZIPFILE) then                                    ''DUPLICATE FOUND
        sDupe = true
        exit for
      end if
    next
  end if
  if (not sDupe) then                                                                 ''DUPLICATE NOT FOUND
    objAPP.namespace(strZIP).copyhere objAPP.namespace(strSRC).items, 16
    ''CHECK FOR COMPLETION OF COMPRESSION
    intLOOP = 0
    do until (sZipFileCount < objAPP.namespace(strZIP).items.count)
      wscript.sleep 500
      objOUT.write "."
      intLOOP = intLOOP + 1
    loop
    objOUT.write vbnewline & vbtab & vbtab & "ZIP COMPLETED" & vbnewline
  end if
  set objZIP = nothing
  if (err.number <> 0) then                                                           ''ERROR RETURNED DURING ZIP ARCHIVE CREATION , 'ERRRET'=61
    call LOGERR(60)
  end if
end function

''SUB-ROUTINES
sub newZIP(strNZIP)                                                                   ''PREPARE NEW ZIP ARCHIVE , 'ERRRET'=61
  Set objNFIL = objFSO.createtextfile(strNZIP)
  objNFIL.write chr(80) & chr(75) & chr(5) & chr(6) & string(18, 0)
  objNFIL.close
  set objNFIL = nothing
  wscript.sleep 500
  if (err.number <> 0) then                                                           ''ERROR CREATING NEW ZIP ARCHIVE , 'ERRRET'=61
    call LOGERR(61)
  end if
end sub

sub CHKAU()																					          ''CHECK FOR SCRIPT UPDATE, DISK_USAGE.VBS, REF #2 , FIXES #45
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
					call FILEDL("https://github.com/CW-Khristos/scripts/raw/dev/disk_usage.vbs", wscript.scriptname)
					''RUN LATEST VERSION
					if (wscript.arguments.count > 0) then               ''ARGUMENTS WERE PASSED
						for x = 0 to (wscript.arguments.count - 1)
							strTMP = strTMP & " " & chr(34) & objARG.item(x) & chr(34)
						next
            objOUT.write vbnewline & now & vbtab & " - RE-EXECUTING  " & objSCR.nodename & " : " & objSCR.text & vbnewline
            objLOG.write vbnewline & now & vbtab & " - RE-EXECUTING  " & objSCR.nodename & " : " & objSCR.text & vbnewline
						objWSH.run "cscript.exe //nologo " & chr(34) & "c:\temp\" & wscript.scriptname & chr(34) & strTMP, 0, false
					elseif (wscript.arguments.count = 0) then           ''NO ARGUMENTS WERE PASSED
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

sub FILEDL(strURL, strFILE)                                   ''CALL HOOK TO DOWNLOAD FILE FROM URL
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

sub HOOK(strCMD)                                              ''CALL HOOK TO MONITOR OUTPUT OF CALLED COMMAND
  on error resume next
  objOUT.write vbnewline & now & vbtab & vbtab & "EXECUTING : " & strCMD
  objLOG.write vbnewline & now & vbtab & vbtab & "EXECUTING : " & strCMD
  set objHOOK = objWSH.exec(strCMD)
  if (instr(1, strCMD, "takeown /F ") = 0) then               ''SUPPRESS 'TAKEOWN' SUCCESS MESSAGES
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

sub LOGERR(intSTG)                                            ''CALL HOOK TO MONITOR OUTPUT OF CALLED COMMAND
  if (err.number <> 0) then
    objOUT.write vbnewline & now & vbtab & vbtab & vbtab & err.number & vbtab & err.description & vbnewline
    objLOG.write vbnewline & now & vbtab & vbtab & vbtab & err.number & vbtab & err.description & vbnewline
		errRET = intSTG
		err.clear
  end if
end sub

sub CLEANUP()                                                 ''SCRIPT CLEANUP
  if (errRET = 0) then                                        ''SCRIPT COMPLETED SUCCESSFULLY
    objOUT.write vbnewline & vbnewline & now & vbtab & " - DISK_USAGE COMPLETE : " & now
    objLOG.write vbnewline & vbnewline & now & vbtab & " - DISK_USAGE COMPLETE : " & now
    err.clear
  elseif (errRET <> 0) then                                   ''SCRIPT FAILED
    objOUT.write vbnewline & vbnewline & now & vbtab & " - DISK_USAGE FAILURE : " & errRET & " : " & now
    objLOG.write vbnewline & vbnewline & now & vbtab & " - DISK_USAGE FAILURE : " & errRET & " : " & now
    ''RAISE CUSTOMIZED ERROR CODE, ERROR CODE WILL BE DEFINED RESTOP NUMBER INDICATING WHICH SECTION FAILED
    call err.raise(vbObjectError + errRET, "DISK_USAGE", "fail")
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
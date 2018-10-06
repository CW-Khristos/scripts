''MSP_ROTATE.VBS
''DESIGNED TO AUTOMATE ARCHIVAL / ROTATION OF MSP BACKUP LSV AND DEBUG LOG DATA
''WRITTEN BY : CJ BLEDSOE / CJ<@>THECOMPUTERWARRIORS.COM
on error resume next
''DEFINE VARIABLES
dim objFSO, objMSP, colFOL, objFOL
dim objIN, objOUT, objARG, objWSH, objSHL
dim errRET, retDEL, strIN
dim strLSV, strDEV, strDLM, intDIFF, blnRUN
''VERSION FOR SCRIPT UPDATE, MSP_ROTATE.VBS, REF #2 , FIXES #26
strVER = 2
''DEFAULT SUCCESS
errRET = 0
''STDIN / STDOUT
set objIN = wscript.stdin
set objOUT = wscript.stdout
set objARG = wscript.arguments
''CREATE OBJECTS
set objWSH = createobject("wscript.shell")
set objSHL = createobject("shell.application")
set objFSO = createobject("scripting.filesystemobject")
''CHECK EXECUTION METHOD OF SCRIPT
strIN = lcase(mid(wscript.fullname, instrrev(wscript.fullname, "\") + 1))
if (strIN <> "cscript.exe") then
  objOUT.write vbnewline & now & " - SCRIPT LAUNCHED VIA EXPLORER, EXECUTING SCRIPT VIA CSCRIPT..."
  objWSH.run "cscript.exe //nologo " & chr(34) & wscript.scriptfullname & chr(34)
  wscript.quit
end if
''PREPARE LOGFILE
if (objFSO.fileexists("C:\temp\MSP_ROTATE")) then     							  ''LOGFILE EXISTS
  objFSO.deletefile "C:\temp\MSP_ROTATE", true
  set objLOG = objFSO.createtextfile("C:\temp\MSP_ROTATE")
  objLOG.close
  set objLOG = objFSO.opentextfile("C:\temp\MSP_ROTATE", 8)
else                                            										  ''LOGFILE NEEDS TO BE CREATED
  set objLOG = objFSO.createtextfile("C:\temp\MSP_ROTATE")
  objLOG.close
  set objLOG = objFSO.opentextfile("C:\temp\MSP_ROTATE", 8)
end if
''READ PASSED COMMANDLINE ARGUMENTS
if (wscript.arguments.count > 0) then        												  ''ARGUMENTS WERE PASSED
  for x = 0 to (wscript.arguments.count - 1)
    objOUT.write vbnewline & now & " - ARGUMENT " & (x + 1) & " (ITEM " & x & ") " & " PASSED : " & ucase(objARG.item(x))
    objLOG.write vbnewline & now & " - ARGUMENT " & (x + 1) & " (ITEM " & x & ") " & " PASSED : " & ucase(objARG.item(x))
  next 
  ''ARGUMENT 0 - FOLDER PATH
  strLSV = objARG.item(0)
  ''ARGUMENT 1 (OPTIONAL) - TARGET FILE AGE / DATE LAST MODIFIED
	''ARGUMENT 2 (OPTIONAL) - SCRIPT EXECUTION FLAG
  if (wscript.arguments.count > 1) then
    intAGE = objARG.item(1)
    blnRUN = objARG.item(2)
  else
    intAGE = 60
    blnRUN = false
  end if
end if
if ((wscript.arguments.count = 0) and (strLSV = vbnullstring)) then    ''NO ARGUMENTS PASSED
  objOUT.write vbnewline & now & " - SCRIPT REQUIRES PATH TO LOCAL MSP BACKUP DESTINATION"
  objLOG.write vbnewline & now & " - SCRIPT REQUIRES PATH TO LOCAL MSP BACKUP DESTINATION"
  intAGE = 60
  blnRUN = false
end if

''------------
''BEGIN SCRIPT
objOUT.write vbnewline & vbnewline & now & vbtab & " - EXECUTING MSP_ROTATE"
objLOG.write vbnewline & vbnewline & now & vbtab & " - EXECUTING MSP_ROTATE"
''AUTOMATIC UPDATE, MSP_ROTATE.VBS, REF #2 , FIXES #26
call CHKAU()
'CHECK LSV IF CONFIGURED, OTHERWISE ATTEMPT TO LOCATE , REF #17
call chkLSV()
''RUN MAIN SCRIPT , REF #17
call rotMSP()
''END SCRIPT
call CLEANUP()
''END SCRIPT
''------------

''ZIP FUNCTIONS , REF #17
function opnZIP(strZIP, strDST)
  if (not objFSO.folderexists(strDST)) then
    objFSO.createfolder(strDST)
  end if
  objSHL.namespace(strDST).copyhere objSHL.namespace(strZIP).items
end function

function makZIP(strSRC, strZIP)
  strSRC = objFSO.getabsolutepathname(strSRC)
  strZIP = objFSO.getabsolutepathname(strZIP)
  if (not objFSO.fileexists(strZIP)) then
    call newZIP(strZIP)
  end if
  ''ENUMERATE FILES
  sDupe = false
  aFileName = split(strSRC, "\")
  sFileName = (aFileName(ubound(aFileName)))
  sZipFileCount = objSHL.namespace(strZIP).items.count
  ''CHECK FOR DUPLICATES
  if (sZipFileCount > 0) then
    for each strZIPFILE in objSHL.namespace(strZIP).items
      if lcase(sFileName) = lcase(strZIPFILE) then                    ''DUPLICATE FOUND
        sDupe = true
        exit for
      end if
    next
  end if
  if (not sDupe) then                                                 ''DUPLICATE NOT FOUND
    objSHL.namespace(strZIP).copyhere objSHL.namespace(strSRC).items, 4
    ''CHECK FOR COMPLETION OF COMPRESSION
    intLOOP = 0
    do until sZipFileCount < objSHL.namespace(strZIP).items.count
      wscript.sleep 15000
        objOUT.write "."
      intLOOP = intLOOP + 1
    loop
    objOUT.write "COMPLETED" & vbnewline
    'set objZIP = objFSO.getfile(strZIP)
    'do
    '  objOUT.write "."
    '    wscript.sleep 500
    '    intMAX = objZIP.size
    'loop while objZIP.size > intMAX 
    'on error goto 0
  end if
  set objZIP = nothing
end function

''SUB-ROUTINES
sub newZIP(strNZIP)
  Set objNFIL = objFSO.createtextfile(strNZIP)
  objNFIL.write chr(80) & chr(75) & chr(5) & chr(6) & string(18, 0)
  objNFIL.close
  set objNFIL = nothing
  wscript.sleep 500
end sub

sub rotMSP()																												  ''MSP_ROTATE MAIN SUB-ROUTINE , REF #17
	''RUN MSP_ROTATE
	objOUT.write vbnewline & vbnewline & now & " - ROTATE MSP BACKUPSET : RUNNING : DELETION : " & blnRUN
	objLOG.write vbnewline & vbnewline & now & " - ROTATE MSP BACKUPSET : RUNNING : DELETION : " & blnRUN
  ''STOP 'BACKUP SERVICE CONTROLLER' SERVICE AND TERMINATE 'BACKUPFP.EXE' PROCESS PRIOR TO ARCHIVE / DELETION
  'call HOOK("net stop " & chr(34) & "backup service controller" & chr(34) & " /y")
  'call HOOK("taskkill /IM " & chr(34) & "BackupFP.exe" & chr(34) & " /F")
	''RETRIEVE MSP BACKUP LSV FOLDER , REF #17
	if (strLSV <> vbnullstring) then
    objOUT.write vbnewline & vbnewline & now & vbtab & " - CHECKING LSV : RUNNING : DELETION : " & blnRUN
    objLOG.write vbnewline & vbnewline & now & vbtab & " - CHECKING LSV : RUNNING : DELETION : " & blnRUN
    set objMSP = objFSO.getfolder(strLSV)
    if (err.number <> 0) then																					''ERROR OBTAINING FOLDER
      objOUT.write vbnewline & vbnewline & now & vbtab & " - ROTATE MSP BACKUPSET : LSV ERROR : " & err.description
      objLOG.write vbnewline & vbnewline & now & vbtab & " - ROTATE MSP BACKUPSET : LSV ERROR : " & err.description
      errRET = 4
    elseif (err.number = 0) then																			''SUCCESSFULLY OBTAINED FOLDER
      errRET = 0
      ''CHECK THROUGH ALL SUB-FOLDERS IN MSP BACKUPSET FOLDER
      ''THIS NEEDS TO BE THE COMPLETE PATH THAT MSP BACKUPS WRITE TO, IE : F:\MSP Backups\jax-dc1_lr0xa_E93A9C6948C0DD5FF5E9\
      call chkFOL(objMSP)
    end if
  end if
  ''RETRIEVE MSP BACKUP DEBUG FOLDERS , REF #17
  objOUT.write vbnewline & vbnewline & now & vbtab & " - CHECKING 'BACKUPFP.PROTOCOL' DEBUG LOGS : RUNNING : DELETION : " & blnRUN
  objLOG.write vbnewline & vbnewline & now & vbtab & " - CHECKING 'BACKUPFP.PROTOCOL' DEBUG LOGS : RUNNING : DELETION : " & blnRUN
  if (objFSO.folderexists("C:\ProgramData\MXB\Backup Manager\logs")) then
    if (objFSO.folderexists("C:\ProgramData\MXB\Backup Manager\logs\BackupFP.Protocol")) then
      set objMSP = objFSO.getfolder("C:\ProgramData\MXB\Backup Manager\logs\BackupFP.Protocol")
    end if
  elseif (objFSO.folderexists("C:\Documents and Settings\All Users\Application Data\MXB\Backup Manager\logs")) then
    if (objFSO.folderexists("C:\Documents and Settings\All Users\Application Data\MXB\Backup Manager\logs\BackupFP.Protocol")) then
      set objMSP = objFSO.getfolder("C:\Documents and Settings\All Users\Application Data\MXB\Backup Manager\logs\BackupFP.Protocol")
    end if
  end if
  if (err.number <> 0) then																					  ''ERROR OBTAINING FOLDER
    objOUT.write vbnewline & vbnewline & now & vbtab & " - ROTATE MSP BACKUPSET : DEBUG ERROR : " & err.description
    objLOG.write vbnewline & vbnewline & now & vbtab & " - ROTATE MSP BACKUPSET : DEBUG ERROR : " & err.description
    errRET = 5
  elseif (err.number = 0) then																			  ''SUCCESSFULLY OBTAINED FOLDER
    errRET = 0
    ''CHECK THROUGH ALL SUB-FOLDERS IN MSP BACKUPSET FOLDER
    ''THIS NEEDS TO BE THE COMPLETE PATH THAT MSP BACKUPS WRITE TO, IE : F:\MSP Backups\jax-dc1_lr0xa_E93A9C6948C0DD5FF5E9\
    call chkFOL(objMSP)
    call chkFIL(objMSP)
  end if
  ''RESTART 'BACKUP SERVICE CONTROLLER' SERVICE
  'call HOOK("net start " & chr(34) & "backup service controller" & chr(34) & " /y")
end sub

sub chkFOL(objMSP)
  set colFOL = objMSP.subfolders
  for each objFOL in colFOL                    										    ''ENUMERATE EACH SUB-FOLDER
    ''DEFAULT FAIL
    retDEL = 6
    ''RETRIEVE SUB-FOLDER LAST DATE MODIFIED
    strDLM = objFOL.datelastmodified
    ''CALCULATE DATE DIFFERENCE (BY VALUE "D"AYS)
    intDIFF = -(datediff("d", now, strDLM))
    if (intDIFF >= cint(intAGE)) then           										  ''FOLDER HAS NOT BEEN MODIFIED IN TARGET AGE
      objOUT.write vbnewline & vbnewline & now & vbtab & vbtab & " - " & objFOL.path
      objOUT.write vbnewline & now & vbtab & vbtab & " - LAST MODIFIED : " & objFOL.DateLastModified & " : " & intDIFF & " Day(s)"
      objLOG.write vbnewline & vbnewline & now & vbtab & vbtab & " - " & objFOL.path
      objLOG.write vbnewline & now & vbtab & vbtab & " - LAST MODIFIED : " & objFOL.DateLastModified & " : " & intDIFF & " Day(s)"
      if (blnRUN) then																		            ''SCRIPT SET TO EXECUTE DELETION
        ''ARCHIVE FOLDER, INCLUDING CONTENT , REF #17
        objOUT.write vbnewline & now & vbtab & vbtab & " - ARCHIVING : " & objFOL.path & " : "
        objLOG.write vbnewline & now & vbtab & vbtab & " - ARCHIVING : " & objFOL.path & " : "
        call makZIP(objFOL.path, objFOL.path & ".zip")
        ''DELETE FOLDER, INCLUDING CONTENT , REF #17
        objOUT.write vbnewline & now & vbtab & vbtab & " - DELETING : " & objFOL.path
        objLOG.write vbnewline & now & vbtab & vbtab & " - DELETING : " & objFOL.path
        'retDEL = objFSO.deletefolder(objFOL.path, true)
        if (retDEL <> 0) then                    									    ''ERROR RETURNED
          objOUT.write vbnewline & now & vbtab & vbtab & " - ERROR DELETING : " & objFOL.path
          objLOG.write vbnewline & now & vbtab & vbtab & " - ERROR DELETING : " & objFOL.path
          errRET = retDEL
        end if
      elseif (not blnRUN) then															          ''SCRIPT NOT SET TO EXECUTE DELETION
        retDEL = 0
      end if
    elseif (intDIFF < cint(intAGE)) then       									      ''FOLDER HAS BEEN MODIFIED MORE RECENT THAN TARGET AGE
      objLOG.write vbnewline & vbnewline & now & vbtab & vbtab & " - EXCLUDED : " & objFOL.path & " : " & intDIFF & " Day(s)"
      retDEL = 0
    end if
  next
end sub

sub chkFIL(objMSP)
  set colFIL = objMSP.files
  for each objFIL in colFIL                    										    ''ENUMERATE EACH SUB-FILE
    ''DEFAULT FAIL
    retDEL = 7
    ''RETRIEVE SUB-FILE LAST DATE MODIFIED
    strDLM = objFIL.datelastmodified
    ''CALCULATE DATE DIFFERENCE (BY VALUE "D"AYS)
    intDIFF = -(datediff("d", now, strDLM))
    if (intDIFF >= cint(intAGE)) then           										  ''FILE HAS NOT BEEN MODIFIED IN TARGET AGE
      objOUT.write vbnewline & vbnewline & now & vbtab & vbtab & " - " & objFIL.path
      objOUT.write vbnewline & now & vbtab & vbtab & " - LAST MODIFIED : " & objFIL.DateLastModified & " : " & intDIFF & " Day(s)"
      objLOG.write vbnewline & vbnewline & now & vbtab & vbtab & " - " & objFIL.path
      objLOG.write vbnewline & now & vbtab & vbtab & " - LAST MODIFIED : " & objFIL.DateLastModified & " : " & intDIFF & " Day(s)"
      if (blnRUN) then																		            ''SCRIPT SET TO EXECUTE DELETION
        ''ARCHIVE FILE , REF #17
        objOUT.write vbnewline & now & vbtab & vbtab & " - ARCHIVING : " & objFIL.path & " : "
        objLOG.write vbnewline & now & vbtab & vbtab & " - ARCHIVING : " & objFIL.path & " : "
        call makZIP(objFIL.path, objFIL.path & ".zip")
        ''DELETE FILE , REF #17
        objOUT.write vbnewline & now & vbtab & vbtab & " - DELETING : " & objFIL.path
        objLOG.write vbnewline & now & vbtab & vbtab & " - DELETING : " & objFIL.path
        'retDEL = objFSO.delete(objFIL.path, true)
        if (retDEL <> 0) then                    									    ''ERROR RETURNED
          objOUT.write vbnewline & now & vbtab & vbtab & " - ERROR DELETING : " & objFIL.path
          objLOG.write vbnewline & now & vbtab & vbtab & " - ERROR DELETING : " & objFIL.path
          errRET = retDEL
        end if
      elseif (not blnRUN) then															          ''SCRIPT NOT SET TO EXECUTE DELETION
        retDEL = 0
      end if
    elseif (intDIFF < cint(intAGE)) then       									      ''FILE HAS BEEN MODIFIED MORE RECENT THAN TARGET AGE
      objLOG.write vbnewline & vbnewline & now & vbtab & vbtab & " - EXCLUDED : " & objFIL.path & " : " & intDIFF & " Day(s)"
      retDEL = 0
    end if
  next
end sub

sub chkLSV()																												  ''CHECK FOR MSP BACKUP LSV DESTINATION , REF #17
	if (strLSV <> vbnullstring) then																	  ''LOCAL SPEEDVAULT VARIABLE SET, CHECK PATH EXISTENCE
		objOUT.write vbnewline & vbnewline & now & " - CHECKING LSV DESTINATION : " & strLSV
		objLOG.write vbnewline & vbnewline & now & " - CHECKING LSV DESTINATION : " & strLSV
		if (not objFSO.folderexistts(strLSV)) then											  ''PATH DOES NOT EXIST
			objOUT.write vbnewline & now & vbtab & " - ERROR ACCESSING : " & strLSV & " : SCRIPT WILL ATTEMPT TO LOCATE LSV"
			objLOG.write vbnewline & now & vbtab & " - ERROR ACCESSING : " & strLSV & " : SCRIPT WILL ATTEMPT TO LOCATE LSV"
			strLSV = vbnullstring
		end if
	end if
	if (strLSV = vbnullstring) then																		  ''LOCAL SPEEDVAULT VARIABLE NOT SET, ATTEMPT TO LOCATE FROM DEVICE "LSV MONITOR" FILE
		objOUT.write vbnewline & vbnewline & now & " - ATTEMPTING TO LOCATE LSV DESTINATION"
		objLOG.write vbnewline & vbnewline & now & " - ATTEMPTING TO LOCATE LSV DESTINATION"
		''DEVICE "LSV MONITOR" FILE EXISTS
		if objFSO.fileexists("C:\temp\lsv.txt") then
			set objMSP = objFSO.opentextfile("C:\temp\lsv.txt")
			while (not objMSP.atendofstream)															  ''READ MSP BACKUP "LSV MONITOR" FILE LINE BY LINE
				strIN = objMSP.readline
				if (instr(1, strIN, "Device ")) then												  ''FOUND MSP BACKUP DEVICE ID
					strDEV = trim(right(strIN, len(strIN) - len("Device") - instrrev(strIN, "Device")))
					objOUT.write vbnewline & now & vbtab & strIN
					objOUT.write vbnewline & now & vbtab & " - FOUND MSP BACKUP DEVICE ID : " & strDEV
					objLOG.write vbnewline & now & vbtab & " - FOUND MSP BACKUP DEVICE ID : " & strDEV
				elseif (instr(1, strIN, "LocalSpeedVaultLocation ")) then		  ''FOUND MSP BACKUP ROOT LSV DESTINATION
					strLSV = trim(right(strIN, len(strIN) - len("LocalSpeedVaultLocation") - instrrev(strIN, "LocalSpeedVaultLocation")))
					objOUT.write vbnewline & now & vbtab & strIN
					objOUT.write vbnewline & now & vbtab & " - FOUND MSP BACKUP ROOT LSV DESTINATION : " & strLSV
					objLOG.write vbnewline & now & vbtab & " - FOUND MSP BACKUP ROOT LSV DESTINATION : " & strLSV
				end if
			wend
			set objMSP = nothing
			objOUT.write vbnewline & vbnewline & now & " - ATTEMPTING TO LOCATE DEVICE SPECIFIC LSV DESTINATION"
			objLOG.write vbnewline & vbnewline & now & " - ATTEMPTING TO LOCATE DEVICE SPECIFIC LSV DESTINATION"
			set objMSP = objFSO.getfolder(strLSV)
			set colFOL = objMSP.subfolders
			for each objFOL in colFOL																			  ''SEARCH EACH SUB-FOLDER IN MSP BACKUP ROOT LSV DESTINATION
				if (instr(1, objFOL.path, strDEV)) then											  ''MSP BACKUP DEVICE ID FOUND IN SUB-FOLDER
					strLSV = objFOL.path
					objOUT.write vbnewline & now & vbtab & " - FOUND DEVICE SPECIFIC LSV DESTINATION : " & strLSV
					objLOG.write vbnewline & now & vbtab & " - FOUND DEVICE SPECIFIC LSV DESTINATION : " & strLSV
					exit for
				end if
			next
			set colFOL = nothing
			set objMSP = nothing
		''DEVICE "LSV MONITOR" FILE DOES NOT EXIST
		elseif (not objFSO.fileexists("C:\temp\lsv.txt")) then
			objOUT.write vbnewline & vbnewline & now & " - MSP BACKUP LSV MONITOR FILE NOT PRESENT. SCRIPT REQUIRES PATH TO LOCAL MSP BACKUP DESTINATION, ENDING"
			objLOG.write vbnewline & vbnewline & now & " - MSP BACKUP LSV MONITOR FILE NOT PRESENT. SCRIPT REQUIRES PATH TO LOCAL MSP BACKUP DESTINATION, ENDING"
			''END SCRIPT
			errRET = 3
			'call CLEANUP()
		end if
	end if
end sub

''SUB-ROUTINES
sub CHKAU()																					                  ''CHECK FOR SCRIPT UPDATE, MSP_ROTATE.VBS, REF #2 , FIXES #26
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
					call FILEDL("https://github.com/CW-Khristos/scripts/raw/master/MSP%20Backups/MSP_Rotate.vbs", wscript.scriptname)
					''RUN LATEST VERSION
					if (wscript.arguments.count > 0) then                       ''ARGUMENTS WERE PASSED
						for x = 0 to (wscript.arguments.count - 1)
							strTMP = strTMP & " " & objARG.item(x)
						next
            objOUT.write vbnewline & now & vbtab & " - RE-EXECUTING  " & objSCR.nodename & " : " & objSCR.text & vbnewline
            objLOG.write vbnewline & now & vbtab & " - RE-EXECUTING  " & objSCR.nodename & " : " & objSCR.text & vbnewline
						objWSH.run "cscript.exe //nologo " & chr(34) & "c:\temp\" & wscript.scriptname & chr(34) & strTMP, 0, false
					elseif (wscript.arguments.count = 0) then                   ''NO ARGUMENTS WERE PASSED
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

sub FILEDL(strURL, strFILE)                   			                  ''CALL HOOK TO DOWNLOAD FILE FROM URL
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

sub HOOK(strCMD)                                                      ''CALL HOOK TO MONITOR OUTPUT OF CALLED COMMAND
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
  'retSTOP = objHOOK.exitcode
  set objHOOK = nothing
  if (err.number <> 0) then
    objOUT.write vbnewline & now & vbtab & vbtab & err.number & vbtab & err.description
    objLOG.write vbnewline & now & vbtab & vbtab & err.number & vbtab & err.description
    errRET = 1
    err.clear
  end if
end sub

sub CLEANUP()                                												  ''SCRIPT CLEANUP
  if (errRET = 0) then                       												  ''NO ERROR RETURNED
    objOUT.write vbnewline & vbnewline & now & " - ROTATE MSP BACKUPSET : COMPLETE : SUCCESS"
    objLOG.write vbnewline & vbnewline & now & " - ROTATE MSP BACKUPSET : COMPLETE : SUCCESS"
    err.clear
  elseif (errRET <> 0) then                  												  ''ERROR RETURNED
    objOUT.write vbnewline & vbnewline & now & " - ROTATE MSP BACKUPSET : COMPLETE : ERROR"
    objLOG.write vbnewline & vbnewline & now & " - ROTATE MSP BACKUPSET : COMPLETE : ERROR"
    ''RAISE CUSTOMIZED ERROR CODE, ERROR CODE WILL BE DEFINE RESTOP NUMBER INDICATING WHICH SECTION FAILED
    call err.raise(vbObjectError + errRET, "ROTATE MSP BACKUPSET", "FAIL")
  end if
  ''EMPTY OBJECTS
  set objFOL = nothing
  set colFOL = nothing
  set objMSP = nothing
  set objLOG = nothing
  set objFSO = nothing
  set objSHL = nothing
  set objWSH = nothing
  set objARG = nothing
  set objOUT = nothing
  set objIN = nothing
  ''END SCRIPT, RETURN ERROR
  wscript.quit err.number
end sub
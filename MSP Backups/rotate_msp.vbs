'on error resume next
''DEFINE VARIABLES
dim objIN, objOUT, objARG, objWSH, objFSO, objMSP, colFOL, objFOL
dim errRET, retDEL, strIN, strLSV, strDEV, strDLM, intDIFF, strRUN
''DEFAULT SUCCESS
errRET = 0
''STDIN / STDOUT
set objIN = wscript.stdin
set objOUT = wscript.stdout
set objARG = wscript.arguments
''CREATE OBJECTS
set objWSH = createobject("wscript.shell")
set objFSO = createobject("scripting.filesystemobject")
''CHECK EXECUTION METHOD OF SCRIPT
strIN = lcase(mid(wscript.fullname, instrrev(wscript.fullname, "\") + 1))
if (strIN <> "cscript.exe") Then
  objOUT.write vbnewline & now & " - SCRIPT LAUNCHED VIA EXPLORER, EXECUTING SCRIPT VIA CSCRIPT..."
  objWSH.run "cscript.exe //nologo " & chr(34) & Wscript.ScriptFullName & chr(34)
  wscript.quit
end if
''PREPARE LOGFILE
if (objFSO.fileexists("C:\temp\rotate_msp")) then     							''LOGFILE EXISTS
  objFSO.deletefile "C:\temp\rotate_msp", true
  set objLOG = objFSO.createtextfile("C:\temp\rotate_msp")
  objLOG.close
  set objLOG = objFSO.opentextfile("C:\temp\rotate_msp", 8)
else                                            										''LOGFILE NEEDS TO BE CREATED
  set objLOG = objFSO.createtextfile("C:\temp\rotate_msp")
  objLOG.close
  set objLOG = objFSO.opentextfile("C:\temp\rotate_msp", 8)
end if
''READ PASSED COMMANDLINE ARGUMENTS
if (wscript.arguments.count > 0) then        												''ARGUMENTS WERE PASSED
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
    strRUN = objARG.item(2)
  else
    intAGE = 60
    strRUN = "false"
  end if
end if
if ((wscript.arguments.count = 0) or (strLSV = vbnullstring)) then  ''NO ARGUMENTS PASSED, END SCRIPT
  objOUT.write vbnewline & now & " - SCRIPT REQUIRES PATH TO LOCAL MSP BACKUP DESTINATION"
  objLOG.write vbnewline & now & " - SCRIPT REQUIRES PATH TO LOCAL MSP BACKUP DESTINATION"
end if
'CHECK LSV IF CONFIGURED, OTHERWISE ATTEMPT TO LOCATE
call chkLSV()
''RUN MAIN SCRIPT
call rotMSP()
''END SCRIPT
call CLEANUP()

''SUB-ROUTINES
sub rotMSP()																												''ROTATE_MSP MAIN SUB-ROUTINE
	''RUN ROTATE_MSP
	objOUT.write vbnewline & vbnewline & now & " - ROTATE MSP BACKUPSET : RUNNING : DELETION : " & strRUN
	objLOG.write vbnewline & vbnewline & now & " - ROTATE MSP BACKUPSET : RUNNING : DELETION : " & strRUN
	''RETRIEVE MSP BACKUPSET FOLDER
	set objMSP = objFSO.getfolder(strLSV)
	if (err.number <> 0) then																					''ERROR OBTAINING FOLDER
		objOUT.write vbnewline & vbnewline & now & " - ROTATE MSP BACKUPSET : ERROR : " & err.description
		objLOG.write vbnewline & vbnewline & now & " - ROTATE MSP BACKUPSET : ERROR : " & err.description
		errRET = 2
	elseif (err.number = 0) then																			''SUCCESSFULLY OBTAINED FOLDER
		''CHECK THROUGH ALL SUB-FOLDERS IN MSP BACKUPSET FOLDER
		''THIS NEEDS TO BE THE COMPLETE PATH THAT MSP BACKUPS WRITE TO, IE : F:\MSP Backups\jax-dc1_lr0xa_E93A9C6948C0DD5FF5E9\
		set colFOL = objMSP.subfolders
		for each objFOL in colFOL                    										''ENUMERATE EACH SUB-FOLDER
			''DEFAULT FAIL
			retDEL = 3
			''RETRIEVE SUB-FOLDER LAST DATE MODIFIED
			strDLM = objFOL.datelastmodified
			''CALCULATE DATE DIFFERENCE (BY VALUE "D"AYS)
			intDIFF = -(datediff("d", now, strDLM))
			if (intDIFF > cint(intAGE)) then           										''FOLDER HAS NOT BEEN MODIFIED IN TARGET AGE
				objOUT.write vbnewline & vbnewline & now & vbtab & " - " & objFOL.path
				objOUT.write vbnewline & now & vbtab & " - " & vbtab & "LAST MODIFIED : " & objFOL.DateLastModified & " : " & intDIFF & " Day(s)"
				objLOG.write vbnewline & vbnewline & now & vbtab & " - " & objFOL.path
				objLOG.write vbnewline & now & vbtab & " - " & vbtab & "LAST MODIFIED : " & objFOL.DateLastModified & " : " & intDIFF & " Day(s)"
				if (strRUN = "true") then																		''SCRIPT SET TO EXECUTE DELETION
					''DELETE FOLDER, INCLUDING CONTENT
					objOUT.write vbnewline & now & vbtab & " - DELETING : " & objFOL.path
					objLOG.write vbnewline & now & vbtab & " - DELETING : " & objFOL.path
					retDEL = objFSO.deletefolder(objFOL.path, true)
					if (retDEL <> 0) then                    									''ERROR RETURNED
						objOUT.write vbnewline & now & vbtab & " - ERROR DELETING : " & objFOL.path
						objLOG.write vbnewline & now & vbtab & " - ERROR DELETING : " & objFOL.path
						errRET = retDEL
					end if
				elseif (strRUN = "false") then															''SCRIPT NOT SET TO EXECUTE DELETION
					retDEL = 0
				end if
			elseif (intDIFF <= cint(intAGE)) then       									''FOLDER HAS BEEN MODIFIED MORE RECENT THAN TARGET AGE
				objLOG.write vbnewline & vbnewline & now & vbtab & " - EXCLUDED : " & objFOL.path & " : " & intDIFF & " Day(s)"
				retDEL = 0
			end if
		next
	end if
end sub

sub chkLSV()																												''CHECK FOR MSP BACKUP LSV DESTINATION
	if (strLSV <> vbnullstring) then																	''LOCAL SPEEDVAULT VARIABLE SET, CHECK PATH EXISTENCE
		objOUT.write vbnewline & vbnewline & now & " - CHECKING LSV DESTINATION : " & strLSV
		objLOG.write vbnewline & vbnewline & now & " - CHECKING LSV DESTINATION : " & strLSV
		if not(objFSO.folderexistts(strLSV)) then												''PATH DOES NOT EXIST
			objOUT.write vbnewline & now & vbtab & " - ERROR ACCESSING : " & strLSV & " : SCRIPT WILL ATTEMPT TO LOCATE LSV"
			objLOG.write vbnewline & now & vbtab & " - ERROR ACCESSING : " & strLSV & " : SCRIPT WILL ATTEMPT TO LOCATE LSV"
			strLSV = vbnullstring
		end if
	end if
	if (strLSV = vbnullstring) then																		''LOCAL SPEEDVAULT VARIABLE NOT SET, ATTEMPT TO LOCATE FROM DEVICE "LSV MONITOR" FILE
		objOUT.write vbnewline & vbnewline & now & " - ATTEMPTING TO LOCATE LSV DESTINATION"
		objLOG.write vbnewline & vbnewline & now & " - ATTEMPTING TO LOCATE LSV DESTINATION"
		''DEVICE "LSV MONITOR" FILE EXISTS
		if objFSO.fileexists("C:\Program Files\Backup Manager\lsv.txt") then
			set objMSP = objFSO.opentextfile("C:\Program Files\Backup Manager\lsv.txt")
			while (not objMSP.atendofstream)															''READ MSP BACKUP "LSV MONITOR" FILE LINE BY LINE
				strIN = objMSP.readline
				if (instr(1, strIN, "Device ")) then												''FOUND MSP BACKUP DEVICE ID
					strDEV = trim(right(strIN, len(strIN) - len("Device") - instrrev(strIN, "Device")))
					objOUT.write vbnewline & now & vbtab & strIN
					objOUT.write vbnewline & now & vbtab & " - FOUND MSP BACKUP DEVICE ID : " & strDEV
					objLOG.write vbnewline & now & vbtab & " - FOUND MSP BACKUP DEVICE ID : " & strDEV
				elseif (instr(1, strIN, "LocalSpeedVaultLocation ")) then		''FOUND MSP BACKUP ROOT LSV DESTINATION
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
			for each objFOL in colFOL																			''SEARCH EACH SUB-FOLDER IN MSP BACKUP ROOT LSV DESTINATION
				if (instr(1, objFOL.path, strDEV)) then											''MSP BACKUP DEVICE ID FOUND IN SUB-FOLDER
					strLSV = objFOL.path
					objOUT.write vbnewline & now & vbtab & " - FOUND DEVICE SPECIFIC LSV DESTINATION : " & strLSV
					objLOG.write vbnewline & now & vbtab & " - FOUND DEVICE SPECIFIC LSV DESTINATION : " & strLSV
					exit for
				end if
			next
			set colFOL = nothing
			set objMSP = nothing
		''DEVICE "LSV MONITOR" FILE DOES NOT EXIST
		elseif (not objFSO.fileexists("C:\Program Files\Backup Manager\lsv.txt")) then
			objOUT.write vbnewline & vbnewline & now & " - MSP BACKUP LSV MONITOR FILE NOT PRESENT. SCRIPT REQUIRES PATH TO LOCAL MSP BACKUP DESTINATION, ENDING"
			objLOG.write vbnewline & vbnewline & now & " - MSP BACKUP LSV MONITOR FILE NOT PRESENT. SCRIPT REQUIRES PATH TO LOCAL MSP BACKUP DESTINATION, ENDING"
			''END SCRIPT
			errRET = 1
			call CLEANUP()
		end if
	end if
end sub

sub CLEANUP()                                												''SCRIPT CLEANUP
  if (errRET = 0) then                       												''NO ERROR RETURNED
    objOUT.write vbnewline & vbnewline & now & " - ROTATE MSP BACKUPSET : COMPLETE : SUCCESS"
    objLOG.write vbnewline & vbnewline & now & " - ROTATE MSP BACKUPSET : COMPLETE : SUCCESS"
    err.clear
  elseif (errRET <> 0) then                  												''ERROR RETURNED
    objOUT.write vbnewline & vbnewline & now & " - ROTATE MSP BACKUPSET : COMPLETE : ERROR"
    objLOG.write vbnewline & vbnewline & now & " - ROTATE MSP BACKUPSET : COMPLETE : ERROR"
    ''RAISE CUSTOMIZED ERROR CODE, ERROR CODE WILL BE DEFINE RESTOP NUMBER INDICATING WHICH SECTION FAILED
    call err.raise(vbObjectError + errRET, "ROTATE MSP BACKUPSET", "FAIL")
  end if
  ''EMPTY OBJECTS
  set objFOL = nothing
  set colFOL = nothing
  set objMSP = nothing
  SET objLOG = nothing
  set objFSO = nothing
  set objWSH = nothing
  set objARG = nothing
  set objOUT = nothing
  set objIN = nothing
  ''END SCRIPT, RETURN ERROR
  wscript.quit err.number
end sub
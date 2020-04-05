		'V--V   exchDEFRAG.vbs   V--V'
		'V--V VampyreMan Studios V--V'
		'V--V  Author : Khristos V--V'
		'Script to dis-mount original edb file
		'Backup edb file to specified file
		'Re-mount the original edb file
		'Defrag the backup edb file
		'Dis-mount/move original edb file
		'Move/mount the defragged backup edb file

	''CREATE VBS SHELL OBJECT AND FILE SYSTEM OBJECT''
dim objWSH, objFSO, objStream
dim blnLog , dPrep, wrtLog, fNum
dim colDrives, objDrive, objFOLp, objUTILp

set objWSH = createobject("wscript.shell")
set objFSO = createobject("scripting.filesystemobject")

	''REMIND USER TO DISMOUNT EXCHANGE STORE FIRST''
dPrep = msgbox("Was the Exchange Store Dis-Mounted Before Executing exchDEFRAG Script?", vbyesno, _
  "exchDEFRAG - Khristos")

	''IF THE STORE WASN'T DISMOUNTED, DISPLAY AN IMAGE''
	''DETAILING STEP BY STEP TO DISMOUNT STORE, EXIT''
	''THE exchDEFRAG SCRIPT''
if dPrep <> vbyes then
  ''Open Image File''
  Cleaner
end if

	''ASK THE USER TO CREATE A LOGFILE''
wrtLog = msgbox("Create Logfile of .edb and .stm Files and Locations?", vbyesno, "exchDEFRAG - Khristos")

	''IF THE USER WANTS TO WRITE A LOGFILE THEN REQUEST A FILE PATH TO WRITE TO''
	''AND WARN THE USER THAT SPECFYING A FILE THAT ALREADY EXISTS WILL OVERWRITE THAT FILE''
if wrtLog = vbyes then

	''REQUEST USER INPUT''
  wrtLog = inputbox("SPECIFYING A LOGFILE THAT ALREADY EXISTS WILL OVERWRITE THAT LOGFILE!" & vbnewline & _
    vbtab & "Input Logfile Directory and File Name","exchDEFRAG - Khristos")

	''VERIFY INPUT IS NOT CANCEL BUTTON (-1) AND IS NOT BLANK''
  if wrtLog <> -1 then
    if wrtLog <> vbnullstring then

	''SET 'blnLog' EQUAL TO "TRUE" FOR USE LATER WHEN WE WANT TO''
	''KNOW IF THE USE SPECIFIED WHETHER TO WRITE A LOGFILE OR NOT''
	''CREATE THE VBS TEXT STREAM AND CREATE THE TEXT FILE TO WRITE''
	''EXTRA OPTION (TRUE) TO OVERWRITE IF FILE EXISTS''
      blnLog = "TRUE"
      set objStream = objFSO.createtextfile(wrtLog, True)

    else

	''IF THE USER DOESN'T WANT A LOGFILE, SET 'blnLog' EQUAL TO "FALSE"''
      blnLog = "FALSE"

    end if
  else
    blnLog = "FALSE"
  end if
end if

	''FOR MULTIPLE DRIVE SUPPORT, ENUMERATE DRIVES, CHECK EACH 'Program Files' ''
set colDrives = objFSO.drives
for each objDrive in colDrives
  objFOLp =  objDrive.driveletter & ":\Program Files"

        ''VARIABLE 'fNum' WILL TRACK THE NUMBER OF EXCHANGE DATA FILES''
	''CALL 'CheckFolder' SUB ROUTINE WITH PARAMETER 'objFSO.getfolder(objFOLp)' ''
	''objFSO.getfolder() USES VBS FILE SYSTEM TO SPECIFY A FOLDER DIRECTORY''
	''IN THIS CASE, THE WINDOWS 'Program Files' DIRECTORY''
  fNum = 0
  CheckFolder (objFSO.getfolder(objFOLp))

next

        ''IF NO ECHANGE DATA FILES WERE FOUND REQUEST A USER SPECIFIED DIRECTORY''
if fNum = 0 then
  specFOL
end if

Cleaner

		''CHECKFOLDER SUB ROUTINE''
sub CheckFolder (objCurrentFolder)

		''SET VARIABLES FOR THE TWO EXCHANGE DATABASE FILE TYPES''
		''EDB AND STM ARE THE TWO FILE TYPES WE WANT TO LOOK FOR''
  dim strSearch1, strSearch2, strTemp
  dim strOutput, backDIR, dExec, dRept

  strSearch1 = ".stm"
  strSearch2 = ".edb"

		''FOR EVERY FILE IN THE FOLDER PASSED TO 'CheckFolder' SUB ROUTINE''
  for each objFile in objCurrentFolder.files

		''SET 'strTemp' EQUAL TO THE LAST 4 CHARACTERS OF THE FILE NAME''
    strTemp = right(objFile.name, 4)

		''COMPARE 'strTemp' TO THE VALUE OF THE VARIABLES DEFINED EARLIER''
		''lcase() WILL MAKE CHARACTERS LOWER CASE SO MATCHING CASE WON'T MATTER''
		''THIS DETERMINES IF THE CURRENT FILE IS ONE OF THE TWO TYPES WE WANT''
    if lcase(strTemp) = lcase(strSearch1) then

                ''ADD 1 TO THE FILE COUNT''
      fNum = fNum + 1

		''IF USER WANTED TO CREATE A LOGFILE, SET 'strOutput' EQUAL TO FILE NAME,''
		''FILE PATH, FILE SIZE, FILE TYPE, LAST ACCESSED. cstr() CONVERTS VALUES TO STRINGS''
      if blnLog = "TRUE" then
        strOutput = cstr(objFile.name) & "," & cstr(objFile.path) & "," & cstr(objFile.size) _ 
          & "," & cstr(objFile.type) & "," & cstr(objFile.datelastaccessed)

		''WRITE 'strOutput' TO FILE USER SPECIFIED EARLIER''
        objStream.writeline strOutput

      end if

		''REQUEST A DIRECTORY TO STORE COPIES OF EXCHANGE FILES''
		''EXCHANGE FILES ARE !VERY! IMPORTANT. WE WANT BACKUPS IF SOMETHING HAPPENS''
      backDIR = inputbox("Enter Directory for Backup of " & objFile.name, "exchDEFRAG - Khristos")

		''CHECK TO ENSURE USER SPECIFIED A DIRECTORY THAT EXISTS''
      if objFSO.folderexists(backDIR) = false then
		''IF THE FOLDER DOESN'T EXIST, CREATE IT''
        objFSO.createfolder(backDIR)

      end iF

		''CHECK FOR ENDING "\" IN DIRECTORY'
      if right(backDIR, 1) <> "\" then

		''IF "\" ISN'T AT THE END, ADD IT''
        backDIR = backDIR & "\"

      end if

		''COPY THE CURRENT FILE TO THE USER SPECIFIED DIRECTORY''
      objFile.copy(backDIR & objFile.name)

    end if

		''THIS IS THE SAME CONDITIONAL AS ABOVE, FOR THE SECOND FILE TYPE''
    if lcase(strTemp) = lcase(strSearch2) then
      fNum = fNum + 1
      if blnLog = "TRUE" then
        strOutput = cstr(objFile.name) & "," & cstr(objFile.path) & "," & cstr(objFile.size) _ 
          & "," + cstr(objFile.type) & "," & cstr(objFile.datelastaccessed)
        objStream.writeline strOutput
      end if
      backDIR = inputbox("Enter directory for backup of " & objFile.name, "exchDEFRAG - Khristos")
      if objFSO.folderexists(backDIR) = false then
        objFSO.createfolder(backDIR)
      end if
      if right(backDIR, 1) <> "\" then
        backDIR = backDIR & "\"
      end if
      objFile.copy(backDIR & objFile.name)

		''ASK IS THE SELECTED FILE IS THE EXCHANGE FILE TO BE DEFRAGED BY ESEUTIL''
      dExec = msgbox("Is the Following File :" & vbnewline & vbtab & objFile.path & vbnewline & _
        "the Exchange File You Wish to Defrag?", vbyesno, "exchDEFRAG - Khristos")

      if dExec = vbyes then

		''SEARCH 'Program Files' FOR ESEUTIL''
        locateUTIL (objFOLp)

		''RUN ESEUTIL DEFRAG ON THE BACKUP LEAVING THE ORIGINAL INTACT FOR REMOUNTING''
		''DEFRAG IS ONLY NECESSARY FOR THE EDB FILE TYPE, THIS LINE IS NOT IN FIRST CONDITIONAL''
        objWSH.run (objUTILp & " /d" & """ & backDIR & objFile.name & """)

		''ASK FOR USER INPUT WHEN ESEUTIL IS FINISHED. EFFECTIVELY MAKES''
                ''SCRIPT WAIT FOR ESEUTIL BEFORE CONTINUING. ASK USER IF THEY WANT''
                ''TO DEFRAG THE NEXT ECHANGE DATA FILE.''
        dRept = msgbox("You May Now Re-Mount the Exhange File :" & vbnewline & vbtab & objFile.path & vbnewline & _
          "While ESEutil Defrags the Copy :" & vbnewline & vbtab & backDIR & objFile.name & vbnewline & _
          "Once ESEutil has Completed Dis-Mount the Store, Replace the old File with the Copy, then " & _
          "Re-Mount the Store Once Again, Maintaining Shorter Service Outages. Do You Wish to Defrag " & _
          "the Other Exchange Data File?", vbyesno, "exchDEFRAG - Khristos")

        if dRept = vbno then
          exit for
        end if
      end if
    end if
	''REPEAT PROCESS FOR THE NEXT FILE''
  next

	''FOR EVERY SUBFOLDER, CALL 'CheckFolder' SUB ROUTINE AGAIN''
  for each objNewFolder in objCurrentFolder.subfolders
    CheckFolder (objNewFolder)
  next

  set objFile = nothing
  set objNewFolder = nothing
  set objCurrentFolder = nothing
end sub

		''ROUTINE TO LOCATE ESEUTIL''
sub locateUTIL (objUTILfol)
  dim blnFind, strUTIL
  blnFind = "FALSE"
  strUTIL = "eseutil"
  for each objFile1 in objUTILfol.files
    if instr(1, objFile1, strUTIL) then
      blnFind = "TRUE"
      objUTILp = objFile1.name
    end if
  next
  if blnFind = "TRUE" then
    set objFile1 = nothing
    set objNewFolder1 = nothing
    set objUTILfol = nothing
    exit sub
  end if
  for each objNewFolder1 in objUTILfol.subfolders
    locateUTIL (objNewFolder1)
  next
  if blnFind = "FALSE" then
    msgbox "ESEutil Not Found!", vbokonly, "exchDEFRAG - Khristos"
    Cleaner
  end if
  set objFile1 = nothing
  set objNewFolder1 = nothing
  set objUTILfol = nothing
end sub

        ''USER SPECIFIED DIRECTORY ROUTINE''
sub specFOL
    objFOLp = inputbox("No Exchange Data Files Found. Specify a Directory?" & vbnewline & vbnewline & _
      "Click Cancel or Enter a Blank Directory to Quit.", "exchDEFRAG - Khristos")
    if objFOLp <> vbcancel then
      if objFOLp <> vbnullstring then
        if objFSO.folderexists(objFOLp) then
          CheckFolder (objFOLp)
        else
          msgbox "Invalid Directory!", vbokonly, "exchDEFRAG - Khristos"
          specFOL
        end if
      else
        Cleaner
      end if
    end if       
end sub

	''SET ALL OBJECTS EQUAL TO NOTHING (CLEAN UP) AND QUIT SCRIPT''
sub Cleaner
  set objUTILp = nothing
  set objStream = nothing
  set objFOLp = nothing
  set objFSO = nothing
  set objWSH = nothing
  wscript.echo "exchDEFRAG has Finished all Processes. Ending now.."
  wscript.quit
end sub
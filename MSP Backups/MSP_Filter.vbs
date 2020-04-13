''MSP_FILTER.VBS
''DESIGNED TO AUTOMATE PASSING OF BACKUP FILTERS TO MSP BACKUP SOFTWARE VIA CLIENTTOOL
''DOWNLOADS 'FILTERS.TXT' FROM GITHUB; THIS FILE CONTAINS EACH BACKUP FILTER IN A 'LINE BY LINE' FORMAT
''DESIGNED TO AUTOMATE PASSING OF BACKUP INCLUSIONS TO MSP BACKUP SOFTWARE VIA CLIENTTOOL
''DOWNLOADS 'INCLUDES.TXT' FROM GITHUB; THIS FILE CONTAINS EACH BACKUP FILTER IN A 'LINE BY LINE' FORMAT
''ACCEPTS 3 PARAMETER , REQUIRES 1 PARAMETER
''REQUIRED PARAMETER 'STROPT' ; STRING VALUE TO INDICATE 'LOCAL' OR 'CLOUD' FILTER OPERATION
''OPTIONAL PARAMETER 'STRFILTER' ; STRING VALUE TO HOLD PASSED 'FILTERS' ; SEPARATE MULTIPLE 'FILTERS' VIA '|'
''OPTIONAL PARAMETER 'STRINCL' ; STRING VALUE TO HOLD PASSED 'INCLUSIONS' ; SEPARATE MULTIPLE 'INCLUSIONS' VIA '|'
''WRITTEN BY : CJ BLEDSOE / CJ<@>THECOMPUTERWARRIORS.COM
on error resume next
''SCRIPT VARIABLES
dim errRET, strVER
''VARIABLES ACCEPTING PARAMETERS
dim strINCL, arrINCL
dim strFILTER, arrFILTER
dim strIN, strOUT, strOPT, strRCMD
''PRE-DEFINED ARRAYS
dim arrUSR(), arrPROT()
''USER AND USER FOLDER ARRAYS
dim objFOL, arrFOL()
''USER FOLDER AND SUB-FOLDER ARRAYS
dim objUFOL, arrUFOL()
''SCRIPT OBJECTS
dim objIN, objOUT, objARG, objWSH, objFSO
dim objLOG, objEXEC, objHOOK, objHTTP, objXML
''VERSION FOR SCRIPT UPDATE , MSP_FILTER.VBS , REF #2
strVER = 5
strREPO = "scripts"
strBRCH = "dev"
strDIR = "MSP Backups"
''DEFAULT SUCCESS
errRET = 0
''STDIN / STDOUT
set objIN = wscript.stdin
set objOUT = wscript.stdout
set objARG = wscript.arguments
''OBJECTS FOR LOCATING FOLDERS
set objWSH = createobject("wscript.shell")
set objFSO = createobject("scripting.filesystemobject")
''PREPARE LOGFILE
if (objFSO.fileexists("C:\temp\msp_filter")) then           ''LOGFILE EXISTS
  objFSO.deletefile "C:\temp\msp_filter", true
  set objLOG = objFSO.createtextfile("C:\temp\msp_filter")
  objLOG.close
  set objLOG = objFSO.opentextfile("C:\temp\msp_filter", 8)
else                                                        ''LOGFILE NEEDS TO BE CREATED
  set objLOG = objFSO.createtextfile("C:\temp\msp_filter")
  objLOG.close
  set objLOG = objFSO.opentextfile("C:\temp\msp_filter", 8)
end if
''READ PASSED COMMANDLINE ARGUMENTS
if (wscript.arguments.count > 0) then                       ''ARGUMENTS WERE PASSED
  for x = 0 to (wscript.arguments.count - 1)
    objOUT.write vbnewline & now & vbtab & " - ARGUMENT " & (x + 1) & " (ITEM " & x & ") " & " PASSED : " & ucase(objARG.item(x))
    objLOG.write vbnewline & now & vbtab & " - ARGUMENT " & (x + 1) & " (ITEM " & x & ") " & " PASSED : " & ucase(objARG.item(x))
  next 
  if (wscript.arguments.count > 0) then                     ''SET VARIABLES ACCEPTING ARGUMENTS
    strOPT = objARG.item(0)                                 ''SET REQUIRE PARAMETER 'STROPT' , 'LOCAL' OR 'CLOUD' OPERATION
    if (wscript.arguments.count > 1) then                   ''SET OPTIONAL PARAMETER 'STRFILTER' , BACKUP FILTERS STRING
      strFILTER = objARG.item(1)
      ''FILL 'ARRFILTER' BACKUP FILTER ARRAY
      objOUT.write vbnewline & vbtab & strFILTER
      arrFILTER = split(strFILTER, "|")
      for intTMP = 0 to ubound(arrFILTER)
        objOUT.write vbnewline & vbtab & ubound(arrFILTER) & vbtab & arrFILTER(intTMP)
      next
    end if
    if (wscript.arguments.count > 2) then                   ''SET OPTIONAL PARAMETER 'STRINCL' , BACKUP INCLUDES STRING
      strINCL = objARG.item(1)
      ''FILL 'ARRINCL' BACKUP INCLUDES ARRAY
      objOUT.write vbnewline & vbtab & strINCL
      arrINCL = split(strINCL, "|")
      for intTMP = 0 to ubound(arrINCL)
        objOUT.write vbnewline & vbtab & ubound(arrINCL) & vbtab & arrINCL(intTMP)
      next
    end if
  end if
else                                                        ''NOT ENOUGH ARGUMENTS PASSED , END SCRIPT , 'ERRRET'=1
  call LOGERR(1)
end if
''UNNEEDED / TO EXCLUDE USER ACCOUNTS
redim arrUSR(9)
arrUSR(0) = "rmmtech"
arrUSR(1) = "Guest"
arrUSR(2) = "__sbs_netsetup__"
arrUSR(3) = "Default"
arrUSR(4) = "Default.migrated"
arrUSR(5) = "Default User"
arrUSR(6) = "DefaultAccount"
arrUSR(7) = "WDAGUtilityAccount"
arrUSR(8) = "UpdatusUser"
''PROTECTED EXT / FILES / DIRECTORIES
redim arrPROT(2)
arrPROT(0) = ".OST"
arrPROT(1) = ".PST"

''------------
''BEGIN SCRIPT
strTMP = vbnullstring
if (errRET <> 0) then                                       ''NO ARGUMENTS PASSED, END SCRIPT , 'ERRRET'=1
  call CLEANUP()
elseif (errRET = 0) then                                    ''ARGUMENTS PASSED, CONTINUE SCRIPT
	objOUT.write vbnewline & vbnewline & now & vbtab & " - EXECUTING MSP_FILTER"
	objLOG.write vbnewline & vbnewline & now & vbtab & " - EXECUTING MSP_FILTER"
	''AUTOMATIC UPDATE, MSP_FILTER.VBS, REF #2
	call CHKAU()
  Select Case lcase(strOPT)
    ''PERFORM 'LOCAL' FILTER CONFIGURATIONS
    Case "local"
      ''DOWNLOAD 'FILTERS.TXT' BACKUP FILTERS DEFINITION FILE , 'ERRRET'=2 , REF #2
      objOUT.write vbnewline & now & vbtab & vbtab & " - DOWNLOADING 'FILTERS.TXT' BACKUP FILTER DEFINITION"
      objLOG.write vbnewline & now & vbtab & vbtab & " - DOWNLOADING 'FILTERS.TXT' BACKUP FILTER DEFINITION"
      ''REMOVE PREVIOUS 'FILTERS.TXT' FILE
      erase arrTMP
      strTMP = vbnullstring
      if (objFSO.fileexists("C:\temp\filters.txt")) then
        objFSO.deletefile "C:\temp\filters.txt", true
      end if
      call FILEDL("https://github.com/CW-Khristos/scripts/raw/master/MSP%20Backups/filters.txt", "filters.txt")
      set objTMP = objFSO.opentextfile("C:\temp\filters.txt", 1)
      while (not objTMP.atendofstream)
        strTMP = strTMP & objTMP.readline
      wend
      objTMP.close
      set objTMP = nothing
      arrTMP = split(strTMP, "|")
      for intTMP = 0 to ubound(arrTMP)
        if (arrTMP(intTMP) <> vbnullstring) then
          objOUT.write vbnewline & now & vbtab & vbtab & "EXECUTING : C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -exclude " & chr(34) & arrTMP(intTMP) & chr(34)
          objLOG.write vbnewline & now & vbtab & vbtab & "EXECUTING : C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -exclude " & chr(34) & arrTMP(intTMP) & chr(34)
          call HOOK("C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -exclude " & chr(34) & arrTMP(intTMP) & chr(34))
        end if
      next
      ''CUSTOM 'FILTER' PASSED
      if (strFILTER <> vbnullstring) then
        for intTMP = 0 to ubound(arrFILTER)
          if (arrFILTER(intTMP) <> vbnullstring) then
            objOUT.write vbnewline & now & vbtab & vbtab & "EXECUTING : C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -exclude " & chr(34) & arrFILTER(intTMP) & chr(34)
            objLOG.write vbnewline & now & vbtab & vbtab & "EXECUTING : C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -exclude " & chr(34) & arrFILTER(intTMP) & chr(34)
            call HOOK("C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -exclude " & chr(34) & arrFILTER(intTMP) & chr(34))
          end if
        next
      end if
      ''DOWNLOAD 'INCLUDES.TXT' BACKUP INCLUDES DEFINITION FILE , 'ERRRET'=2 , REF #2
      objOUT.write vbnewline & now & vbtab & vbtab & " - DOWNLOADING 'INCLUDES.TXT' BACKUP INCLUDES DEFINITION"
      objLOG.write vbnewline & now & vbtab & vbtab & " - DOWNLOADING 'INCLUDES.TXT' BACKUP INCLUDES DEFINITION"
      ''REMOVE PREVIOUS 'INCLUDES.TXT' FILE
      erase arrTMP
      strTMP = vbnullstring
      if (objFSO.fileexists("C:\temp\includes.txt")) then
        objFSO.deletefile "C:\temp\includes.txt", true
      end if
      call FILEDL("https://github.com/CW-Khristos/scripts/raw/master/MSP%20Backups/includes.txt", "includes.txt")
      set objTMP = objFSO.opentextfile("C:\temp\includes.txt", 1)
      while (not objTMP.atendofstream)
        strTMP = strTMP & objTMP.readline
      wend
      objTMP.close
      set objTMP = nothing
      arrTMP = split(strTMP, "|")
      for intTMP = 0 to ubound(arrTMP)
        if (arrTMP(intTMP) <> vbnullstring) then
          objOUT.write vbnewline & now & vbtab & vbtab & "EXECUTING : C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -include " & chr(34) & arrTMP(intTMP) & chr(34)
          objLOG.write vbnewline & now & vbtab & vbtab & "EXECUTING : C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -include " & chr(34) & arrTMP(intTMP) & chr(34)
          call HOOK("C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -include " & chr(34) & arrTMP(intTMP) & chr(34))
        end if
      next
      ''CUSTOM 'INCLUDE' PASSED
      if (strINCL <> vbnullstring) then
        for intTMP = 0 to ubound(arrINCL)
          if (arrINCL(intTMP) <> vbnullstring) then
            objOUT.write vbnewline & now & vbtab & vbtab & "EXECUTING : C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -include " & chr(34) & arrINCL(intTMP) & chr(34)
            objLOG.write vbnewline & now & vbtab & vbtab & "EXECUTING : C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -include " & chr(34) & arrINCL(intTMP) & chr(34)
            call HOOK("C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -include " & chr(34) & arrINCL(intTMP) & chr(34))
          end if
        next
      end if
    ''PERFORM 'CLOUD' FILTER CONFIGURATIONS
    case "cloud"
      ''RESET CURRENT BACKUP INCLUDES , REF #2
      objOUT.write vbnewline & now & vbtab & vbtab & " - RESETTING CURRENT MSP BACKUP INCLUDES"
      objLOG.write vbnewline & now & vbtab & vbtab & " - RESETTING CURRENT MSP BACKUP INCLUDES"
      call HOOK("C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -include C:\")
      wscript.sleep 5000
      ''DOWNLOAD 'CLOUD_FILTERS.TXT' BACKUP FILTERS DEFINITION FILE , 'ERRRET'=2 , REF #2
      objOUT.write vbnewline & now & vbtab & vbtab & " - DOWNLOADING 'CLOUD_FILTERS.TXT' BACKUP FILTER DEFINITION"
      objLOG.write vbnewline & now & vbtab & vbtab & " - DOWNLOADING 'CLOUD_FILTERS.TXT' BACKUP FILTER DEFINITION"
      ''REMOVE PREVIOUS 'FILTERS.TXT' FILE
      erase arrTMP
      strTMP = vbnullstring
      if (objFSO.fileexists("C:\temp\cloud_filters.txt")) then
        objFSO.deletefile "C:\temp\cloud_filters.txt", true
      end if
      call FILEDL("https://github.com/CW-Khristos/scripts/raw/master/MSP%20Backups/cloud_filters.txt", "cloud_filters.txt")
      set objTMP = objFSO.opentextfile("C:\temp\cloud_filters.txt", 1)
      while (not objTMP.atendofstream)
        strTMP = strTMP & objTMP.readline
      wend
      objTMP.close
      set objTMP = nothing
      arrTMP = split(strTMP, "|")
      for intTMP = 0 to ubound(arrTMP)
        if (arrTMP(intTMP) <> vbnullstring) then
          objOUT.write vbnewline & now & vbtab & vbtab & "EXECUTING : C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -exclude " & chr(34) & arrTMP(intTMP) & chr(34)
          objLOG.write vbnewline & now & vbtab & vbtab & "EXECUTING : C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -exclude " & chr(34) & arrTMP(intTMP) & chr(34)
          call HOOK("C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -exclude " & chr(34) & arrTMP(intTMP) & chr(34))
        end if
      next
      ''CUSTOM 'FILTER' PASSED
      if (strFILTER <> vbnullstring) then
        for intTMP = 0 to ubound(arrFILTER)
          if (arrFILTER(intTMP) <> vbnullstring) then
            objOUT.write vbnewline & now & vbtab & vbtab & "EXECUTING : C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -exclude " & chr(34) & arrFILTER(intTMP) & chr(34)
            objLOG.write vbnewline & now & vbtab & vbtab & "EXECUTING : C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -exclude " & chr(34) & arrFILTER(intTMP) & chr(34)
            call HOOK("C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -exclude " & chr(34) & arrFILTER(intTMP) & chr(34))
          end if
        next
      end if
      ''DOWNLOAD 'INCLUDES.TXT' BACKUP INCLUDES DEFINITION FILE , 'ERRRET'=2 , REF #2
      objOUT.write vbnewline & now & vbtab & vbtab & " - DOWNLOADING 'CLOUD_INCLUDES.TXT' BACKUP INCLUDES DEFINITION"
      objLOG.write vbnewline & now & vbtab & vbtab & " - DOWNLOADING 'CLOUD_INCLUDES.TXT' BACKUP INCLUDES DEFINITION"
      ''REMOVE PREVIOUS 'INCLUDES.TXT' FILE
      erase arrTMP
      strTMP = vbnullstring
      if (objFSO.fileexists("C:\temp\cloud_includes.txt")) then
        objFSO.deletefile "C:\temp\cloud_includes.txt", true
      end if
      call FILEDL("https://github.com/CW-Khristos/scripts/raw/master/MSP%20Backups/cloud_includes.txt", "cloud_includes.txt")
      set objTMP = objFSO.opentextfile("C:\temp\cloud_includes.txt", 1)
      while (not objTMP.atendofstream)
        strTMP = strTMP & objTMP.readline
      wend
      objTMP.close
      set objTMP = nothing
      arrTMP = split(strTMP, "|")
      for intTMP = 0 to ubound(arrTMP)
        if (arrTMP(intTMP) <> vbnullstring) then
          objOUT.write vbnewline & now & vbtab & vbtab & "EXECUTING : C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -include " & chr(34) & arrTMP(intTMP) & chr(34)
          objLOG.write vbnewline & now & vbtab & vbtab & "EXECUTING : C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -include " & chr(34) & arrTMP(intTMP) & chr(34)
          call HOOK("C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -include " & chr(34) & arrTMP(intTMP) & chr(34))
        end if
      next
      ''CUSTOM 'INCLUDE' PASSED
      if (strINCL <> vbnullstring) then
        for intTMP = 0 to ubound(arrINCL)
          if (arrINCL(intTMP) <> vbnullstring) then
            objOUT.write vbnewline & now & vbtab & vbtab & "EXECUTING : C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -include " & chr(34) & arrINCL(intTMP) & chr(34)
            objLOG.write vbnewline & now & vbtab & vbtab & "EXECUTING : C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -include " & chr(34) & arrINCL(intTMP) & chr(34)
            call HOOK("C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -include " & chr(34) & arrINCL(intTMP) & chr(34))
          end if
        next
      end if
      ''PERFORM FINAL EXCLUDES
      objOUT.write vbnewline & now & vbtab & vbtab & " - PERFORMING FINAL EXCLUDES"
      objLOG.write vbnewline & now & vbtab & vbtab & " - PERFORMING FINAL EXCLUDES"
      ''ENUMERATE 'C:\USERS' SUB-FOLDERS
      objOUT.write vbnewline & now & vbtab & vbtab & " - CHECKING USER FOLDERS"
      objLOG.write vbnewline & now & vbtab & vbtab & " - CHECKING USER FOLDERS"
      set objFOL = objFSO.getfolder("C:\Users")
      set colFOL = objFOL.subfolders
      intFOL = 0
      for each subFOL in colFOL
        redim preserve arrFOL(intFOL + 1)
        arrFOL(intFOL) = subFOL.path
        intFOL = intFOL + 1
      next
      set colFOL = nothing
      set objFOL = nothing
      intFOL = 0
      ''CHECK EACH 'C:\USERS\<USERNAME>' FOLDER
      for intFOL = 0 to ubound(arrFOL)
        intCOL = 0
        blnFND = false
        strFOL = arrFOL(intFOL)
        if (strFOL <> vbnullstring) then
        
            ''ENUMERATE THROUGH AND MAKE SURE THIS ISN'T ONE OF THE 'UNNEEDED / TO EXCLUDE' USER ACCOUNTS
            for intCOL = 0 to ubound(arrUSR)
              blnFND = false
              if (arrUSR(intCOL) <> vbnullstring) then
                '' 'UNNEEDED / TO EXCLUDE' USER ACCOUNT 'ARRUSR' FOUND IN FOLDER PATH
                if (instr(1, lcase(strFOL), lcase(arrUSR(intCOL)))) then
                  objOUT.write vbnewline & now & vbtab & vbtab & vbtab & "UNNEEDED / TO EXCLUDE : " & arrUSR(intCOL)
                  objLOG.write vbnewline & now & vbtab & vbtab & vbtab & "UNNEEDED / TO EXCLUDE : " & arrUSR(intCOL)
                  ''MARK 'UNNEEDED / TO EXCLUDE'
                  blnFND = true
                  ''PROCEED WITH EXCLUDING ENTIRE USER DIRECTORY
                  
                  exit for
                end if
              end if
              ''AN 'UNNEEDED / TO EXCLUDE' USER ACCOUNT WAS PASSED TO 'STRUSR'
              if (wscript.arguments.count > 0) then
                '' PASSED 'UNNEEDED / TO EXCLUDE' USER ACCOUNT 'ARRUSR'
                if (instr(1, lcase(strFOL), lcase(objARG.item(0)))) then
                  objOUT.write vbnewline & now & vbtab & vbtab & vbtab & "UNNEEDED / TO EXCLUDE : " & objARG.item(0)
                  objLOG.write vbnewline & now & vbtab & vbtab & vbtab & "UNNEEDED / TO EXCLUDE : " & objARG.item(0)
                  ''MARK 'UNNEEDED / TO EXCLUDE'
                  blnFND = true
                  ''PROCEED WITH EXCLUDING ENTIRE USER DIRECTORY
                  
                  exit for
                end if          
              end if
            next
            ''NO MATCH TO 'UNNEEDED / TO EXCLUDE' USER ACCOUNTS
            if (not (blnFND)) then
              ''CHECK FOR USER FOLDER
              if (objFSO.folderexists(strFOL)) then
                objOUT.write vbnewline & now & vbtab & vbtab & vbtab & "ENUMERATING : " & strFOL
                objLOG.write vbnewline & now & vbtab & vbtab & vbtab & "ENUMERATING : " & strFOL
                ''EXCLUDE FROM BACKUPS
                blnFND = false
                
                ''ENUMERATE 'C:\USERS\<USERNAME>' SUB-FOLDERS
                set objUFOL = objFSO.getfolder(strFOL)
                set colUFOL = objUFOL.subfolders
                intUFOL = 0
                for each subUFOL in colUFOL
                  redim preserve arrUFOL(intUFOL + 1)
                  arrUFOL(intUFOL) = subUFOL.path
                  intUFOL = intUFOL + 1
                next
                set colUFOL = nothing
                set objUFOL = nothing
                intUFOL = 0
                ''!---- THE BELOW WILL NEED TO BE USED AS A CALLABLE FUNCTION WITH RETURN VALUE                     ----!''
                ''!---- THIS WILL ALLOW RECURSION THROUGH EACH SUB-FOLDER OF 'C:\USERS\<USERNAME>'                  ----!''
                ''!---- ONCE DONE; FURTHER SUB-FOLDER DIRECTORIES AND FILES WILL BE ABLE TO BE RECURSIVELY CHECKED  ----!''
                ''CHECK EACH 'C:\USERS\<USERNAME>' SUB-FOLDER
                for intUFOL = 0 to ubound(arrUFOL)
                  intUCOL = 0
                  blnFND = false
                  strUFOL = arrUFOL(intUFOL)
                  if (strUFOL <> vbnullstring) then            
                    ''ENUMERATE THROUGH AND MAKE SURE THIS ISN'T ONE OF THE 'PROTECTED' EXT / FILES / DIRECTORIES
                    for intPCOL = 0 to ubound(arrPROT)
                      blnFND = false
                      if (arrPROT(intPCOL) <> vbnullstring) then
                        '' 'PRTOTECTED' USER ACCOUNT 'arrPROT' FOUND IN FOLDER PATH
                        if (instr(1, lcase(strUFOL), lcase(arrPROT(intPCOL)))) then
                          objOUT.write vbnewline & now & vbtab & vbtab & vbtab & "PROTECTED : " & arrPROT(intPCOL)
                          objLOG.write vbnewline & now & vbtab & vbtab & vbtab & "PROTECTED : " & arrPROT(intPCOL)
                          ''MARK 'PROTECTED'
                          blnFND = true
                          exit for
                        end if
                      end if
                      ''A 'UNNEEDED / TO EXCLUDE' USER ACCOUNT WAS PASSED TO 'STRUSR'
                      if (wscript.arguments.count > 0) then
                        '' PASSED 'PRTOTECTED' USER ACCOUNT 'ARRUSR'
                        if (instr(1, lcase(strFOL), lcase(objARG.item(0)))) then
                          objOUT.write vbnewline & now & vbtab & vbtab & vbtab & "UNNEEDED / TO EXCLUDE : " & objARG.item(0)
                          objLOG.write vbnewline & now & vbtab & vbtab & vbtab & "UNNEEDED / TO EXCLUDE : " & objARG.item(0)
                          ''MARK 'UNNEEDED / TO EXCLUDE'
                          blnFND = true
                          exit for
                        end if          
                      end if
                    next
                    ''NO MATCH TO 'PROTECTED' EXT / FILES / DIRECTORIES
                    if (not (blnFND)) then

                    end if
                  next
                end if
              next
        next
  end select
end if
''END SCRIPT
call CLEANUP()
''END SCRIPT
''------------

''SUB-ROUTINES
sub CHKAU()																									''CHECK FOR SCRIPT UPDATE , 'ERRRET'=10 , MSP_FILTER.VBS , REF #2
  ''REMOVE WINDOWS AGENT CACHED VERSION OF SCRIPT
  if (objFSO.fileexists("C:\Program Files (x86)\N-Able Technologies\Windows Agent\Temp\Script\" & wscript.scriptname)) then
    objFSO.deletefile "C:\Program Files (x86)\N-Able Technologies\Windows Agent\Temp\Script\" & wscript.scriptname, true
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
        objOUT.write vbnewline & now & vbtab & " - MSP_FILTER :  " & strVER & " : GitHub : " & objSCR.text & vbnewline
        objLOG.write vbnewline & now & vbtab & " - MSP_FILTER :  " & strVER & " : GitHub : " & objSCR.text & vbnewline
				if (cint(objSCR.text) > cint(strVER)) then
					objOUT.write vbnewline & now & vbtab & " - UPDATING " & objSCR.nodename & " : " & objSCR.text & vbnewline
					objLOG.write vbnewline & now & vbtab & " - UPDATING " & objSCR.nodename & " : " & objSCR.text & vbnewline
					''DOWNLOAD LATEST VERSION OF SCRIPT
					call FILEDL("https://github.com/CW-Khristos/scripts/raw/master/MSP%20Backups/MSP_Filter.vbs", wscript.scriptname)
					''RUN LATEST VERSION
					if (wscript.arguments.count > 0) then             ''ARGUMENTS WERE PASSED
						for x = 0 to (wscript.arguments.count - 1)
							strTMP = strTMP & " " & chr(34) & objARG.item(x) & chr(34)
						next
            objOUT.write vbnewline & now & vbtab & " - RE-EXECUTING " & objSCR.nodename & " : " & objSCR.text & vbnewline
            objLOG.write vbnewline & now & vbtab & " - RE-EXECUTING  " & objSCR.nodename & " : " & objSCR.text & vbnewline
						objWSH.run "cscript.exe //nologo " & chr(34) & "c:\temp\" & wscript.scriptname & chr(34) & strTMP, 0, false
					elseif (wscript.arguments.count = 0) then         ''NO ARGUMENTS WERE PASSED
            objOUT.write vbnewline & now & vbtab & " - RE-EXECUTING  " & objSCR.nodename & " : " & objSCR.text & vbnewline
            objLOG.write vbnewline & now & vbtab & " - RE-EXECUTING  " & objSCR.nodename & " : " & objSCR.text & vbnewline
						objWSH.run "cscript.exe //nologo " & chr(34) & "c:\temp\" & wscript.scriptname & chr(34), 0, false
					end if
          if (err.number <> 0) then
            call LOGERR(10)
          end if
					''END SCRIPT
					call CLEANUP()
				end if
			end if
		next
	end if
	set colVER = nothing
	set objXML = nothing
  if (err.number <> 0) then                                 ''ERROR RETURNED DURING UPDATE CHECK , 'ERRRET'=10
    call LOGERR(10)
  end if
end sub

sub FILEDL(strURL, strFILE)                                 ''CALL HOOK TO DOWNLOAD FILE FROM URL , 'ERRRET'=11
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
  end if
	set objHTTP = nothing
  if (err.number <> 0) then                                 ''ERROR RETURNED , 'ERRRET'=11
    call LOGERR(11)
  end if
end sub

sub HOOK(strCMD)                                            ''CALL HOOK TO MONITOR OUTPUT OF CALLED COMMAND , 'ERRRET'=12
  on error resume next
  set objHOOK = objWSH.exec(strCMD)
  while (not objHOOK.stdout.atendofstream)
		strIN = objHOOK.stdout.readline
		if (strIN <> vbnullstring) then
			objOUT.write vbnewline & now & vbtab & vbtab & strIN 
			objLOG.write vbnewline & now & vbtab & vbtab & strIN 
		end if
	wend
	wscript.sleep 10
  strIN = objHOOK.stdout.readall
  if (strIN <> vbnullstring) then
    objOUT.write vbnewline & now & vbtab & vbtab & strIN 
    objLOG.write vbnewline & now & vbtab & vbtab & strIN 
  end if
  set objHOOK = nothing
  if (err.number <> 0) then                                 ''ERROR RETURNED , 'ERRRET'=12
    call LOGERR(12)
  end if
end sub

sub LOGERR(intSTG)                                          ''CALL HOOK TO MONITOR OUTPUT OF CALLED COMMAND
  if (err.number <> 0) then
    objOUT.write vbnewline & now & vbtab & vbtab & vbtab & err.number & vbtab & err.description & vbnewline
    objLOG.write vbnewline & now & vbtab & vbtab & vbtab & err.number & vbtab & err.description & vbnewline
		err.clear
  end if
  select case intSTG
    case 1                                                  '' 'ERRRET'=1 - NOT ENOUGH ARGUMENTS
  end select
  errRET = intSTG
end sub

sub CLEANUP()                                               ''SCRIPT CLEANUP
  if (errRET = 0) then         															''MSP_FILTER COMPLETED SUCCESSFULLY
    objOUT.write vbnewline & "MSP_FILTER SUCCESSFUL : " & errRET & " : " & now
    err.clear
  elseif (errRET <> 0) then    															''MSP_FILTER FAILED
    objOUT.write vbnewline & "MSP_FILTER FAILURE : " & errRET & " : " & now
    ''RAISE CUSTOMIZED ERROR CODE, ERROR CODE WILL BE DEFINE RESTOP NUMBER INDICATING WHICH SECTION FAILED
    call err.raise(vbObjectError + errRET, "MSP_FILTER", "FAILURE")
  end if
  objOUT.write vbnewline & vbnewline & now & " - MSP_FILTER COMPLETE" & vbnewline
  objLOG.write vbnewline & vbnewline & now & " - MSP_FILTER COMPLETE" & vbnewline
  objLOG.close
  ''EMPTY OBJECTS
  set objLOG = nothing
  set objFSO = nothing
  set objWSH = nothing
  set objARG = nothing
  set objOUT = nothing
  set objIN = nothing
  ''END SCRIPT, RETURN ERROR NUMBER
  wscript.quit err.number
end sub
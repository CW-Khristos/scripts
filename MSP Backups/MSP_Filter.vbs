''MSP_FILTER.VBS
''DESIGNED TO AUTOMATE PASSING OF BACKUP FILTERS TO MSP BACKUP SOFTWARE VIA CLIENTTOOL
''DOWNLOADS 'FILTERS.TXT' FROM GITHUB; THIS FILE CONTAINS EACH BACKUP FILTER IN A 'LINE BY LINE' FORMAT
''DESIGNED TO AUTOMATE PASSING OF BACKUP INCLUSIONS TO MSP BACKUP SOFTWARE VIA CLIENTTOOL
''DOWNLOADS 'INCLUDES.TXT' FROM GITHUB; THIS FILE CONTAINS EACH BACKUP FILTER IN A 'LINE BY LINE' FORMAT
''ACCEPTS 4 PARAMETER , REQUIRES 1 PARAMETER
''REQUIRED PARAMETER 'STROPT' ; STRING VALUE TO INDICATE 'LOCAL' OR 'CLOUD' FILTER OPERATION
''OPTIONAL PARAMETER 'STRFILTER' ; STRING VALUE TO HOLD PASSED 'FILTERS' ; SEPARATE MULTIPLE 'FILTERS' VIA '|'
''OPTIONAL PARAMETER 'STRINCL' ; STRING VALUE TO HOLD PASSED 'INCLUSIONS' ; SEPARATE MULTIPLE 'INCLUSIONS' VIA '|'
''OPTIONAL PARAMETER 'STRUSR' ; STRING VALUE TO HOLD PASSED 'USER ACCOUNT' TO EXCLUDE
''WRITTEN BY : CJ BLEDSOE / CJ<@>THECOMPUTERWARRIORS.COM
on error resume next
''SCRIPT VARIABLES
dim errRET, strVER
''VARIABLES ACCEPTING PARAMETERS
dim strOPT, strUSR
dim strINCL, arrINCL
dim strFILTER, arrFILTER
dim strIN, strOUT, strRCMD
''USER AND USER FOLDER ARRAYS
dim objFOL, arrFOL()
''USER FOLDER AND SUB-FOLDER ARRAYS
dim objUFOL, arrUFOL()
''PRE-DEFINED ARRAYS
dim arrEXCL(), arrPUSR(), arrPFOL(), arrAPP()
''SCRIPT OBJECTS
dim objIN, objOUT, objARG, objWSH, objFSO
dim objLOG, objEXEC, objHOOK, objHTTP, objXML
''VERSION FOR SCRIPT UPDATE , MSP_FILTER.VBS , REF #2 , REF #68 , REF #69
strVER = 7
strREPO = "scripts"
strBRCH = "master"
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
''CHECK 'PERSISTENT' FOLDERS , REF #2 , REF #73
if (not (objFSO.folderexists("c:\temp"))) then
  objFSO.createfolder("c:\temp")
end if
if (not (objFSO.folderexists("C:\IT\"))) then
  objFSO.createfolder("C:\IT\")
end if
if (not (objFSO.folderexists("C:\IT\Scripts\"))) then
  objFSO.createfolder("C:\IT\Scripts\")
end if
''PREPARE LOGFILE
if (objFSO.fileexists("C:\temp\msp_filter")) then             ''LOGFILE EXISTS
  objFSO.deletefile "C:\temp\msp_filter", true
  set objLOG = objFSO.createtextfile("C:\temp\msp_filter")
  objLOG.close
  set objLOG = objFSO.opentextfile("C:\temp\msp_filter", 8)
else                                                          ''LOGFILE NEEDS TO BE CREATED
  set objLOG = objFSO.createtextfile("C:\temp\msp_filter")
  objLOG.close
  set objLOG = objFSO.opentextfile("C:\temp\msp_filter", 8)
end if
''READ PASSED COMMANDLINE ARGUMENTS
if (wscript.arguments.count > 0) then                         ''ARGUMENTS WERE PASSED
  for x = 0 to (wscript.arguments.count - 1)
    objOUT.write vbnewline & now & vbtab & " - ARGUMENT " & (x + 1) & " (ITEM " & x & ") " & " PASSED : " & ucase(objARG.item(x))
    objLOG.write vbnewline & now & vbtab & " - ARGUMENT " & (x + 1) & " (ITEM " & x & ") " & " PASSED : " & ucase(objARG.item(x))
  next 
  if (wscript.arguments.count > 0) then                       ''SET VARIABLES ACCEPTING ARGUMENTS
    strOPT = objARG.item(0)                                   ''SET REQUIRE PARAMETER 'STROPT' , 'LOCAL' OR 'CLOUD' OPERATION
    if (wscript.arguments.count > 1) then                     ''SET OPTIONAL PARAMETER 'STRFILTER' , BACKUP FILTERS STRING
      strFILTER = objARG.item(1)
      ''FILL 'ARRFILTER' BACKUP FILTER ARRAY
      'objOUT.write vbnewline & vbtab & strFILTER
      arrFILTER = split(strFILTER, "|")
      for intTMP = 0 to ubound(arrFILTER)
        objOUT.write vbnewline & vbtab & ubound(arrFILTER) & vbtab & arrFILTER(intTMP)
      next
    end if
    if (wscript.arguments.count > 2) then                     ''SET OPTIONAL PARAMETER 'STRINCL' , BACKUP INCLUDES STRING
      strINCL = objARG.item(2)
      ''FILL 'ARRINCL' BACKUP INCLUDES ARRAY
      'objOUT.write vbnewline & vbtab & strINCL
      arrINCL = split(strINCL, "|")
      for intTMP = 0 to ubound(arrINCL)
        objOUT.write vbnewline & vbtab & ubound(arrINCL) & vbtab & arrINCL(intTMP)
      next
    end if
    if (wscript.arguments.count > 3) then                     ''SET OPTIONAL PARAMETER 'STRUSR' , USER ACCOUNT TO EXCLUDE
      strUSR = objARG.item(3)
    end if
  end if
elseif (wscript.arguments.count = 0) then                     ''NOT ENOUGH ARGUMENTS PASSED , END SCRIPT , 'ERRRET'=1
  call LOGERR(1)
end if
''UNNEEDED / TO EXCLUDE USER ACCOUNTS
redim arrEXCL(1)
arrEXCL(0) = "rmmtech"
''PROTECTED USER ACCOUNTS
redim arrPUSR(4)
arrPUSR(0) = "MSSQL"
arrPUSR(1) = "Public"
arrPUSR(2) = "Default"
arrPUSR(3) = "Default.migrated"
''PROTECTED EXT / FILES / DIRECTORIES
redim arrPFOL(2)
arrPFOL(0) = ".PST"
arrPFOL(1) = "Outlook\Roamcache"
''APPDATA FILES / FOLDERS
redim arrAPP(40)
arrAPP(0) = "\AppData\Local\CrashDumps"
arrAPP(1) = "\AppData\Local\D3DSCache"
arrAPP(2) = "\AppData\Local\Google\Chrome\User Data\Crashpad"
arrAPP(3) = "\AppData\Local\Google\Chrome\User Data\Default\Application Cache"
arrAPP(4) = "\AppData\Local\Google\Chrome\User Data\Default\Cache"
arrAPP(5) = "\AppData\Local\Google\Chrome\User Data\Default\Code Cache"
arrAPP(6) = "\AppData\Local\Google\Chrome\User Data\Default\GPUCache"
arrAPP(7) = "\AppData\Local\Google\Chrome\User Data\FontLookupTableCache"
arrAPP(8) = "\AppData\Local\Google\Chrome\User Data\ShaderCache"
arrAPP(9) = "\AppData\Local\Google\Chrome\User Data\PnaclTranslationCache"
arrAPP(10) = "\AppData\Local\Google\Chrome\User Data\SwReporter"
arrAPP(11) = "\AppData\Local\Google\CrashReports"
arrAPP(12) = "\AppData\Local\Google\Software Reporter Tool"
arrAPP(13) = "\AppData\Local\GWX"
arrAPP(14) = "\AppData\Local\Microsoft\Feeds Cache"
arrAPP(15) = "\AppData\Local\Microsoft\FontCache"
arrAPP(16) = "\AppData\Local\Microsoft\SquirrelTemp"
arrAPP(17) = "\AppData\Local\Microsoft\Terminal Server Client\Cache"
arrAPP(18) = "\AppData\Local\Microsoft\Windows\ActionCenterCache"
arrAPP(19) = "\AppData\Local\Microsoft\Windows\AppCache"
arrAPP(20) = "\AppData\Local\Microsoft\Windows\Caches"
arrAPP(21) = "\AppData\Local\Microsoft\Windows\Explorer\IconCacheToDelete"
arrAPP(22) = "\AppData\Local\Microsoft\Windows\IECompatCache"
arrAPP(23) = "\AppData\Local\Microsoft\Windows\IECompatUaCache"
arrAPP(24) = "\AppData\Local\Microsoft\Windows\INetCache"
arrAPP(25) = "\AppData\Local\Microsoft\Windows\PPBCompatCache"
arrAPP(26) = "\AppData\Local\Microsoft\Windows\PPBCompatUaCache"
arrAPP(27) = "\AppData\Local\Microsoft\Windows\PRICache"
arrAPP(28) = "\AppData\Local\Microsoft\Windows\SchCache"
arrAPP(29) = "\AppData\Local\Microsoft\Windows\WER"
arrAPP(30) = "\AppData\Local\Microsoft\Windows\WebCache"
arrAPP(31) = "\AppData\Local\Mozilla"
arrAPP(32) = "\AppData\Local\SquirrelTemp"
arrAPP(33) = "\AppData\Local\Temp"
arrAPP(34) = "\AppData\Local\IconCache.db"
arrAPP(35) = "\AppData\Local\Microsoft\Outlook\*.ost"
arrAPP(36) = "\AppData\Local\Microsoft\Outlook\*.tmp"
arrAPP(37) = "\AppData\Local\Microsoft\Windows\Explorer\iconcache*.db"
arrAPP(38) = "\AppData\Local\Microsoft\Windows\Explorer\thumbcache*.db"
arrAPP(39) = "\AppData\Local\Microsoft\MicrosoftEdge\SharedCacheContainers"

''------------
''BEGIN SCRIPT
strTMP = vbnullstring
if (errRET = 0) then                                          ''ARGUMENTS PASSED, CONTINUE SCRIPT
	objOUT.write vbnewline & vbnewline & now & vbtab & " - EXECUTING MSP_FILTER"
	objLOG.write vbnewline & vbnewline & now & vbtab & " - EXECUTING MSP_FILTER"
	''AUTOMATIC UPDATE, MSP_FILTER.VBS, REF #2 , REF #69 , REF #68
  ''DOWNLOAD CHKAU.VBS SCRIPT, REF #2 , REF #69 , REF #68
  call FILEDL("https://raw.githubusercontent.com/CW-Khristos/scripts/master/chkAU.vbs", "C:\IT\Scripts", "chkAU.vbs")
  ''EXECUTE CHKAU.VBS SCRIPT, REF #69
  objOUT.write vbnewline & now & vbtab & vbtab & " - CHECKING FOR UPDATE : MSP_FILTER : " & strVER
  objLOG.write vbnewline & now & vbtab & vbtab & " - CHECKING FOR UPDATE : MSP_FILTER : " & strVER
  intRET = objWSH.run ("cmd.exe /C " & chr(34) & "cscript.exe " & chr(34) & "C:\IT\Scripts\chkAU.vbs" & chr(34) & " " & _
    chr(34) & strREPO & chr(34) & " " & chr(34) & strBRCH & chr(34) & " " & chr(34) & strDIR & chr(34) & " " & _
    chr(34) & wscript.scriptname & chr(34) & " " & chr(34) & strVER & chr(34) & " " & _
    chr(34) & strOPT & "|" & strFILTER & "|" & strINCL & chr(34) & chr(34), 0, true)
  ''CHKAU RETURNED - NO UPDATE FOUND , REF #2 , REF #69 , REF #68
  objOUT.write vbnewline & "errRET='" & intRET & "'"
  objLOG.write vbnewline & "errRET='" & intRET & "'"
  intRET = (intRET - vbObjectError)
  objOUT.write vbnewline & "errRET='" & intRET & "'"
  objLOG.write vbnewline & "errRET='" & intRET & "'"
  if ((intRET = 4) or (intRET = 10) or (intRET = 11) or (intRET = 1) or (intRET = 2147221505) or (intRET = 2147221517)) then
    Select Case lcase(strOPT)
      ''PERFORM 'LOCAL' FILTER CONFIGURATIONS
      Case "local"
        ''DISABLED TO PREVENT OVER-WRITE OF TECHNICIAN SELECTIONS AT A LATER TIME
        ''RESET CURRENT BACKUP INCLUDES , REF #2
        'objOUT.write vbnewline & now & vbtab & vbtab & " - RESETTING CURRENT MSP BACKUP INCLUDES"
        'objLOG.write vbnewline & now & vbtab & vbtab & " - RESETTING CURRENT MSP BACKUP INCLUDES"
        'call HOOK("C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -include C:\")
        'wscript.sleep 5000
        ''DOWNLOAD 'FILTERS.TXT' BACKUP FILTERS DEFINITION FILE , 'ERRRET'=2 , REF #2
        objOUT.write vbnewline & now & vbtab & vbtab & " - DOWNLOADING 'FILTERS.TXT' BACKUP FILTER DEFINITION"
        objLOG.write vbnewline & now & vbtab & vbtab & " - DOWNLOADING 'FILTERS.TXT' BACKUP FILTER DEFINITION"
        ''REMOVE PREVIOUS 'FILTERS.TXT' FILE
        strTMP = vbnullstring
        if (objFSO.fileexists("C:\IT\Scripts\filters.txt")) then
          objFSO.deletefile "C:\IT\Scripts\filters.txt", true
        end if
        call FILEDL("https://raw.githubusercontent.com/CW-Khristos/scripts/master/MSP%20Backups/filters.txt", "C:\IT\Scripts", "filters.txt")
        set objTMP = objFSO.opentextfile("C:\IT\Scripts\filters.txt", 1)
        while (not objTMP.atendofstream)
          strTMP = strTMP & objTMP.readline
        wend
        objTMP.close
        set objTMP = nothing
        arrTMP = split(strTMP, "|")
        for intTMP = 0 to ubound(arrTMP)
          if (arrTMP(intTMP) <> vbnullstring) then
            objOUT.write vbnewline & now & vbtab & vbtab & _
              "EXECUTING : C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -exclude " & chr(34) & arrTMP(intTMP) & chr(34)
            objLOG.write vbnewline & now & vbtab & vbtab & _
              "EXECUTING : C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -exclude " & chr(34) & arrTMP(intTMP) & chr(34)
            call HOOK("C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -exclude " & chr(34) & arrTMP(intTMP) & chr(34))
            wscript.sleep 200
          end if
        next
        ''CUSTOM 'FILTER' PASSED
        if (strFILTER <> vbnullstring) then
          for intTMP = 0 to ubound(arrFILTER)
            if (arrFILTER(intTMP) <> vbnullstring) then
              objOUT.write vbnewline & now & vbtab & vbtab & _
                "EXECUTING : C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -exclude " & chr(34) & arrFILTER(intTMP) & chr(34)
              objLOG.write vbnewline & now & vbtab & vbtab & _
                "EXECUTING : C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -exclude " & chr(34) & arrFILTER(intTMP) & chr(34)
              call HOOK("C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -exclude " & chr(34) & arrFILTER(intTMP) & chr(34))
              wscript.sleep 200
            end if
          next
        end if
        ''DOWNLOAD 'INCLUDES.TXT' BACKUP INCLUDES DEFINITION FILE , 'ERRRET'=2 , REF #2
        objOUT.write vbnewline & now & vbtab & vbtab & " - DOWNLOADING 'INCLUDES.TXT' BACKUP INCLUDES DEFINITION"
        objLOG.write vbnewline & now & vbtab & vbtab & " - DOWNLOADING 'INCLUDES.TXT' BACKUP INCLUDES DEFINITION"
        ''REMOVE PREVIOUS 'INCLUDES.TXT' FILE
        strTMP = vbnullstring
        if (objFSO.fileexists("C:\IT\Scripts\includes.txt")) then
          objFSO.deletefile "C:\IT\Scripts\includes.txt", true
        end if
        call FILEDL("https://raw.githubusercontent.com/CW-Khristos/scripts/master/MSP%20Backups/includes.txt", "C:\IT\Scripts", "includes.txt")
        set objTMP = objFSO.opentextfile("C:\IT\Scripts\includes.txt", 1)
        while (not objTMP.atendofstream)
          strTMP = strTMP & objTMP.readline
        wend
        objTMP.close
        set objTMP = nothing
        arrTMP = split(strTMP, "|")
        for intTMP = 0 to ubound(arrTMP)
          if (arrTMP(intTMP) <> vbnullstring) then
            objOUT.write vbnewline & now & vbtab & vbtab & _
              "EXECUTING : C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -include " & chr(34) & arrTMP(intTMP) & chr(34)
            objLOG.write vbnewline & now & vbtab & vbtab & _
              "EXECUTING : C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -include " & chr(34) & arrTMP(intTMP) & chr(34)
            call HOOK("C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -include " & chr(34) & arrTMP(intTMP) & chr(34))
            wscript.sleep 200
          end if
        next
        ''CUSTOM 'INCLUDE' PASSED
        if (strINCL <> vbnullstring) then
          for intTMP = 0 to ubound(arrINCL)
            if (arrINCL(intTMP) <> vbnullstring) then
              objOUT.write vbnewline & now & vbtab & vbtab & _
                "EXECUTING : C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -include " & chr(34) & arrINCL(intTMP) & chr(34)
              objLOG.write vbnewline & now & vbtab & vbtab & _
                "EXECUTING : C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -include " & chr(34) & arrINCL(intTMP) & chr(34)
              call HOOK("C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -include " & chr(34) & arrINCL(intTMP) & chr(34))
              wscript.sleep 200
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
        if (objFSO.fileexists("C:\IT\Scripts\cloud_filters.txt")) then
          objFSO.deletefile "C:\IT\Scripts\cloud_filters.txt", true
        end if
        call FILEDL("https://raw.githubusercontent.com/CW-Khristos/scripts/master/MSP%20Backups/cloud_filters.txt", "C:\IT\Scripts", "cloud_filters.txt")
        set objTMP = objFSO.opentextfile("C:\IT\Scripts\cloud_filters.txt", 1)
        while (not objTMP.atendofstream)
          strTMP = strTMP & objTMP.readline
        wend
        objTMP.close
        set objTMP = nothing
        arrTMP = split(strTMP, "|")
        for intTMP = 0 to ubound(arrTMP)
          if (arrTMP(intTMP) <> vbnullstring) then
            objOUT.write vbnewline & now & vbtab & vbtab & _
              "EXECUTING : C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -exclude " & chr(34) & arrTMP(intTMP) & chr(34)
            objLOG.write vbnewline & now & vbtab & vbtab & _
              "EXECUTING : C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -exclude " & chr(34) & arrTMP(intTMP) & chr(34)
            call HOOK("C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -exclude " & chr(34) & arrTMP(intTMP) & chr(34))
            wscript.sleep 200
          end if
        next
        ''CUSTOM 'FILTER' PASSED
        if (strFILTER <> vbnullstring) then
          for intTMP = 0 to ubound(arrFILTER)
            if (arrFILTER(intTMP) <> vbnullstring) then
              objOUT.write vbnewline & now & vbtab & vbtab & _
                "EXECUTING : C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -exclude " & chr(34) & arrFILTER(intTMP) & chr(34)
              objLOG.write vbnewline & now & vbtab & vbtab & _
                "EXECUTING : C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -exclude " & chr(34) & arrFILTER(intTMP) & chr(34)
              call HOOK("C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -exclude " & chr(34) & arrFILTER(intTMP) & chr(34))
              wscript.sleep 200
            end if
          next
        end if
        ''DOWNLOAD 'INCLUDES.TXT' BACKUP INCLUDES DEFINITION FILE , 'ERRRET'=2 , REF #2
        objOUT.write vbnewline & now & vbtab & vbtab & " - DOWNLOADING 'CLOUD_INCLUDES.TXT' BACKUP INCLUDES DEFINITION"
        objLOG.write vbnewline & now & vbtab & vbtab & " - DOWNLOADING 'CLOUD_INCLUDES.TXT' BACKUP INCLUDES DEFINITION"
        ''REMOVE PREVIOUS 'INCLUDES.TXT' FILE
        erase arrTMP
        strTMP = vbnullstring
        if (objFSO.fileexists("C:\IT\Scripts\cloud_includes.txt")) then
          objFSO.deletefile "C:\IT\Scripts\cloud_includes.txt", true
        end if
        call FILEDL("https://raw.githubusercontent.com/CW-Khristos/scripts/master/MSP%20Backups/cloud_includes.txt", "C:\IT\Scripts", "cloud_includes.txt")
        set objTMP = objFSO.opentextfile("C:\IT\Scripts\cloud_includes.txt", 1)
        while (not objTMP.atendofstream)
          strTMP = strTMP & objTMP.readline
        wend
        objTMP.close
        set objTMP = nothing
        arrTMP = split(strTMP, "|")
        for intTMP = 0 to ubound(arrTMP)
          if (arrTMP(intTMP) <> vbnullstring) then
            objOUT.write vbnewline & now & vbtab & vbtab & _
              "EXECUTING : C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -include " & chr(34) & arrTMP(intTMP) & chr(34)
            objLOG.write vbnewline & now & vbtab & vbtab & _
              "EXECUTING : C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -include " & chr(34) & arrTMP(intTMP) & chr(34)
            call HOOK("C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -include " & chr(34) & arrTMP(intTMP) & chr(34))
            wscript.sleep 200
          end if
        next
        ''CUSTOM 'INCLUDE' PASSED
        if (strINCL <> vbnullstring) then
          for intTMP = 0 to ubound(arrINCL)
            if (arrINCL(intTMP) <> vbnullstring) then
              objOUT.write vbnewline & now & vbtab & vbtab & _
                "EXECUTING : C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -include " & chr(34) & arrINCL(intTMP) & chr(34)
              objLOG.write vbnewline & now & vbtab & vbtab & _
                "EXECUTING : C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -include " & chr(34) & arrINCL(intTMP) & chr(34)
              call HOOK("C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -include " & chr(34) & arrINCL(intTMP) & chr(34))
              wscript.sleep 200
            end if
          next
        end if
    end select
    ''PERFORM FINAL EXCLUDES
    objOUT.write vbnewline & now & vbtab & vbtab & " - PERFORMING FINAL EXCLUDES"
    objLOG.write vbnewline & now & vbtab & vbtab & " - PERFORMING FINAL EXCLUDES"
    ''DEFAULT EXCLUDES
    for intEXCL = 65 to 90
      ''PROCEED WITH EXCLUDING DEFAULTS
      objOUT.write vbnewline & now & vbtab & vbtab & _
        "EXECUTING : C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -exclude " & chr(34) & chr(intEXCL) & ":\Temp" & chr(34)
      objLOG.write vbnewline & now & vbtab & vbtab & _
        "EXECUTING : C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -exclude " & chr(34) & chr(intEXCL) & ":\Temp" & chr(34)
      call HOOK("C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -exclude " & chr(34) & chr(intEXCL) & ":\Temp" & chr(34))
      objOUT.write vbnewline & now & vbtab & vbtab & _
        "EXECUTING : C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -exclude " & chr(34) & chr(intEXCL) & ":\Recovery" & chr(34)
      objLOG.write vbnewline & now & vbtab & vbtab & _
        "EXECUTING : C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -exclude " & chr(34) & chr(intEXCL) & ":\Recovery" & chr(34)
      call HOOK("C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -exclude " & chr(34) & chr(intEXCL) & ":\Recovery" & chr(34))
      objOUT.write vbnewline & now & vbtab & vbtab & _
        "EXECUTING : C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -exclude " & chr(34) & chr(intEXCL) & ":\RECYCLED" & chr(34)
      objLOG.write vbnewline & now & vbtab & vbtab & _
        "EXECUTING : C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -exclude " & chr(34) & chr(intEXCL) & ":\RECYCLED" & chr(34)
      call HOOK("C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -exclude " & chr(34) & chr(intEXCL) & ":\RECYCLED" & chr(34))
      objOUT.write vbnewline & now & vbtab & vbtab & _
        "EXECUTING : C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -exclude " & chr(34) & chr(intEXCL) & ":\$AV_ASW" & chr(34)
      objLOG.write vbnewline & now & vbtab & vbtab & _
        "EXECUTING : C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -exclude " & chr(34) & chr(intEXCL) & ":\$AV_ASW" & chr(34)
      call HOOK("C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -exclude " & chr(34) & chr(intEXCL) & ":\$AV_ASW" & chr(34))
      objOUT.write vbnewline & now & vbtab & vbtab & _
        "EXECUTING : C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -exclude " & chr(34) & chr(intEXCL) & ":\$GetCurrent" & chr(34)
      objLOG.write vbnewline & now & vbtab & vbtab & _
        "EXECUTING : C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -exclude " & chr(34) & chr(intEXCL) & ":\$GetCurrent" & chr(34)
      call HOOK("C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -exclude " & chr(34) & chr(intEXCL) & ":\$GetCurrent" & chr(34))
      objOUT.write vbnewline & now & vbtab & vbtab & _
        "EXECUTING : C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -exclude " & chr(34) & chr(intEXCL) & ":\$Recycle.Bin" & chr(34)
      objLOG.write vbnewline & now & vbtab & vbtab & _
        "EXECUTING : C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -exclude " & chr(34) & chr(intEXCL) & ":\$Recycle.Bin" & chr(34)
      call HOOK("C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -exclude " & chr(34) & chr(intEXCL) & ":\$Recycle.Bin" & chr(34))
      objOUT.write vbnewline & now & vbtab & vbtab & _
        "EXECUTING : C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -exclude " & chr(34) & chr(intEXCL) & ":\$Windows.~BT" & chr(34)
      objLOG.write vbnewline & now & vbtab & vbtab & _
        "EXECUTING : C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -exclude " & chr(34) & chr(intEXCL) & ":\$Windows.~BT" & chr(34)
      call HOOK("C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -exclude " & chr(34) & chr(intEXCL) & ":\$Windows.~BT" & chr(34))
      objOUT.write vbnewline & now & vbtab & vbtab & _
        "EXECUTING : C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -exclude " & chr(34) & chr(intEXCL) & ":\$Windows.~WS" & chr(34)
      objLOG.write vbnewline & now & vbtab & vbtab & _
        "EXECUTING : C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -exclude " & chr(34) & chr(intEXCL) & ":\$Windows.~WS" & chr(34)
      call HOOK("C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -exclude " & chr(34) & chr(intEXCL) & ":\$Windows.~WS" & chr(34))
      objOUT.write vbnewline & now & vbtab & vbtab & _
        "EXECUTING : C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -exclude " & chr(34) & chr(intEXCL) & ":\Windows10Upgrade" & chr(34)
      objLOG.write vbnewline & now & vbtab & vbtab & _
        "EXECUTING : C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -exclude " & chr(34) & chr(intEXCL) & ":\Windows10Upgrade" & chr(34)
      call HOOK("C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -exclude " & chr(34) & chr(intEXCL) & ":\Windows10Upgrade" & chr(34))
      objOUT.write vbnewline & now & vbtab & vbtab & _
        "EXECUTING : C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -exclude " & chr(34) & chr(intEXCL) & ":\hiberfil.sys" & chr(34)
      objLOG.write vbnewline & now & vbtab & vbtab & _
        "EXECUTING : C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -exclude " & chr(34) & chr(intEXCL) & ":\hiberfil.sys" & chr(34)
      call HOOK("C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -exclude " & chr(34) & chr(intEXCL) & ":\hiberfil.sys" & chr(34))
      objOUT.write vbnewline & now & vbtab & vbtab & _
        "EXECUTING : C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -exclude " & chr(34) & chr(intEXCL) & ":\pagefile.sys" & chr(34)
      objLOG.write vbnewline & now & vbtab & vbtab & _
        "EXECUTING : C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -exclude " & chr(34) & chr(intEXCL) & ":\pagefile.sys" & chr(34)
      call HOOK("C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -exclude " & chr(34) & chr(intEXCL) & ":\pagefile.sys" & chr(34))
      objOUT.write vbnewline & now & vbtab & vbtab & _
        "EXECUTING : C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -exclude " & chr(34) & chr(intEXCL) & ":\swapfile.sys" & chr(34)
      objLOG.write vbnewline & now & vbtab & vbtab & _
        "EXECUTING : C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -exclude " & chr(34) & chr(intEXCL) & ":\swapfile.sys" & chr(34)
      call HOOK("C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -exclude " & chr(34) & chr(intEXCL) & ":\swapfile.sys" & chr(34))
      objOUT.write vbnewline & now & vbtab & vbtab & _
        "EXECUTING : C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -exclude " & chr(34) & chr(intEXCL) & ":\System Volume Information" & chr(34)
      objLOG.write vbnewline & now & vbtab & vbtab & _
        "EXECUTING : C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -exclude " & chr(34) & chr(intEXCL) & ":\System Volume Information" & chr(34)
      call HOOK("C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -exclude " & chr(34) & chr(intEXCL) & ":\System Volume Information" & chr(34))
      wscript.sleep 200
    next
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
        for intCOL = 0 to ubound(arrEXCL)
          blnFND = false
          if (arrEXCL(intCOL) <> vbnullstring) then
            '' 'UNNEEDED / TO EXCLUDE' USER ACCOUNT 'ARREXCL' FOUND IN FOLDER PATH
            if (instr(1, lcase(strFOL), lcase(arrEXCL(intCOL)))) then
              objOUT.write vbnewline & now & vbtab & vbtab & vbtab & "UNNEEDED / TO EXCLUDE USER : " & arrEXCL(intCOL)
              objLOG.write vbnewline & now & vbtab & vbtab & vbtab & "UNNEEDED / TO EXCLUDE USER : " & arrEXCL(intCOL)
              ''MARK 'UNNEEDED / TO EXCLUDE'
              blnFND = true
              ''PROCEED WITH INCLUDING ENTIRE USER DIRECTORY
              objOUT.write vbnewline & now & vbtab & vbtab & _
                "EXECUTING : C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -include " & chr(34) & strFOL & chr(34)
              objLOG.write vbnewline & now & vbtab & vbtab & _
                "EXECUTING : C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -include " & chr(34) & strFOL & chr(34)
              call HOOK("C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -include " & chr(34) & strFOL & chr(34))
              wscript.sleep 200
              ''EXCLUDE USER FOLDER SUB-FOLDERS
              ''ENUMERATE 'C:\USERS\<USERNAME>' SUB-FOLDERS
              set objUFOL = objFSO.getfolder(strFOL)
              set colUFOL = objUFOL.subfolders
              for each subUFOL in colUFOL
                ''PROCEED WITH EXCLUDING USER DIRECTORY SUB-FOLDERS
                objOUT.write vbnewline & now & vbtab & vbtab & _
                  "EXECUTING : C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -exclude " & chr(34) & subUFOL.path & chr(34)
                objLOG.write vbnewline & now & vbtab & vbtab & _
                  "EXECUTING : C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -exclude " & chr(34) & subUFOL.path & chr(34)
                'call HOOK("C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -exclude " & chr(34) & subUFOL.path & chr(34))
                ''INCLUDE 'SUB-FOLDER\DESKTOP.INI' FOR EACH SUB-FOLDER TO RETAIN ORIGINAL FOLDER STRUCTURE
                objOUT.write vbnewline & now & vbtab & vbtab & _
                  "EXECUTING : C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -include " & chr(34) & subUFOL.path & "\desktop.ini" & chr(34)
                objLOG.write vbnewline & now & vbtab & vbtab & _
                  "EXECUTING : C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -include " & chr(34) & subUFOL.path & "\desktop.ini" & chr(34)
                call HOOK("C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -include " & chr(34) & subUFOL.path & "\desktop.ini" & chr(34))
                wscript.sleep 200
              next
              set objUFOL = nothing
              set colUFOL = nothing
              exit for
            end if
          end if
          ''AN 'UNNEEDED / TO EXCLUDE' USER ACCOUNT WAS PASSED TO 'STRUSR'
          'if (wscript.arguments.count > 0) then
          '  '' PASSED 'UNNEEDED / TO EXCLUDE' USER ACCOUNT 'ARREXCL'
          '  if (instr(1, lcase(strFOL), lcase(objARG.item(0)))) then
          '    objOUT.write vbnewline & now & vbtab & vbtab & vbtab & "UNNEEDED / TO EXCLUDE : " & objARG.item(0)
          '    objLOG.write vbnewline & now & vbtab & vbtab & vbtab & "UNNEEDED / TO EXCLUDE : " & objARG.item(0)
          '    ''MARK 'UNNEEDED / TO EXCLUDE'
          '    blnFND = true
          '    ''PROCEED WITH EXCLUDING ENTIRE USER DIRECTORY
          '    objOUT.write vbnewline & now & vbtab & vbtab & _
          '      "EXECUTING : C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -exclude " & chr(34) & strFOL & chr(34)
          '    objLOG.write vbnewline & now & vbtab & vbtab & _
          '      "EXECUTING : C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -exclude " & chr(34) & strFOL & chr(34)
          '    'call HOOK("C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -exclude " & chr(34) & strFOL & chr(34))
          '    exit for
          '  end if          
          'end if
        next
        ''NO MATCH TO 'UNNEEDED / TO EXCLUDE' USER ACCOUNTS
        if (not (blnFND)) then
          ''ENUMERATE THROUGH AND MAKE SURE THIS ISN'T ONE OF THE 'PROTECTED' USER ACCOUNTS
          intPCOL = 0
          for intPCOL = 0 to ubound(arrPUSR)
            blnFND = false
            if (arrPUSR(intPCOL) <> vbnullstring) then
              'objOUT.write vbnewline & arrPUSR(intPCOL)
              '' 'PRTOTECTED' USER ACCOUNTS DIRECTORIES 'ARRPUSR' FOUND IN FOLDER PATH
              if (instr(1, lcase(strFOL), lcase(arrPUSR(intPCOL)))) then
                objOUT.write vbnewline & now & vbtab & vbtab & vbtab & "PROTECTED : " & arrPUSR(intPCOL)
                objLOG.write vbnewline & now & vbtab & vbtab & vbtab & "PROTECTED : " & arrPUSR(intPCOL)
                ''PROCEED WITH INCLUDING ENTIRE USER DIRECTORY
                objOUT.write vbnewline & now & vbtab & vbtab & _
                  "EXECUTING : C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -include " & chr(34) & strFOL & chr(34)
                objLOG.write vbnewline & now & vbtab & vbtab & _
                  "EXECUTING : C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -include " & chr(34) & strFOL & chr(34)
                call HOOK("C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -include " & chr(34) & strFOL & chr(34))
                wscript.sleep 200
                ''MARK 'PROTECTED'
                blnFND = true
                exit for
              end if
            end if
          next
          ''NO MATCH TO 'PROTECTED' USER ACCOUNTS
          if (not (blnFND)) then
            ''CHECK FOR USER FOLDER
            if (objFSO.folderexists(strFOL)) then
              objOUT.write vbnewline & now & vbtab & vbtab & vbtab & "ENUMERATING : " & strFOL
              objLOG.write vbnewline & now & vbtab & vbtab & vbtab & "ENUMERATING : " & strFOL
              ''PROCEED WITH INCLUDING ENTIRE USER DIRECTORY
              objOUT.write vbnewline & now & vbtab & vbtab & _
                "EXECUTING : C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -include " & chr(34) & strFOL & chr(34)
              objLOG.write vbnewline & now & vbtab & vbtab & _
                "EXECUTING : C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -include " & chr(34) & strFOL & chr(34)
              call HOOK("C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -include " & chr(34) & strFOL & chr(34))
              wscript.sleep 200
              ''ENUMERATE 'C:\USERS\<USERNAME>\APPDATA' SUB-FOLDERS
              for intUFOL = 0 to ubound(arrAPP)
                'objOUT.write vbnewline & arrAPP(intUFOL)
                if (arrAPP(intUFOL) <> vbnullstring) then
                  'objOUT.write vbnewline & intUFOL
                  call chkSFOL(strFOL & arrAPP(intUFOL))
                end if
              next
            end if
          end if
        end if
      end if
    next
  end if
elseif (errRET <> 0) then                                     ''NO ARGUMENTS PASSED, END SCRIPT , 'ERRRET'=1
  call LOGERR(errRET)
end if
''END SCRIPT
call CLEANUP()
''END SCRIPT
''------------

'FUNCTIONS
function chkSFOL(strSFOL)
  on error resume next
  ''CHECK EACH 'C:\USERS\<USERNAME>' SUB-FOLDER
  blnFND = false
  if (strSFOL <> vbnullstring) then            
    ''ENUMERATE THROUGH AND MAKE SURE THIS ISN'T ONE OF THE 'PROTECTED' EXT / FILES / DIRECTORIES
    intPCOL = 0
    for intPCOL = 0 to ubound(arrPFOL)
      blnFND = false
      if (arrPFOL(intPCOL) <> vbnullstring) then
        '' 'PRTOTECTED' EXT / FILES / DIRECTORIES 'ARRPFOL' FOUND IN FOLDER PATH
        if (instr(1, lcase(strSFOL), lcase(arrPFOL(intPCOL)))) then
          objOUT.write vbnewline & now & vbtab & vbtab & vbtab & "PROTECTED : " & arrPFOL(intPCOL)
          objLOG.write vbnewline & now & vbtab & vbtab & vbtab & "PROTECTED : " & arrPFOL(intPCOL)
          ''PROCEED WITH INCLUDING ENTIRE USER DIRECTORY
          objOUT.write vbnewline & now & vbtab & vbtab & _
            "EXECUTING : C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -include " & chr(34) & strSFOL & chr(34)
          objLOG.write vbnewline & now & vbtab & vbtab & _
            "EXECUTING : C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -include " & chr(34) & strSFOL & chr(34)
          call HOOK("C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -include " & chr(34) & strSFOL & chr(34))
          wscript.sleep 200
          ''MARK 'PROTECTED'
          blnFND = true
          exit for
        end if
      end if
      ''A 'UNNEEDED / TO EXCLUDE' USER ACCOUNT WAS PASSED TO 'STRUSR'
      'if (wscript.arguments.count > 0) then
      '  '' PASSED 'PRTOTECTED' USER ACCOUNT 'ARREXCL'
      '  if (instr(1, lcase(strSFOL), lcase(objARG.item(0)))) then
      '    objOUT.write vbnewline & now & vbtab & vbtab & vbtab & "UNNEEDED / TO EXCLUDE : " & objARG.item(0)
      '    objLOG.write vbnewline & now & vbtab & vbtab & vbtab & "UNNEEDED / TO EXCLUDE : " & objARG.item(0)
      '    ''MARK 'UNNEEDED / TO EXCLUDE'
      '    blnFND = true
      '    exit for
      '  end if          
      'end if
    next
    ''NO MATCH TO 'PROTECTED' EXT / FILES / DIRECTORIES
    if (not (blnFND)) then
      ''OUTLOOK OST / TMP  AND ICONCACHE / THUMBCACHE EXCLUSIONS
      if (instr(1, strSFOL, "*")) then
        strTMP = vbnullstring
        arrTMP = split(strSFOL, "\")
        for intTMP = 0 to (ubound(arrTMP) - 1)
          strTMP = strTMP & arrTMP(intTMP) & "\"
        next
        set objSFOL = objFSO.getfolder(strTMP)
        set colSFIL = objSFOL.files
        for each subFIL in colSFIL
          if (instr(1, lcase(subFIL.path), lcase(split(strSFOL, "*")(0)))) then
            if (instr(1, lcase(subFIL.path), lcase(split(strSFOL, "*")(1)))) then
              'objOUT.write vbnewline & "FILE : " & lcase(subFIL.path)
              'objOUT.write vbnewline & "MATCH : " & lcase(split(strSFOL, "*")(0))
              'objOUT.write vbnewline & "MATCH : " & lcase(split(strSFOL, "*")(1))
              ''EXCLUDE FOLDER / FILE
              objOUT.write vbnewline & now & vbtab & vbtab & _
                "EXECUTING : C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -exclude " & chr(34) & subFIL.path & chr(34)
              objLOG.write vbnewline & now & vbtab & vbtab & _
                "EXECUTING : C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -exclude " & chr(34) & subFIL.path & chr(34)
              call HOOK("C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -exclude " & chr(34) & subFIL.path & chr(34))
              wscript.sleep 200
            end if
          end if
          wscript.sleep 100
        next
        set colSFIL = nothing
        set objSFOL = nothing
      elseif (instr(1, strSFOL, "*") = 0) then
        ''EXCLUDE FOLDER / FILE
        objOUT.write vbnewline & now & vbtab & vbtab & _
          "EXECUTING : C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -exclude " & chr(34) & strSFOL & chr(34)
        objLOG.write vbnewline & now & vbtab & vbtab & _
          "EXECUTING : C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -exclude " & chr(34) & strSFOL & chr(34)
        call HOOK("C:\Program Files\Backup Manager\clienttool.exe control.selection.modify -datasource FileSystem -exclude " & chr(34) & strSFOL & chr(34))
        wscript.sleep 200
      end if
    end if
  end if
  ''USE TO CHECK FURTHER SUB-FOLDERS / FILES
  'set objSFOL = objFSO.getfolder(strSFOL)
  'set colSFOL = objSFOL.subfolders
  'for each subSFOL in colSFOL
  '  call chkSFOL(subSFOL.path)
  'next
  'set colSFOL = nothing
  'set objSFOL = nothing
end function

''SUB-ROUTINES
sub FILEDL(strURL, strDL, strFILE)                            ''CALL HOOK TO DOWNLOAD FILE FROM URL , 'ERRRET'=11
  strSAV = vbnullstring
  ''SET DOWNLOAD PATH
  strSAV = strDL & "\" & strFILE
  objOUT.write vbnewline & now & vbtab & vbtab & "HTTPDOWNLOAD-------------DOWNLOAD : " & strURL & " : SAVE AS :  " & strSAV
  objLOG.write vbnewline & now & vbtab & vbtab & "HTTPDOWNLOAD-------------DOWNLOAD : " & strURL & " : SAVE AS :  " & strSAV
  ''CHECK IF FILE ALREADY EXISTS
  if objFSO.fileexists(strSAV) then
    ''DELETE FILE FOR OVERWRITE
    objFSO.deletefile(strSAV)
  end if
  ''CREATE HTTP OBJECT
  set objHTTP = createobject("WinHttp.WinHttpRequest.5.1")
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
    objOUT.write vbnewline & vbnewline & now & vbtab & " - DOWNLOAD : " & strSAV & " : SUCCESSFUL"
    objLOG.write vbnewline & vbnewline & now & vbtab & " - DOWNLOAD : " & strSAV & " : SUCCESSFUL"
  end if
	set objHTTP = nothing
  if ((err.number <> 0) and (err.number <> 58)) then          ''ERROR RETURNED DURING UPDATE CHECK , 'ERRRET'=11
    call LOGERR(11)
  end if
end sub

sub HOOK(strCMD)                                              ''CALL HOOK TO MONITOR OUTPUT OF CALLED COMMAND , 'ERRRET'=12
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
  if (err.number <> 0) then                                   ''ERROR RETURNED , 'ERRRET'=12
    call LOGERR(12)
  end if
end sub

sub LOGERR(intSTG)                                            ''CALL HOOK TO MONITOR OUTPUT OF CALLED COMMAND
  errRET = intSTG
  if (err.number <> 0) then
    objOUT.write vbnewline & now & vbtab & vbtab & vbtab & err.number & vbtab & err.description & vbnewline
    objLOG.write vbnewline & now & vbtab & vbtab & vbtab & err.number & vbtab & err.description & vbnewline
		err.clear
  end if
  select case intSTG
    case 1                                                  '' 'ERRRET'=1 - NOT ENOUGH ARGUMENTS
  end select
end sub

sub CLEANUP()                                               ''SCRIPT CLEANUP
  on error resume next
  if (errRET = 0) then         															''MSP_FILTER COMPLETED SUCCESSFULLY
    objOUT.write vbnewline & vbnewline & now & vbtab & " - MSP_FILTER SUCCESSFUL : " & now
    objLOG.write vbnewline & vbnewline & now & vbtab & " - MSP_FILTER SUCCESSFUL : " & now
    err.clear
  elseif (errRET <> 0) then    															''MSP_FILTER FAILED
    objOUT.write vbnewline & vbnewline & now & vbtab & " - MSP_FILTER FAILURE : " & errRET & " : " & now
    objLOG.write vbnewline & vbnewline & now & vbtab & " - MSP_FILTER FAILURE : " & errRET & " : " & now
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
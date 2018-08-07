'on error resume next
''ALWAYS RIGHT-CLICK SCRIPT, CHOOSE "PROPERTIES", CLICK "UNBLOCK"
''SCRIPT VARIABLES
dim strIN, strNUL, strSEL, strRUN
dim strNAM, strACT, strDAT, strDAY, strMON, strTIM, strARC
dim objIN, objOUT, objARG, objWSH, objFSO, objSCR, objLOG, objHOOK
''STDIN / STDOUT
set objIN = wscript.stdin
set objOUT = wscript.stdout
set objARG = wscript.arguments
''OBJECTS FOR LOCATING FOLDERS
set objWSH = createobject("wscript.shell")
set objFSO = createobject("scripting.filesystemobject")
''SET EXECUTION FLAG
strRUN = "false"
''PREPARE LOGFILE
if (objFSO.fileexists("C:\temp\set_msparchive")) then      ''LOGFILE EXISTS
  objFSO.deletefile "C:\temp\set_msparchive", true
  set objLOG = objFSO.createtextfile("C:\temp\set_msparchive")
  objLOG.close
  set objLOG = objFSO.opentextfile("C:\temp\set_msparchive", 8)
else                                                  ''LOGFILE NEEDS TO BE CREATED
  set objLOG = objFSO.createtextfile("C:\temp\set_msparchive")
  objLOG.close
  set objLOG = objFSO.opentextfile("C:\temp\set_msparchive", 8)
end if
''READ PASSED COMMANDLINE ARGUMENTS
if (wscript.arguments.count > 0) then                 ''ARGUMENTS WERE PASSED
  for x = 0 to (wscript.arguments.count - 1)
    objOUT.write vbnewline & now & vbtab & "ARGUMENT " & (x + 1) & " (ITEM " & x & ") " & " PASSED : " & objARG.item(x)
    objLOG.write vbnewline & now & vbtab & "ARGUMENT " & (x + 1) & " (ITEM " & x & ") " & " PASSED : " & objARG.item(x)
  next
  ''SET ARCHIVE SCHEDULE NAME
  strNAM = objARG.item(0)
  ''SET SERVICE / CONFIGURATION OPTIONS
  if (wscript.arguments.count = 7) then               ''ALL ARGUMENTS PASSED
    ''SET ARCHIVE SCHEDULE ACTIVE
    strACT = objARG.item(1)
    ''SET ARCHIVE DATASOURCES
    strDAT = objARG.item(2)
    ''SET ARCHIVE DAYS OF MONTH
    strDAY = objARG.item(3)
    ''SET ARCHIVE MONTHS
    strMON = objARG.item(4)
    ''SET ARCHIVE TIME
    strTIM = objARG.item(5)
    ''SET SCRIPT RUN LEVEL
    strRUN = objARG.item(6)
    if (strRUN = "true") then
      strARC = "Y"
    end if
  end if
else                                                  ''NO ARGUMENTS PASSED
  objOUT.write vbnewline & now & vbtab & " - NO ARGUMENTS PASSED. SCRIPT WILL REQUEST SETTINGS DURING EXECUTION"
  objLOG.write vbnewline & now & vbtab & " - NO ARGUMENTS PASSED. SCRIPT WILL REQUEST SETTINGS DURING EXECUTION"
end if
objOUT.write vbnewline & now & " - STARTING SET_MSPARCHIVE" & vbnewline
objLOG.write vbnewline & now & " - STARTING SET_MSPARCHIVE" & vbnewline
''ENTER CALL VERIFY LOOP
call VERIFY()
''END SCRIPT
call CLEANUP()

''SUB-ROUTINES
sub VERIFY()                                          ''CALL HOOK TO VERIFY SCRIPT CONFIGURATIONS, RUNS IN A LOOP UNTIL EXECUTION / TERMINATION
  if (wscript.arguments.count = 0) then               ''SCRIPT NOT PRE-CONFIGURED
    strSEL = vbnullstring
    ''ENABLE MSP BACKUP ARCHIVES
    if (strARC = vbnullstring) then
      objOUT.write vbnewline & vbtab & "ENABLE MSP BACKUP ARCHIVES (Y / N)?"
      objLOG.write vbnewline & vbtab & "ENABLE MSP BACKUP ARCHIVES (Y / N)?"
      strSEL = objIN.readline
      strARC = strSEL
      if (ucase(strSEL) = "Y") then
        objOUT.write vbnewline & vbtab & "ENABLE MSP BACKUP ARCHIVES : " & strARC
        objLOG.write vbnewline & vbtab & "ENABLE MSP BACKUP ARCHIVES : " & strARC
      end if
    elseif (strARC <> vbnullstring) then
      objOUT.write vbnewline & vbtab & "ENABLE MSP BACKUP ARCHIVES : " & strARC
      objLOG.write vbnewline & vbtab & "ENABLE MSP BACKUP ARCHIVES : " & strARC
    end if
    strSEL = vbnullstring
  elseif (wscript.arguments.count > 0) then           ''SCRIPT PRE-CONFIGURED
    ''PLACEHOLDER
  end if
  ''EXECUTION CHECK
  if (lcase(strRUN) <> "true") then
    objOUT.write vbnewline & vbnewline & vbtab & "EXECUTE SCRIPT WITH CONFIGURED SETTINGS (Y / N)"
    objLOG.write vbnewline & vbnewline & vbtab & "EXECUTE SCRIPT WITH CONFIGURED SETTINGS (Y / N)"
    strSEL = objIN.readline
    if (ucase(strSEL) = "Y") then
      strRUN = "true"
      objOUT.write vbnewline & vbnewline & now & " - EXECUTING SET_MSPARCHIVE SCRIPT" & vbnewline
      objLOG.write vbnewline & vbnewline & now & " - EXECUTING SET_MSPARCHIVE SCRIPT" & vbnewline
      ''EXIT VERIFY LOOP, RUN SCRIPT EXECUTION
      call EXECUTE()
    elseif (ucase(strSEL) = "N") then
      strRUN = "false"
      objOUT.write vbnewline & vbnewline & now & " - SKIPPING SCRIPT EXECUTION" & vbnewline
      objLOG.write vbnewline & vbnewline & now & " - SKIPPING SCRIPT EXECUTION" & vbnewline
      ''RETURN CALL TO VERIFY LOOP
      call VERIFY()
    end if
  elseif (lcase(strRUN) = "true") then
    ''EXIT VERIFY LOOP, RUN SCRIPT EXECUTION
    call EXECUTE()
  end if
end sub

sub EXECUTE()                                         ''CALL HOOK TO EXECUTE SCRIPT CHANGES
  if (strRUN = "true") then
    ''PAUSE TO ENABLE MSP BACKUP ARCHIVES
    if (lcase(strARC) = "y") then
      ''REQUEST LOCAL SPEEDVAULT PATH
      'objOUT.write vbnewline & vbnewline & now & vbtab & " - ENTER LOCAL SPEEDVAULT PATH : "
      'objLOG.write vbnewline & vbnewline & now & vbtab & " - ENTER LOCAL SPEEDVAULT PATH : "
      'strLSVL = objIN.readline
      ''REQUEST RMMTECH CREDENTIALS FOR LOCAL SPEEDVAULT
      'objOUT.write vbnewline & vbnewline & now & vbtab & " - ENTER RMMTECH USERNAME FOR LSV ACCESS : "
      'objLOG.write vbnewline & vbnewline & now & vbtab & " - ENTER RMMTECH USERNAME FOR LSV ACCESS : "
      'strLSVU = objIN.readline
      'objOUT.write vbnewline & vbnewline & now & vbtab & " - ENTER RMMTECH PASSWORD FOR LSV ACCESS : "
      'objLOG.write vbnewline & vbnewline & now & vbtab & " - ENTER RMMTECH PASSWORD FOR LSV ACCESS : "
      'strLSVP = objIN.readline
      ''SET DEFAULT 'CW_DEFAULT_MSPARCHIVE" ARCHIVING SCHEDULE
      objOUT.write vbnewline & vbnewline & now & vbtab & " - APPLYING MSP BACKUP ARCHIVE SCHEDULE '" & strNAM & "'"
      objLOG.write vbnewline & vbnewline & now & vbtab & " - APPLYING MSP BACKUP ARCHIVE SCHEDULE '" & strNAM & "'"
      call HOOK("C:\Program Files\Backup Manager\ClientTool.exe control.archiving.add -name " & chr(34) & strNAM & chr(34) & _
        " -active " & chr(34) & strACT & chr(34) & " -datasources " & chr(34) & strDAT & chr(34) & " -days-of-month " & chr(34) & strDAY & chr(34) & _
        " -months " & chr(34) & strMON & chr(34) & " -time " & chr(34) & strTIM & chr(34))
      ''STEP TO VERIFY MSP BACKUP ARCHIVING SCHEDULE
      objOUT.write vbnewline & vbnewline & now & vbtab & " - PLEASE VERIFY MSP BACKUP ARCHIVE SCHEDULE 'CW_DEFAULT_MSPARCHIVE'"
      objOUT.write vbnewline & now & vbtab & " - VIA N-CENTRAL>CONFIGURATION>BACKUP MANAGER>MSP BACKUPS>DASHBOARD>LAUNCH BACKUP CLIENT>PREFERENCES>ARCHIVING"
      objOUT.write vbnewline & now & vbtab & " - PRESS 'ENTER' WHEN READY"
      objLOG.write vbnewline & vbnewline & now & vbtab & " - PLEASE VERIFY MSP BACKUP ARCHIVE SCHEDULE 'CW_DEFAULT_MSPARCHIVE'"
      objLOG.write vbnewline & now & vbtab & " - VIA N-CENTRAL>CONFIGURATION>BACKUP MANAGER>MSP BACKUPS>DASHBOARD>LAUNCH BACKUP CLIENT>PREFERENCES>ARCHIVING"
      if (wscript.arguments.count <> 7) then               ''NOT ALL ARGUMENTS PASSED
        objLOG.write vbnewline & now & vbtab & " - PRESS 'ENTER' WHEN READY"
        strNUL = objIN.readline
      end if
    end if
  end if
end sub

sub HOOK(strCMD)                                      ''CALL HOOK TO MONITOR OUTPUT OF CALLED COMMAND
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
  end if
end sub

sub CLEANUP()                                         ''SCRIPT CLEANUP
  objOUT.write vbnewline & vbnewline & now & " - SET_MSPARCHIVE COMPLETE. PLEASE VERIFY ALL MONITORING AND SERVICES!" & vbnewline
  objLOG.write vbnewline & vbnewline & now & " - SET_MSPARCHIVE COMPLETE. PLEASE VERIFY ALL MONITORING AND SERVICES!" & vbnewline
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
'on error resume next
''ALWAYS RIGHT-CLICK SCRIPT, CHOOSE "PROPERTIES", CLICK "UNBLOCK"
''SCRIPT VARIABLES
''STANDARD VARIABLES
dim objIN, objOUT, objARG, objWSH, objFSO, objSCR, objLOG, objHOOK
''CONFIGURABLE VARIABLES
dim objSIN, objSOUT, strORG, strREP   ''SVCLOGON RIGHTS VARIABLES
dim strPATH, strSAV, strNUSR, strNPWD,  strSEL, strRUN
dim strAGT, strAGTdl, strAVD, strAVDdl, strPWD, strUPD, strPRB, strPRBdl
dim strMSP, strMSPdl, strMSPVD, strMSPVDdl, strLSV, strLSVL, strLSVU, strLSVP, strARC
''STDIN / STDOUT
set objIN = wscript.stdin
set objOUT = wscript.stdout
set objARG = wscript.arguments
''OBJECTS FOR LOCATING FOLDERS
set objWSH = createobject("wscript.shell")
set objFSO = createobject("scripting.filesystemobject")
Set objSCR = objFSO.getfile(wscript.scriptfullname)
strPATH = objFSO.getparentfoldername(objSCR)
if (right(strPATH,1) <> "\") then
  strPATH = strPATH & "\"
end if
''PREPARE LOGFILE
if (objFSO.fileexists("C:\temp\auto_plan")) then      ''LOGFILE EXISTS
  objFSO.deletefile "C:\temp\auto_plan", true
  set objLOG = objFSO.createtextfile("C:\temp\auto_plan")
  objLOG.close
  set objLOG = objFSO.opentextfile("C:\temp\auto_plan", 8)
else                                                  ''LOGFILE NEEDS TO BE CREATED
  set objLOG = objFSO.createtextfile("C:\temp\auto_plan")
  objLOG.close
  set objLOG = objFSO.opentextfile("C:\temp\auto_plan", 8)
end if
''SET EXECUTION FLAG
strRUN = "false"
''READ PASSED COMMANDLINE ARGUMENTS
if (wscript.arguments.count > 0) then                 ''ARGUMENTS WERE PASSED
  for x = 0 to (wscript.arguments.count - 1)
    objOUT.write vbnewline & now & vbtab & "ARGUMENT " & (x + 1) & " (ITEM " & x & ") " & " PASSED : " & objARG.item(x)
    objLOG.write vbnewline & now & vbtab & "ARGUMENT " & (x + 1) & " (ITEM " & x & ") " & " PASSED : " & objARG.item(x)
  next
  ''SET RMMTECH PASSWORD
  strPWD = objARG.item(0)
  ''SET SERVICE / CONFIGURATION OPTIONS
  if (wscript.arguments.count = 10) then               ''ALL ARGUMENTS PASSED
    ''SET AGENT INSTALL
    strAGT = objARG.item(1)
    ''SET AV DEFENDER INSTALL
    strAVD = objARG.item(2)
    ''PAUSE TO ENABLE PATCHING
    strUPD = objARG.item(3)
    ''SET PROBE INSTALL
    strPRB = objARG.item(4)
    ''SET MSP BACKUP INSTALL
    strMSP = objARG.item(5)
    ''PAUSE TO ENABLE MSP BACKUP LOCAL SPEEDVAULT
    strLSV = objARG.item(6)
    ''PAUSE TO ENABLE MSP BACKUP ARCHIVES
    strARC = objARG.item(7)
    ''SET MSP BACKUP VIRTUAL DRIVE INSTALL
    strMSPVD = objARG.item(8)
    ''SET EXECUTION FLAG
    strRUN = objARG.item(9)
    objOUT.write vbnewline & now & vbtab & " - ALL ARGUMENTS PASSED, SCRIPT WILL SET : RMMTECH PASSWORD TO " & chr(34) & strPWD & chr(34) & _
      " : INSTALL WINDOWS AGENT : " & strAGT & " : INSTALL AV DEFENDER : " & strAVD & " : ENABLE PATCHING : " & strUPD & " : INSTALL WINDOWS PROBE : " & strPRB & _
      " : INSTALL MSP BACKUP : " & strMSP & " : ENABLE MSP BACKUP LSV : " & strLSV & " : ENABLE MSP BACKUP ARCHVIVES : " & strARC & " : INSTALL MSP BACKUP VDRIVE : " & strMSPVD
    objLOG.write vbnewline & now & vbtab & " - ALL ARGUMENTS PASSED, SCRIPT WILL SET : RMMTECH PASSWORD : INSTALL WINDOWS AGENT : INSTALL AV DEFENDER : ENABLE PATCHING : INSTALL WINDOWS PROBE" & _
      " : INSTALL MSP BACKUP : ENABLE MSP BACKUP LSV : ENABLE MSP BACKUP ARCHIVES : INSTALL MSP BACKUP VDRIVE"
  else                                                ''NOT ENOUGH ARGUMENTS PASSED
    if (wscript.arguments.count < 2) then
      objOUT.write vbnewline & now & vbtab & " - ONLY 1 ARGUMENT PASSED, SCRIPT WILL ONLY SET : RMMTECH PASSWORD TO " & chr(34) & strPWD & chr(34)
      objLOG.write vbnewline & now & vbtab & " - ONLY 1 ARGUMENT PASSED, SCRIPT WILL ONLY SET : RMMTECH PASSWORD"
    elseif (wscript.arguments.count < 3) then
      strAGT = objARG.item(1)
      objOUT.write vbnewline & now & vbtab & " - ONLY 2 ARGUMENTS PASSED, SCRIPT WILL ONLY SET : RMMTECH PASSWORD TO " & chr(34) & strPWD & chr(34) & _
        " : INSTALL WINDOWS AGENT : " & strAGT
      objLOG.write vbnewline & now & vbtab & " - ONLY 2 ARGUMENTS PASSED, SCRIPT WILL ONLY SET : RMMTECH PASSWORD : INSTALL WINDOWS AGENT"
    elseif (wscript.arguments.count < 4) then
      strAGT = objARG.item(1)
      strAVD = objARG.item(2)
      objOUT.write vbnewline & now & vbtab & " - ONLY 3 ARGUMENTS PASSED, SCRIPT WILL ONLY SET : RMMTECH PASSWORD TO " & chr(34) & strPWD & chr(34) & _
        " : INSTALL WINDOWS AGENT : " & strAGT & " : INSTALL AV DEFENDER : " & strAVD
      objLOG.write vbnewline & now & vbtab & " - ONLY 3 ARGUMENTS PASSED, SCRIPT WILL ONLY SET : RMMTECH PASSWORD : INSTALL WINDOWS AGENT : INSTALL AV DEFENDER"
    elseif (wscript.arguments.count < 5) then
      strAGT = objARG.item(1)
      strAVD = objARG.item(2)
      strUPD = objARG.item(3)
      objOUT.write vbnewline & now & vbtab & " - ONLY 4 ARGUMENTS PASSED, SCRIPT WILL ONLY SET : RMMTECH PASSWORD TO " & chr(34) & strPWD & chr(34) & _
        " : INSTALL WINDOWS AGENT : " & strAGT & " : INSTALL AV DEFENDER : " & strAVD & " : ENABLE PATCHING : " & strUPD
      objLOG.write vbnewline & now & vbtab & " - ONLY 4 ARGUMENTS PASSED, SCRIPT WILL ONLY SET ; RMMTECH PASSWORD : INSTALL WINDOWS AGENT : INSTALL AV DEFENDER : ENABLE PATCHING"
    elseif (wscript.arguments.count < 6) then
      strAGT = objARG.item(1)
      strAVD = objARG.item(2)
      strUPD = objARG.item(3)
      strPRB = objARG.item(4)
      objOUT.write vbnewline & now & vbtab & " - ONLY 5 ARGUMENTS PASSED, SCRIPT WILL ONLY SET : RMMTECH PASSWORD TO " & chr(34) & strPWD & chr(34) & _
        " : INSTALL WINDOWS AGENT : " & strAGT & " : INSTALL AV DEFENDER : " & strAVD & " : ENABLE PATCHING : " & strUPD & " : INSTALL WINDOWS PROBE : " & strPRB
      objLOG.write vbnewline & now & vbtab & " - ONLY 5 ARGUMENTS PASSED, SCRIPT WILL ONLY SET : RMMTECH PASSWORD : INSTALL WINDOWS AGENT : INSTALL AV DEFENDER : ENABLE PATCHING : INSTALL WINDOWS PROBE"
    elseif (wscript.arguments.count < 7) then
      strAGT = objARG.item(1)
      strAVD = objARG.item(2)
      strUPD = objARG.item(3)
      strPRB = objARG.item(4)
      strMSP = objARG.item(5)
      objOUT.write vbnewline & now & vbtab & " - ONLY 6 ARGUMENTS PASSED, SCRIPT WILL ONLY SET : RMMTECH PASSWORD TO " & chr(34) & strPWD & chr(34) & _
        " : INSTALL WINDOWS AGENT : " & strAGT & " : INSTALL AV DEFENDER : " & strAVD & " : ENABLE PATCHING : " & strUPD & " : INSTALL WINDOWS PROBE : " & strPRB & _
        " : INSTALL MSP BACKUP " & strMSP
      objLOG.write vbnewline & now & vbtab & " - ONLY 6 ARGUMENTS PASSED, SCRIPT WILL ONLY SET : RMMTECH PASSWORD : INSTALL WINDOWS AGENT : INSTALL AV DEFENDER : ENABLE PATCHING : INSTALL WINDOWS PROBE" & _
        ": INSTALL MSP BACKUP"
    elseif (wscript.arguments.count < 8) then
      strAGT = objARG.item(1)
      strAVD = objARG.item(2)
      strUPD = objARG.item(3)
      strPRB = objARG.item(4)
      strMSP = objARG.item(5)
      strLSV = objARG.item(6)
      objOUT.write vbnewline & now & vbtab & " - ONLY 7 ARGUMENTS PASSED, SCRIPT WILL ONLY SET : RMMTECH PASSWORD TO " & chr(34) & strPWD & chr(34) & _
        " : INSTALL WINDOWS AGENT : " & strAGT & " : INSTALL AV DEFENDER : " & strAVD & " : ENABLE PATCHING : " & strUPD & " : INSTALL WINDOWS PROBE : " & strPRB & _
        " : INSTALL MSP BACKUP " & strMSP & " : ENABLE MSP BACKUP LSV : " & strLSV
      objLOG.write vbnewline & now & vbtab & " - ONLY 7 ARGUMENTS PASSED, SCRIPT WILL ONLY SET : RMMTECH PASSWORD : INSTALL WINDOWS AGENT : INSTALL AV DEFENDER : ENABLE PATCHING : INSTALL WINDOWS PROBE" & _
        ": INSTALL MSP BACKUP : ENABLE MSP BACKUP LSV"
    elseif (wscript.arguments.count < 9) then
      strAGT = objARG.item(1)
      strAVD = objARG.item(2)
      strUPD = objARG.item(3)
      strPRB = objARG.item(4)
      strMSP = objARG.item(5)
      strLSV = objARG.item(6)
      strARC = objARG.item(7)
      objOUT.write vbnewline & now & vbtab & " - ONLY 8 ARGUMENTS PASSED, SCRIPT WILL ONLY SET : RMMTECH PASSWORD TO " & chr(34) & strPWD & chr(34) & _
        " : INSTALL WINDOWS AGENT : " & strAGT & " : INSTALL AV DEFENDER : " & strAVD & " : ENABLE PATCHING : " & strUPD & " : INSTALL WINDOWS PROBE : " & strPRB & _
        " : INSTALL MSP BACKUP " & strMSP & " : ENABLE MSP BACKUP LSV : " & strLSV & " : ENABLE MSP BACKUP ARCHIVES : " & strARC
      objLOG.write vbnewline & now & vbtab & " - ONLY 8 ARGUMENTS PASSED, SCRIPT WILL ONLY SET : RMMTECH PASSWORD : INSTALL WINDOWS AGENT : INSTALL AV DEFENDER : ENABLE PATCHING : INSTALL WINDOWS PROBE" & _
        " : INSTALL MSP BACKUP : ENABLE MSP BACKUP LSV : ENABLE MSP BACKUP ARCHIVES"
    end if
  end if
else                                                  ''NO ARGUMENTS PASSED
  objOUT.write vbnewline & now & vbtab & " - NO ARGUMENTS PASSED. SCRIPT WILL REQUEST SETTINGS DURING EXECUTION"
  objLOG.write vbnewline & now & vbtab & " - NO ARGUMENTS PASSED. SCRIPT WILL REQUEST SETTINGS DURING EXECUTION"
end if
objOUT.write vbnewline & now & " - STARTING AUTO_PLAN" & vbnewline
objLOG.write vbnewline & now & " - STARTING AUTO_PLAN" & vbnewline
''CHANGE ACTIVE POWER PLAN
objOUT.write vbnewline & now & " - SETTING ACTIVE POWER PLAN : HIGH PERFORMANCE" & vbnewline
objLOG.write vbnewline & now & " - SETTING ACTIVE POWER PLAN : HIGH PERFORMANCE" & vbnewline
call HOOK("powercfg.exe /setactive 8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c")
''DISABLE HIBERNATION
objOUT.write vbnewline & now & " - DISABLING HIBERNATION" & vbnewline
objLOG.write vbnewline & now & " - DISABLING HIBERNATION" & vbnewline
call HOOK("powercfg –h off")
''VERIFY SCRIPT SETTINGS
objOUT.write vbnewline & now & " - CHECKING SCRIPT SETTINGS" & vbnewline
objLOG.write vbnewline & now & " - CHECKING SCRIPT SETTINGS" & vbnewline
if (wscript.arguments.count <> 10) then
  ''ENTER CALL VERIFY LOOP, DO NOT SKIP RMMTECH PASSWORD SETTINGS
  call VERIFY("false")
elseif (wscript.arguments.count = 10) then
  ''ENTER CALL VERIFY LOOP, SKIP RMMTECH PASSWORD SETTINGS
  call VERIFY("true")
end if
''END SCRIPT
call CLEANUP()

''SUB-ROUTINES
sub VERIFY(skipPW)                                     ''CALL HOOK TO VERIFY SCRIPT CONFIGURATIONS, RUNS IN A LOOP UNTIL EXECUTION / TERMINATION
  'if (wscript.arguments.count = 0) then               ''SCRIPT NOT PRE-CONFIGURED
    strSEL = vbnullstring
    ''CREATE AND UPDATE RMMTECH
    if (skipPW = "false") then
      if (strPWD = vbnullstring) then
        objOUT.write vbnewline & vbnewline & vbtab & "CREATE AND UPDATE RMMTECH USER (Y / N)?"
        objLOG.write vbnewline & vbnewline & vbtab & "CREATE AND UPDATE RMMTECH USER (Y / N)?"
        strSEL = objIN.readline
        if (ucase(strSEL) = "Y") then
          objOUT.write vbnewline & vbtab & "ENTER NEW PASSWORD : "
          objLOG.write vbnewline & vbtab & "ENTER NEW PASSWORD : "
          strPWD = objIN.readline
        end if
      elseif ((strPWD <> vbnullstring) and (lcase(left(strSEL, 1)) <> "n")) then
        objOUT.write vbnewline & vbtab & "CREATE AND UPDATE RMMTECH : " & strPWD
        objLOG.write vbnewline & vbtab & "CREATE AND UPDATE RMMTECH : Y"
      end if
    end if
    strSEL = vbnullstring
    ''INSTALL WINDOWS AGENT
    if (strAGT = vbnullstring) then
      objOUT.write vbnewline & vbtab & "INSTALL WINDOWS AGENT (Y / N)?"
      objLOG.write vbnewline & vbtab & "INSTALL WINDOWS AGENT (Y / N)?"
      strSEL = objIN.readline
      strAGT = strSEL
      if (ucase(strSEL) = "Y") then
        objOUT.write vbnewline & vbtab & "ENTER WINDOWS AGENT URL : "
        objLOG.write vbnewline & vbtab & "ENTER WINDOWS AGENT URL : "
        strAGTdl = objIN.readline
      end if
    elseif (strAGT <> vbnullstring) then
      objOUT.write vbnewline & vbtab & "INSTALL WINDOWS AGENT : " & strAGT & " : " & strAGTdl
      objLOG.write vbnewline & vbtab & "INSTALL WINDOWS AGENT : " & strAGT & " : " & strAGTdl
    end if
    strSEL = vbnullstring
    ''INSTALL AV DEFENDER
    if (strAVD = vbnullstring) then
      objOUT.write vbnewline & vbtab & "INSTALL AV DEFENDER (Y / N)?"
      objLOG.write vbnewline & vbtab & "INSTALL AV DEFENDER (Y / N)?"
      strSEL = objIN.readline
      strAVD = strSEL
      if (ucase(strSEL) = "Y") then
        objOUT.write vbnewline & vbtab & "ENTER AV DEFENDER URL : "
        objLOG.write vbnewline & vbtab & "ENTER AV DEFENDER URL : "
        strAVDdl = objIN.readline
      end if
    elseif (strAVD <> vbnullstring) then
      objOUT.write vbnewline & vbtab & "INSTALL AV DEFENDER : " & strAVD & " : " & strAVDdl
      objLOG.write vbnewline & vbtab & "INSTALL AV DEFENDER : " & strAVD & " : " & strAVDdl
    end if
    strSEL = vbnullstring
    ''ENABLE PATCHING
    if (strUPD = vbnullstring) then
      objOUT.write vbnewline & vbtab & "ENABLE PATCHING (Y / N)?"
      objLOG.write vbnewline & vbtab & "ENABLE PATCHING (Y / N)?"
      strSEL = objIN.readline
      strUPD = strSEL
      if (ucase(strSEL) = "Y") then
        objOUT.write vbnewline & vbtab & "ENABLE PATCHING : " & strUPD
        objLOG.write vbnewline & vbtab & "ENABLE PATCHING : " & strUPD
      end if
    elseif (strUPD <> vbnullstring) then
      objOUT.write vbnewline & vbtab & "ENABLE PATCHING : " & strUPD
      objLOG.write vbnewline & vbtab & "ENABLE PATCHING : " & strUPD
    end if
    strSEL = vbnullstring
    ''INSTALL PROBE
    if (strPRB = vbnullstring) then
      objOUT.write vbnewline & vbtab & "INSTALL WINDOWS PROBE (Y / N)?"
      objLOG.write vbnewline & vbtab & "INSTALL WINDOWS PROBE (Y / N)?"
      strSEL = objIN.readline
      strPRB = strSEL
      if (ucase(strSEL) = "Y") then
        objOUT.write vbnewline & vbtab & "ENTER WINDOWS PROBE URL : "
        objLOG.write vbnewline & vbtab & "ENTER WINDOWS PROBE URL : "
        strPRBdl = objIN.readline
      end if
    elseif (strPRB <> vbnullstring) then
      objOUT.write vbnewline & vbtab & "INSTALL WINDOWS PROBE : " & strPRB & " : " & strPRBdl
      objLOG.write vbnewline & vbtab & "INSTALL WINDOWS PROBE : " & strPRB & " : " & strPRBdl
    end if
    strSEL = vbnullstring
    ''INSTALL MSP BACKUP
    if (strMSP = vbnullstring) then
      objOUT.write vbnewline & vbtab & "INSTALL MSP BACKUP (Y / N)?"
      objLOG.write vbnewline & vbtab & "INSTALL MSP BACKUP (Y / N)?"
      strSEL = objIN.readline
      strMSP = strSEL
      if (ucase(strSEL) = "Y") then
        objOUT.write vbnewline & vbtab & "ENTER MSP BACKUP URL : "
        objLOG.write vbnewline & vbtab & "ENTER MSP BACKUP URL : "
        strMSPdl = objIN.readline
      end if
    elseif (strMSP <> vbnullstring) then
      objOUT.write vbnewline & vbtab & "INSTALL MSP BACKUP : " & strMSP & " : " & strMSPdl
      objLOG.write vbnewline & vbtab & "INSTALL MSP BACKUP : " & strMSP & " : " & strMSPdl
    end if
    strSEL = vbnullstring
    ''ENABLE MSP BACKUP LOCAL SPEEDVAULT
    if (strLSV = vbnullstring) then
      objOUT.write vbnewline & vbtab & "ENABLE MSP BACKUP LSV (Y / N)?"
      objLOG.write vbnewline & vbtab & "ENABLE MSP BACKUP LSV (Y / N)?"
      strSEL = objIN.readline
      strLSV = strSEL
      if (ucase(strSEL) = "Y") then
        objOUT.write vbnewline & vbtab & "ENABLE MSP BACKUP LSV : " & strLSV
        objLOG.write vbnewline & vbtab & "ENABLE MSP BACKUP LSV : " & strLSV
      end if
    elseif (strLSV <> vbnullstring) then
      objOUT.write vbnewline & vbtab & "ENABLE MSP BACKUP LSV : " & strLSV
      objLOG.write vbnewline & vbtab & "ENABLE MSP BACKUP LSV : " & strLSV
    end if
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
    ''INSTALL MSP BACKUP VIRTUAL DRIVE
    if (strMSPVD = vbnullstring) then
      objOUT.write vbnewline & vbtab & "INSTALL MSP BACKUP VIRTUAL DRIVE (Y / N)?"
      objLOG.write vbnewline & vbtab & "INSTALL MSP BACKUP VIRTUAL DRIVE (Y / N)?"
      strSEL = objIN.readline
      strMSPVD = strSEL
      if (ucase(strSEL) = "Y") then
        objOUT.write vbnewline & vbtab & "ENTER MSP BACKUP VIRTUAL DRIVE URL : "
        objLOG.write vbnewline & vbtab & "ENTER MSP BACKUP VIRTUAL DRIVE URL : "
        strMSPVDdl = objIN.readline
      end if
    elseif (strMSPVD <> vbnullstring) then
      objOUT.write vbnewline & vbtab & "INSTALL MSP BACKUP VIRTUAL DRIVE : " & strMSPVD & " : " & strMSPVDdl
      objLOG.write vbnewline & vbtab & "INSTALL MSP BACKUP VIRTUAL DRIVE : " & strMSPVD & " : " & strMSPVDdl
    end if
    strSEL = vbnullstring
  'elseif (wscript.arguments.count > 0) then           ''SCRIPT PRE-CONFIGURED
    ''PLACEHOLDER
  'end if
  ''EXECUTION CHECK
  if (lcase(strRUN) = "false") then
    objOUT.write vbnewline & vbnewline & vbtab & "EXECUTE SCRIPT WITH CONFIGURED SETTINGS (Y / N)"
    objLOG.write vbnewline & vbnewline & vbtab & "EXECUTE SCRIPT WITH CONFIGURED SETTINGS (Y / N)"
    strSEL = objIN.readline
    if (ucase(strSEL) = "Y") then
      strRUN = "true"
      objOUT.write vbnewline & vbnewline & now & " - EXECUTING AUTO_PLAN SCRIPT" & vbnewline
      objLOG.write vbnewline & vbnewline & now & " - EXECUTING AUTO_PLAN SCRIPT" & vbnewline
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
    objOUT.write vbnewline & vbnewline & now & " - EXECUTION='" & ucase(strRUN) & "', EXECUTING AUTO_PLAN SCRIPT" & vbnewline
    objLOG.write vbnewline & vbnewline & now & " - EXECUTION='" & ucase(strRUN) & "', EXECUTING AUTO_PLAN SCRIPT" & vbnewline
    ''EXIT VERIFY LOOP, RUN SCRIPT EXECUTION
    call EXECUTE()
  end if
end sub

sub EXECUTE()                                          ''CALL HOOK TO EXECUTE SCRIPT CHANGES
  if (strRUN = "true") then
    ''UPDATE RMMTECH USER
    if (strPWD <> vbnullstring) then
      ''CREATE RMMTECH USER
      objOUT.write vbnewline & now & vbtab & " - CREATING RMMTECH USER"
      objLOG.write vbnewline & now & vbtab & " - CREATING RMMTECH USER"
      call HOOK("net user " & chr(34) & "RMMTech" & chr(34) & " " & chr(34) & strPWD & chr(34) & "  /add /active:yes /expires:never /passwordchg:yes /passwordreq:yes /Y")
      ''SET PASSWORD TO NEVER EXPIRE
      objOUT.write vbnewline & now & vbtab & " - SETTING RMMTECH PASSWORD TO NEVER EXPIRE"
      objLOG.write vbnewline & now & vbtab & " - SETTING RMMTECH PASSWORD TO NEVER EXPIRE"
      call HOOK("wmic useraccount where Name='rmmtech' set PasswordExpires=FALSE")
      ''ADD RMMTECH TO LOCAL ADMINISTRATORS GROUP
      objOUT.write vbnewline & now & vbtab & " - ADDING RMMTECH TO LOCAL ADMINISTRATORS GROUP"
      objLOG.write vbnewline & now & vbtab & " - ADDING RMMTECH TO LOCAL ADMINISTRATORS GROUP"
      call HOOK("net localgroup " & chr(34) & "Administrators" & chr(34) & " " & chr(34) & "RMMTech" & chr(34) & " /add")
      ''GRANT 'LOGON AS A SERVICE' TO RMMTECH USER
      objOUT.write vbnewline & now & vbtab & " - GRANT LONGON AS SERVICE : RMMTECH"
      objLOG.write vbnewline & now & vbtab & " - GRANT LONGON AS SERVICE : RMMTECH"
      strORG = "SeServiceLogonRight ="
      strREP = "SeServiceLogonRight = " & "*" & strSID & ","
      ''EXPORT CURRENT SECURITY DATABASE CONFIGS
      call HOOK("secedit /export /cfg c:\temp\config.inf")
      ''READ CURRENT EXPORTED SECURITY DATABASE CONFIGS
      set objSIN = objFSO.opentextfile("c:\temp\config.inf", 1, 1, -1)
      strIN = objSIN.readall
      objSIN.close
      set objSIN = nothing
      ''WRITE SECURITY DATABASE CONFIGS WITH 'SeServiceLogonRight' FOR RMMTECH
      set objSOUT = objFSO.opentextfile("c:\temp\config.inf", 2, 1, -1)
      objSOUT.write (replace(strIN,strORG,strREP))
      objSOUT.close
      set objSOUT = nothing
      ''APPLY NEW SECURITY DATABASE CONFIGS
      call HOOK("secedit /import /db secedit.sdb /cfg c:\temp\config.inf")
      call HOOK("secedit /configure /db secedit.sdb")
      call HOOK("gpupdate /force")
      ''REMOVE TEMP FILES
      'objFSO.deletefile("c:\temp\config.inf") 
      objOUT.write vbnewline & now & vbtab & " - LOGON AS SERVICE GRANTED : RMMTECH"
      objLOG.write vbnewline & now & vbtab & " - LOGON AS SERVICE GRANTED : RMMTECH"    
    end if
    ''DOWNLOAD AND INSTALL WINDOWS AGENT
    if (strAGTdl <> vbnullstring) then
      objOUT.write vbnewline & vbnewline & now & vbtab & " - DOWNLOADING WINDOWS AGENT"
      objLOG.write vbnewline & vbnewline & now & vbtab & " - DOWNLOADING WINDOWS AGENT"
      call FILEDL(strAGTdl,"WindowsAgent.exe")
      ''STEP TO VERIFY WINDOWS AGENT / MONITORING
      objOUT.write vbnewline & vbnewline & now & vbtab & " - PLEASE VERIFY IMPORT / WINDOWS AGENT MONITORING AND SET APPROPRIATE LICENSE"
      objOUT.write vbnewline & now & vbtab & " - VIA N-CENTRAL>CUSTOMER>ALL DEVICES. THIS NEW DEVICE SHOULD SHOW UP WITHIN MINUTES AS THE AGENT CHECKS IN"
      objOUT.write vbnewline & now & vbtab & " - VIA N-CENTRAL>CUSTOMER>DEVICE DETAILS>MONITORING. RE-APPLY '<DEVICE CLASS> - WINDOWS' TEMPLATE IF NEEDED"
      objOUT.write vbnewline & now & vbtab & " - PRESS 'ENTER' WHEN READY"
      objLOG.write vbnewline & vbnewline & now & vbtab & " - PLEASE VERIFY IMPORT / WINDOWS AGENT MONITORING AND SET APPROPRIATE LICENSE"
      objLOG.write vbnewline & now & vbtab & " - VIA N-CENTRAL>CUSTOMER>ALL DEVICES. THIS NEW DEVICE SHOULD SHOW UP WITHIN MINUTES AS THE AGENT CHECKS IN"
      objLOG.write vbnewline & now & vbtab & " - VIA N-CENTRAL>CUSTOMER>DEVICE DETAILS>MONITORING. RE-APPLY '<DEVICE CLASS> - WINDOWS' TEMPLATE IF NEEDED"
      objLOG.write vbnewline & now & vbtab & " - PRESS 'ENTER' WHEN READY"
      strNUL = objIN.readline
    end if
    ''DOWNLOAD AND INSTALL AV DEFENDER
    if (strAVDdl <> vbnullstring) then
      objOUT.write vbnewline & vbnewline & now & vbtab & " - DOWNLOADING AV DEFENDER"
      objOUT.write vbnewline & now & vbtab & " - ONCE INSTALLED, PLEASE ENABLE AV DEFENDER VIA N-CENTRAL>DEVICE DETAILS>SETTINGS>SECURITY MANAGER"
      objLOG.write vbnewline & vbnewline & now & vbtab & " - DOWNLOADING AV DEFENDER"
      objLOG.write vbnewline & now & vbtab & " - ONCE INSTALLED, PLEASE ENABLE AV DEFENDER VIA N-CENTRAL>DEVICE DETAILS>SETTINGS>SECURITY MANAGER"
      call FILEDL(strAVDdl,"AVDefender.exe")
      ''STEP TO VERIFY AV DEFENDER / MONITORING
      objOUT.write vbnewline & vbnewline & now & vbtab & " - PLEASE VERIFY AV DEFENDER / MONITORING"
      objOUT.write vbnewline & now & vbtab & " - VIA N-CENTRAL>CUSTOMER>DEVICE DETAILS>MONITORING. RE-APPLY 'AV DEFENDER REQUIRED SERVICES - <DEVICE CLASS>' TEMPLATE IF NEEDED"
      objOUT.write vbnewline & now & vbtab & " - PRESS 'ENTER' WHEN READY"
      objLOG.write vbnewline & vbnewline & now & vbtab & " - PLEASE VERIFY AV DEFENDER / MONITORING"
      objLOG.write vbnewline & now & vbtab & " - VIA N-CENTRAL>CUSTOMER>DEVICE DETAILS>MONITORING. RE-APPLY 'AV DEFENDER REQUIRED SERVICES - <DEVICE CLASS>' TEMPLATE IF NEEDED"
      objLOG.write vbnewline & now & vbtab & " - PRESS 'ENTER' WHEN READY"
      strNUL = objIN.readline
    end if
    ''PAUSE TO ENABLE PATCHING
    if (lcase(strUPD) = "y") then
      objOUT.write vbnewline & vbnewline & now & vbtab & " - PLEASE ENABLE PATCHING VIA N-CENTRAL>DEVICE DETAILS>SETTINGS>PATCH MANAGEMENT"
      objOUT.write vbnewline & now & vbtab & " - PRESS 'ENTER' WHEN READY"
      objLOG.write vbnewline & vbnewline & now & vbtab & " - PLEASE ENABLE PATCHING VIA N-CENTRAL>DEVICE DETAILS>SETTINGS>PATCH MANAGEMENT"
      objLOG.write vbnewline & now & vbtab & " - PRESS 'ENTER' WHEN READY"
      strNUL = objIN.readline
      ''STEP TO VERIFY PATCHING / MONITORING
      objOUT.write vbnewline & vbnewline & now & vbtab & " - PLEASE VERIFY PATCHING / MONITORING"
      objOUT.write vbnewline & now & vbtab & " - VIA N-CENTRAL>CUSTOMER>DEVICE DETAILS>MONITORING. RE-APPLY 'PATCH MANAGEMENT - <DEVICE CLASS>' TEMPLATE IF NEEDED"
      objOUT.write vbnewline & now & vbtab & " - PRESS 'ENTER' WHEN READY"
      objLOG.write vbnewline & vbnewline & now & vbtab & " - PLEASE VERIFY PATCHING / MONITORING"
      objLOG.write vbnewline & now & vbtab & " - VIA N-CENTRAL>CUSTOMER>DEVICE DETAILS>MONITORING. RE-APPLY 'PATCH MANAGEMENT - <DEVICE CLASS>' TEMPLATE IF NEEDED"
      objLOG.write vbnewline & now & vbtab & " - PRESS 'ENTER' WHEN READY"
      strNUL = objIN.readline
    end if
    ''DOWNLOAD AND INSTALL WINDOWS PROBE
    if (strPRBdl <> vbnullstring) then
      objOUT.write vbnewline & vbnewline & now & vbtab & " - DOWNLOADING WINDOWS PROBE"
      objLOG.write vbnewline & vbnewline & now & vbtab & " - DOWNLOADING WINDOWS PROBE"
      call FILEDL(strPRBdl,"WindowsProbe.exe")
      ''STEP TO VERIFY WINDOWS PROBE / MONITORING
      objOUT.write vbnewline & vbnewline & now & vbtab & " - PLEASE VERIFY WINDOWS PROBE / MONITORING"
      objOUT.write vbnewline & now & vbtab & " - VIA N-CENTRAL>CUSTOMER>DEVICE DETAILS>MONITORING. RE-APPLY 'WINDOWS PROBE - <DEVICE CLASS>' TEMPLATE IF NEEDED"
      objOUT.write vbnewline & now & vbtab & " - PRESS 'ENTER' WHEN READY"
      objLOG.write vbnewline & vbnewline & now & vbtab & " - PLEASE VERIFY WINDOWS PROBE / MONITORING"
      objLOG.write vbnewline & now & vbtab & " - VIA N-CENTRAL>CUSTOMER>DEVICE DETAILS>MONITORING. RE-APPLY 'WINDOWS PROBE - <DEVICE CLASS>' TEMPLATE IF NEEDED"
      objLOG.write vbnewline & now & vbtab & " - PRESS 'ENTER' WHEN READY"
      strNUL = objIN.readline
    end if
    ''DOWNLOAD AND INSTALL MSP BACKUP
    if (strMSPdl <> vbnullstring) then
      objOUT.write vbnewline & vbnewline & now & vbtab & " - DOWNLOADING MSP BACKUP"
      objOUT.write vbnewline & now & vbtab & " - ONCE INSTALLED, PLEASE ENABLE MSP BACKUP VIA N-CENTRAL>DEVICE DETAILS>SETTINGS>BACKUP MANAGEMENT"
      objLOG.write vbnewline & vbnewline & now & vbtab & " - DOWNLOADING MSP BACKUP"
      objLOG.write vbnewline & now & vbtab & " - ONCE INSTALLED, PLEASE ENABLE MSP BACKUP VIA N-CENTRAL>DEVICE DETAILS>SETTINGS>BACKUP MANAGEMENT"
      call FILEDL(strMSPdl, "MSPBackup.exe")
      ''STEP TO VERIFY MSP BACKUP / MONITORING
      objOUT.write vbnewline & vbnewline & now & vbtab & " - PLEASE VERIFY MSP BACKUP / MONITORING"
      objOUT.write vbnewline & now & vbtab & " - VIA N-CENTRAL>CUSTOMER>DEVICE DETAILS>MONITORING. RE-APPLY 'MSP BAKCUP MANAGER - <DEVICE CLASS>' TEMPLATE IF NEEDED"
      objOUT.write vbnewline & now & vbtab & " - PRESS 'ENTER' WHEN READY"
      objLOG.write vbnewline & vbnewline & now & vbtab & " - PLEASE VERIFY MSP BACKUP / MONITORING"
      objLOG.write vbnewline & now & vbtab & " - VIA N-CENTRAL>CUSTOMER>DEVICE DETAILS>MONITORING. RE-APPLY 'MSP BACKUP MANAGER - <DEVICE CLASS>' TEMPLATE IF NEEDED"
      objLOG.write vbnewline & now & vbtab & " - PRESS 'ENTER' WHEN READY"
      strNUL = objIN.readline
    end if
    ''PAUSE TO ENABLE MSP BACKUP LOCAL SPEEDVAULT
    if (lcase(strLSV) = "y") then
      ''STEP TO CONNECT LOCAL SPEEDVAULT DRIVE
      objOUT.write vbnewline & vbnewline & now & vbtab & " - PLEASE CONNECT MSP BACKUP LOCAL SPEEDVAULT DRIVE"
      objOUT.write vbnewline & now & vbtab & " - PRESS 'ENTER' WHEN READY"
      objLOG.write vbnewline & vbnewline & now & vbtab & " - PLEASE CONNECT MSP BACKUP LOCAL SPEEDVAULT DRIVE"
      objLOG.write vbnewline & now & vbtab & " - PRESS 'ENTER' WHEN READY"
      strNUL = objIN.readline
      ''REQUEST LOCAL SPEEDVAULT PATH
      objOUT.write vbnewline & vbnewline & now & vbtab & " - ENTER LOCAL SPEEDVAULT PATH : "
      objLOG.write vbnewline & vbnewline & now & vbtab & " - ENTER LOCAL SPEEDVAULT PATH : "
      strLSVL = objIN.readline
      ''REQUEST RMMTECH CREDENTIALS FOR LOCAL SPEEDVAULT
      objOUT.write vbnewline & vbnewline & now & vbtab & " - ENTER RMMTECH USERNAME FOR LSV ACCESS : "
      objLOG.write vbnewline & vbnewline & now & vbtab & " - ENTER RMMTECH USERNAME FOR LSV ACCESS : "
      strLSVU = objIN.readline
      objOUT.write vbnewline & vbnewline & now & vbtab & " - ENTER RMMTECH PASSWORD FOR LSV ACCESS : "
      objLOG.write vbnewline & vbnewline & now & vbtab & " - ENTER RMMTECH PASSWORD FOR LSV ACCESS : "
      strLSVP = objIN.readline
      ''SET MSP BACKUP LOCAL SPEEDVAULT SETTINGS
      objOUT.write vbnewline & vbnewline & now & vbtab & " - APPLYING MSP BACKUP LOCAL SPEEDVAULT SETTINGS"
      objLOG.write vbnewline & vbnewline & now & vbtab & " - APPLYING MSP BACKUP LOCAL SPEEDVAULT SETTINGS"
      call HOOK("C:\Program Files\Backup Manager>ClientTool.exe control.setting.modify -name LocalSpeedVaultEnabled -value 1 -name LocalSpeedVaultLocation -value " & _
        chr(34) & strLSVL & chr(34) & " -name LocalSpeedVaultPassword -value " & chr(34) & strLSVP & chr(34) & " -name LocalSpeedVaultUser -value " & chr(34) & strLSVU & chr(34))
      ''STEP TO VERIFY LOCAL SPEEDVAULT SETTINGS / MONITORING
      objOUT.write vbnewline & vbnewline & now & vbtab & " - PLEASE VERIFY MSP BACKUP LOCAL SPEEDVAULT SETTINGS"
      objOUT.write vbnewline & now & vbtab & " - VIA N-CENTRAL>CUSTOMER>DEVICE DETAILS>MONITORING. RE-APPLY 'MSP BAKCUP MANAGER - <DEVICE CLASS>' TEMPLATE IF NEEDED"
      objOUT.write vbnewline & now & vbatb & " - VIA N-CENTRAL>CONFIGURATION>BACKUP MANAGER>MSP BACKUPS>DASHBOARD>LAUNCH BACKUP CLIENT>PREFERENCES>LOCAL SPEEDVAULT"
      objOUT.write vbnewline & now & vbtab & " - PRESS 'ENTER' WHEN READY"
      objLOG.write vbnewline & vbnewline & now & vbtab & " - PLEASE VERIFY MSP BACKUP LOCAL SPEEDVAULT SETTINGS"
      objLOG.write vbnewline & now & vbtab & " - VIA N-CENTRAL>CUSTOMER>DEVICE DETAILS>MONITORING. RE-APPLY 'MSP BAKCUP MANAGER - <DEVICE CLASS>' TEMPLATE IF NEEDED"
      objLOG.write vbnewline & now & vbtab & " - VIA N-CENTRAL>CONFIGURATION>BACKUP MANAGER>MSP BACKUPS>DASHBOARD>LAUNCH BACKUP CLIENT>PREFERENCES>LOCAL SPEEDVAULT"
      objLOG.write vbnewline & now & vbtab & " - PRESS 'ENTER' WHEN READY"
      strNUL = objIN.readline
    end if
    ''PAUSE TO ENABLE MSP BACKUP ARCHIVES
    if (lcase(strARC) = "y") then
      ''SET DEFAULT 'CW_DEFAULT_MSPARCHIVE" ARCHIVING SCHEDULE
      objOUT.write vbnewline & vbnewline & now & vbtab & " - APPLYING MSP BACKUP ARCHIVE SCHEDULE 'CW_DEFAULT_MSPARCHIVE'"
      objLOG.write vbnewline & vbnewline & now & vbtab & " - APPLYING MSP BACKUP ARCHIVE SCHEDULE 'CW_DEFAULT_MSPARCHIVE'"
      call HOOK("C:\Program Files\Backup Manager\ClientTool.exe control.archiving.add -name " & chr(34) & "CW_DEFAULT_MSPARCHIVE" & chr(34) & _
        " -active 1 -datasources All -days-of-month 1,16,Last -months All -time 22:00")
      ''STEP TO VERIFY MSP BACKUP ARCHIVING SCHEDULE
      objOUT.write vbnewline & vbnewline & now & vbtab & " - PLEASE VERIFY MSP BACKUP ARCHIVE SCHEDULE 'CW_DEFAULT_MSPARCHIVE'"
      objOUT.write vbnewline & now & vbtab & " - VIA N-CENTRAL>CONFIGURATION>BACKUP MANAGER>MSP BACKUPS>DASHBOARD>LAUNCH BACKUP CLIENT>PREFERENCES>ARCHIVING"
      objOUT.write vbnewline & now & vbtab & " - PRESS 'ENTER' WHEN READY"
      objLOG.write vbnewline & vbnewline & now & vbtab & " - PLEASE VERIFY MSP BACKUP ARCHIVE SCHEDULE 'CW_DEFAULT_MSPARCHIVE'"
      objLOG.write vbnewline & now & vbtab & " - VIA N-CENTRAL>CONFIGURATION>BACKUP MANAGER>MSP BACKUPS>DASHBOARD>LAUNCH BACKUP CLIENT>PREFERENCES>ARCHIVING"
      objLOG.write vbnewline & now & vbtab & " - PRESS 'ENTER' WHEN READY"
      strNUL = objIN.readline
    end if
    ''DOWNLOAD AND INSTALL MSP BACKUP VIRTUAL DRIVE
    if (strMSPVDdl <> vbnullstring) then
      objOUT.write vbnewline & vbnewline & now & vbtab & " - DOWNLOADING MSP BACKUP VIRTUAL DRIVE"
      objOUT.write vbnewline & now & vbtab & " - ONCE INSTALLED PLEASE EDUCATE CUSTOMER"
      objLOG.write vbnewline & vbnewline & now & vbtab & " - DOWNLOADING MSP BACKUP VIRTUAL DRIVE"
      objLOG.write vbnewline & now & vbtab & " - ONCE INSTALLED PLEASE EDUCATE CUSTOMER"
      call FILEDL(strMSPVDdl, "MSPBackupVD.exe")
    end if
  end if
end sub

sub FILEDL(strURL, strFILE)                           ''CALL HOOK TO DOWNLOAD FILE FROM URL
  strSAV = vbnullstring
  ''SET DOWNLOAD PATH
  strSAV = strPATH & strFILE
  objOUT.write vbnewline & now & vbtab & vbtab & "HTTPDOWNLOAD-------------DOWNLOAD : " & strURL & " : SAVE AS :  " & strSAV
  objLOG.write vbnewline & now & vbtab & vbtab & "HTTPDOWNLOAD-------------DOWNLOAD : " & strURL & " : SAVE AS :  " & strSAV
  ''CREATE HTTP OBJECT
  set objHTTP = createobject( "WinHttp.WinHttpRequest.5.1" )
  ''DOWNLOAD FROM URL
  objHTTP.open "GET", strURL, false
  objHTTP.send
  ''SET N-CENTRAL CREDENTIALS, IF NEEDED
  if ((instr(1, strURL, "ilmcw.dyndns.biz") or instr(1, strURL, "sis.n-able.com"))) then
    if ((strNUSR = vbnullstring) or (strNPWD = vbnullstring)) then
      objOUT.write vbnewline & vbnewline & vbtab & vbtab & "N-CENTRAL ADDRESS, PLEASE ENTER N-CENTRAL USERNAME : "
      objLOG.write vbnewline & vbnewline & vbtab & vbtab & "N-CENTRAL ADDRESS, PLEASE ENTER N-CENTRAL USERNAME : "
      strNUSR = objIN.readline
      objOUT.write vbnewline & vbtab & vbtab & "N-CENTRAL ADDRESS, PLEASE ENTER N-CENTRAL PASSWORD : "
      objLOG.write vbnewline & vbtab & vbtab & "N-CENTRAL ADDRESS, PLEASE ENTER N-CENTRAL PASSWORD : "
      strNPWD = objIN.readline
    end if
    if ((strNUSR <> vbnullstring) and (strNPWD <> vbnullstring)) then
      objHTTP.setcredentials strNUSR, strNPWD, 0
    end if
  end if
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
    ''EXECUTE DOWNLOADED INSTALL
    objOUT.write vbnewline & now & vbtab & " - INSTALLING : " & strSAV
    objLOG.write vbnewline & now & vbtab & " - INSTALLING : " & strSAV
    call HOOK(strSAV)
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
  objOUT.write vbnewline & vbnewline & now & " - AUTO_PLAN COMPLETE. PLEASE VERIFY ALL MONITORING AND SERVICES!" & vbnewline
  objLOG.write vbnewline & vbnewline & now & " - AUTO_PLAN COMPLETE. PLEASE VERIFY ALL MONITORING AND SERVICES!" & vbnewline
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
''AUTO_PLANv2.VBS
''DESIGNED TO AUTOMATE PROTECTION PLAN SETUP
''RUNS CUSTOMIZED MODULE "STAGES" REPRESENTING SECTIONS OF PLAN SETUP
''RUN ON LOCAL DEVICE WITH ADMINISTRATIVE PRIVILEGES
''COMPUTER RENAME WILL REQUIRE REBOOT AND RE-RUN OF SCRIPT
''CURRENTLY ONLY CREATES / UPDATES LOCAL RMMTECH USER
''WRITTEN BY : CJ BLEDSOE / CJ<@>THECOMPUTERWARRIORS.COM
on error resume next
''SCRIPT VARIABLES
dim errRET, strVER, strSEL
''VARIABLES ACCEPTING PARAMETERS - CONFIGURES SNMP TRAP AND COMMUNITY STRING
dim strSNMP, strTRP
''SCRIPT OBJECTS
dim objLOG, objHOOK, objHTTP, objXML
dim objIN, objOUT, objARG, objWSH, objFSO
''VERSION FOR SCRIPT UPDATE, AUTO_PLAN.VBS, REF #2 , REF #6 , FIXES #5 , FIXES #7
strVER = 6
''DEFAULT SUCCESS
errRET = 0
''STDIN / STDOUT
set objIN = wscript.stdin
set objOUT = wscript.stdout
set objARG = wscript.arguments
''OBJECTS FOR LOCATING FOLDERS
set objWSH = createobject("wscript.shell")
set objFSO = createobject("scripting.filesystemobject")
''CONNECT TO WMI REGISTRY PROVIDER
strCOMP = "."
Set objWMI = createobject("winmgmts:{impersonationLevel=impersonate}!\\" & strCOMP & "\root\cimv2")
''PREPARE LOGFILE
if (objFSO.fileexists("C:\temp\auto_planv2")) then          ''LOGFILE EXISTS
  objFSO.deletefile "C:\temp\auto_planv2", true
  set objLOG = objFSO.createtextfile("C:\temp\auto_planv2")
  objLOG.close
  set objLOG = objFSO.opentextfile("C:\temp\auto_planv2", 8)
else                                                  	    ''LOGFILE NEEDS TO BE CREATED
  set objLOG = objFSO.createtextfile("C:\temp\auto_planv2")
  objLOG.close
  set objLOG = objFSO.opentextfile("C:\temp\auto_planv2", 8)
end if
''READ PASSED COMMANDLINE ARGUMENTS - REQUIRES (AT LEAST) 2 ARGUMENTS
if (wscript.arguments.count > 0) then                 	    ''ARGUMENTS WERE PASSED
  for x = 0 to (wscript.arguments.count - 1)
    objOUT.write vbnewline & now & vbtab & "ARGUMENT " & (x + 1) & " (ITEM " & x & ") " & " PASSED : " & ucase(objARG.item(x))
  next
  if (wscript.arguments.count > 1) then
  else
  end if
else
end if

''------------
''BEGIN SCRIPT
objOUT.write vbnewline & vbnewline & now & " - STARTING AUTO_PLANv2" & vbnewline
objLOG.write vbnewline & vbnewline & now & " - STARTING AUTO_PLANv2" & vbnewline
''AUTOMATIC UPDATE, AUTO_PLAN.VBS, REF #2 , REF #6 , FIXES #5
call CHKAU()
''PRE-MATURE END SCRIPT, TESTING AUTOMATIC UPDATE AUTO_PLAN.VBS, REF #2
'call CLEANUP()
''------------
''MAINLOOP - LOOPED ENTRY PERMITS SELECTING OPTIONS IN SEQUENCE
''ALLOWS FOR STAGES TO BE SELECTED IN A LOOP
''ENTERING 'Q' OR 'QUIT' WILL END THE LOOP
''SELECTING STAGE2 WILL FORCE REBOOT AND SCRIPT WILL NEED TO BE RUN AGAIN
blnEND = false
while (blnEND = false)
  strSEL = vbnullstring
  objOUT.write vbnewline & now & vbtab & " - SELECT WHICH STAGE TO RUN"
  objLOG.write vbnewline & now & vbtab & " - SELECT WHICH STAGE TO RUN"
  objOUT.write vbnewline & vbtab & vbtab & " - (1)STAGE1 - SET HIGH PERF. POWER PLAN" & vbnewline & vbtab & vbtab & " - (2)STAGE2 - RENAME COMPUTER (RESTART REQ.)"
  objLOG.write vbnewline & vbtab & vbtab & " - (1)STAGE1 - SET HIGH PERF. POWER PLAN" & vbnewline & vbtab & vbtab & " - (2)STAGE2 - RENAME COMPUTER (RESTART REQ.)"
  objOUT.write vbnewline & vbtab & vbtab & " - (3)STAGE3 - LOCAL RMMTECH (PWD & SVC LOGON)" & vbnewline & vbtab & vbtab & " - (4)STAGE4 - INSTALL & CONFIGURE SNMP"
  objLOG.write vbnewline & vbtab & vbtab & " - (3)STAGE3 - LOCAL RMMTECH (PWD & SVC LOGON)" & vbnewline & vbtab & vbtab & " - (4)STAGE4 - INSTALL & CONFIGURE SNMP"
  objOUT.write vbnewline & vbtab & vbtab & " - (5)STAGE5 - SETUP WINDOWS AGENT" & vbnewline & vbtab & vbtab & " - (6)STAGE6 - SETUP WINDOWS PROBE"
  objLOG.write vbnewline & vbtab & vbtab & " - (5)STAGE5 - SETUP WINDOWS AGENT" & vbnewline & vbtab & vbtab & " - (6)STAGE6 - SETUP WINDOWS PROBE"
  objOUT.write vbnewline & vbtab & vbtab & " - (7)STAGE7 - AV MONITORING AND SERVICES" & vbnewline & vbtab & vbtab & " - (8)STAGE8 - PATCHING MONITORING AND SERVICES"
  objLOG.write vbnewline & vbtab & vbtab & " - (7)STAGE7 - AV MONITORING AND SERVICES" & vbnewline & vbtab & vbtab & " - (8)STAGE8 - PATCHING MONITORING AND SERVICES"
  objOUT.write vbnewline & vbtab & vbtab & " - (9)STAGE9 - BACKUP MONITORING AND SERVICES"
  objLOG.write vbnewline & vbtab & vbtab & " - (9)STAGE9 - BACKUP MONITORING AND SERVICES"
  objOUT.write vbnewline & vbtab & vbtab & " - (Q)QUIT - END SCRIPT" & vbnewline & vbtab & vbtab
  objLOG.write vbnewline & vbtab & vbtab & " - (Q)QUIT - END SCRIPT" & vbnewline & vbtab & vbtab
  strSEL = objIN.readline
  ''CHECK FOR QUIT
  if ((lcase(strSEL) = "q") or (lcase(strSEL) = "quit")) then
    strSEL = 10
    blnEND = true
  end if
  select case strSEL
    case 1
      call STAGE1()
    case 2
      call STAGE2()
    case 3
      call STAGE3()
    case 4
      call STAGE4()
    case 5
      call STAGE5()
    case 6
      call STAGE6()
    case 7
      call STAGE7()
    case 8
      call STAGE8()
    case 9
      call STAGE9()
  end select
  strSEL = vbnullstring
wend
''END MAINLOOP
''------------
''CLEANUP
call CLEANUP()
''END SCRIPT
''------------

''SUB-ROUTINES
sub STAGE1()
  ''------------REF #6
  ''STAGE1
  ''CHANGE ACTIVE POWER PLAN
  objOUT.write vbnewline & now & vbtab & " - SETTING ACTIVE POWER PLAN : HIGH PERFORMANCE" & vbnewline
  objLOG.write vbnewline & now & vbtab & " - SETTING ACTIVE POWER PLAN : HIGH PERFORMANCE" & vbnewline
  call HOOK("powercfg.exe /setactive 8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c")
  ''DISABLE HIBERNATION
  objOUT.write vbnewline & now & vbtab & " - DISABLING HIBERNATION" & vbnewline
  objLOG.write vbnewline & now & vbtab & " - DISABLING HIBERNATION" & vbnewline
  call HOOK("powercfg –h off")
  strSEL = vbnullstring
  if (err.number <> 0) then
    call LOGERR(1)
  end if
end sub

sub STAGE2()
  ''------------REF #6
  ''STAGE2 - REQUIRES TECHNICIAN INPUT / PASSED PARAMETERS
  ''RENAME COMPUTER - REQUIRES RESTART; REQUIRES 'STRNEWPC'; REQUIRES TECHNICIAN INPUT / PASSED PARAMETERS
  objOUT.write vbnewline & vbnewline & vbtab & "RENAME COMPUTER? (WILL REQUIRE RESTART PRIOR TO CONTINUING, Y / N)"
  objLOG.write vbnewline & vbnewline & vbtab & "RENAME COMPUTER? (WILL REQUIRE RESTART PRIOR TO CONTINUING, Y / N)"
  strSEL = objIN.readline
  ''DEFAULT NO REBOOT
  blnRBT = false
  if (ucase(strSEL) = "Y") then
    objOUT.write vbnewline & vbtab & vbtab & "ENTER NEW COMPUTER NAME : " & vbnewline & vbtab & vbtab & _
      "RECOMMENDED FOLLOWING '<CO INITIALS–DEVICE TYPE–NAME>' FORMAT"
    objLOG.write vbnewline & vbtab & vbtab & "ENTER NEW COMPUTER NAME : " & vbnewline & vbtab & vbtab & _
      "RECOMMENDED FOLLOWING '<CO INITIALS–DEVICE TYPE–NAME>' FORMAT"
    strNEWPC = objIN.readline
    if (strNEWPC <> vbnullstring) then
      set colCOMP = objWMI.execquery ("select * from Win32_ComputerSystem")
      for each objCOMP in colCOMP
        intERR = objCOMP.rename(strNEWPC)
        if (intERR <> 0) then
          objOUT.write vbnewline & vbtab & " - ERROR RENAMING COMPUTER : " & strNEWPC & vbnewline & vbtab & _
            "PLEASE RESTART AND TRY AGAIN / CHECK PERMISSIONS"
          objLOG.write vbnewline & vbtab & " - ERROR RENAMING COMPUTER : " & strNEWPC & vbnewline & vbtab & _
            "PLEASE RESTART AND TRY AGAIN / CHECK PERMISSIONS"
        elseif (intERR = 0) then
          objOUT.write vbnewline & vbtab & " - SUCCESSFULLY RENAMED COMPUTER : " & strNEWPC & vbnewline & vbtab & _
            "COMPUTER WILL NOW RESTART, PLEASE RUN SCRIPT AGAIN AND SKIP THIS STEP"
          objLOG.write vbnewline & vbtab & " - SUCCESSFULLY RENAMED COMPUTER : " & strNEWPC & vbnewline & vbtab & _
            "COMPUTER WILL NOW RESTART, PLEASE RUN SCRIPT AGAIN AND SKIP THIS STEP"
          blnRBT = true
        end if
      next
      set objCOMP = nothing
      set colCOMP = nothing
      ''STEP TO VERIFY ADMIN USER CREDENTIALS
      objOUT.write vbnewline & vbnewline & now & vbtab & " - PLEASE VERIFY DEVICE NAME"
      objLOG.write vbnewline & vbnewline & now & vbtab & " - PLEASE VERIFY DEVICE NAME"
      objOUT.write vbnewline & now & vbtab & " - VIA N-CENTRAL>CUSTOMER>ALL DEVICES>DEVICE DETAILS>SETTINGS>PROPERTIES>GIVEN NAME"
      objLOG.write vbnewline & now & vbtab & " - VIA N-CENTRAL>CUSTOMER>ALL DEVICES>DEVICE DETAILS>SETTINGS>PROPERTIES>GIVEN NAME"
      objOUT.write vbnewline & now & vbtab & " - ENTER NEW CONFIGURED DEVICE NAME"
      objLOG.write vbnewline & now & vbtab & " - ENTER NEW CONFIGURED DEVICE NAME"
      objOUT.write vbnewline & now & vbtab & " - PRESS 'ENTER' WHEN READY, DEVICE WILL RESTART AFTERWARDS, PLEASE RE-RUN SCRIPT"
      objLOG.write vbnewline & now & vbtab & " - PRESS 'ENTER' WHEN READY, DEVICE WILL RESTART AFTERWARDS, PLEASE RE-RUN SCRIPT"
      strNUL = objIN.readline
      if (blnRBT) then
        ''RESTART COMPUTER - PROVIDES REASON
        call HOOK("shutdown /r /t 10 /d:p /c " & chr(34) & "AUTO_PANv2 - COMPUTER RENAME : " & strNEWPC & chr(34))
        if (err.number <> 0) then
          call LOGERR(21)
        end if
      end if
    end if
  end if
  ''CLEAR INPUT
  strSEL = vbnullstring
  if (err.number <> 0) then
    call LOGERR(2)
  end if
end sub

sub STAGE3()
  ''------------REF #6
  ''STAGE3 - REQUIRES TECHNICIAN INPUT / PASSED PARAMETERS
  ''UPDATE RMMTECH USER (LOCAL ONLY) - REQUIRES 'STRPWD'; REQUIRES TECHNICIAN INPUT / PASSED PARAMETERS
  if (strPWD = vbnullstring) then
    objOUT.write vbnewline & vbnewline & vbtab & "CREATE AND UPDATE RMMTECH USER (Y / N)?"
    objLOG.write vbnewline & vbnewline & vbtab & "CREATE AND UPDATE RMMTECH USER (Y / N)?"
    strSEL = objIN.readline
    if (ucase(strSEL) = "Y") then
      objOUT.write vbnewline & vbtab & vbtab & "ENTER NEW PASSWORD :"
      objLOG.write vbnewline & vbtab & vbtab & "ENTER NEW PASSWORD :"
      strPWD = objIN.readline
      if (strPWD <> vbnullstring) then
        ''CREATE RMMTECH USER
        objOUT.write vbnewline & now & vbtab & vbtab & " - CREATING RMMTECH USER"
        objLOG.write vbnewline & now & vbtab & vbtab & " - CREATING RMMTECH USER"
        call HOOK("net user " & chr(34) & "RMMTech" & chr(34) & " " & chr(34) & strPWD & chr(34) & _
          "  /add /active:yes /expires:never /passwordchg:yes /passwordreq:yes /Y")
        ''SET PASSWORD TO NEVER EXPIRE
        objOUT.write vbnewline & now & vbtab & vbtab & " - SETTING RMMTECH PASSWORD TO NEVER EXPIRE"
        objLOG.write vbnewline & now & vbtab & vbtab & " - SETTING RMMTECH PASSWORD TO NEVER EXPIRE"
        call HOOK("wmic useraccount where Name='rmmtech' set PasswordExpires=FALSE")
        ''ADD RMMTECH TO LOCAL ADMINISTRATORS GROUP
        objOUT.write vbnewline & now & vbtab & vbtab & " - ADDING RMMTECH TO LOCAL ADMINISTRATORS GROUP"
        objLOG.write vbnewline & now & vbtab & vbtab & " - ADDING RMMTECH TO LOCAL ADMINISTRATORS GROUP"
        call HOOK("net localgroup " & chr(34) & "Administrators" & chr(34) & " " & chr(34) & "RMMTech" & chr(34) & " /add")
        ''GRANT 'LOGON AS A SERVICE' TO RMMTECH USER
        ''GET SIDS OF ALL USERS
        intUSR = 0
        intSID = 0
        objOUT.write vbnewline & now & vbtab & vbtab & " - ENUMERATING USERNAMES AND SIDS"
        objLOG.write vbnewline & now & vbtab & vbtab & " - ENUMERATING USERNAMES AND SIDS"
        set objEXEC = objWSH.exec("wmic useraccount get name,sid /format:csv")
        while (not objEXEC.stdout.atendofstream)
          strIN = objEXEC.stdout.readline
          if ((trim(strIN) <> vbnullstring) and (instr(1, strIN, ","))) then
            if ((trim(split(strIN, ",")(1)) <> vbnullstring) and (trim(split(strIN, ",")(1)) <> "Name")) then
              redim preserve colUSR(intUSR + 1)
              redim preserve colSID(intSID + 1)
              colUSR(intUSR) = trim(split(strIN, ",")(1))
              colSID(intSID) = trim(split(strIN, ",")(2))
              ''SAVE RMMTECH USER SID
              if (lcase(colUSR(intUSR)) = "rmmtech") then 
                strSID = colSID(intUSR)
              end if
              intUSR = (intUSR + 1)
              intSID = (intSID + 1)
            end if
          end if
          if (err.number <> 0) then
            objOUT.write vbnewline & now & vbtab & vbtab & vbtab & err.number & vbtab & err.description
            objLOG.write vbnewline & now & vbtab & vbtab & vbtab & err.number & vbtab & err.description
          end if
        wend
        err.clear
        objOUT.write vbnewline & now & vbtab & vbtab & " - GRANT LONGON AS SERVICE : RMMTECH"
        objLOG.write vbnewline & now & vbtab & vbtab & " - GRANT LONGON AS SERVICE : RMMTECH"
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
        objOUT.write vbnewline & now & vbtab & vbtab & " - LOGON AS SERVICE GRANTED : RMMTECH"
        objLOG.write vbnewline & now & vbtab & vbtab & " - LOGON AS SERVICE GRANTED : RMMTECH"    
      end if
      ''STEP TO VERIFY ADMIN USER CREDENTIALS
      objOUT.write vbnewline & vbnewline & now & vbtab & " - PLEASE VERIFY ADMIN USER CREDENTIALS"
      objLOG.write vbnewline & vbnewline & now & vbtab & " - PLEASE VERIFY ADMIN USER CREDENTIALS"
      objOUT.write vbnewline & now & vbtab & " - VIA N-CENTRAL>CUSTOMER>ADMINISTRATION>DEFAULTS>APPLIANCE SETTINGS>CREDENTIALS"
      objLOG.write vbnewline & now & vbtab & " - VIA N-CENTRAL>CUSTOMER>ADMINISTRATION>DEFAULTS>APPLIANCE SETTINGS>CREDENTIALS"
      objOUT.write vbnewline & now & vbtab & " - ENTER RMMTECH AND RMMTECH PASSWORD, CHECK BOX UNDER 'PROPAGATE'"
      objLOG.write vbnewline & now & vbtab & " - ENTER RMMTECH AND RMMTECH PASSWORD, CHECK BOX UNDER 'PROPAGATE'"
      objOUT.write vbnewline & now & vbtab & " - VIA N-CENTRAL>CUSTOMER>ALL DEVICES>DEVICE DETAILS>SETTINGS>PROPERTIES>CREDENTIALS"
      objOUT.write vbnewline & now & vbtab & " - VIA N-CENTRAL>CUSTOMER>ALL DEVICES>DEVICE DETAILS>SETTINGS>PROPERTIES>CREDENTIALS"
      objOUT.write vbnewline & now & vbtab & " - PRESS 'ENTER' WHEN READY"
      objLOG.write vbnewline & now & vbtab & " - PRESS 'ENTER' WHEN READY"
      strNUL = objIN.readline
    end if
  end if
  strSEL = vbnullstring
  if (err.number <> 0) then
    call LOGERR(3)
  end if
end sub

sub STAGE4()
  ''------------REF #6
  ''STAGE4 - REQUIRES TECHNICIAN INPUT / PASSED PARAMETERS
  ''INSTALL AND CONFIGURE SNMP - REQUIRES 'STRTRP', 'STRSNMP'; REQUIRES TECHNICIAN INPUT / PASSED PARAMETERS
  objOUT.write vbnewline & vbnewline & vbtab & "INSTALLING AND CONFIGURING SNMP"
  objLOG.write vbnewline & vbnewline & vbtab & "INSTALLING AND CONFIGURING SNMP"
  objOUT.write vbnewline & vbtab & vbtab & "ENTER WINDOWS SOFTWARE PROBE IP / SNMP MONITOR AGENT IP (SEPARATE MULTIPLE ENTRIES WITH ',') :"
  objLOG.write vbnewline & vbtab & vbtab & "ENTER WINDOWS SOFTWARE PROBE IP / SNMP MONITOR AGENT IP (SEPARATE MULTIPLE ENTRIES WITH ',') :"
  strTRP = objIN.readline
  objOUT.write vbnewline & vbtab & vbtab & "ENTER SNMP COMMUNITY STRING (DO NOT USE 'PUBLIC') :"
  objLOG.write vbnewline & vbtab & vbtab & "ENTER SNMP COMMUNITY STRING (DO NOT USE 'PUBLIC') :"
  strSNMP = objIN.readline
  if ((strTRP <> vbnullstring) and (strSNMP <> vbnullstring)) then
    ''DOWNLOAD SNMP SETUP : SNMPPARAM, REF #6 , FIXES #15
    objOUT.write vbnewline & now & vbtab & vbtab & " - DOWNLOADING SNMP SETUP : SNMPPARAM"
    objLOG.write vbnewline & now & vbtab & vbtab & " - DOWNLOADING SNMP SETUP : SNMPPARAM"
    call FILEDL("https://github.com/CW-Khristos/scripts/raw/master/SNMPparam.vbs", "SNMPparam.vbs")
    if (objFSO.fileexists("c:\temp\SNMPparam.vbs")) then
      ''INSTALL SNMP VIA SNMPPARAM , REF #6 , FIXES #15
      objOUT.write vbnewline & now & vbtab & vbtab & " - SNMP SETUP : SNMPPARAM"
      objLOG.write vbnewline & now & vbtab & vbtab & " - SNMP SETUP : SNMPPARAM"
      call HOOK("cscript.exe //nologo " & chr(34) & "c:\temp\SNMPparam.vbs" & chr(34) & " " & chr(34) & "modify" & chr(34) & _
        " " & chr(34) & strSNMP & chr(34) & " " & chr(34) & strTRP & chr(34))
    elseif (not objFSO.fileexists("c:\temp\SNMPparam.vbs")) then
      objOUT.write vbnewline & now & vbtab & vbtab & " - DOWNLOAD UNSUCCESSFUL"
      objLOG.write vbnewline & now & vbtab & vbtab & " - DOWNLOAD UNSUCCESSFUL"
      call LOGERR(41)
    end if
    objOUT.write vbnewline & now & vbtab & "SNMP CONFIGURATIONS COMPLETED"
    objLOG.write vbnewline & now & vbtab & "SNMP CONFIGURATIONS COMPLETED"
    objOUT.write vbnewline & now & vbtab & "PLEASE REVIEW SNMP CONFIGURATIONS :"
    objLOG.write vbnewline & now & vbtab & "PLEASE REVIEW SNMP CONFIGURATIONS :"    
    call HOOK("reg query " & chr(34) & "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\services\SNMP\Parameters" & chr(34) & " /s")
    if (err.number <> 0) then
      call LOGERR(42)
    end if 
    ''STEP TO VERIFY SNMP MONITORING
    objOUT.write vbnewline & vbnewline & now & vbtab & " - PLEASE VERIFY SNMP MONITORING"
    objLOG.write vbnewline & vbnewline & now & vbtab & " - PLEASE VERIFY IMPORT / WINDOWS AGENT MONITORING"
    objOUT.write vbnewline & now & vbtab & " - VIA N-CENTRAL>CUSTOMER>ADMINISTRATION>DEFAULTS>BACKUP AND SNMP DEFAULTS"
    objLOG.write vbnewline & now & vbtab & " - VIA N-CENTRAL>CUSTOMER>ADMINISTRATION>DEFAULTS>BACKUP AND SNMP DEFAULTS"
    objOUT.write vbnewline & now & vbtab & " - DISABLE THE 'SYSTEM DEFAULT' SNMP PROFILE, USE 'ADD SNMP CREDENTIALS' TO CREATE NEW SNMPv1 PROFILE, USE CUSTOMER SNMP SETTINGS"
    objLOG.write vbnewline & now & vbtab & " - DISABLE THE 'SYSTEM DEFAULT' SNMP PROFILE, USE 'ADD SNMP CREDENTIALS' TO CREATE NEW SNMPv1 PROFILE, USE CUSTOMER SNMP SETTINGS"
    objOUT.write vbnewline & now & vbtab & " - VIA N-CENTRAL>CUSTOMER>ALL DEVICES>DEVICE DETAILS>SETTINGS>MONITORING OPTIONS CHECK 'USE SNMP', SETTINGS FROM PREVIOUS STEP SHOULD BE POPULATED"
    objLOG.write vbnewline & now & vbtab & " - VIA N-CENTRAL>CUSTOMER>ALL DEVICES>DEVICE DETAILS>SETTINGS>MONITORING OPTIONS CHECK 'USE SNMP', SETTINGS FROM PREVIOUS STEP SHOULD BE POPULATED"
    objOUT.write vbnewline & now & vbtab & " - VIA N-CENTRAL>CUSTOMER>ALL DEVICES>DEVICE DETAILS>MONITORING. RE-APPLY '<DEVICE CLASS> - WINDOWS' TEMPLATE IF NEEDED"
    objOUT.write vbnewline & now & vbtab & " - VIA N-CENTRAL>CUSTOMER>ALL DEVICES>DEVICE DETAILS>MONITORING. RE-APPLY '<DEVICE CLASS> - WINDOWS' TEMPLATE IF NEEDED"
    objOUT.write vbnewline & now & vbtab & " - PRESS 'ENTER' WHEN READY"
    objLOG.write vbnewline & now & vbtab & " - PRESS 'ENTER' WHEN READY"
    strNUL = objIN.readline
  end if
  strSEL = vbnullstring
  if (err.number <> 0) then
    call LOGERR(4)
  end if
end sub

sub STAGE5()
  ''------------REF #6
  ''STAGE5 - REQUIRES TECHNICIAN INPUT / PASSED PARAMETERS
  ''INSTALL WINDOWS AGENT - REQUIRES 'STRCID', 'STRCNAM'; REQUIRES TECHNICIAN INPUT / PASSED PARAMETERS
  if (strAGT = vbnullstring) then
    objOUT.write vbnewline & vbnewline & vbtab & "INSTALL WINDOWS AGENT (Y / N)?"
    objLOG.write vbnewline & vbnewline & vbtab & "INSTALL WINDOWS AGENT (Y / N)?"
    strSEL = objIN.readline
    if (ucase(strSEL) = "Y") then
      ''CUSTOMER ID
      objOUT.write vbnewline & vbtab & vbtab & "ENTER CUSTOMER ID :"
      objLOG.write vbnewline & vbtab & vbtab & "ENTER CUSTOMER ID :"
      strCID = objIN.readline
      ''CUSTOMER NAME
      objOUT.write vbnewline & vbtab & vbtab & "ENTER CUSTOMER NAME :"
      objLOG.write vbnewline & vbtab & vbtab & "ENTER CUSTOMER NAME :"
      strCNAM = objIN.readline
      if ((strCID <> vbnullstring) and (strCNAM <> vbnullstring)) then
        ''DOWNLOAD WINDOWS AGENT SETUP : RE-AGENT , FIXES #7
        objOUT.write vbnewline & now & vbtab & vbtab & " - DOWNLOADING WINDOWS AGENT SETUP : RE-AGENT"
        objLOG.write vbnewline & now & vbtab & vbtab & " - DOWNLOADING WINDOWS AGENT SETUP : RE-AGENT"
        call FILEDL("https://github.com/CW-Khristos/CW_MSI/raw/master/reagent.vbs", "reagent.vbs")
        if (objFSO.fileexists("c:\temp\reagent.vbs")) then
          ''INSTALL WINDOWS AGENT VIA RE-AGENT , FIXES #7
          objOUT.write vbnewline & now & vbtab & vbtab & " - WINDOWS AGENT SETUP : RE-AGENT"
          objLOG.write vbnewline & now & vbtab & vbtab & " - WINDOWS AGENT SETUP : RE-AGENT"
          call HOOK("cscript.exe //nologo " & chr(34) & "c:\temp\reagent.vbs" & chr(34) & " " & chr(34) & strCID & chr(34) & " " & chr(34) & strCNAM & chr(34))
        elseif (not objFSO.fileexists("c:\temp\reprobe.vbs")) then
          objOUT.write vbnewline & now & vbtab & vbtab & " - DOWNLOAD UNSUCCESSFUL"
          objLOG.write vbnewline & now & vbtab & vbtab & " - DOWNLOAD UNSUCCESSFUL"
          call LOGERR(51)
        end if
      end if
      ''STEP TO VERIFY WINDOWS AGENT / MONITORING
      objOUT.write vbnewline & vbnewline & now & vbtab & " - PLEASE VERIFY IMPORT / WINDOWS AGENT MONITORING"
      objLOG.write vbnewline & vbnewline & now & vbtab & " - PLEASE VERIFY IMPORT / WINDOWS AGENT MONITORING"
      objOUT.write vbnewline & now & vbtab & " - VIA N-CENTRAL>CUSTOMER>ALL DEVICES. THIS NEW DEVICE SHOULD SHOW UP WITHIN MINUTES AS THE AGENT CHECKS IN"
      objLOG.write vbnewline & now & vbtab & " - VIA N-CENTRAL>CUSTOMER>ALL DEVICES. THIS NEW DEVICE SHOULD SHOW UP WITHIN MINUTES AS THE AGENT CHECKS IN"
      objOUT.write vbnewline & now & vbtab & " - SET APPROPRIATE LICENSE VIA N-CENTRAL>CUSTOMER>ALL DEVICES>DEVICE DETAILS>SETTINGS>CUSTOM DETAILS>SERVICE CODE CUSTOM DEVICE PROPERTY"
      objLOG.write vbnewline & now & vbtab & " - SET APPROPRIATE LICENSE VIA N-CENTRAL>CUSTOMER>ALL DEVICES>DEVICE DETAILS>SETTINGS>CUSTOM DETAILS>SERVICE CODE CUSTOM DEVICE PROPERTY"
      objOUT.write vbnewline & now & vbtab & " - VIA N-CENTRAL>CUSTOMER>ALL DEVICES>DEVICE DETAILS>MONITORING. RE-APPLY '<DEVICE CLASS> - WINDOWS' TEMPLATE IF NEEDED"
      objLOG.write vbnewline & now & vbtab & " - VIA N-CENTRAL>CUSTOMER>ALL DEVICES>DEVICE DETAILS>MONITORING. RE-APPLY '<DEVICE CLASS> - WINDOWS' TEMPLATE IF NEEDED"
      objOUT.write vbnewline & now & vbtab & " - THIS WILL AUTOMATICALLY ASSIGN AV, PATCHING, AND MSP BACKUPS PER CONTRACT 'SERVICE CODE' SOP"
      objOUT.write vbnewline & now & vbtab & " - THIS WILL AUTOMATICALLY ASSIGN AV, PATCHING, AND MSP BACKUPS PER CONTRACT 'SERVICE CODE' SOP"
      objOUT.write vbnewline & now & vbtab & " - PRESS 'ENTER' WHEN READY"
      objLOG.write vbnewline & now & vbtab & " - PRESS 'ENTER' WHEN READY"
      strNUL = objIN.readline
    end if
  end if
  strSEL = vbnullstring
  strCID = vbnullstring
  strCNAM = vbnullstring
  if (err.number <> 0) then
    call LOGERR(5)
  end if
end sub

sub STAGE6()
  ''------------REF #6
  ''STAGE6 - REQUIRES TECHNICIAN INPUT / PASSED PARAMETERS - NEED DOMAIN / WORKGROUP SELECTION
  ''INSTALL PROBE - REQUIRES 'STRCID', 'STRCNAM', 'STRDMN', 'STRDUSR', 'STRDPWD'; REQUIRES TECHNICIAN INPUT / PASSED PARAMETERS
  if (strPRB = vbnullstring) then
    objOUT.write vbnewline & vbnewline & vbtab & "INSTALL WINDOWS PROBE (Y / N)?"
    objLOG.write vbnewline & vbnewline & vbtab & "INSTALL WINDOWS PROBE (Y / N)?"
    strSEL = objIN.readline
    if (ucase(strSEL) = "Y") then
      ''PROBE TYPE - Workgroup_Windows / Network_Windows
      objOUT.write vbnewline & vbtab & vbtab & "SELECT PROBE TYPE : (1) Workgroup_Windows / (2) Network_Windows :"
      objLOG.write vbnewline & vbtab & vbtab & "SELECT PROBE TYPE : (1) Workgroup_Windows / (2) Network_Windows :"
      strTYP = objIN.readline
      if (strTYP = 1) then
        strTYP = "Workgroup_Windows"
      elseif (strTYP = 2) then
        strTYP = "Network_Windows"
      elseif ((strTYP <> 1) and (strTYP <> 2)) then
        call LOGERR(61)
      end if
      if (errRET <> 61) then
        ''CUSTOMER ID
        objOUT.write vbnewline & vbtab & vbtab & "ENTER CUSTOMER ID :"
        objLOG.write vbnewline & vbtab & vbtab & "ENTER CUSTOMER ID :"
        strCID = objIN.readline
        ''CUSTOMER NAME
        objOUT.write vbnewline & vbtab & vbtab & "ENTER CUSTOMER NAME :"
        objLOG.write vbnewline & vbtab & vbtab & "ENTER CUSTOMER NAME :"
        strCNAM = objIN.readline
        ''DOMAIN
        objOUT.write vbnewline & vbtab & vbtab & "ENTER DOMAIN (DO NOT INCLUDE '\', MAY BE BLANK IF WORKGROUP) :"
        objLOG.write vbnewline & vbtab & vbtab & "ENTER DOMAIN (DO NOT INCLUDE '\', MAY BE BLANK IF WORKGROUP) :"
        strDMN = objIN.readline
        if (instr(1, strDMN, "\")) then
          strDMN = replace(strDMN, "\", vbnullstring)
        end if
        ''DOMAIN USER
        objOUT.write vbnewline & vbtab & vbtab & "ENTER DOMAIN USER :"
        objLOG.write vbnewline & vbtab & vbtab & "ENTER DOMAIN USER :"
        strDUSR = objIN.readline
        ''DOMAIN USER PASSWORD
        objOUT.write vbnewline & vbtab & vbtab & "ENTER DOMAIN USER PASSWORD :"
        objLOG.write vbnewline & vbtab & vbtab & "ENTER DOMAIN USER PASSWORD :"
        strDPWD = objIN.readline
        if ((strCID <> vbnullstring) and (strCNAM <> vbnullstring) and _
          (strDMN <> vbnullstring) and (strDUSR <> vbnullstring) and (strDPWD <> vbnullstring)) then
            ''DOWNLOAD WINDOWS PROBE SETUP : RE-PROBE , FIXES #7
            objOUT.write vbnewline & now & vbtab & vbtab & " - DOWNLOADING WINDOWS PROBE SETUP : RE-PROBE"
            objLOG.write vbnewline & now & vbtab & vbtab & " - DOWNLOADING WINDOWS PROBE SETUP : RE-PROBE"
            call FILEDL("https://github.com/CW-Khristos/CW_MSI/raw/master/reprobe.vbs", "reprobe.vbs")
            if (objFSO.fileexists("c:\temp\reprobe.vbs")) then
              ''INSTALL WINDOWS PROBE VIA RE-PROBE , FIXES #7
              objOUT.write vbnewline & now & vbtab & vbtab & " - WINDOWS PROBE SETUP : RE-PROBE"
              objLOG.write vbnewline & now & vbtab & vbtab & " - WINDOWS PROBE SETUP : RE-PROBE"
              call HOOK("cscript.exe //nologo " & chr(34) & "c:\temp\reprobe.vbs" & chr(34) & " " & chr(34) & strCID & chr(34) & " " & chr(34) & strCNAM & chr(34) & _
                " " & chr(34) & strTYP & chr(34) & " " & chr(34) & strDMN & chr(34) & " " & chr(34) & strDUSR & chr(34) & " " & chr(34) & strDPWD & chr(34))
            elseif (not objFSO.fileexists("c:\temp\reprobe.vbs")) then
              objOUT.write vbnewline & now & vbtab & vbtab & " - DOWNLOAD UNSUCCESSFUL"
              objLOG.write vbnewline & now & vbtab & vbtab & " - DOWNLOAD UNSUCCESSFUL"
              call LOGERR(62)
            end if
        end if
      end if
      ''STEP TO VERIFY WINDOWS PROBE / MONITORING
      objOUT.write vbnewline & vbnewline & now & vbtab & " - PLEASE VERIFY WINDOWS PROBE / MONITORING"
      objLOG.write vbnewline & vbnewline & now & vbtab & " - PLEASE VERIFY WINDOWS PROBE / MONITORING"
      objOUT.write vbnewline & now & vbtab & " - VIA N-CENTRAL>CUSTOMER>ALL DEVICES>DEVICE DETAILS>MONITORING. RE-APPLY 'WINDOWS PROBE - <DEVICE CLASS>' TEMPLATE IF NEEDED"
      objLOG.write vbnewline & now & vbtab & " - VIA N-CENTRAL>CUSTOMER>ALL DEVICES>DEVICE DETAILS>MONITORING. RE-APPLY 'WINDOWS PROBE - <DEVICE CLASS>' TEMPLATE IF NEEDED"
      objOUT.write vbnewline & now & vbtab & " - VIA N-CENTRAL>CUSTOMER>ADMINISTRATION>PROBES, USE 'TRANSFER TASKS' TO ASSIGN DEVICE MONITORING TO PROBE"
      objLOG.write vbnewline & now & vbtab & " - VIA N-CENTRAL>CUSTOMER>ADMINISTRATION>PROBES, USE 'TRANSFER TASKS' TO ASSIGN DEVICE MONITORING TO PROBE"
      objOUT.write vbnewline & now & vbtab & " - PRESS 'ENTER' WHEN READY"
      objLOG.write vbnewline & now & vbtab & " - PRESS 'ENTER' WHEN READY"
      strNUL = objIN.readline
    end if
  end if
  strSEL = vbnullstring
   if (err.number <> 0) then
    call LOGERR(6)
  end if
end sub

sub STAGE7()
  ''------------REF #6
  ''STAGE7 - REQUIRES TECHNICIAN INPUT / PASSED PARAMETERS
  ''INSTALL AV DEFENDER - REQUIRES 'ENTER' TO RESUME AFTER PAUSE
  ''DOWNLOAD AND INSTALL AV DEFENDER, REF #6 , FIXES #14
  objOUT.write vbnewline & vbnewline & now & vbtab & " - DOWNLOADING AV DEFENDER"
  objLOG.write vbnewline & vbnewline & now & vbtab & " - DOWNLOADING AV DEFENDER"
  objOUT.write vbnewline & now & vbtab & " - ONCE INSTALLED, PLEASE ENABLE AV DEFENDER VIA N-CENTRAL>CUSTOMER>ALL DEVICES>DEVICE DETAILS>SETTINGS>SECURITY MANAGER"
  objLOG.write vbnewline & now & vbtab & " - ONCE INSTALLED, PLEASE ENABLE AV DEFENDER VIA N-CENTRAL>CUSTOMER>ALL DEVICES>DEVICE DETAILS>SETTINGS>SECURITY MANAGER"
  'call FILEDL(strAVDdl,"AVDefender.exe")
  ''STEP TO VERIFY AV DEFENDER / MONITORING, REF #6 , FIXES #14
  objOUT.write vbnewline & vbnewline & now & vbtab & " - PLEASE VERIFY AV DEFENDER / MONITORING"
  objLOG.write vbnewline & vbnewline & now & vbtab & " - PLEASE VERIFY AV DEFENDER / MONITORING"
  objOUT.write vbnewline & now & vbtab & " - VIA N-CENTRAL>CUSTOMER>ALL DEVICES>DEVICE DETAILS>MONITORING. RE-APPLY 'AV DEFENDER REQUIRED SERVICES - <DEVICE CLASS>' TEMPLATE IF NEEDED"
  objLOG.write vbnewline & now & vbtab & " - VIA N-CENTRAL>CUSTOMER>ALL DEVICES>DEVICE DETAILS>MONITORING. RE-APPLY 'AV DEFENDER REQUIRED SERVICES - <DEVICE CLASS>' TEMPLATE IF NEEDED"
  objOUT.write vbnewline & now & vbtab & " - PRESS 'ENTER' WHEN READY"
  objLOG.write vbnewline & now & vbtab & " - PRESS 'ENTER' WHEN READY"
  strNUL = objIN.readline
  if (err.number <> 0) then
    call LOGERR(7)
  end if
end sub

sub STAGE8()
  ''------------REF #6
  ''STAGE8 - REQUIRES TECHNICIAN INPUT / PASSED PARAMETERS
  ''ENABLE PATCHING - REQUIRES 'ENTER' TO RESUME AFTER PAUSE
  ''PAUSE TO ENABLE PATCHING, REF #6 , FIXES #13
  objOUT.write vbnewline & vbnewline & now & vbtab & " - PLEASE ENABLE PATCHING VIA N-CENTRAL>CUSTOMER>ALL DEVICES>DEVICE DETAILS>SETTINGS>PATCH MANAGEMENT"
  objLOG.write vbnewline & vbnewline & now & vbtab & " - PLEASE ENABLE PATCHING VIA N-CENTRAL>CUSTOMER>ALL DEVICES>DEVICE DETAILS>SETTINGS>PATCH MANAGEMENT"
  objOUT.write vbnewline & now & vbtab & " - PRESS 'ENTER' WHEN READY"
  objLOG.write vbnewline & now & vbtab & " - PRESS 'ENTER' WHEN READY"
  strNUL = objIN.readline
  ''STEP TO VERIFY PATCHING / MONITORING, REF #6 , FIXES #13
  objOUT.write vbnewline & vbnewline & now & vbtab & " - PLEASE VERIFY PATCHING / MONITORING"
  objLOG.write vbnewline & vbnewline & now & vbtab & " - PLEASE VERIFY PATCHING / MONITORING"
  objOUT.write vbnewline & now & vbtab & " - VIA N-CENTRAL>CUSTOMER>ALL DEVICES>DEVICE DETAILS>MONITORING. RE-APPLY 'PATCH MANAGEMENT - <DEVICE CLASS>' TEMPLATE IF NEEDED"
  objLOG.write vbnewline & now & vbtab & " - VIA N-CENTRAL>CUSTOMER>ALL DEVICES>DEVICE DETAILS>MONITORING. RE-APPLY 'PATCH MANAGEMENT - <DEVICE CLASS>' TEMPLATE IF NEEDED"
  objOUT.write vbnewline & now & vbtab & " - PRESS 'ENTER' WHEN READY"
  objLOG.write vbnewline & now & vbtab & " - PRESS 'ENTER' WHEN READY"
  strNUL = objIN.readline
  if (err.number <> 0) then
    call LOGERR(8)
  end if
end sub

sub STAGE9()
  ''------------REF #6
  ''STAGE9 - REQUIRES TECHNICIAN INPUT / PASSED PARAMETERS
  ''ENABLE MSP BACKUP - REQUIRES 'ENTER' TO RESUME AFTER PAUSE
  ''DOWNLOAD AND INSTALL MSP BACKUP, REF #6 , FIXES #12
  objOUT.write vbnewline & vbnewline & now & vbtab & " - DOWNLOADING MSP BACKUP"
  objLOG.write vbnewline & vbnewline & now & vbtab & " - DOWNLOADING MSP BACKUP"
  objOUT.write vbnewline & now & vbtab & " - ONCE INSTALLED, PLEASE ENABLE MSP BACKUP VIA N-CENTRAL>DEVICE DETAILS>SETTINGS>BACKUP MANAGEMENT"
  objLOG.write vbnewline & now & vbtab & " - ONCE INSTALLED, PLEASE ENABLE MSP BACKUP VIA N-CENTRAL>DEVICE DETAILS>SETTINGS>BACKUP MANAGEMENT"
  'call FILEDL(strMSPdl, "MSPBackup.exe")
  ''STEP TO VERIFY MSP BACKUP / MONITORING, REF #6 , FIXES #12
  objOUT.write vbnewline & vbnewline & now & vbtab & " - PLEASE VERIFY MSP BACKUP / MONITORING"
  objLOG.write vbnewline & vbnewline & now & vbtab & " - PLEASE VERIFY MSP BACKUP / MONITORING"
  objOUT.write vbnewline & now & vbtab & " - VIA N-CENTRAL>CUSTOMER>DEVICE DETAILS>MONITORING. RE-APPLY 'MSP BAKCUP MANAGER - <DEVICE CLASS>' TEMPLATE IF NEEDED"
  objLOG.write vbnewline & now & vbtab & " - VIA N-CENTRAL>CUSTOMER>DEVICE DETAILS>MONITORING. RE-APPLY 'MSP BACKUP MANAGER - <DEVICE CLASS>' TEMPLATE IF NEEDED"
  objOUT.write vbnewline & now & vbtab & " - PRESS 'ENTER' WHEN READY"
  objLOG.write vbnewline & now & vbtab & " - PRESS 'ENTER' WHEN READY"
  strNUL = objIN.readline
  ''PAUSE TO ENABLE MSP BACKUP LOCAL SPEEDVAULT
  objOUT.write vbnewline & vbtab & "ENABLE MSP BACKUP LSV (Y / N)?"
  objLOG.write vbnewline & vbtab & "ENABLE MSP BACKUP LSV (Y / N)?"
  strLSV = objIN.readline
  if (lcase(strLSV) = "y") then
    ''STEP TO CONNECT LOCAL SPEEDVAULT DRIVE
    objOUT.write vbnewline & vbnewline & now & vbtab & " - PLEASE CONNECT MSP BACKUP LOCAL SPEEDVAULT DRIVE"
    objLOG.write vbnewline & vbnewline & now & vbtab & " - PLEASE CONNECT MSP BACKUP LOCAL SPEEDVAULT DRIVE"
    objOUT.write vbnewline & now & vbtab & " - PRESS 'ENTER' WHEN READY"
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
    ''SET MSP BACKUP LOCAL SPEEDVAULT SETTINGS, REF #6 , FIXES #12
    objOUT.write vbnewline & vbnewline & now & vbtab & " - APPLYING MSP BACKUP LOCAL SPEEDVAULT SETTINGS"
    objLOG.write vbnewline & vbnewline & now & vbtab & " - APPLYING MSP BACKUP LOCAL SPEEDVAULT SETTINGS"
    call HOOK("C:\Program Files\Backup Manager\ClientTool.exe control.setting.modify -name LocalSpeedVaultEnabled -value 1 -name LocalSpeedVaultLocation -value " & _
      chr(34) & strLSVL & chr(34) & " -name LocalSpeedVaultPassword -value " & chr(34) & strLSVP & chr(34) & " -name LocalSpeedVaultUser -value " & chr(34) & strLSVU & chr(34))
    if (err.number <> 0) then
      call LOGERR(91)
    end if
    ''RESTRICT MSP BACKUP LOCAL SPEEDVAULT PERMISSIONS
    objOUT.write vbnewline & vbnewline & now & vbtab & " - RESTRICTING MSP BACKUP LOCAL SPEEDVAULT PERMISSIONS"
    objLOG.write vbnewline & vbnewline & now & vbtab & " - RESTRICTING MSP BACKUP LOCAL SPEEDVAULT PERMISSIONS"
    if ((strLSVL <> vbnullstring) and _
      (strLSVU <> vbnullstring) and (strLSVP <> vbnullstring)) then
        ''DOWNLOAD LSV PERMISSIONS SETUP : LSVPERM, REF #6 , FIXES #12
        objOUT.write vbnewline & now & vbtab & vbtab & " - DOWNLOADING LSV PERMISSIONS SETUP : LSVPERM"
        objLOG.write vbnewline & now & vbtab & vbtab & " - DOWNLOADING LSV PERMISSIONS SETUP : LSVPERM"
        call FILEDL("https://github.com/CW-Khristos/scripts/raw/master/MSP Backups/LSVperm.vbs", "LSVperm.vbs")
        if (objFSO.fileexists("c:\temp\LSVperm.vbs")) then
          ''RESTRICT LSV PERMISSIONS VIA LSVPERM, REF #6 , FIXES #12
          objOUT.write vbnewline & now & vbtab & vbtab & " - RESTRICT LSV PERMISSIONS SETUP : LSVPERM"
          objLOG.write vbnewline & now & vbtab & vbtab & " - RESTRICT LSV PERMISSIONS SETUP : LSVPERM"
          call HOOK("cscript.exe //nologo " & chr(34) & "c:\temp\LSVperm.vbs" & chr(34) & " " & chr(34) & strLSVL & chr(34) & " " & chr(34) & strLSVU & chr(34) & _
            " " & chr(34) & strLSVP)
        elseif (not objFSO.fileexists("c:\temp\LSVperm.vbs")) then
          objOUT.write vbnewline & now & vbtab & vbtab & " - DOWNLOAD UNSUCCESSFUL"
          objLOG.write vbnewline & now & vbtab & vbtab & " - DOWNLOAD UNSUCCESSFUL"
          call LOGERR(92)
        end if
        if (err.number <> 0) then
          call LOGERR(93)
        end if
    end if
    ''STEP TO VERIFY LOCAL SPEEDVAULT SETTINGS / MONITORING
    objOUT.write vbnewline & vbnewline & now & vbtab & " - PLEASE VERIFY MSP BACKUP LOCAL SPEEDVAULT SETTINGS"
    objLOG.write vbnewline & vbnewline & now & vbtab & " - PLEASE VERIFY MSP BACKUP LOCAL SPEEDVAULT SETTINGS"
    objOUT.write vbnewline & now & vbtab & " - VIA N-CENTRAL>CUSTOMER>DEVICE DETAILS>MONITORING. RE-APPLY 'MSP BAKCUP MANAGER - <DEVICE CLASS>' TEMPLATE IF NEEDED"
    objLOG.write vbnewline & now & vbtab & " - VIA N-CENTRAL>CUSTOMER>DEVICE DETAILS>MONITORING. RE-APPLY 'MSP BAKCUP MANAGER - <DEVICE CLASS>' TEMPLATE IF NEEDED"
    objOUT.write vbnewline & now & vbatb & " - VIA N-CENTRAL>CUSTOMER>CONFIGURATION>BACKUP MANAGER>MSP BACKUPS>DASHBOARD>LAUNCH BACKUP CLIENT>PREFERENCES>LOCAL SPEEDVAULT"
    objLOG.write vbnewline & now & vbtab & " - VIA N-CENTRAL>CUSTOMER>CONFIGURATION>BACKUP MANAGER>MSP BACKUPS>DASHBOARD>LAUNCH BACKUP CLIENT>PREFERENCES>LOCAL SPEEDVAULT"
    objOUT.write vbnewline & now & vbtab & " - PRESS 'ENTER' WHEN READY"
    objLOG.write vbnewline & now & vbtab & " - PRESS 'ENTER' WHEN READY"
    strNUL = objIN.readline
  end if
  ''PAUSE TO ENABLE MSP BACKUP ARCHIVES, REF #6 , FIXES #12
  objOUT.write vbnewline & vbtab & "ENABLE MSP BACKUP ARCHIVES (Y / N)?"
  objLOG.write vbnewline & vbtab & "ENABLE MSP BACKUP ARCHIVES (Y / N)?"
  strARC = objIN.readline
  if (lcase(strARC) = "y") then
    ''SET DEFAULT 'CW_DEFAULT_MSPARCHIVE" ARCHIVING SCHEDULE, REF #6 , FIXES #12
    objOUT.write vbnewline & vbnewline & now & vbtab & " - APPLYING MSP BACKUP ARCHIVE SCHEDULE 'CW_DEFAULT_MSPARCHIVE'"
    objLOG.write vbnewline & vbnewline & now & vbtab & " - APPLYING MSP BACKUP ARCHIVE SCHEDULE 'CW_DEFAULT_MSPARCHIVE'"
    call HOOK("C:\Program Files\Backup Manager\ClientTool.exe control.archiving.add -name " & chr(34) & "CW_DEFAULT_MSPARCHIVE" & chr(34) & _
      " -active 1 -datasources All -days-of-month 1,15 -months All -time 22:00")
    if (err.number <> 0) then
      call LOGERR(94)
    end if
    ''STEP TO VERIFY MSP BACKUP ARCHIVING SCHEDULE
    objOUT.write vbnewline & vbnewline & now & vbtab & " - PLEASE VERIFY MSP BACKUP ARCHIVE SCHEDULE 'CW_DEFAULT_MSPARCHIVE'"
    objLOG.write vbnewline & vbnewline & now & vbtab & " - PLEASE VERIFY MSP BACKUP ARCHIVE SCHEDULE 'CW_DEFAULT_MSPARCHIVE'"
    objOUT.write vbnewline & now & vbtab & " - VIA N-CENTRAL>CUSTOMER>CONFIGURATION>BACKUP MANAGER>MSP BACKUPS>DASHBOARD>LAUNCH BACKUP CLIENT>PREFERENCES>ARCHIVING"
    objLOG.write vbnewline & now & vbtab & " - VIA N-CENTRAL>CUSTOMER>CONFIGURATION>BACKUP MANAGER>MSP BACKUPS>DASHBOARD>LAUNCH BACKUP CLIENT>PREFERENCES>ARCHIVING"
    objOUT.write vbnewline & now & vbtab & " - PRESS 'ENTER' WHEN READY"
    objLOG.write vbnewline & now & vbtab & " - PRESS 'ENTER' WHEN READY"
    strNUL = objIN.readline
  end if
  ''DOWNLOAD AND INSTALL MSP BACKUP VIRTUAL DRIVE, REF #6 , FIXES #12
  objOUT.write vbnewline & vbtab & "INSTALL MSP BACKUP VIRTUAL DRIVE (THIS WILL USE DRIVE B:) (Y / N)?"
  objLOG.write vbnewline & vbtab & "INSTALL MSP BACKUP VIRTUAL DRIVE (THIS WILL USE DRIVE B:) (Y / N)?"
  strMSPVD = objIN.readline
  if (ucase(strMSPVD) = "Y") then
    objOUT.write vbnewline & vbtab & "SELECT MSP BACKUP VIRTUAL DRIVE URL : (1) - (X86) / (2) - (X64)"
    objLOG.write vbnewline & vbtab & "SELECT MSP BACKUP VIRTUAL DRIVE URL : (1) - (X86) / (2) - (X64)"
    strMSPVDdl = objIN.readline
  end if
  ''SELECT URL
  select case strMSPVDdl
    case 1                    ''(x86) DOWNLOAD URL
      strMSPVDdl = "http://cdn.cloudbackup.management/maxdownloads/mxb-vd-windows-x86.exe"
    case 2                    ''(x64) DOWNLOAD URL
      strMSPVDdl = "http://cdn.cloudbackup.management/maxdownloads/mxb-vd-windows-x64.exe"
  end select
  if (strMSPVDdl <> vbnullstring) then
    objOUT.write vbnewline & vbnewline & now & vbtab & " - DOWNLOADING MSP BACKUP VIRTUAL DRIVE"
    objLOG.write vbnewline & vbnewline & now & vbtab & " - DOWNLOADING MSP BACKUP VIRTUAL DRIVE"
    objOUT.write vbnewline & now & vbtab & " - ONCE INSTALLED PLEASE EDUCATE CUSTOMER"
    objLOG.write vbnewline & now & vbtab & " - ONCE INSTALLED PLEASE EDUCATE CUSTOMER"
    ''DOWNLOAD MSP BACKUP VIRTUAL DRIVE
    call FILEDL(strMSPVDdl, "MSPBackupVD.exe")
    ''INSTALL MSP BACKUP VIRTUAL DRIVE
    call HOOK("C:\temp\MSPBackupVD.exe")
    if (err.number <> 0) then
      call LOGERR(95)
    end if
  end if
end sub

sub CHKAU()																									''CHECK FOR SCRIPT UPDATE, AUTO_PLAN.VBS, REF #2 , REF #6 , FIXES #5
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
					call FILEDL("https://github.com/CW-Khristos/scripts/raw/Auto_Plan/auto_plan.vbs", wscript.scriptname)
					''RUN LATEST VERSION
					if (wscript.arguments.count > 0) then             ''ARGUMENTS WERE PASSED
						for x = 0 to (wscript.arguments.count - 1)
							strTMP = strTMP & " " & objARG.item(x)
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
  if (err.number <> 0) then
    call LOGERR(10)
  end if
end sub

sub FILEDL(strURL, strFILE)                                 ''CALL HOOK TO DOWNLOAD FILE FROM URL
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
    call LOGERR(11)
  end if
end sub

sub HOOK(strCMD)                                            ''CALL HOOK TO MONITOR OUTPUT OF CALLED COMMAND
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
  if (err.number <> 0) then
    call LOGERR(12)
  end if
end sub

sub LOGERR(intSTG)                                          ''CALL HOOK TO MONITOR OUTPUT OF CALLED COMMAND
  if (err.number <> 0) then
    objOUT.write vbnewline & now & vbtab & vbtab & vbtab & err.number & vbtab & err.description
    objLOG.write vbnewline & now & vbtab & vbtab & vbtab & err.number & vbtab & err.description
		errRET = intSTG
		err.clear
  end if
end sub

sub CLEANUP()                                               ''SCRIPT CLEANUP
  if (errRET = 0) then         												      ''AUTO_PLANv2 COMPLETED SUCCESSFULLY
    objOUT.write vbnewline & "AUTO_PLANv2 SUCCESSFUL : " & NOW
  elseif (errRET <> 0) then    												      ''AUTO_PLANv2 FAILED
    objOUT.write vbnewline & "AUTO_PLANv2 FAILURE : " & NOW & " : " & errRET
    ''RAISE CUSTOMIZED ERROR CODE, ERROR CODE WILL BE DEFINE RESTOP NUMBER INDICATING WHICH SECTION FAILED
    call err.raise(vbObjectError + errRET, "AUTO_PLANv2", "FAILURE")
  end if
  objOUT.write vbnewline & vbnewline & now & " - AUTO_PLANv2 COMPLETE. PLEASE VERIFY ALL MONITORING AND SERVICES!" & vbnewline
  objLOG.write vbnewline & vbnewline & now & " - AUTO_PLANv2 COMPLETE. PLEASE VERIFY ALL MONITORING AND SERVICES!" & vbnewline
  objLOG.close
  ''EMPTY OBJECTS
  set objLOG = nothing
  set objFSO = nothing
  set objWSH = nothing
  set objARG = nothing
  set objOUT = nothing
  set objIN = nothing
  ''END SCRIPT, RETURN ERROR NUMBER
  wscript.quit errRET
end sub
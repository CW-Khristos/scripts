''AUTO_PLANv2.VBS
''DESIGNED TO AUTOMATE PROTECTION PLAN SETUP
''RUNS CUSTOMIZED MODULE "STAGES" REPRESENTING SECTIONS OF PLAN SETUP
''RUN ON LOCAL DEVICE WITH ADMINISTRATIVE PRIVILEGES
''COMPUTER RENAME WILL REQUIRE REBOOT AND RE-RUN OF SCRIPT
''CURRENTLY ONLY CREATES / UPDATES LOCAL RMMTECH USER
''WRITTEN BY : CJ BLEDSOE / CJ<@>THECOMPUTERWARRIORS.COM
on error resume next
''SCRIPT VARIABLES
dim errRET, strVER
''VARIABLES ACCEPTING PARAMETERS - CONFIGURES SNMP TRAP AND COMMUNITY STRING
dim strSNMP, strTRP
''SCRIPT OBJECTS
dim objLOG, objHOOK, objHTTP, objXML
dim objIN, objOUT, objARG, objWSH, objFSO
''VERSION FOR SCRIPT UPDATE, AUTO_PLAN.VBS, REF #2 , FIXES #5
strVER = 2
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
if (objFSO.fileexists("C:\temp\auto_planv2")) then      ''LOGFILE EXISTS
  objFSO.deletefile "C:\temp\auto_planv2", true
  set objLOG = objFSO.createtextfile("C:\temp\auto_planv2")
  objLOG.close
  set objLOG = objFSO.opentextfile("C:\temp\auto_planv2", 8)
else                                                  	''LOGFILE NEEDS TO BE CREATED
  set objLOG = objFSO.createtextfile("C:\temp\auto_planv2")
  objLOG.close
  set objLOG = objFSO.opentextfile("C:\temp\auto_planv2", 8)
end if
''READ PASSED COMMANDLINE ARGUMENTS - REQUIRES (AT LEAST) 2 ARGUMENTS
if (wscript.arguments.count > 0) then                 	''ARGUMENTS WERE PASSED
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
''AUTOMATIC UPDATE, AUTO_PLAN.VBS, REF #2 , FIXES #5
call CHKAU()
''PRE-MATURE END SCRIPT, TESTING AUTOMATIC UPDATE AUTO_PLAN.VBS, REF #2
call CLEANUP()
''------------
''STAGE1
''CHANGE ACTIVE POWER PLAN
objOUT.write vbnewline & now & vbtab & " - SETTING ACTIVE POWER PLAN : HIGH PERFORMANCE" & vbnewline
objLOG.write vbnewline & now & vbtab & " - SETTING ACTIVE POWER PLAN : HIGH PERFORMANCE" & vbnewline
call HOOK("powercfg.exe /setactive 8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c")
''DISABLE HIBERNATION
objOUT.write vbnewline & now & vbtab & " - DISABLING HIBERNATION" & vbnewline
objLOG.write vbnewline & now & vbtab & " - DISABLING HIBERNATION" & vbnewline
call HOOK("powercfg –h off")
''------------
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
    if (blnRBT) then
      ''RESTART COMPUTER - PROVIDES REASON
      call HOOK("shutdown /r /t 10 /d:p /c " & chr(34) & "AUTO_PANv2 - COMPUTER RENAME : " & strNEWPC & chr(34))
    end if
  end if
end if
''CLEAR INPUT
strSEL = vbnullstring
''------------
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
  end if
end if
strSEL = vbnullstring
''------------
''STAGE4 - REQUIRES TECHNICIAN INPUT / PASSED PARAMETERS
''INSTALL AND CONFIGURE SNMP - REQUIRES 'STRTRP', 'STRSNMP'; REQUIRES TECHNICIAN INPUT / PASSED PARAMETERS
objOUT.write vbnewline & vbnewline & vbtab & "INSTALLING AND CONFIGURING SNMP"
objLOG.write vbnewline & vbnewline & vbtab & "INSTALLING AND CONFIGURING SNMP"
objOUT.write vbnewline & vbtab & vbtab & "ENTER WINDOWS SOFTWARE PROBE IP / SNMP MONITOR AGENT IP :"
objLOG.write vbnewline & vbtab & vbtab & "ENTER WINDOWS SOFTWARE PROBE IP / SNMP MONITOR AGENT IP :"
strTRP = objIN.readline
objOUT.write vbnewline & vbtab & vbtab & "ENTER SNMP COMMUNITY STRING (DO NOT USE 'PUBLIC') :"
objLOG.write vbnewline & vbtab & vbtab & "ENTER SNMP COMMUNITY STRING (DO NOT USE 'PUBLIC') :"
strSNMP = objIN.readline
if ((strTRP <> vbnullstring) and (strSNMP <> vbnullstring)) then
  ''CLEAR PREVIOUS SNMP CONFIGURATIONS
  objOUT.write vbnewline & now & vbtab & vbtab & " - REMOVING PREVIOUS SNMP CONFIGURATIONS"
  objLOG.write vbnewline & now & vbtab & vbtab & " - REMOVING PREVIOUS SNMP CONFIGURATIONS"    
  call HOOK("reg delete " & chr(34) & "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\services\SNMP\Parameters" & chr(34) & " /va /f")
  ''MODIFY SNMP REGISTRY VALUES
  objOUT.write vbnewline & now & vbtab & "CHECKING SNMP STATUS"
  objLOG.write vbnewline & now & vbtab & "CHECKING SNMP STATUS" 
  set objDSM = objWSH.exec("DISM /online /get-features /format:table")
  while (not objDSM.stdout.atendofstream)
    strRET = objDSM.stdout.readline
    if (strRET <> vbnullstring) then
      if (instr(1,strRET,"SNMP") and instr(1,strRET,"Disabled")) then
        objOUT.write vbnewline & now & vbtab & "SNMP NOT INSTALLED, INSTALLING"
        objLOG.write vbnewline & now & vbtab & "SNMP NOT INSTALLED, INSTALLING"
        ''INSTALL SNMP
        call HOOK("DISM /online /enable-feature /featurename:SNMP")   
        call HOOK("powershell " & chr(34) & "Install-WindowsFeature RSAT-SNMP" & chr(34))
        objOUT.write vbnewline & now & vbtab & "SNMP INSTALLED"
        objLOG.write vbnewline & now & vbtab & "SNMP INSTALLED"            
      end if
      strRET = vbnullstring
    end if
  wend
  set objDSM = nothing
  ''ADD SNMP REGISTRY VALUES
  objOUT.write vbnewline & now & vbtab & "ADDING SNMP CONFIGURATIONS"
  objLOG.write vbnewline & now & vbtab & "ADDING SNMP CONFIGURATIONS"
  call HOOK("reg add " & chr(34) & "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\services\SNMP\Parameters" & chr(34) & _
		" /v EnableAuthenticationTraps /t REG_DWORD /d 0 /f")
  call HOOK("reg add " & chr(34) & "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\SNMP\Parameters\TrapConfiguration" & chr(34) & _
		" /f")
  call HOOK("reg add " & chr(34) & "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\SNMP\Parameters\TrapConfiguration\" & _
		strSNMP & chr(34) & " /f")
  call HOOK("reg add " & chr(34) & "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\services\SNMP\Parameters\ValidCommunities" & chr(34) & _
		" /v " & strSNMP & " /t REG_DWORD /d 4 /f")
  if (instr(1, strTRP, ",")) then ''HANDLE MULTIPLE SNMP TRAP AGENTS
    arrTRP = split(strTRP, ",")
    for intTRP = 0 to ubound(arrTRP)
      if (arrTRP(intTRP) <> vbnullstring) then
				'wscript.echo "reg add " & chr(34) & "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\services\SNMP\Parameters\PermittedManagers" & chr(34) & _
				'	" /v " & (intTRP + 1) & " /t REG_SZ /d " & arrTRP(intTRP) & " /f"
        call HOOK("reg add " & chr(34) & "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\services\SNMP\Parameters\PermittedManagers" & chr(34) & _
					" /v " & (intTRP + 1) & " /t REG_SZ /d " & arrTRP(intTRP) & " /f")
        call HOOK("reg add " & chr(34) & "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\services\SNMP\Parameters\PermittedManagers" & chr(34) & _
					" /v " & (intTRP + 1) & " /t REG_SZ /d " & arrTRP(intTRP) & " /f")
        call HOOK("reg add " & chr(34) & "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\services\SNMP\Parameters\TrapConfiguration\" & _
					strSNMP & chr(34) & " /v " & (intTRP + 1) & " /t REG_SZ /d " & arrTRP(intTRP) & " /f")
      end if
    next
  else  ''HANDLE SINGLE SNMP TRAP AGENT
    call HOOK("reg add " & chr(34) & "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\services\SNMP\Parameters\PermittedManagers" & chr(34) & _
			" /v 1 /t REG_SZ /d " & strTRP & " /f")
    call HOOK("reg add " & chr(34) & "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\services\SNMP\Parameters\PermittedManagers" & chr(34) & _
			" /v 2 /t REG_SZ /d " & strTRP & " /f")
    call HOOK("reg add " & chr(34) & "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\services\SNMP\Parameters\TrapConfiguration\" & _
			strSNMP & chr(34) & " /v 1 /t REG_SZ /d " & strTRP & " /f")
  end if
  objOUT.write vbnewline & now & vbtab & "SNMP CONFIGURATIONS COMPLETED"
  objLOG.write vbnewline & now & vbtab & "SNMP CONFIGURATIONS COMPLETED"
  objOUT.write vbnewline & now & vbtab & "PLEASE REVIEW SNMP CONFIGURATIONS :"
  objLOG.write vbnewline & now & vbtab & "PLEASE REVIEW SNMP CONFIGURATIONS :"    
  call HOOK("reg query " & chr(34) & "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\services\SNMP\Parameters" & chr(34) & " /s")
  if (err.number <> 0) then
    errRET = 1
    objOUT.write vbnewline & now & vbtab & "KEY NOT FOUND / ACCESS DENIED"
    objLOG.write vbnewline & now & vbtab & "KEY NOT FOUND / ACCESS DENIED"
    err.clear
  end if 
end if
''------------
''STAGE5 - REQUIRES TECHNICIAN INPUT / PASSED PARAMETERS
''INSTALL WINDOWS AGENT - REQUIRES 'STRCID', 'STRCNAM'; REQUIRES TECHNICIAN INPUT / PASSED PARAMETERS
''http://ilmcw.dyndns.biz/downloadFileServlet.download?relativePathToFile=%2Fdownload%2Frepository%2F669519631%2FWindows+Agent.msi
''msiexec /i "c:\temp\windows agent.msi" /qn CUSTOMERID=487 CUSTOMERNAME="Intermodal Logistics" SERVERPROTOCOL=https & _
''SERVERADDRESS=ilmcw.dyndns.biz SERVERPORT=443 /l*v c:\temp\install.log ALLUSERS=2
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
      ''DOWNLOAD WINDOWS AGENT MSI
      objOUT.write vbnewline & now & vbtab & vbtab & " - DOWNLOADING WINDOWS AGENT MSI"
      objLOG.write vbnewline & now & vbtab & vbtab & " - DOWNLOADING WINDOWS AGENT MSI"
      call FILEDL("http://download1047.mediafire.com/vxhb6ggzgs7g/4o9kbgmba0t3o7f/Windows+Agent.msi", "windows agent.msi")
      ''INSTALL WINDOWS AGENT
      objOUT.write vbnewline & now & vbtab & vbtab & " - INSTALLING WINDOWS AGENT"
      objLOG.write vbnewline & now & vbtab & vbtab & " - INSTALLING WINDOWS AGENT"
      call HOOK("msiexec /i " & chr(34) & "c:\temp\windows agent.msi" & chr(34) & " /qn CUSTOMERID=" & strCID & " CUSTOMERNAME=" & _
				chr(34) & strCNAM & chr(34) & "  SERVERPROTOCOL=https SERVERADDRESS=ilmcw.dyndns.biz SERVERPORT=443 /l*v c:\temp\install.log ALLUSERS=2")
    end if
  end if
end if
strSEL = vbnullstring
strCID = vbnullstring
strCNAM = vbnullstring
''------------
''STAGE6 - REQUIRES TECHNICIAN INPUT / PASSED PARAMETERS - NEED DOMAIN / WORKGROUP SELECTION
''INSTALL PROBE - REQUIRES 'STRCID', 'STRCNAM', 'STRDMN', 'STRDUSR', 'STRDPWD'; REQUIRES TECHNICIAN INPUT / PASSED PARAMETERS
''http://ilmcw.dyndns.biz/downloadFileServlet.download?relativePathToFile=%2Fdownload%2Frepository%2F679064808%2FWindows+Software+Probe.msi
''msiexec /i "c:\temp\windows software probe.msi" /qn CUSTOMERID=231 CUSTOMERNAME="Teabo & Sons Stucco" & _
''SERVERPROTOCOL="HTTPS://" SERVERADDRESS="ilmcw.dyndns.biz" SERVERPORT=443 PROBETYPE="Workgroup_Windows" & _ 
''AGENTDOMAIN=".\" AGENTUSERNAME="RMMTech" AGENTPASSWORD=""
if (strPRB = vbnullstring) then
  objOUT.write vbnewline & vbnewline & vbtab & "INSTALL WINDOWS PROBE (Y / N)?"
  objLOG.write vbnewline & vbnewline & vbtab & "INSTALL WINDOWS PROBE (Y / N)?"
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
    ''DOMAIN
    objOUT.write vbnewline & vbtab & vbtab & "ENTER DOMAIN :"
    objLOG.write vbnewline & vbtab & vbtab & "ENTER DOMAIN :"
    strDMN = objIN.readline
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
        ''DOWNLOAD WINDOWS PROBE MSI
        objOUT.write vbnewline & now & vbtab & vbtab & " - DOWNLOADING WINDOWS PROBE MSI"
        objLOG.write vbnewline & now & vbtab & vbtab & " - DOWNLOADING WINDOWS PROBE MSI"
        call FILEDL("http://download794.mediafire.com/zjr2jwie89ug/cla52wsyp957s6w/Windows+Software+Probe.msi", "windows software probe.msi")
        ''INSTALL WINDOWS PROBE
        objOUT.write vbnewline & now & vbtab & vbtab & " - INSTALLING WINDOWS PROBE"
        objLOG.write vbnewline & now & vbtab & vbtab & " - INSTALLING WINDOWS PROBE"
        call HOOK("msiexec /i " & chr(34) & "c:\temp\windows software probe.msi" & chr(34) & " /qn CUSTOMERID=" & strCID & " CUSTOMERNAME=" & _
					chr(34) & strCNAM & chr(34) & " SERVERPROTOCOL=" & chr(34) & "HTTPS://" & chr(34) & " SERVERADDRESS=" & chr(34) & "ilmcw.dyndns.biz" & chr(34) & _
					" SERVERPORT=443" & " PROBETYPE=" & chr(34) & "Network_Windows" & chr(34) & " AGENTDOMAIN=" & chr(34) & strDMN & chr(34) & " AGENTUSERNAME=" & _
					chr(34) & strDUSR & chr(34) & " AGENTPASSWORD=" & chr(34) & strDPWD & chr(34) & " /l*v c:\temp\probe_install.log ALLUSERS=2")
    end if
  end if
end if
''CLEANUP
call CLEANUP()
''END SCRIPT
''------------

''SUB-ROUTINES
sub CHKAU()																									''CHECK FOR SCRIPT UPDATE, AUTO_PLAN.VBS, REF #2 , FIXES #5
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
					call FILEDL("https://github.com/CW-Khristos/CW_MSI/raw/master/auto_plan.vbs", wscript.scriptname)
					''RUN LATEST VERSION
					if (wscript.arguments.count > 0) then             ''ARGUMENTS WERE PASSED
						for x = 0 to (wscript.arguments.count - 1)
							strTMP = strTMP & " " & objARG.item(x)
						next
						objWSH.run "cscript.exe //nologo " & chr(34) & "c:\temp\" & wscript.scriptname & chr(34) & strTMP, 0, false
					elseif (wscript.arguments.count = 0) then         ''NO ARGUMENTS WERE PASSED
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

sub FILEDL(strURL, strFILE)                           ''CALL HOOK TO DOWNLOAD FILE FROM URL
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

sub HOOK(strCMD)                                      ''CALL HOOK TO MONITOR OUTPUT OF CALLED COMMAND
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
    objOUT.write vbnewline & now & vbtab & vbtab & vbtab & err.number & vbtab & err.description
    objLOG.write vbnewline & now & vbtab & vbtab & vbtab & err.number & vbtab & err.description
		errRET = 3
		err.clear
  end if
end sub

sub CLEANUP()                                         ''SCRIPT CLEANUP
  if (errRET = 0) then         												''AUTO_PLANv2 COMPLETED SUCCESSFULLY
    objOUT.write vbnewline & "AUTO_PLANv2 SUCCESSFUL : " & NOW
  elseif (errRET <> 0) then    												''AUTO_PLANv2 FAILED
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
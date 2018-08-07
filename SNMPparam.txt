''SNMPPARAM.VBS
''DESIGNED TO CONFIGURE / MONITOR SNMP CONFIGURATIONS
''SCRIPT 'MODE' IS SET BY VARIABLE 'STRMOD', THIS MUST BE THE FIRST PARAMETER PASSED
''SNMP COMMUNITY STRING SET BY VARIABLE 'STRSNMP', THIS MUST BE THE SECOND PARAMETER PASSED
''SNMP TRAP AGENT SET BY VARIABLE 'STRTRP', THIS MUST BE THE THIRD PARAMETER PASSED, SEPARATE MULTIPLE TRAP DESTINATIONS WITH A ','
''WRITTEN BY : CJ BLEDSOE / CJ<@>THECOMPUTERWARRIORS.COM
on error resume next
''SCRIPT VARIABLES
''STANDARD VARIABLES
dim errRET
dim objIN, objOUT, objARG, objWSH, objFSO, objLOG, objHOOK
''VARIABLES ACCEPTING PARAMETERS - CONFIGURES SNMP TRAP AND COMMUNITY STRING
dim strMOD, strTRP, strSNMP
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
if (objFSO.fileexists("C:\temp\snmpparam")) then      ''LOGFILE EXISTS
  objFSO.deletefile "C:\temp\snmpparam", true
  set objLOG = objFSO.createtextfile("C:\temp\snmpparam")
  objLOG.close
  set objLOG = objFSO.opentextfile("C:\temp\snmpparam", 8)
else                                                  ''LOGFILE NEEDS TO BE CREATED
  set objLOG = objFSO.createtextfile("C:\temp\snmpparam")
  objLOG.close
  set objLOG = objFSO.opentextfile("C:\temp\snmpparam", 8)
end if
''READ PASSED COMMANDLINE ARGUMENTS - REQUIRES (AT LEAST) 2 ARGUMENTS
if (wscript.arguments.count > 0) then                 ''ARGUMENTS WERE PASSED
  for x = 0 to (wscript.arguments.count - 1)
    objOUT.write vbnewline & now & vbtab & "ARGUMENT " & (x + 1) & " (ITEM " & x & ") " & " PASSED : " & ucase(objARG.item(x))
  next
  ''SCRIPT MODE OF OPERATION
  strMOD = objARG.item(0)
  if (wscript.arguments.count > 1) then
    strSNMP = objARG.item(1)
    strTRP = objARG.item(2)
  else
  end if
else
end if

''------------
''BEGIN SCRIPT
objOUT.write vbnewline & vbnewline & now & " - STARTING SNMPPARAM" & vbnewline
objLOG.write vbnewline & vbnewline & now & " - STARTING SNMPPARAM" & vbnewline
''SELECT MODE "QUERY" / "MODIFY"
select case lcase(strMOD)
  ''QUERY
  case vbnullstring
    ''QUERY SNMP REGISTRY VALUES
    objOUT.write vbnewline & now & vbtab & "QUERYING SNMP CONFIGURATIONS"
    objLOG.write vbnewline & now & vbtab & "QUERYING SNMP CONFIGURATIONS"
    call HOOK("reg query " & chr(34) & "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\services\SNMP\Parameters" & chr(34) & " /s")
  ''CLEAR
  case "clear"
      ''CLEAR PREVIOUS SNMP CONFIGURATIONS
    objOUT.write vbnewline & now & vbtab & "REMOVING PREVIOUS SNMP CONFIGURATIONS"
    objLOG.write vbnewline & now & vbtab & "REMOVING PREVIOUS SNMP CONFIGURATIONS"    
    call HOOK("reg delete " & chr(34) & "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\services\SNMP\Parameters" & chr(34) & " /va /f")
  ''MODIFY
  case "modify"
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
    call HOOK("reg add " & chr(34) & "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\services\SNMP\Parameters" & chr(34) & " /v EnableAuthenticationTraps /t REG_DWORD /d 0 /f")
    call HOOK("reg add " & chr(34) & "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\SNMP\Parameters\TrapConfiguration" & chr(34) & " /f")
    call HOOK("reg add " & chr(34) & "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\SNMP\Parameters\TrapConfiguration\" & strSNMP & chr(34) & " /f")
    call HOOK("reg add " & chr(34) & "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\services\SNMP\Parameters\ValidCommunities" & chr(34) & " /v " & strSNMP & " /t REG_DWORD /d 4 /f")
    if (instr(1, strTRP, ",")) then ''HANDLE MULTIPLE SNMP TRAP AGENTS
      arrTRP = split(strTRP, ",")
      for intTRP = 0 to ubound(arrTRP)
        if (arrTRP(intTRP) <> vbnullstring) then
        wscript.echo "reg add " & chr(34) & "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\services\SNMP\Parameters\PermittedManagers" & chr(34) & " /v " & (intTRP + 1) & " /t REG_SZ /d " & arrTRP(intTRP) & " /f"
          call HOOK("reg add " & chr(34) & "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\services\SNMP\Parameters\PermittedManagers" & chr(34) & " /v " & (intTRP + 1) & " /t REG_SZ /d " & arrTRP(intTRP) & " /f")
          call HOOK("reg add " & chr(34) & "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\services\SNMP\Parameters\PermittedManagers" & chr(34) & " /v " & (intTRP + 1) & " /t REG_SZ /d " & arrTRP(intTRP) & " /f")
          call HOOK("reg add " & chr(34) & "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\services\SNMP\Parameters\TrapConfiguration\" & strSNMP & chr(34) & " /v " & (intTRP + 1) & " /t REG_SZ /d " & arrTRP(intTRP) & " /f")
        end if
      next
    else  ''HANDLE SINGLE SNMP TRAP AGENT
      call HOOK("reg add " & chr(34) & "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\services\SNMP\Parameters\PermittedManagers" & chr(34) & " /v 1 /t REG_SZ /d " & strTRP & " /f")
      call HOOK("reg add " & chr(34) & "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\services\SNMP\Parameters\PermittedManagers" & chr(34) & " /v 2 /t REG_SZ /d " & strTRP & " /f")
      call HOOK("reg add " & chr(34) & "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\services\SNMP\Parameters\TrapConfiguration\" & strSNMP & chr(34) & " /v 1 /t REG_SZ /d " & strTRP & " /f")
    end if
    objOUT.write vbnewline & now & vbtab & "SNMP CONFIGURATIONS COMPLETED"
    objLOG.write vbnewline & now & vbtab & "SNMP CONFIGURATIONS COMPLETED"
    objOUT.write vbnewline & now & vbtab & "PLEASE REVIEW SNMP CONFIGURATIONS :"
    objLOG.write vbnewline & now & vbtab & "PLEASE REVIEW SNMP CONFIGURATIONS :"    
    call HOOK("reg query " & chr(34) & "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\services\SNMP\Parameters" & chr(34) & " /s")
end select
if (err.number <> 0) then
  errRET = 1
  objOUT.write vbnewline & now & vbtab & "KEY NOT FOUND / ACCESS DENIED"
  objLOG.write vbnewline & now & vbtab & "KEY NOT FOUND / ACCESS DENIED"
  err.clear
end if
''CLEANUP
call CLEANUP
''END SCRIPT
''------------

''SUB-ROUTINES
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
  objOUT.write vbnewline & vbnewline & now & " - SNMPPARAM COMPLETE. PLEASE VERIFY ALL MONITORING AND SERVICES!" & vbnewline
  objLOG.write vbnewline & vbnewline & now & " - SNMPPARAM COMPLETE. PLEASE VERIFY ALL MONITORING AND SERVICES!" & vbnewline
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
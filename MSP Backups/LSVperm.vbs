''LSVPERM.VBS
''DESIGNED TO RESTRICT LSV PERMISSIONS TO PASSED USER ONLY AND UPDATE BACKUP SERVICE CONTROLLER SERVICE LOGON
''SCRIPT WILL CHECK STATUS OF BACKUPS PRIOR TO EXECUTION; IF BACKUPS ARE IN PROGRESS, SCRIPT WILL NOT PROCEED
''ACCEPTS 3 PARAMETERS , REQUIRES 3 PARAMETER
''REQUIRED PARAMETER : 'STRLSV' , STRING TO IDENTIFY 'ROOT' LSV BACKUP DESTINATION PATH
''REQUIRED PARAMETER : 'STRUSR' , STRING TO SET TARGET USER FOR LSV PERMISSIONS AND SERVICE LOGON
''REQUIRED PARAMETER : 'STRPWD' , STRING TO SET TARGET USER PASSWORD
''WRITTEN BY : CJ BLEDSOE / CJ<@>THECOMPUTERWARRIORS.COM
on error resume next
''SCRIPT OBJECTS
dim objSIN, objSOUT
dim objLOG, objEXEC, objHOOK
dim objIN, objOUT, objARG, objWSH, objFSO
''SCRIPT VARIABLES
dim colUSR(), colSID()
dim strIN, strOUT, strVER, errRET
dim strORG, strREP, strSID, strDMN
''VARIABLES ACCEPTING PARAMETERS
dim strLSV, strUSR, strPWD
''VERSION FOR SCRIPT UPDATE , LSVPERM.VBS , REF #2 , FIXES #32
strVER = 5
''DEFAULT SUCCESS
errRET = 0
''STDIN / STDOUT
set objIN = wscript.stdin
set objOUT = wscript.stdout
set objARG = wscript.arguments
''CREATE SCRIPTING OBJECTS
set objWSH = createobject("wscript.shell")
set objFSO = createobject("scripting.filesystemobject")
''PREPARE LOGFILE
if (objFSO.fileexists("C:\temp\LSVperm")) then                              ''LOGFILE EXISTS
  objFSO.deletefile "C:\temp\LSVperm", true
  set objLOG = objFSO.createtextfile("C:\temp\LSVperm")
  objLOG.close
  set objLOG = objFSO.opentextfile("C:\temp\LSVperm", 8)
else                                                                        ''LOGFILE NEEDS TO BE CREATED
  set objLOG = objFSO.createtextfile("C:\temp\LSVperm")
  objLOG.close
  set objLOG = objFSO.opentextfile("C:\temp\LSVperm", 8)
end if
''READ PASSED COMMANDLINE ARGUMENTS
if (wscript.arguments.count > 0) then                                       ''ARGUMENTS WERE PASSED
  for x = 0 to (wscript.arguments.count - 1)
    objOUT.write vbnewline & now & vbtab & " - ARGUMENT " & (x + 1) & " (ITEM " & x & ") " & " PASSED : " & ucase(objARG.item(x))
    objLOG.write vbnewline & now & vbtab & " - ARGUMENT " & (x + 1) & " (ITEM " & x & ") " & " PASSED : " & ucase(objARG.item(x))
  next 
  ''ARGUMENT 0 - FOLDER PATH
  strLSV = objARG.item(0)
  if (wscript.arguments.count > 1) then                                     ''SET RMMTECH LOGON ARGUMENTS FOR UPDATING 'BACKUP SERVICE CONTROLLER' LOGON
    strUSR = objARG.item(1)
    strPWD = objARG.item(2)
  elseif (wscript.arguments.count <= 1) then                                ''NOT ENOUGH ARGUMENTS PASSED , END SCRIPT , 'ERRRET'=1
    call LOGERR(1)
  end if
else                                                                        ''NO ARGUMENTS PASSED , END SCRIPT , 'ERRRET'=1
  call LOGERR(1)
end if

''------------
''BEGIN SCRIPT
if (errRET <> 0) then                                                       ''NO ARGUMENTS PASSED , END SCRIPT , 'ERRRET'=1
  call CLEANUP()
elseif (errRET = 0) then                                                    ''ARGUMENTS PASSED , CONTINUE SCRIPT
  objOUT.write vbnewline & vbnewline & now & vbtab & " - EXECUTING LSVPERM"
  objLOG.write vbnewline & vbnewline & now & vbtab & " - EXECUTING LSVPERM"
  ''AUTOMATIC UPDATE , 'ERRRET'=10 , LSVPERM.VBS , REF #2 , FIXES #32
  call CHKAU()
  ''CHECK MSP BACKUP STATUS VIA MSP BACKUP CLIENTTOOL UTILITY
  objOUT.write vbnewline & now & vbtab & " - CHECKING MSP BACKUP STATUS"
  objLOG.write vbnewline & now & vbtab & " - CHECKING MSP BACKUP STATUS"
  set objHOOK = objWSH.exec(chr(34) & "c:\Program Files\Backup Manager\ClientTool.exe" & chr(34) & " control.status.get")
  strIDL = objHOOK.stdout.readall
  objOUT.write vbnewline & now & vbtab & vbtab & vbtab & strIDL
  objLOG.write vbnewline & now & vbtab & vbtab & vbtab & strIDL
  set objHOOK = nothing
  if ((instr(1, strIDL, "Idle")) or (instr(1, strIDL, "RegSync"))) then     ''BACKUPS NOT IN PROGRESS , CONTINUE SCRIPT
    ''GET SIDS OF ALL USERS , 'ERRRET'=20
    intUSR = 0
    intSID = 0
    objOUT.write vbnewline & vbnewline & now & vbtab & vbtab & " - ENUMERATING USERNAMES AND SIDS, THIS MAY TAKE A FEW MOMENTS"
    objLOG.write vbnewline & vbnewline & now & vbtab & vbtab & " - ENUMERATING USERNAMES AND SIDS, THIS MAY TAKE A FEW MOMENTS"
    set objEXEC = objWSH.exec("wmic useraccount get name,sid /format:csv")
    while (not objEXEC.stdout.atendofstream)
      strIN = objEXEC.stdout.readline
      'objOUT.write vbnewline & now & vbtab & vbtab & strIN
      'objLOG.write vbnewline & now & vbtab & vbtab & strIN
      if ((trim(strIN) <> vbnullstring) and (instr(1, strIN, ","))) then
        if ((trim(split(strIN, ",")(1)) <> vbnullstring) and (trim(split(strIN, ",")(1)) <> "Name")) then
          redim preserve colUSR(intUSR + 1)
          redim preserve colSID(intSID + 1)
          colUSR(intUSR) = trim(split(strIN, ",")(1))
          colSID(intSID) = trim(split(strIN, ",")(2))
          intUSR = (intUSR + 1)
          intSID = (intSID + 1)
        end if
      end if
      if (err.number <> 0) then
        call LOGERR(2)
      end if
    wend
    err.clear
    ''VALIDATE COLLECTED USERNAMES AND SIDS
    intUSR = 0
    intSID = 0
    objOUT.write vbnewline & vbnewline & now & vbtab & vbtab & " - COLLECTED USERNAMES AND SIDS"
    objLOG.write vbnewline & vbnewline & now & vbtab & vbtab & " - COLLECTED USERNAMES AND SIDS"
    for intUSR = 0 to ubound(colUSR)
      ''FIND USER/S MATCHING PASSED 'STRUSR' TARGET USER
      ''HANDLE '\' IS PASSED TARGET USERNAME 'STRUSR' , REF #37
      if (instr(1, lcase(strUSR), "\")) then
        if (instr(1, lcase(colUSR(intUSR)), lcase(split(strUSR, "\")(1)))) then
          redim preserve arrSID(intSID + 1)
          arrSID(intSID) = colSID(intUSR)
          intSID = intSID + 1
          objOUT.write vbnewline & now & vbtab & vbtab & vbtab & "MARKED : " & colUSR(intUSR) & " : " & arrSID(intSID - 1)
          objLOG.write vbnewline & now & vbtab & vbtab & vbtab & "MARKED : " & colUSR(intUSR) & " : " & arrSID(intSID - 1)
        end if
      ''HANDLE WITHOUT '\' IN PASSED TARGET USERNAME 'STRUSR' , REF #37
      elseif (instr(1, lcase(strUSR), "\") = 0) then
        if (instr(1, lcase(colUSR(intUSR)), lcase(strUSR))) then
          redim preserve arrSID(intSID + 1)
          arrSID(intSID) = colSID(intUSR)
          intSID = intSID + 1
          objOUT.write vbnewline & now & vbtab & vbtab & vbtab & "MARKED : " & colUSR(intUSR) & " : " & arrSID(intSID - 1)
          objLOG.write vbnewline & now & vbtab & vbtab & vbtab & "MARKED : " & colUSR(intUSR) & " : " & arrSID(intSID - 1)
        end if
      end if
      objOUT.write vbnewline & now & vbtab & vbtab & vbtab & colUSR(intUSR) & " : " & colSID(intUSR)
      objLOG.write vbnewline & now & vbtab & vbtab & vbtab & colUSR(intUSR) & " : " & colSID(intUSR)
    next
    ''RETRIEVE WORKGROUP / DOMAIN INFORMATION FROM NETWORK
    if (instr(1, strUSR, "\") = 0) then
      strDMN = objWSH.ExpandEnvironmentStrings("%USERDOMAIN%")
      if (lcase(strDMN) = "workgroup") then
        strDMN = ".\"
        strUSR = strDMN & strUSR
      elseif (lcase(strDMN) <> "workgroup") then
        strUSR = strDMN & "\" & strUSR
      else
        strDMN = ".\"
        strUSR = strDMN & strUSR
      end if
    end if
    ''STOP 'BACKUP SERVICE CONTROLLER' AND UPDATE ACCOUNT LOGON TO RMMTECH
    objOUT.write vbnewline & vbnewline & now & vbtab & vbtab & " - UPDATING BACKUP SERVICE AND LSV PERMISSIONS"
    objLOG.write vbnewline & vbnewline & now & vbtab & vbtab & " - UPDATING BACKUP SERVICE AND LSV PERMISSIONS"
    call HOOK("sc.exe stop " & chr(34) & "Backup Service Controller" & chr(34))
    ''RESTRICT FILE-SYSTEM PERMISSIONS PRIOR TO APPLYING SERVICE LOGON AND RESTARTING SERVICE , 'ERRRET'=21 , REF #2 , REF #32
    ''TAKEOWN USING CURRENT USER, THIS SHOULD BE RMMTECH
    objOUT.write vbnewline & vbnewline & now & vbtab & vbtab & " - ASSIGNING " & strUSR & " OWNERSHIP"
    objLOG.write vbnewline & vbnewline & now & vbtab & vbtab & " - ASSIGNING " & strUSR & " OWNERSHIP"
    call HOOK("takeown /F " & chr(34) & strLSV & chr(34) & " /R /D Y")
    if (errRET <> 0) then
      call LOGERR(21)
    end if
    ''ADD RMMTECH USER EXPLICIT FULL CONTROL , 'ERRRET'=22
    objOUT.write vbnewline & now & vbtab & vbtab & " - ASSIGNING " & strUSR & " FULL CONTROL"
    objLOG.write vbnewline & now & vbtab & vbtab & " - ASSIGNING " & strUSR & " FULL CONTROL"
    for intUSR = 0 to ubound(colUSR)
      intSID = intUSR
      if (instr(1, lcase(colUSR(intUSR)), strUSR)) then
        call HOOK("icacls " & chr(34) & strLSV & chr(34) & " /grant " & colUSR(intUSR) & ":(OI)(CI)F /T /C /Q")
        call HOOK("icacls " & chr(34) & strLSV & chr(34) & " /grant *" & colSID(intSID) & ":(OI)(CI)F /T /C /Q")
      end if
    next
    if (errRET <> 0) then
      call LOGERR(22)
    end if
    ''DISABLE INHERITANCE ON LSV DESTINATION, AND ONLY THE ROOT MSP BACKUP LSV DESTINATION , 'ERRRET'=23
    ''THIS MUST BE DONE LAST, BEFORE REMOVING ALL OTHER USER PERMISSIONS SO RMMTECH PERMISSIONS FULLY PROPAGATE TO ALL FILES / FOLDERS
    objOUT.write vbnewline & vbnewline & now & vbtab & vbtab & " - REMOVING INHERITANCE ON LSV DESTINATION"
    objLOG.write vbnewline & vbnewline & now & vbtab & vbtab & " - REMOVING INHERITANCE ON LSV DESTINATION"
    call HOOK("icacls " & chr(34) & strLSV & chr(34) & " /inheritance:r /C")
    if (errRET <> 0) then
      call LOGERR(23)
    end if
    ''REMOVE USER PERMISSIONS, EXCLUDE RMMTECH FROM REMOVAL , 'ERRRET'=24
    objOUT.write vbnewline & vbnewline & now & vbtab & vbtab & " - REMOVING ALL OTHER ENUMERATED USERS' PERMISSIONS"
    objLOG.write vbnewline & vbnewline & now & vbtab & vbtab & " - REMOVING ALL OTHER ENUMERATED USERS' PERMISSIONS"
    for intUSR = 0 to ubound(colUSR)
      intSID = intUSR
      if ((colUSR(intUSR) <> vbnullstring) and (instr(1, lcase(colUSR(intUSR)), "rmmtech") = 0)) then
        call HOOK("icacls " & chr(34) & strLSV & chr(34) & " /remove:g " & colUSR(intUSR) & " /T /C /Q")
        call HOOK("icacls " & chr(34) & strLSV & chr(34) & " /remove:g *" & colSID(intSID) & " /T /C /Q")
      end if
    next
    if (errRET <> 0) then
      call LOGERR(24)
    end if
    ''DOWNLOAD SVCPERM.VBS SCRIPT TO GRANT USER SERVICE LOGON , 'ERRRET'=30 , REF #2 , FIXES #32
    objOUT.write vbnewline & vbnewline & now & vbtab & vbtab & " - DOWNLOADING SERVICE LOGON SCRIPT : SVCPERM"
    objLOG.write vbnewline & vbnewline & now & vbtab & vbtab & " - DOWNLOADING SERVICE LOGON SCRIPT : SVCPERM"
    call FILEDL("https://github.com/CW-Khristos/scripts/raw/master/SVCperm.vbs", "SVCperm.vbs")
    if (errRET <> 0) then
      call LOGERR(30)
    end if
    ''EXECUTE SERVICE LOGON SCRIPT : SVCPERM , 'ERRRET'=31 , REF #2 , FIXES #32
    if (objFSO.fileexists("c:\temp\svcperm.vbs")) then                                  ''SVCPERM.VBS DOWNLOAD SUCCESSFUL
      objOUT.write vbnewline & vbnewline & now & vbtab & vbtab & " - EXECUTING SERVICE LOGON SCRIPT : SVCPERM : THIS MAY TAKE A FEW MOMENTS"
      objLOG.write vbnewline & vbnewline & now & vbtab & vbtab & " - EXECUTING SERVICE LOGON SCRIPT : SVCPERM : THIS MAY TAKE A FEW MOMENTS"
      call HOOK("cscript.exe //nologo " & chr(34) & "c:\temp\svcperm.vbs" & chr(34) & " " & chr(34) & strUSR & chr(34) & _
        " " & chr(34) & strPWD & chr(34) & " " & chr(34) & "Backup Service Controller" & chr(34))
    elseif (not objFSO.fileexists("c:\temp\svcperm.vbs")) then                          ''SVCPERM.VBS DOWNLOAD UNSUCCESSFUL , 'ERRRET'=31
      call LOGERR(31)
    end if
    if (errRET = 0) then                                                                ''SERVICE PERMISSIONS UPDATE SUCCESSFUL
      objOUT.write vbnewline & now & vbtab & vbtab & " - LOGON AS SERVICE GRANTED : " & strUSR
      objLOG.write vbnewline & now & vbtab & vbtab & " - LOGON AS SERVICE GRANTED : " & strUSR
      objOUT.write vbnewline & now & vbtab & vbtab & " - BACKUP SERVICE CONTROLLER UPDATED"
      objLOG.write vbnewline & now & vbtab & vbtab & " - BACKUP SERVICE CONTROLLER UPDATED"
      ''RESTART 'BACKUP SERVICE CONTROLLER'
      wscript.sleep 90
      objOUT.write vbnewline & vbnewline & now & vbtab & vbtab & " - RESTARTING BACKUP SERVICE CONTROLLER"
      objLOG.write vbnewline & vbnewline & now & vbtab & vbtab & " - RESTARTING BACKUP SERVICE CONTROLLER"
      call HOOK("sc.exe start " & chr(34) & "Backup Service Controller" & chr(34))
    elseif (errRET <> 0) then                                                           ''SERVICE PERMISSIONS UPDATE UNSUCCESFUL , 'ERRRET'=32
      call LOGERR(32)
    end if
  elseif ((instr(1, strIDL, "Idle") = 0) and (instr(1, strIDL, "RegSync") = 0)) then    ''BACKUPS IN PROGRESS , 'ERRRET'=2
    call LOGERR(2)
  end if
end if
''END SCRIPT
call CLEANUP()
''END SCRIPT
''------------

''SUB-ROUTINES
sub CHKAU()																					                                    ''CHECK FOR SCRIPT UPDATE, 'ERRRET'=10 , LSVPERM.VBS , REF #2
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
  if objXML.load("https://github.com/CW-Khristos/scripts/raw/dev/version.xml") then
    set colVER = objXML.documentelement
    for each objSCR in colVER.ChildNodes
      ''LOCATE CURRENTLY RUNNING SCRIPT
      if (lcase(objSCR.nodename) = lcase(wscript.scriptname)) then
        ''CHECK LATEST VERSION
        if (cint(objSCR.text) > cint(strVER)) then
          objOUT.write vbnewline & now & " - UPDATING " & objSCR.nodename & " : " & objSCR.text & vbnewline
          objLOG.write vbnewline & now & " - UPDATING " & objSCR.nodename & " : " & objSCR.text & vbnewline
          ''DOWNLOAD LATEST VERSION OF SCRIPT
          call FILEDL("https://github.com/CW-Khristos/scripts/raw/dev/MSP%20Backups/LSVperm.vbs", wscript.scriptname)
          ''RUN LATEST VERSION
          if (wscript.arguments.count > 0) then                                         ''ARGUMENTS WERE PASSED
            for x = 0 to (wscript.arguments.count - 1)
              strTMP = strTMP & " " & chr(34) & objARG.item(x) & chr(34)
            next
            objOUT.write vbnewline & now & vbtab & " - RE-EXECUTING  " & objSCR.nodename & " : " & objSCR.text & vbnewline
            objLOG.write vbnewline & now & vbtab & " - RE-EXECUTING  " & objSCR.nodename & " : " & objSCR.text & vbnewline
            objWSH.run "cscript.exe //nologo " & chr(34) & "c:\temp\" & wscript.scriptname & chr(34) & strTMP, 0, false
          elseif (wscript.arguments.count = 0) then                                     ''NO ARGUMENTS WERE PASSED
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
  if (err.number <> 0) then                                                             ''ERROR RETURNED DURING UPDATE CHECK , 'ERRRET'=10
    call LOGERR(10)
  end if
end sub

sub FILEDL(strURL, strFILE)                   			                                    ''CALL HOOK TO DOWNLOAD FILE FROM URL , 'ERRRET'=11
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
  if (objFSO.fileexists(strSAV)) then
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
  if (objFSO.fileexists(strSAV)) then
    objOUT.write vbnewline & vbnewline & now & vbtab & " - DOWNLOAD : " & strSAV & " : SUCCESSFUL"
    objLOG.write vbnewline & vbnewline & now & vbtab & " - DOWNLOAD : " & strSAV & " : SUCCESSFUL"
    set objHTTP = nothing
  end if
  if (err.number <> 0) then                                                             ''ERROR RETURNED , 'ERRRET'=11
    call LOGERR(11)
  end if
end sub

sub HOOK(strCMD)                                                                        ''CALL HOOK TO MONITOR OUTPUT OF CALLED COMMAND , 'ERRRET'=12
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
  if (err.number <> 0) then                                                             ''ERROR RETURNED , 'ERRRET'=12
    call LOGERR(12)
  end if
end sub

sub LOGERR(intSTG)                                                                      ''CALL HOOK TO LOG AND SET ERRORS
  if (err.number <> 0) then
    objOUT.write vbnewline & now & vbtab & vbtab & vbtab & err.number & vbtab & err.description & vbnewline
    objLOG.write vbnewline & now & vbtab & vbtab & vbtab & err.number & vbtab & err.description & vbnewline
    errRET = intSTG
    err.clear
  end if
  select case intSTG
    case 1                                                                              '' 'ERRRET'=1 - NOT ENOUGH ARGUMENTS
      objOUT.write vbnewline & vbnewline & now & vbtab & " - SCRIPT REQUIRES PATH TO MSP LSV DESTINATION, LSV USER, LSV PASSWORD"
      objLOG.write vbnewline & vbnewline & now & vbtab & " - SCRIPT REQUIRES PATH TO MSP LSV DESTINATION, LSV USER, LSV PASSWORD"
    case 2                                                                              '' 'ERRRET'=2 - BACKUPS IN PROGRESS
      objOUT.write vbnewline & now & vbtab & vbtab & " - BACKUPS IN PROGRESS, ENDING LSVPERM"
      objLOG.write vbnewline & now & vbtab & vbtab & " - BACKUPS IN PROGRESS, ENDING LSVPERM"
  end select
end sub

sub CLEANUP()                                                                           ''SCRIPT CLEANUP
  if (errRET = 0) then                                                                  ''SCRIPT COMPLETED SUCCESSFULLY
    objOUT.write vbnewline & vbnewline & now & vbtab & " - LSVPERM COMPLETE : " & errRET & " : " & now
    objLOG.write vbnewline & vbnewline & now & vbtab & " - LSVPERM COMPLETE : " & errRET & " : " & now
    err.clear
  elseif (errRET <> 0) then                                                             ''SCRIPT FAILED
    objOUT.write vbnewline & vbnewline & now & vbtab & " - LSVPERM FAILURE : " & errRET & " : " & now
    objLOG.write vbnewline & vbnewline & now & vbtab & " - LSVPERM FAILURE : " & errRET & " : " & now
    ''RAISE CUSTOMIZED ERROR CODE, ERROR CODE WILL BE DEFINED RESTOP NUMBER INDICATING WHICH SECTION FAILED
    call err.raise(vbObjectError + errRET, "LSVPERM", "FAILURE")
  end if
  ''EMPTY OBJECTS
  set objEXEC = nothing
  set objLOG = nothing
  set objFSO = nothing
  set objWSH = nothing
  set objARG = nothing
  set objOUT = nothing
  set objIN = nothing
  ''END SCRIPT, RETURN ERROR
  wscript.quit err.number
end sub
''PBXUPLOAD.VBS
''DESIGNED TO AUTOMATE SSH CONNECTION AND UPLOAD OF FILES TO PBX
''ACCEPTS 3 PARAMETERS , REQUIRES 2 PARAMETERS
''REQUIRED PARAMETER 'STRUSR' ; STRING VALUE TO HOLD PASSED 'USER' ; PBX USER LOGIN
''REQUIRED PARAMETER 'STRPWD' ; STRING VALUE TO HOLD PASSED 'PASSWORD' ; PBX USER PASSWORD
''OPTIONAL PARAMETER 'STRIP' ; STRING VALUE TO HOLD PASSED 'IP' ; TARGET PBX IP
''WRITTEN BY : CJ BLEDSOE / CBLEDSOE<@>IPMCOMPUTERS.COM
on error resume next
''SCRIPT VARIABLES
dim errRET, strVER
''VARIABLES ACCEPTING PARAMETERS
dim arrPEM()
dim strIP, strUSR, strPWD
''SCRIPT OBJECTS
dim objFSO, objLOG, objHOOK
dim objIN, objOUT, objARG, objWSH
''PBX LIST
redim arrPEM(0)
''VERSION FOR SCRIPT UPDATE, PBXUPLOAD.VBS, REF #2 , REF #68 , REF #69
strVER = 1
strREPO = "scripts"
strBRCH = "dev"
strDIR = vbnullstring
''DEFAULT SUCCESS
errRET = 0
''STDIN / STDOUT
set objIN = wscript.stdin
set objOUT = wscript.stdout
set objARG = wscript.arguments
''OBJECTS FOR LOCATING FOLDERS
set objWSH = createobject("wscript.shell")
set objFSO = createobject("scripting.filesystemobject")
strPBX = "C:\Users\CBledsoe\IPM-Github\pbxlist.txt"
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
if (objFSO.fileexists("C:\temp\PBXUPLOAD")) then            ''LOGFILE EXISTS
  objFSO.deletefile "C:\temp\PBXUPLOAD", true
  set objLOG = objFSO.createtextfile("C:\temp\PBXUPLOAD")
  objLOG.close
  set objLOG = objFSO.opentextfile("C:\temp\PBXUPLOAD", 8)
else                                                        ''LOGFILE NEEDS TO BE CREATED
  set objLOG = objFSO.createtextfile("C:\temp\PBXUPLOAD")
  objLOG.close
  set objLOG = objFSO.opentextfile("C:\temp\PBXUPLOAD", 8)
end if
''READ PASSED COMMANDLINE ARGUMENTS - REQUIRES (AT LEAST) 2 ARGUMENTS
if (wscript.arguments.count > 0) then                       ''ARGUMENTS WERE PASSED
  ''ARGUMENT OUTPUT DISABLED TO SANITIZE
  'for x = 0 to (wscript.arguments.count - 1)
  '  objOUT.write vbnewline & now & vbtab & "ARGUMENT " & (x + 1) & " (ITEM " & x & ") " & " PASSED : " & ucase(objARG.item(x))
  'next
  ''SCRIPT MODE OF OPERATION
  if (wscript.arguments.count > 1) then
    strUSR = objARG.item(0)
    strPWD = objARG.item(1)
    if (wscript.arguments.count > 2) then
      strIP = objARG.item(2)
    end if
  elseif (wscript.arguments.count < 2) then                 ''NOT ENOUGH ARGUMENTS PASSED , END SCRIPT , 'ERRRET'=1
    call LOGERR(1)
  end if
elseif (wscript.arguments.count = 0) then                   ''NO ARGUMENTS PASSED , END SCRIPT , 'ERRRET'=1
  call LOGERR(1)
end if

''------------
''BEGIN SCRIPT
if (errRET = 0) then                                        ''NO ERRORS DURING INITIAL START
  objOUT.write vbnewline & vbnewline & now & vbtab & " - STARTING PBXUPLOAD" & vbnewline
  objLOG.write vbnewline & vbnewline & now & vbtab & " - STARTING PBXUPLOAD" & vbnewline
	''AUTOMATIC UPDATE, SNMPARAM.VBS, REF #2 , REF #69 , REF #68 , FIXES #9
  objOUT.write vbnewline & now & vbtab & vbtab & " - CHECKING FOR UPDATE : PBXUPLOAD : " & strVER
  objLOG.write vbnewline & now & vbtab & vbtab & " - CHECKING FOR UPDATE : PBXUPLOAD : " & strVER
  ''DOWNLOAD CHKAU.VBS SCRIPT, REF #2 , REF #69 , REF #68
  call FILEDL("https://raw.githubusercontent.com/CW-Khristos/scripts/master/chkAU.vbs", "C:\IT\Scripts", "chkAU.vbs")
  ''EXECUTE CHKAU.VBS SCRIPT, REF #69
  intRET = objWSH.run ("cmd.exe /C " & chr(34) & "cscript.exe " & chr(34) & "C:\IT\Scripts\chkAU.vbs" & chr(34) & " " & _
    chr(34) & strREPO & chr(34) & " " & chr(34) & strBRCH & chr(34) & " " & chr(34) & strDIR & chr(34) & " " & _
    chr(34) & wscript.scriptname & chr(34) & " " & chr(34) & strVER & chr(34) & " " & _
    chr(34) & strMOD & "|" & strSNMP & "|" & strTRP & chr(34) & chr(34), 0, true)
  ''CHKAU RETURNED - NO UPDATE FOUND , REF #2 , REF #69 , REF #68
  objOUT.write vbnewline & "errRET='" & intRET & "'"
  objLOG.write vbnewline & "errRET='" & intRET & "'"
  intRET = (intRET - vbObjectError)
  objOUT.write vbnewline & "errRET='" & intRET & "'"
  objLOG.write vbnewline & "errRET='" & intRET & "'"
  if ((intRET = 4) or (intRET = 10) or (intRET = 11) or (intRET = 1) or (intRET = 2147221505) or (intRET = 2147221517)) then
    ''COLLECT CERTIFICATE FILES
    set objSRC = objFSO.getfolder("C:\3cx")
    set colFILE = objSRC.files
    for each objFILE in colFILE
      arrPEM(intPEM) = objFILE.path
      redim preserve arrPEM(intPEM + 1)
      intPEM = (intPEM + 1)
    next
    set colFILE = nothing
    set objSRC = nothing
    ''CONNECT TO PBX
    if (objARG.count = 2) then
      set objTMP = objFSO.opentextfile(strPBX, 1)
      while (not objTMP.atendofstream)
        strTMP = objTMP.readline
        if ((strTMP <> vbnullstring) and (instr(1, strTMP, "PBXLIST.TXT") = 0)) then
          intPEM = 0
          arrTMP = split(strTMP, "|")
          strIP = arrTMP(0)
          for intPEM = 0 to ubound(arrPEM)
            if (arrPEM(intPEM) <> vbnullstring) then
              if (instr(1, arrPEM(intPEM), "_")) then
                strPEM = "C:\3cx\upload\" & arrTMP(1) & "-key.pem"
                strRCMD = "cmd.exe /c copy /Y " & arrPEM(intPEM) & " " & strPEM
                objOUT.write vbnewline & now & vbtab & vbtab & " - COPYING CERT : " & strPEM
                objLOG.write vbnewline & now & vbtab & vbtab & " - COPYING CERT : " & strPEM
                'objOUT.write vbnewline & vbnewline & strRCMD
                call HOOK(strRCMD)
                wscript.sleep 1000
                objOUT.write vbnewline & now & vbtab & vbtab & " - UPLOADING CERT : " & strPEM
                objLOG.write vbnewline & now & vbtab & vbtab & " - UPLOADING CERT : " & strPEM
                strRCMD = "C:\Users\CBledsoe\AppData\Local\Programs\WinSCP\winscp.com /command " & chr(34) & "open scp://" & strUSR & ":" & strPWD & "@" & strIP & ":22/ -hostkey=acceptnew" & chr(34) & " " & _
                  chr(34) & "put " & strPEM & " /var/lib/3cxpbx/Bin/nginx/conf/Instance1/" & chr(34) & " " & chr(34) & "exit" & chr(34) & " /log=" & chr(34) & "C:\temp\pbx_winscp.log" & chr(34) & " /loglevel=0"
                'objOUT.write vbnewline & vbnewline & strRCMD
                call HOOK(strRCMD)
              elseif (instr(1, arrPEM(intPEM), "_") = 0) then
                strPEM = "C:\3cx\upload\" & arrTMP(1) & "-crt.pem"
                strRCMD = "cmd.exe /c copy /Y " & arrPEM(intPEM) & " " & strPEM
                objOUT.write vbnewline & now & vbtab & vbtab & " - COPYING CERT : " & strPEM
                objLOG.write vbnewline & now & vbtab & vbtab & " - COPYING CERT : " & strPEM
                'objOUT.write vbnewline & vbnewline & strRCMD
                call HOOK(strRCMD)
                wscript.sleep 1000
                objOUT.write vbnewline & now & vbtab & vbtab & " - UPLOADING CERT : " & strPEM
                objLOG.write vbnewline & now & vbtab & vbtab & " - UPLOADING CERT : " & strPEM
                strRCMD = "C:\Users\CBledsoe\AppData\Local\Programs\WinSCP\winscp.com /command " & chr(34) & "open scp://" & strUSR & ":" & strPWD & "@" & strIP & ":22/ -hostkey=acceptnew" & chr(34) & " " & _
                  chr(34) & "put " & strPEM & " /var/lib/3cxpbx/Bin/nginx/conf/Instance1/" & chr(34) & " " & chr(34) & "exit" & chr(34) & " /log=" & chr(34) & "C:\temp\pbx_winscp.log" & chr(34) & " /loglevel=0"
                'objOUT.write vbnewline & vbnewline & strRCMD
                call HOOK(strRCMD)
              end if
            end if
          next
          ''service '3CX PhoneSystem Nginx Server' restart
          objOUT.write vbnewline & now & vbtab & vbtab & " - RESTARTING PBX NGINX SERVICE"
          objLOG.write vbnewline & now & vbtab & vbtab & " - RESTARTING PBX NGINX SERVICE"
          strRCMD = "C:\Putty\putty.exe -ssh " & strUSR & "@" & strIP & " -pw " & strPWD & " 22"
          objWSH.run strRCMD, 1, false
          wscript.sleep 2000
          objWSH.sendkeys "{RIGHT}{ENTER}"
          wscript.sleep 1000
          objWSH.sendkeys "service nginx restart{ENTER}"
          wscript.sleep 4000
          objWSH.sendkeys "exit{ENTER}"
          'objOUT.write vbnewline & vbnewline & strRCMD
          'call HOOK(strRCMD)
        end if
        wscript.sleep 1000
      wend
      objTMP.close
      set objTMP = nothing
    elseif (objARG.count = 3) then
      set objTMP = objFSO.opentextfile(strPBX, 1)
      while (not objTMP.atendofstream)
        strTMP = objTMP.readline
        if ((strTMP <> vbnullstring) and (instr(1, strTMP, "PBXLIST.TXT") = 0)) then
          intPEM = 0
          arrTMP = split(strTMP, "|")
          strIP = arrTMP(0)
          if (objARG.item(2) = strIP) then
            set objSRC = objFSO.getfolder("C:\3cx")
            set colFILE = objSRC.files
            for each objFILE in colFILE
              arrPEM(intPEM) = objFILE.path
              redim preserve arrPEM(intPEM + 1)
              intPEM = (intPEM + 1)
            next
            set colFILE = nothing
            set objSRC = nothing
            for intPEM = 0 to ubound(arrPEM)
              if (arrPEM(intPEM) <> vbnullstring) then
                if (instr(1, arrPEM(intPEM), "_")) then
                  strPEM = "C:\3cx\upload\" & arrTMP(1) & "-key.pem"
                  strRCMD = "cmd.exe /c copy /Y " & arrPEM(intPEM) & " " & strPEM
                  objOUT.write vbnewline & now & vbtab & vbtab & " - COPYING CERT : " & strPEM
                  objLOG.write vbnewline & now & vbtab & vbtab & " - COPYING CERT : " & strPEM
                  'objOUT.write vbnewline & vbnewline & strRCMD
                  call HOOK(strRCMD)
                  wscript.sleep 1000
                  objOUT.write vbnewline & now & vbtab & vbtab & " - UPLOADING CERT : " & strPEM
                  objLOG.write vbnewline & now & vbtab & vbtab & " - UPLOADING CERT : " & strPEM
                  strRCMD = "C:\Users\CBledsoe\AppData\Local\Programs\WinSCP\winscp.com /command " & chr(34) & "open scp://" & strUSR & ":" & strPWD & "@" & strIP & ":22/ -hostkey=acceptnew" & chr(34) & " " & _
                    chr(34) & "put " & strPEM & " /var/lib/3cxpbx/Bin/nginx/conf/Instance1/" & chr(34) & " " & chr(34) & "exit" & chr(34) & " /log=" & chr(34) & "C:\temp\pbx_winscp.log" & chr(34) & " /loglevel=0"
                  'objOUT.write vbnewline & vbnewline & strRCMD
                  call HOOK(strRCMD)
                elseif (instr(1, arrPEM(intPEM), "_") = 0) then
                  strPEM = "C:\3cx\upload\" & arrTMP(1) & "-crt.pem"
                  strRCMD = "cmd.exe /c copy /Y " & arrPEM(intPEM) & " " & strPEM
                  objOUT.write vbnewline & now & vbtab & vbtab & " - COPYING CERT : " & strPEM
                  objLOG.write vbnewline & now & vbtab & vbtab & " - COPYING CERT : " & strPEM
                  'objOUT.write vbnewline & vbnewline & strRCMD
                  call HOOK(strRCMD)
                  wscript.sleep 1000
                  objOUT.write vbnewline & now & vbtab & vbtab & " - UPLOADING CERT : " & strPEM
                  objLOG.write vbnewline & now & vbtab & vbtab & " - UPLOADING CERT : " & strPEM
                  strRCMD = "C:\Users\CBledsoe\AppData\Local\Programs\WinSCP\winscp.com /command " & chr(34) & "open scp://" & strUSR & ":" & strPWD & "@" & strIP & ":22/ -hostkey=acceptnew" & chr(34) & " " & _
                    chr(34) & "put " & strPEM & " /var/lib/3cxpbx/Bin/nginx/conf/Instance1/" & chr(34) & " " & chr(34) & "exit" & chr(34) & " /log=" & chr(34) & "C:\temp\pbx_winscp.log" & chr(34) & " /loglevel=0"
                  'objOUT.write vbnewline & vbnewline & strRCMD
                  call HOOK(strRCMD)
                end if
              end if
            next
            ''service '3CX PhoneSystem Nginx Server' restart
            objOUT.write vbnewline & now & vbtab & vbtab & " - RESTARTING PBX NGINX SERVICE"
            objLOG.write vbnewline & now & vbtab & vbtab & " - RESTARTING PBX NGINX SERVICE"
            strRCMD = "C:\Putty\putty.exe -ssh " & strUSR & "@" & strIP & " -pw " & strPWD & " 22"
            objWSH.run strRCMD, 1, false
            wscript.sleep 2000
            objWSH.sendkeys "{RIGHT}{ENTER}"
            wscript.sleep 1000
            objWSH.sendkeys "service nginx restart{ENTER}"
            wscript.sleep 4000
            objWSH.sendkeys "exit{ENTER}"
            'objOUT.write vbnewline & vbnewline & strRCMD
            'call HOOK(strRCMD)
          end if
        end if
        wscript.sleep 1000
      wend
      objTMP.close
      set objTMP = nothing
    end if
  end if 
elseif (errRET <> 0) then                                   ''ERRORS ENCOUNTERED DURING INITIAL START
  call LOGERR(errRET)
end if
''END SCRIPT
call CLEANUP()
''END SCRIPT
''------------

''SUB-ROUTINES
sub FILEDL(strURL, strDL, strFILE)                          ''CALL HOOK TO DOWNLOAD FILE FROM URL , 'ERRRET'=11
  strSAV = vbnullstring
  ''SET DOWNLOAD PATH
  strSAV = strDL & "\" & strFILE
  objOUT.write vbnewline & vbnewline & now & vbtab & vbtab & " - HTTPDOWNLOAD-------------DOWNLOAD : " & strURL & " : SAVE AS :  " & strSAV
  objLOG.write vbnewline & vbnewline & now & vbtab & vbtab & " - HTTPDOWNLOAD-------------DOWNLOAD : " & strURL & " : SAVE AS :  " & strSAV
  ''ADD WINHTTP SECURE CHANNEL TLS REGISTRY KEYS
  call HOOK("reg add " & chr(34) & "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\WinHttp" & chr(34) & _
    " /f /v DefaultSecureProtocols /t REG_DWORD /d 0x00000A00 /reg:32")
  call HOOK("reg add " & chr(34) & "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\WinHttp" & chr(34) & _
    " /f /v DefaultSecureProtocols /t REG_DWORD /d 0x00000A00 /reg:64")
  ''CHECK IF FILE ALREADY EXISTS
  if (objFSO.fileexists(strSAV)) then
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
  if (objFSO.fileexists(strSAV)) then
    objOUT.write vbnewline & now & vbtab & vbtab & " - DOWNLOAD : " & strSAV & " : SUCCESSFUL"
    objLOG.write vbnewline & now & vbtab & vbtab & " - DOWNLOAD : " & strSAV & " : SUCCESSFUL"
  end if
	set objHTTP = nothing
  if ((err.number <> 0) and (err.number <> 58)) then        ''ERROR RETURNED DURING DOWNLOAD , 'ERRRET'=11
    call LOGERR(11)
  end if
end sub

sub HOOK(strCMD)                                            ''CALL HOOK TO MONITOR OUTPUT OF CALLED COMMAND , 'ERRRET'=12
  on error resume next
  objOUT.write vbnewline & now & vbtab & vbtab & " - EXECUTING : HOOK(" & strCMD & ")"
  objLOG.write vbnewline & now & vbtab & vbtab & " - EXECUTING : HOOK(" & strCMD & ")"
  set objHOOK = objWSH.exec(strCMD)
  if (instr(1, strCMD, "takeown /F ") = 0) then             ''SUPPRESS 'TAKEOWN' SUCCESS MESSAGES
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
  end if
  set objHOOK = nothing
  if (err.number <> 0) then                                 ''ERROR RETURNED DURING UPDATE CHECK , 'ERRRET'=12
    call LOGERR(12)
  end if
end sub

sub LOGERR(intSTG)                                          ''CALL HOOK TO MONITOR OUTPUT OF CALLED COMMAND
  errRET = intSTG
  if (err.number <> 0) then
    objOUT.write vbnewline & now & vbtab & vbtab & vbtab & err.number & vbtab & err.description & vbnewline
    objLOG.write vbnewline & now & vbtab & vbtab & vbtab & err.number & vbtab & err.description & vbnewline
		err.clear
  end if
  ''CUSTOM ERROR CODES
  select case intSTG
    case 1                                                  '' 'ERRRET'=1 - NOT ENOUGH ARGUMENTS
      objOUT.write vbnewline & vbnewline & now & vbtab & " - NOT ENOUGH ARGUMENTS"
      objLOG.write vbnewline & vbnewline & now & vbtab & " - NOT ENOUGH ARGUMENTS"
  end select
end sub

sub CLEANUP()                                               ''SCRIPT CLEANUP
  on error resume next
  if (errRET = 0) then         															''PBXUPLOAD COMPLETED SUCCESSFULLY
    objOUT.write vbnewline & vbnewline & now & vbtab & " - PBXUPLOAD SUCCESSFUL : " & now
    objLOG.write vbnewline & vbnewline & now & vbtab & " - PBXUPLOAD SUCCESSFUL : " & now
    err.clear
  elseif (errRET <> 0) then    															''PBXUPLOAD FAILED
    objOUT.write vbnewline & vbnewline & now & vbtab & " - PBXUPLOAD FAILURE : " & errRET & " : " & now
    objLOG.write vbnewline & vbnewline & now & vbtab & " - PBXUPLOAD FAILURE : " & errRET & " : " & now
    ''RAISE CUSTOMIZED ERROR CODE, ERROR CODE WILL BE DEFINE RESTOP NUMBER INDICATING WHICH SECTION FAILED
    call err.raise(vbObjectError + errRET, "PBXUPLOAD", "FAILURE")
  end if
  objOUT.write vbnewline & vbnewline & now & " - PBXUPLOAD COMPLETE" & vbnewline
  objLOG.write vbnewline & vbnewline & now & " - PBXUPLOAD COMPLETE" & vbnewline
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
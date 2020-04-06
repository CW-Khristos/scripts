''CHKAU.VBS
''DESIGNED TO AUTOMATE UPDATING AND RE-EXECUTION OF CW SCRIPTS
''CHECKS FOR UPDATE TO PASSED SCRIPT; BASED ON PASSED 'BRANCH','SCRIPT','VERSION'
''HANDLES PASSING OF ANY 'ARGUMENTS' GIVEN TO ORIGINAL SCRIPT TO UPDATED COPY
''ACCEPTS 6 PARAMETER , REQUIRES 5 PARAMETERS
''REQUIRED PARAMETER 'STRREPO' ; STRING VALUE TO HOLD PASSED 'REPO' ; TARGET GITHUB REPOSITORY
''REQUIRED PARAMETER 'STRBRCH' ; STRING VALUE TO HOLD PASSED 'BRANCH' ; TARGET GITHUB BRANCH
''REQUIRED PARAMETER 'STRDIR' ; STRING VALUE TO HOLD PASSED 'DIRECTORY' ; TARGET GITHUB SCRIPT DIRECTORY
''REQUIRED PARAMETER 'STRSCR' ; STRING VALURE TO HOLD PASSED 'SCRIPTNAME' ; WSCRIPT.SCRIPTNAME
''REQUIRED PARAMETER 'STRSVER' ; STRING VALURE TO HOLD PASSED 'VERSION' ; 'STRVER'
''OPTIONAL PARAMETER 'STRARG' ; STRING VALURE TO HOLD PASSED 'ARGUMENTS' ; SEPARATE MULTIPLE 'ARGUMENTS' VIA '|'
''WRITTEN BY : CJ BLEDSOE / CJ<@>THECOMPUTERWARRIORS.COM
on error resume next
''SCRIPT VARIABLES
dim errRET, strVER
''VARIABLES ACCEPTING PARAMETERS
dim strARG, arrARG
dim strSCR, strSVER
dim strREPO, strBRCH, strDIR
dim strIN, strOUT, strOPT, strRCMD
''SCRIPT OBJECTS
dim objIN, objOUT, objARG, objWSH, objFSO
dim objLOG, objEXEC, objHOOK, objHTTP, objXML
''VERSION FOR SCRIPT UPDATE , CHKAU.VBS , REF #2
strVER = 1
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
if (objFSO.fileexists("C:\temp\CHKAU")) then                ''LOGFILE EXISTS
  objFSO.deletefile "C:\temp\CHKAU", true
  set objLOG = objFSO.createtextfile("C:\temp\CHKAU")
  objLOG.close
  set objLOG = objFSO.opentextfile("C:\temp\CHKAU", 8)
else                                                        ''LOGFILE NEEDS TO BE CREATED
  set objLOG = objFSO.createtextfile("C:\temp\CHKAU")
  objLOG.close
  set objLOG = objFSO.opentextfile("C:\temp\CHKAU", 8)
end if
''READ PASSED COMMANDLINE ARGUMENTS
if (wscript.arguments.count > 4) then                       ''ARGUMENTS WERE PASSED
  for x = 0 to (wscript.arguments.count - 1)
    objOUT.write vbnewline & now & vbtab & " - ARGUMENT " & (x + 1) & " (ITEM " & x & ") " & " PASSED : " & ucase(objARG.item(x))
    objLOG.write vbnewline & now & vbtab & " - ARGUMENT " & (x + 1) & " (ITEM " & x & ") " & " PASSED : " & ucase(objARG.item(x))
  next 
  if (wscript.arguments.count > 4) then                     ''SET VARIABLES ACCEPTING ARGUMENTS
    strREPO = objARG.item(0)                                ''SET REQUIRED PARAMETER 'STRREPO' , TARGET 'REPO' TO UPDATE FROM ON GITHUB
    strBRCH = objARG.item(1)                                ''SET REQUIRED PARAMETER 'STRBRCH' , TARGET 'BRANCH' TO UPDATE FROM ON GITHUB
    strDIR = objARG.item(2)                                 ''SET REQUIRED PARAMETER 'STRDIR' , TARGET 'DIRECTORY' TO UPDATE FROM ON GITHUB
    strSCR = objARG.item(3)                                 ''SET REQUIRED PARAMETER 'STRSCR' , TARGET 'SCRIPTNAME' TO UPDATE
    strSVER = objARG.item(4)                                ''SET REQUIRED PARAMETER 'STRSVER' , REQUESTING SCRIPT 'VERSION' TO COMPARE
    if (wscript.arguments.count > 5) then                   ''SET OPTIONAL PARAMETERS
      strARG = objARG.item(5)                               ''SET OPTIONAL PARAMETER 'STRARG' , ORIGINAL 'ARGUMENTS' FROM REQUESTING SCRIPT ; SEPARATE MULTIPLE 'ARGUMENTS' VIA '|'
      ''FILL 'ARRARG' ORIGINAL 'ARGUMENTS'
      objOUT.write vbnewline & vbtab & strARG
      arrARG = split(strARG, "|")
      for intTMP = 0 to ubound(arrARG)
        objOUT.write vbnewline & vbtab & ubound(arrARG) & vbtab & arrARG(intTMP)
      next
    end if
  end if
else                                                        ''NOT ENOUGH ARGUMENTS PASSED , END SCRIPT , 'ERRRET'=1
  call LOGERR(1)
end if

''------------
''BEGIN SCRIPT
strTMP = vbnullstring
if (errRET <> 0) then                                       ''NO ARGUMENTS PASSED, END SCRIPT , 'ERRRET'=1
  call CLEANUP()
elseif (errRET = 0) then                                    ''ARGUMENTS PASSED, CONTINUE SCRIPT
	objOUT.write vbnewline & vbnewline & now & vbtab & " - EXECUTING CHKAU : " & strVER
	objLOG.write vbnewline & vbnewline & now & vbtab & " - EXECUTING CHKAU : " & strVER
	''AUTOMATIC UPDATE, CHKAU.VBS, REF #2
	intRET = (CHKAU(wscript.scriptname, strVER, _
    strREPO & "|" & strBRCH & "|" & strDIR & "|" & strSCR & "|" & strSVER & "|" & strARG))
  if (intRET = true) then
    objOUT.write vbnewline & vbnewline & now & vbtab & " - CHKAU UPDATED - RE-EXECUTED : " & wscript.scriptname & " : " & _
      strREPO & "|" & strBRCH & "|" & strDIR & "|" & strSCR & "|" & strSVER & "|" & strARG
    objLOG.write vbnewline & vbnewline & now & vbtab & " - CHKAU UPDATED - RE-EXECUTED : " & wscript.scriptname & " : " & _ 
      strREPO & "|" & strBRCH & "|" & strDIR & "|" & strSCR & "|" & strSVER & "|" & strARG
  elseif (intRET = false) then
    objOUT.write vbnewline & vbnewline & now & vbtab & " - CHKAU NO UPDATE - EXITING : CHKAU : SELF-UPDATE"
    objLOG.write vbnewline & vbnewline & now & vbtab & " - CHKAU NO UPDATE - EXITING : CHKAU : SELF-UPDATE"
  end if
  ''AUTOMATIC UPDATE, REQUESTING SCRIPT 'STRSCR', REF #2
  intRET = (CHKAU(strSCR, strSVER, strARG))
  if (intRET = true) then
    errRET = 2
    objOUT.write vbnewline & vbnewline & now & vbtab & " - CHKAU UPDATED - RE-EXECUTED : " & strSCR & strARG
    objLOG.write vbnewline & vbnewline & now & vbtab & " - CHKAU UPDATED - RE-EXECUTED : " & strSCR & strARG
  elseif (intRET = false) then
    errRET = 3
    objOUT.write vbnewline & vbnewline & now & vbtab & " - CHKAU NO UPDATE - EXITING : "
    objLOG.write vbnewline & vbnewline & now & vbtab & " - CHKAU NO UPDATE - EXITING : "    
  end if
end if
''END SCRIPT
call CLEANUP()
''END SCRIPT
''------------

''CHKAU FUNCTIONS
function CHKAU(strSCR, strSVER, strARG)                     ''CHECK FOR SCRIPT UPDATE , 'ERRRET'=10 , CHKAU.VBS , REF #2
  ''REMOVE WINDOWS AGENT CACHED VERSION OF SCRIPT
  if (objFSO.fileexists("C:\Program Files (x86)\N-Able Technologies\Windows Agent\Temp\Script\" & strSCR)) then
    objFSO.deletefile "C:\Program Files (x86)\N-Able Technologies\Windows Agent\Temp\Script\" & strSCR, true
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
	if objXML.load("https://github.com/CW-Khristos/scripts/raw/" & strBRCH & "/version.xml") then
		set colVER = objXML.documentelement
    for each objSCR in colVER.ChildNodes
      ''LOCATE ORIGINAL RUNNING SCRIPT
      if (ucase(objSCR.nodename) = ucase(strSCR)) then
        if (ucase(strSCR) <> "CHKAU.VBS") then              ''REQUESTING SCRIPT IS NOT 'CHKAU.VBS' , UPDATE CW SCRIPT , RE-EXECUTE WITH ORIGINAL 'ARGUMENTS'
          ''CHECK LATEST VERSION
          objOUT.write vbnewline & now & vbtab & " - CHKAU :  " & strSVER & " : GitHub - " & strBRCH & " : " & objSCR.text & vbnewline
          objLOG.write vbnewline & now & vbtab & " - CHKAU :  " & strSVER & " : GitHub - " & strBRCH & " : " & objSCR.text & vbnewline
          if (cint(objSCR.text) > cint(strSVER)) then
            objOUT.write vbnewline & now & vbtab & " - UPDATING " & objSCR.nodename & " : " & objSCR.text & vbnewline
            objLOG.write vbnewline & now & vbtab & " - UPDATING " & objSCR.nodename & " : " & objSCR.text & vbnewline
            ''DOWNLOAD LATEST VERSION OF ORIGINAL SCRIPT
            call FILEDL("https://github.com/CW-Khristos/" & strREPO & "/raw/" & strBRCH & strDIR & "/" & strSCR, strSCR)
            ''RUN LATEST VERSION OF ORIGINAL SCRIPT
            if (ubound(arrARG) > 0) then                    ''ARGUMENTS WERE PASSED
              strTMP = vbnullstring
              for x = 0 to (ubound(arrARG))
                if (arrARG(x) <> vbnullstring) then
                  strTMP = strTMP & " " & chr(34) & arrARG(x) & chr(34)
                end if
              next
              objOUT.write vbnewline & now & vbtab & " - RE-EXECUTING " & objSCR.nodename & " : " & objSCR.text & vbnewline
              objLOG.write vbnewline & now & vbtab & " - RE-EXECUTING  " & objSCR.nodename & " : " & objSCR.text & vbnewline
              intRET = objWSH.run("cscript.exe //nologo " & chr(34) & "c:\temp\" & strSCR & chr(34) & strTMP, 0, false)
              if (intRET = 0) then
                CHKAU = true
              end if
            elseif (wscript.arguments.count = 0) then       ''NO ARGUMENTS WERE PASSED
              objOUT.write vbnewline & now & vbtab & " - RE-EXECUTING  " & objSCR.nodename & " : " & objSCR.text & vbnewline
              objLOG.write vbnewline & now & vbtab & " - RE-EXECUTING  " & objSCR.nodename & " : " & objSCR.text & vbnewline
              intRET = objWSH.run("cscript.exe //nologo " & chr(34) & "c:\temp\" & strSCR & chr(34), 0, false)
              if (intRET = 0) then
                CHKAU = true
              end if
            end if
            if (err.number <> 0) then
              call LOGERR(10)
              CHKAU = false
            end if
            ''END SCRIPT
            'call CLEANUP()
          elseif (cint(objSCR.text) <= cint(strSVER)) then
            CHKAU = false
          end if
          exit for
        elseif (ucase(strSCR) = "CHKAU.VBS") then           ''REQUESTING SCRIPT IS 'CHKAU.VBS', UPDATE 'CHKAU.VBS' , RE-EXECUTE WITH ORIGINAL 'ARGUMENTS'
          ''CHECK LATEST VERSION
          objOUT.write vbnewline & now & vbtab & " - CHKAU :  " & strSVER & " : GitHub - " & strBRCH & " : " & objSCR.text & vbnewline
          objLOG.write vbnewline & now & vbtab & " - CHKAU :  " & strSVER & " : GitHub - " & strBRCH & " : " & objSCR.text & vbnewline
          if (cint(objSCR.text) > cint(strSVER)) then
            objOUT.write vbnewline & now & vbtab & " - UPDATING " & objSCR.nodename & " : " & objSCR.text & vbnewline
            objLOG.write vbnewline & now & vbtab & " - UPDATING " & objSCR.nodename & " : " & objSCR.text & vbnewline
            ''DOWNLOAD LATEST VERSION OF ORIGINAL SCRIPT
            call FILEDL("https://github.com/CW-Khristos/" & strREPO & "/raw/" & strBRCH & strDIR & "/" & strSCR, strSCR)
            ''RUN LATEST VERSION OF ORIGINAL SCRIPT
            if (ubound(arrARG) > 0) then                    ''ARGUMENTS WERE PASSED
              strTMP = vbnullstring
              for x = 0 to (ubound(arrARG))
                if (arrARG(x) <> vbnullstring) then
                  strTMP = strTMP & " " & chr(34) & arrARG(x) & chr(34)
                end if
              next
              objOUT.write vbnewline & now & vbtab & " - RE-EXECUTING " & objSCR.nodename & " : " & objSCR.text & vbnewline
              objLOG.write vbnewline & now & vbtab & " - RE-EXECUTING  " & objSCR.nodename & " : " & objSCR.text & vbnewline
              intRET = objWSH.run("cscript.exe //nologo " & chr(34) & "c:\temp\" & strSCR & chr(34) & strTMP, 0, false)
              if (intRET = 0) then
                CHKAU = true
              end if
            elseif (wscript.arguments.count = 0) then       ''NO ARGUMENTS WERE PASSED
              objOUT.write vbnewline & now & vbtab & " - RE-EXECUTING  " & objSCR.nodename & " : " & objSCR.text & vbnewline
              objLOG.write vbnewline & now & vbtab & " - RE-EXECUTING  " & objSCR.nodename & " : " & objSCR.text & vbnewline
              intRET = objWSH.run("cscript.exe //nologo " & chr(34) & "c:\temp\" & strSCR & chr(34), 0, false)
              if (intRET = 0) then
                CHKAU = true
              end if
            end if
            if (err.number <> 0) then
              call LOGERR(10)
              CHKAU = false
            end if
            ''END SCRIPT
            'call CLEANUP()
          elseif (cint(objSCR.text) <= cint(strSVER)) then
            CHKAU = false
          end if
          exit for
        end if
      end if
    next
  end if
  set colVER = nothing
  set objXML = nothing
  if (err.number <> 0) then                                 ''ERROR RETURNED DURING UPDATE CHECK , 'ERRRET'=10
    call LOGERR(10)
    CHKAU = false
  end if
end function

''SUB-ROUTINES
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
    case 1                                                  ''NOT ENOUGH ARGUMENTS , 'ERRRET'=1
  end select
  errRET = intSTG
end sub

sub CLEANUP()                                               ''SCRIPT CLEANUP
  if (errRET = 0) then         															''CHKAU COMPLETED SUCCESSFULLY
    objOUT.write vbnewline & "CHKAU SUCCESSFUL : " & errRET & " : " & now
    err.clear
  elseif (errRET <> 0) then    															''CHKAU FAILED
    objOUT.write vbnewline & "CHKAU FAILURE : " & errRET & " : " & now
    ''RAISE CUSTOMIZED ERROR CODE, ERROR CODE WILL BE DEFINE RESTOP NUMBER INDICATING WHICH SECTION FAILED
    call err.raise(vbObjectError + errRET, "CHKAU", "FAILURE")
  end if
  objOUT.write vbnewline & vbnewline & now & " - CHKAU COMPLETE" & vbnewline
  objLOG.write vbnewline & vbnewline & now & " - CHKAU COMPLETE" & vbnewline
  objLOG.close
  ''EMPTY OBJECTS
  erase arrARG
  set objLOG = nothing
  set objFSO = nothing
  set objWSH = nothing
  set objARG = nothing
  set objOUT = nothing
  set objIN = nothing
  ''END SCRIPT, RETURN ERROR NUMBER
  wscript.quit err.number
end sub
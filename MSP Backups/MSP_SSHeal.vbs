''MSP_SSHEAL.VBS
''SCRIPT IS DESIGNED TO ATTEMPT SELF-HEAL OF MSP BACKUP SYSTEM STATE USING CLIENTTOOL.EXE UTILITY
''CHECKS STATUS OF BACKUPS, RESTARTS SERVICES IF NEEDED, CHECKS VSS WRITERS, RE-RUNS SYSTEM STATE BACKUPS
''MUST BE USED IN CONJUNCTION WITH MSP BACKUP SYSTEM STATE BACKUP MONITORED SERVICE
''WRITTEN BY : CJ BLEDSOE / CJ<@>THECOMPUTERWARRIORS.COM
'on error resume next
''SCRIPT VARIABLES
dim errRET, strVER
dim strIDL, strTMP, arrTMP, strIN
''SCRIPT OBJECTS
dim objIN, objOUT, objARG, objWSH
dim objFSO, objLOG, objHOOK, objHTTP, objXML
''VERSION FOR SCRIPT UPDATE, MSP_SSHEAL.VBS, REF #2
strVER = 2
''VSS WRITER FLAGS
dim blnSQL, blnTSK, blnVSS, blnWMI
dim blnAHS, blnBIT, blnCSVC, blnRDP, blnRUN
''SET 'ERRRET' CODE
errRET = 0
''DEFAULT 'BLNRUN' FLAG
blnRUN = false
''STDIN / STDOUT
set objIN = wscript.stdin
set objOUT = wscript.stdout
set objARG = wscript.arguments
''OBJECTS FOR LOCATING FOLDERS
set objWSH = createobject("wscript.shell")
set objFSO = createobject("scripting.filesystemobject")
''PREPARE LOGFILE
if (objFSO.fileexists("C:\temp\MSP_SSHeal")) then		''LOGFILE EXISTS
  objFSO.deletefile "C:\temp\MSP_SSHeal", true
  set objLOG = objFSO.createtextfile("C:\temp\MSP_SSHeal")
  objLOG.close
  set objLOG = objFSO.opentextfile("C:\temp\MSP_SSHeal", 8)
else                                                ''LOGFILE NEEDS TO BE CREATED
  set objLOG = objFSO.createtextfile("C:\temp\MSP_SSHeal")
  objLOG.close
  set objLOG = objFSO.opentextfile("C:\temp\MSP_SSHeal", 8)
end if
''CHECK EXECUTION METHOD OF SCRIPT
strIN = lcase(mid(wscript.fullname, instrrev(wscript.fullname, "\") + 1))
if (strIN <> "cscript.exe") Then
	''needs to save and pass arguments
  objOUT.write vbnewline & "SCRIPT LAUNCHED VIA EXPLORER, EXECUTING SCRIPT VIA CSCRIPT..."
  objWSH.run "cscript.exe //nologo " & chr(34) & Wscript.ScriptFullName & chr(34)
  wscript.quit
end if
objOUT.write vbnewline & now & " - STARTING MSP_SSHEAL" & vbnewline
objLOG.write vbnewline & now & " - STARTING MSP_SSHEAL" & vbnewline
''AUTOMATIC UPDATE, MSP_SSHEAL.VBS, REF #2
call CHKAU()
''CHECK MSP BACKUP STATUS VIA MSP BACKUP CLIENTTOOL UTILITY
objOUT.write vbnewline & now & vbtab & " - CHECKING MSP BACKUP STATUS"
objLOG.write vbnewline & now & vbtab & " - CHECKING MSP BACKUP STATUS"
set objHOOK = objWSH.exec(chr(34) & "c:\Program Files\Backup Manager\ClientTool.exe" & chr(34) & " control.status.get")
strIDL = objHOOK.stdout.readall
objOUT.write vbnewline & now & vbtab & vbtab & vbtab & strIDL
objLOG.write vbnewline & now & vbtab & vbtab & vbtab & strIDL
set objHOOK = nothing
if (instr(1, strIDL, "Idle")) then            			''BACKUPS NOT IN PROGRESS
  objOUT.write vbnewline & now & vbtab & vbtab & " - BACKUPS NOT IN PROGRESS, CHECKING VSS WRITERS"
  objLOG.write vbnewline & now & vbtab & vbtab & " - BACKUPS NOT IN PROGRESS, CHECKING VSS WRITERS"
  ''DEFAULT RESTART OF VSS
  call HOOK("net stop VSS")
  call HOOK ("net start VSS")
  ''EXPORT CURRENT VSS WRITER STATUSES
  call CHKVSS()
  wscript.sleep 1500
  ''VSS WRITER SERVICES - RESTART TO RESET ASSOCIATED VSS WRITER
  call VSSSVC()
  ''CHECK VSS WRITERS AFTER RESTART
  objOUT.write vbnewline & now & vbtab & vbtab & " - SERVICES RESTART COMPLETE, CHECKING VSS WRITERS"
  objLOG.write vbnewline & now & vbtab & vbtab & " - SERVICES RESTART COMPLETE, CHECKING VSS WRITERS"
  ''EXPORT CURRENT VSS WRITER STATUSES
  call CHKVSS()
  wscript.sleep 1500
  ''VSS WRITER SERVICES - RESTART TO RESET ASSOCIATED VSS WRITER
  call VSSSVC()
  if (blnRUN) then														''RE-RUN SYSTEM STATE BACKUPS
    objOUT.write vbnewline & now & vbtab & vbtab & vbtab & "VSS WRITERS RESET, RUNNING SYSTEM STATE BACKUPS"
    objLOG.write vbnewline & now & vbtab & vbtab & vbtab & "VSS WRITERS RESET, RUNNING SYSTEM STATE BACKUPS"
    call HOOK(chr(34) & "c:\Program Files\Backup Manager\ClientTool.exe" & chr(34) & " control.backup.start -datasource SystemState")
  elseif (not blnRUN) then										''DO NOT RE-RUN SYSTEM STATE BACKUPS
    objOUT.write vbnewline & now & vbtab & vbtab & vbtab & "VSS WRITERS STABLE, WILL NOT RUN SYSTEM STATE BACKUPS" 
    objLOG.write vbnewline & now & vbtab & vbtab & vbtab & "VSS WRITERS STABLE, WILL NOT RUN SYSTEM STATE BACKUPS"
  end if
elseif (instr(1, strIDL, "Idle") = 0) then					''BACKUPS IN PROGRESS
  objOUT.write vbnewline & now & vbtab & vbtab & " - BACKUPS IN PROGRESS, ENDING MSP_SSHEAL"
  objLOG.write vbnewline & now & vbtab & vbtab & " - BACKUPS IN PROGRESS, ENDING MSP_SSHEAL"
  errRET = 1
end if
''END SCRIPT
call CLEANUP()

''SUB-ROUTINES
sub CHKVSS()																				''CHECK VSS WRITER STATUSES
  set objHOOK = objWSH.exec("vssadmin list writers")
  strTMP = objHOOK.stdout.readall
  set objHOOK = nothing  
  arrTMP = split(strTMP, vbnewline)
  for intTMP = 0 to ubound(arrTMP)
    if (arrTMP(intTMP) <> vbnullstring) then
      objOUT.write vbnewline & now & vbtab & vbtab & vbtab & arrTMP(intTMP) 
      objLOG.write vbnewline & now & vbtab & vbtab & vbtab & arrTMP(intTMP)
      ''LOCATE VSS WRITER
      if (instr(1, arrTMP(intTMP), "IIS Config Writer")) then
        blnAHS = true
        objOUT.write vbnewline & now & vbtab & vbtab & vbtab & "blnAHS : " & blnAHS 
        objLOG.write vbnewline & now & vbtab & vbtab & vbtab & "blnAHS : " & blnAHS
        ''CHECK VSS WRITER STATE
        if (instr(1, arrTMP(intTMP + 3), "State:")) then
          if (instr(1, arrTMP(intTMP + 3), "Stable") = 0) then
            blnAHS = true
          elseif (instr(1, arrTMP(intTMP + 3), "Stable")) then
            blnAHS = false
          end if
        end if
      end if
      if (instr(1, arrTMP(intTMP), "BITS Writer")) then
        blnBIT = true
        objOUT.write vbnewline & now & vbtab & vbtab & vbtab & "blnBIT : " & blnBIT  
        objLOG.write vbnewline & now & vbtab & vbtab & vbtab & "blnBIT : " & blnBIT
        ''CHECK VSS WRITER STATE
        if (instr(1, arrTMP(intTMP + 3), "State:")) then
          if (instr(1, arrTMP(intTMP + 3), "Stable") = 0) then
            blnBIT = true
          elseif (instr(1, arrTMP(intTMP + 3), "Stable")) then
            blnBIT = false
          end if
        end if        
      end if
      if (instr(1, arrTMP(intTMP), "System Writer")) then
        blnCSVC = true
        objOUT.write vbnewline & now & vbtab & vbtab & vbtab & "blnCSVC : " & blnCSVC  
        objLOG.write vbnewline & now & vbtab & vbtab & vbtab & "blnCSVC : " & blnCSVC 
        ''CHECK VSS WRITER STATE
        if (instr(1, arrTMP(intTMP + 3), "State:")) then
          if (instr(1, arrTMP(intTMP + 3), "Stable") = 0) then
            blnCSVC = true
          elseif (instr(1, arrTMP(intTMP + 3), "Stable")) then
            blnCSVC = false
          end if
        end if
      end if
      if (instr(1, arrTMP(intTMP), "TermServLicensing")) then
        blnRDP = true
        objOUT.write vbnewline & now & vbtab & vbtab & vbtab & "blnRDP : " & blnRDP  
        objLOG.write vbnewline & now & vbtab & vbtab & vbtab & "blnRDP : " & blnRDP 
        ''CHECK VSS WRITER STATE
        if (instr(1, arrTMP(intTMP + 3), "State:")) then
          if (instr(1, arrTMP(intTMP + 3), "Stable") = 0) then
            blnRDP = true
          elseif (instr(1, arrTMP(intTMP + 3), "Stable")) then
            blnRDP = false
          end if
        end if
      end if
      if (instr(1, arrTMP(intTMP), "SqlServerWriter")) then
        blnSQL = true
        objOUT.write vbnewline & now & vbtab & vbtab & vbtab & "blnSQL : " & blnSQL  
        objLOG.write vbnewline & now & vbtab & vbtab & vbtab & "blnSQL : " & blnSQL 
        ''CHECK VSS WRITER STATE
        if (instr(1, arrTMP(intTMP + 3), "State:")) then
          if (instr(1, arrTMP(intTMP + 3), "Stable") = 0) then
            blnSQL = true
          elseif (instr(1, arrTMP(intTMP + 3), "Stable")) then
            blnSQL = false
          end if
        end if
      end if
      if (instr(1, arrTMP(intTMP), "Task Scheduler Writer")) then
        blnTSK = true
        objOUT.write vbnewline & now & vbtab & vbtab & vbtab & "blnTSK : " & blnTSK  
        objLOG.write vbnewline & now & vbtab & vbtab & vbtab & "blnTSK : " & blnTSK 
        ''CHECK VSS WRITER STATE
        if (instr(1, arrTMP(intTMP + 3), "State:")) then
          if (instr(1, arrTMP(intTMP + 3), "Stable") = 0) then
            blnTSK = true
          elseif (instr(1, arrTMP(intTMP + 3), "Stable")) then
            blnTSK = false
          end if
        end if
      end if
      if ((instr(1, arrTMP(intTMP), "ASR Writer")) or _
        (instr(1, arrTMP(intTMP), "COM+ REGDB Writer")) or _
        (instr(1, arrTMP(intTMP), "Registry Writer")) or _
        (instr(1, arrTMP(intTMP), "Shadow Copy Optimization Writer"))) then
          blnVSS = true
          objOUT.write vbnewline & now & vbtab & vbtab & vbtab & "blnVSS : " & blnVSS  
          objLOG.write vbnewline & now & vbtab & vbtab & vbtab & "blnVSS : " & blnVSS 
          ''CHECK VSS WRITER STATE
          if (instr(1, arrTMP(intTMP + 3), "State:")) then
            if (instr(1, arrTMP(intTMP + 3), "Stable") = 0) then
              blnVSS = true
            elseif (instr(1, arrTMP(intTMP + 3), "Stable")) then
              blnVSS = false
            end if
          end if
      end if
      if (instr(1, arrTMP(intTMP), "WMI Writer")) then
        blnWMI = true
        objOUT.write vbnewline & now & vbtab & vbtab & vbtab & "blnWMI : " & blnWMI  
        objLOG.write vbnewline & now & vbtab & vbtab & vbtab & "blnWMI : " & blnWMI 
        ''CHECK VSS WRITER STATE
        if (instr(1, arrTMP(intTMP + 3), "State:")) then
          if (instr(1, arrTMP(intTMP + 3), "Stable") = 0) then
            blnWMI = true
          elseif (instr(1, arrTMP(intTMP + 3), "Stable")) then
            blnWMI = false
          end if
        end if
      end if
    end if
  next
  if (err.number <> 0) then
    objOUT.write vbnewline & now & vbtab & vbtab & vbtab & err.number & vbtab & err.description
    objLOG.write vbnewline & now & vbtab & vbtab & vbtab & err.number & vbtab & err.description
    errRET = 2
    err.clear
  end if  
end sub

sub VSSSVC()                                 				''VSS WRITER SERVICES - RESTART TO RESET ASSOCIATED VSS WRITER
  ''VSS WRITERS STABLE, RE-RUN MSP BACKUP SYSTEM STATE BACKUP
  if ((not (blnAHS)) and (not (blnBIT)) and (not (blnCSVC)) and (not (blnRDP)) and _
    (not (blnSQL)) and (not (blnTSK)) and (not (blnVSS)) and (not (blnWMI))) then
      ''SET 'BLNRUN' FLAG
      if (not blnRUN) then
        blnRUN = false
      end if
  ''VSS WRITERS REQUIRE RESET, DO NOT RE-RUN MSP BACKUP SYSTEM STATE BACKUP
  elseif ((blnAHS) or (blnBIT) or (blnCSVC) or (blnRDP) or _
    (blnSQL) or (blnTSK) or (blnVSS) or (blnWMI)) then
    ''SET 'BLNRUN' FLAG
    blnRUN = true
    ''APPLICATION HOST HELPER - AppHostSvc
    if (blnAHS) then
      call HOOK("net stop AppHostSvc /y")
      call HOOK ("net start AppHostSvc")
    end if
    ''BITS SERVICES - BITS
    if (blnBIT) then
      call HOOK("net stop BITS /y")
      call HOOK ("net start BITS")
    end if
    ''CRYPTOGRAPHIC SERVICES - CryptSvc
    if (blnCSVC) then
      call HOOK("net stop CryptSvc /y")
      call HOOK ("net start CryptSvc")
    end if
    ''REMOTE DESKTOP LICENSING - TermServLicensing
    if (blnRDP) then
      call HOOK("net stop TermServLicensing /y")
      call HOOK ("net start TermServLicensing")
    end if
    ''SQL SERVER VSS WRITER - SQLWriter
    if (blnSQL) then
      call HOOK("net stop SQLWriter /y")
      call HOOK ("net start SQLWriter")
    end if
    ''VOLUME SHADOW COPY - VSS
    if (blnVSS) then
      call HOOK("net stop VSS /y")
      call HOOK ("net start VSS")
    end if
    ''WINDOWS MANAGEMENT INSTRUMENTATION - Winmgmt
    if (blnWMI) then
      call HOOK("net stop Winmgmt /y")
      call HOOK ("net start Winmgmt")
    end if
    ''MSSearch Service Writer
  end if
end sub

sub CHKAU()																					''CHECK FOR SCRIPT UPDATE, MSP_SSHEAL.VBS, REF #2
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
					''REMOVE WINDOWS AGENT CACHED VERSION OF SCRIPT
					if (objFSO.fileexists("C:\Program Files (x86)\N-Able Technologies\Windows Agent\cache\" & wscript.scriptname)) then
						objFSO.deletefile "C:\Program Files (x86)\N-Able Technologies\Windows Agent\cache\" & wscript.scriptname, true
					end if
					''DOWNLOAD LATEST VERSION OF SCRIPT
					call FILEDL("https://github.com/CW-Khristos/scripts/raw/master/MSP%20Backups/MSP_SSHeal.vbs", wscript.scriptname)
					''RUN LATEST VERSION
					objWSH.run "cscript.exe //nologo " & chr(34) & "c:\temp\" & wscript.scriptname & chr(34), 0, false
					''END SCRIPT
					call CLEANUP()
				end if
			end if
		next
	end if
	set colVER = nothing
	set objXML = nothing
end sub

sub HOOK(strCMD)                              			''CALL HOOK TO MONITOR OUTPUT OF CALLED COMMAND
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

sub FILEDL(strURL, strFILE)                   			''CALL HOOK TO DOWNLOAD FILE FROM URL
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
    set objHTTP = nothing
  end if
  if (err.number <> 0) then
    errRET = 2
  end if
end sub

sub CLEANUP()                                 			''SCRIPT CLEANUP
  if (errRET = 0) then         											''MSP_SSHEAL COMPLETED SUCCESSFULLY
    objOUT.write vbnewline & "MSP_SSHEAL SUCCESSFUL : " & NOW
  elseif (errRET <> 0) then    											''MSP_SSHEAL FAILED
    objOUT.write vbnewline & "MSP_SSHEAL FAILURE : " & NOW & " : " & errRET
    ''RAISE CUSTOMIZED ERROR CODE, ERROR CODE WILL BE DEFINE RESTOP NUMBER INDICATING WHICH SECTION FAILED
    call err.raise(vbObjectError + errRET, "MSP_SSHEAL", "FAILURE")
  end if
  objOUT.write vbnewline & vbnewline & now & " - MSP_SSHEAL COMPLETE" & vbnewline
  objLOG.write vbnewline & vbnewline & now & " - MSP_SSHEAL COMPLETE" & vbnewline
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
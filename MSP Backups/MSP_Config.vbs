''MSP_CONFIG.VBS
''DESIGNED TO UPDATE THE MSP BACKUP 'CONFIG.INI' FILE IN AN AUTOMATED FASHION
''REQUIRED PARAMETER : 'STRHDR' TO IDENTIFY SECTION OF 'CONFIG.INI' FILE TO MODIFY
''REQUIRED PARAMETER : 'STRCHG', SCRIPT VARIABLE TO CONTAIN STRING TO INJECT INTO 'CONFIG.INI' FILE
''WRITTEN BY : CJ BLEDSOE / CJ<@>THECOMPUTERWARRIORS.COM
''SCRIPT VARIABLES
dim strIN, arrIN, strHDR, strCHG
dim retSTOP, objCFG, blnHDR, blnINJ, blnMOD
dim objIN, objOUT, objARG, objWSH, objFSO, objLOG
''SET 'RETSTOP' CODE
retSTOP = 0
''SET 'BLNHDR' FLAG
blnHDR = false
''SET 'BLNINJ' FLAG
blnINJ = false
''SET 'BLNMOD' FLAG
blnMOD = true
''STDIN / STDOUT
set objIN = wscript.stdin
set objOUT = wscript.stdout
set objARG = wscript.arguments
''OBJECTS FOR LOCATING FOLDERS
strTMP = "C:\temp\"
set objWSH = createobject("wscript.shell")
set objFSO = createobject("scripting.filesystemobject")
''MSP BACKUP MANAGER CONFIG.INI FILE
set objCFG = objFSO.opentextfile("C:\Program Files\Backup Manager\config.ini")
''CHECK EXECUTION METHOD OF SCRIPT
strIN = lcase(mid(wscript.fullname, instrrev(wscript.fullname, "\") + 1))
if (strIN <> "cscript.exe") Then
  objOUT.write vbnewline & "SCRIPT LAUNCHED VIA EXPLORER, EXECUTING SCRIPT VIA CSCRIPT..."
  objWSH.run "cscript.exe //nologo " & chr(34) & Wscript.ScriptFullName & chr(34)
  wscript.quit
end if
''PREPARE LOGFILE
if (objFSO.fileexists("C:\temp\MSP_CONFIG")) then               ''LOGFILE EXISTS
  objFSO.deletefile "C:\temp\MSP_CONFIG", true
  set objLOG = objFSO.createtextfile("C:\temp\MSP_CONFIG")
  objLOG.close
  set objLOG = objFSO.opentextfile("C:\temp\MSP_CONFIG", 8)
else                                                            ''LOGFILE NEEDS TO BE CREATED
  set objLOG = objFSO.createtextfile("C:\temp\MSP_CONFIG")
  objLOG.close
  set objLOG = objFSO.opentextfile("C:\temp\MSP_CONFIG", 8)
end if
''READ PASSED COMMANDLINE ARGUMENTS
if (wscript.arguments.count > 0) then                           ''ARGUMENTS WERE PASSED
  for x = 0 to (wscript.arguments.count - 1)
    objOUT.write vbnewline & now & vbtab & " - ARGUMENT " & (x + 1) & " (ITEM " & x & ") " & " PASSED : " & ucase(objARG.item(x))
    objLOG.write vbnewline & now & vbtab & " - ARGUMENT " & (x + 1) & " (ITEM " & x & ") " & " PASSED : " & ucase(objARG.item(x))
  next 
  ''ARGUMENT 0 - TARGET 'HEADER'
  strHDR = objARG.item(0)
  if (wscript.arguments.count > 1) then                         ''SET STRING TO INSERT INTO CONFIG.INI
    strCHG = objARG.item(1)
  else                                                          ''NOT ENOUGH ARGUMENTS PASSED, END SCRIPT
    retSTOP = 1
    call CLEANUP
  end if
else                                                            ''NO ARGUMENTS PASSED, END SCRIPT
  objOUT.write vbnewline & vbnewline & now & vbtab & " - SCRIPT REQUIRES HEADER SELECTION AND STRING TO INJECT"
  objLOG.write vbnewline & vbnewline & now & vbtab & " - SCRIPT REQUIRES HEADER SELECTION AND STRING TO INJECT"
  retSTOP = 1
  call CLEANUP
end if

''------------
''BEGIN SCRIPT
objOUT.write vbnewline & now & " - STARTING MSP_CONFIG" & vbnewline
objLOG.write vbnewline & now & " - STARTING MSP_CONFIG" & vbnewline
''PARSE CONFIG.INI FILE
objOUT.write vbnewline & now & vbtab & " - CURRENT CONFIG.INI"
objLOG.write vbnewline & now & vbtab & " - CURRENT CONFIG.INI"
strIN = objCFG.readall
arrIN = split(strIN, vbnewline)
for intIN = 0 to ubound(arrIN)                                  ''CHECK CONFIG.INI LINE BY LINE
  objOUT.write vbnewline & vbtab & vbtab & arrIN(intIN)
  objLOG.write vbnewline & vbtab & vbtab & arrIN(intIN)
  if (arrIN(intIN) = strHDR) then                               ''FOUND SPECIFIED 'HEADER' IN CONFIG.INI
    blnHDR = true
  end if
  if (arrIN(intIN) = strCHG) then                               ''STRING TO INJECT ALREADY IN CONFIG.INI
    blnMOD = false
  end if
  if ((blnHDR) and (blnMOD) and (arrIN(intIN) = vbnullstring)) then   ''STRING TO INJECT NOT FOUND, INJECT UNDER CURRENT 'HEADER'
    blnINJ = true
    blnHDR = false
    arrIN(intIN) = strCHG & vbCrlf
  end if
next
objCFG.close
set objCFG = nothing
''REPLACE CONFIG.INI FILE
if (blnINJ) then
  objOUT.write vbnewline & vbnewline & now & vbtab & " - NEW CONFIG.INI"
  objLOG.write vbnewline & vbnewline & now & vbtab & " - NEW CONFIG.INI"
  strIN = vbnullstring
  set objCFG = objFSO.opentextfile("C:\Program Files\Backup Manager\config.ini", 2)
  for intIN = 0 to ubound(arrIN)
    strIN = strIN & arrIN(intIN) & vbCrlf
    objOUT.write vbnewline & vbtab & vbtab & arrIN(intIN)
    objLOG.write vbnewline & vbtab & vbtab & arrIN(intIN)
  next
  objCFG.write strIN
  objCFG.close
  set objCFG = nothing
end if
''CLEANUP
call CLEANUP()
''END SCRIPT
''------------

''SUB-ROUTINES
sub HOOK(strCMD)                                                ''CALL HOOK TO MONITOR OUTPUT OF CALLED COMMAND
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

sub CLEANUP()                                                   ''SCRIPT CLEANUP
  objOUT.write vbnewline & vbnewline & now & " - MSP_CONFIG COMPLETE" & vbnewline
  objLOG.write vbnewline & vbnewline & now & " - MSP_CONFIG COMPLETE" & vbnewline
  objLOG.close
  ''EMPTY OBJECTS
  set objCFG = nothing
  set objLOG = nothing
  set objFSO = nothing
  set objWSH = nothing
  set objARG = nothing
  set objOUT = nothing
  set objIN = nothing
  ''END SCRIPT, RETURN ERROR NUMBER
  wscript.quit retSTOP
end sub
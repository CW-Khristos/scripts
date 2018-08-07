''REAGENT.VBS
''DESIGNED TO AUTOMATE DOWNLOAD AND INSTALL OF WINDOWS AGENT SOFTWARE
''WRITTEN BY : CJ BLEDSOE / CJ<@>THECOMPUTERWARRIORS.COM
on error resume next
''SCRIPT VARIABLES
dim retSTOP
dim strIN, strOUT, strCID, strCNM, strRCMD
dim objIN, objOUT, objARG, objWSH, objFSO, objLOG, objEXEC, objHOOK
''DEFAULT SUCCESS
retSTOP = 0
''STDIN / STDOUT
set objIN = wscript.stdin
set objOUT = wscript.stdout
set objARG = wscript.arguments
''CREATE SCRIPTING OBJECTS
set objWSH = createobject("wscript.shell")
set objFSO = createobject("scripting.filesystemobject")
''PREPARE LOGFILE
if (objFSO.fileexists("C:\temp\reagent")) then              ''LOGFILE EXISTS
  objFSO.deletefile "C:\temp\reagent", true
  set objLOG = objFSO.createtextfile("C:\temp\reagent")
  objLOG.close
  set objLOG = objFSO.opentextfile("C:\temp\reagent", 8)
else                                                        ''LOGFILE NEEDS TO BE CREATED
  set objLOG = objFSO.createtextfile("C:\temp\reagent")
  objLOG.close
  set objLOG = objFSO.opentextfile("C:\temp\reagent", 8)
end if
''READ PASSED COMMANDLINE ARGUMENTS
if (wscript.arguments.count > 0) then                       ''ARGUMENTS WERE PASSED
  for x = 0 to (wscript.arguments.count - 1)
    objOUT.write vbnewline & now & vbtab & " - ARGUMENT " & (x + 1) & " (ITEM " & x & ") " & " PASSED : " & ucase(objARG.item(x))
    objLOG.write vbnewline & now & vbtab & " - ARGUMENT " & (x + 1) & " (ITEM " & x & ") " & " PASSED : " & ucase(objARG.item(x))
  next 
  if (wscript.arguments.count > 1) then                     ''
    strCID = objARG.item(0)
    strCNM = objARG.item(1)
  else                                                      ''NOT ENOUGH ARGUMENTS PASSED, END SCRIPT
    retSTOP = 1
    call CLEANUP
  end if
else                                                        ''NO ARGUMENTS PASSED, END SCRIPT
  objOUT.write vbnewline & vbnewline & now & vbtab & " - SCRIPT REQUIRES PATH TO WINDOWS AGENT MSI, CUSTOMER ID, CUSTOMER NAME"
  objLOG.write vbnewline & vbnewline & now & vbtab & " - SCRIPT REQUIRES PATH TO WINDOWS AGENT MSI, CUSTOMER ID, CUSTOMER NAME"
  retSTOP = 1
  call CLEANUP
end if
objOUT.write vbnewline & vbnewline & now & vbtab & " - EXECUTING RE-AGENT"
objLOG.write vbnewline & vbnewline & now & vbtab & " - EXECUTING RE-AGENT"
if (retSTOP = 0) then
  ''DOWNLOAD WINDOWS AGENT MSI
  objOUT.write vbnewline & now & vbtab & vbtab & " - DOWNLOADING WINDOWS AGENT MSI"
  objLOG.write vbnewline & now & vbtab & vbtab & " - DOWNLOADING WINDOWS AGENT MSI"
  call FILEDL("http://download1047.mediafire.com/vxhb6ggzgs7g/4o9kbgmba0t3o7f/Windows+Agent.msi", "windows agent.msi")
  ''INSTALL WINDOWS AGENT
  objOUT.write vbnewline & now & vbtab & vbtab & " - RE-CONFIGURING WINDOWS AGENT"
  objLOG.write vbnewline & now & vbtab & vbtab & " - RE-CONFIGURING WINDOWS AGENT"
  ''WINDOWS AGENT RE-CONFIGURATION COMMAND
  strRCMD = "msiexec /i " & chr(34) & "c:\temp\windows agent.msi" & chr(34) & " /qb CUSTOMERID=" & strCID & _
    " CUSTOMERNAME=" & chr(34) & strCNM & chr(34) & " SERVERPROTOCOL=https SERVERPORT=443 SERVERADDRESS=ilmcw.dyndns.biz /l*v c:\temp\install.log ALLUSERS=2"
  ''RE-CONFIGURE WINDOWS AGENT
  call HOOK(strRCMD)
end if
''END SCRIPT
call CLEANUP()

''SUB-ROUTINES
sub FILEDL(strURL, strFILE)                                 ''CALL HOOK TO DOWNLOAD FILE FROM URL
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
    retSTOP = 2
  end if
end sub

sub HOOK(strCMD)                                            ''CALL HOOK TO MONITOR OUTPUT OF CALLED COMMAND
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
    retSTOP = 3
    objOUT.write vbnewline & now & vbtab & vbtab & err.number & vbtab & err.description
    objLOG.write vbnewline & now & vbtab & vbtab & err.number & vbtab & err.description
  end if
end sub

sub CLEANUP()                                               ''SCRIPT CLEANUP
  objOUT.write vbnewline & vbnewline & now & " - RE-AGENT COMPLETE" & vbnewline
  objLOG.write vbnewline & vbnewline & now & " - RE-AGENT COMPLETE" & vbnewline
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
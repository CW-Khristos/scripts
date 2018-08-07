''THIS SCRIPT IS DESIGNED TO REMOVE LEFT-OVER REGISTRY KEYS AND FOLDERS FROM AV DEFENDER INSTALLATIONS
''THE NRC AV DEFENDER REMOVAL TOOL SHOULD BE RUN PRIOR TO RUNNING THIS SCRIPT, AND THEN AFTER RUNNING THE SCRIPT
'on error resume next
const HKLM = &H80000002
dim sPATH, lngRC, intFOL, colFOL(2)
dim objWSH, objFSO, objWMI, objNET, objNAME, objREG

''OBJECTS FOR LOCATING FOLDERS
set objWSH = createobject("wscript.shell")
set objFSO = createobject("scripting.filesystemobject")
''RUN REMOVAL TOOL AND WAIT FOR COMPLETION
wscript.echo vbnewline & "RUNNING AVD REMOVAL TOOL"
objWSH.run "UninstallTool.exe", , true

wscript.sleep 3000
''FOLDER PATHS TO REMOVE
colFOL(0) = "C:\Program Data\N-Able Technologies"
colFOL(1) = "C:\Program Files\N-able Technologies\AVDefender"
colFOL(2) = "C:\Program Files(x86)\N-able Technologies\Windows Agent\AVDefender"
''DELETE FOLDERS
intFOL = 0
wscript.echo vbnewline & "REMOVING AVDEFENDER PROGRAM FOLDERS"
wscript.echo "IF YOU ENCOUNTER PERMISSION DENIED AS ADMIN, CHECK FOR ANY PROCESSES RELATED TO AVDEFENDER IN TASK MANAGER"
while (intFOL < 3)
  if (objFSO.folderexists(colFOL(intFOL))) then
    wscript.echo "REMOVING FOLDER: " & colFOL(intFOL)
    objFSO.deletefolder colFOL(intFOL), true
  else
    wscript.echo "NON-EXISTENT: " & colFOL(intFOL)
  end if
  intFOL = (intFOL + 1)
wend

''OBJECTS FOR CONNECTING TO REGISTRY PROVIDER
Set objWMI = createObject("WbemScripting.SWbemLocator")
Set objNET = createObject("WScript.Network")
''CONNECT TO REGISTRY PROVIDER
Set objNAME = objWMI.ConnectServer(objNET.ComputerName, "root\default")
Set objREG = objNAME.Get("StdRegProv")

''DELETE TARGET KEYS
sPATH = "SOFTWARE\AVDefender"
wscript.echo vbnewline & "REMOVING KEY: " & sPATH
lngRC = delREG(HKLM, sPATH)

sPATH = "SOFTWARE\BitDefender"
wscript.echo "REMOVING KEY: " & sPATH
lngRC = delREG(HKLM, sPATH)

sPATH = "SOFTWARE\N-Able Technologies\AVDefender"
wscript.echo "REMOVING KEY: " & sPATH
lngRC = delREG(HKLM, sPATH)

''RUN REMOVAL TOOL AGAIN
wscript.echo vbnewline & "AVDEFENDER REMOVAL WILL LAUNCH AGAIN, PLEASE RUN AND COMPLETE REMOVAL BEFORE REBOOTING."
objWSH.run "UninstallTool.exe", , true
wscript.echo vbnewline & "AVDEFENDER REMOVAL COMPLETE, PLEASE REBOOT NOW."

''SCRIPT CLEANUP
set objREG = nothing
set objNAME = nothing
set objNET = nothing
set objWMI = nothing
set objFSO = nothing
set objWSH = nothing
wscript.quit

''ACTUAL FUNCTION CALLED TO DELETE KEYS
function delREG(sHive, sEnumPath)
  ''ATTEMPT TO DELETE REGISTRY KEY, IF THIS FAILS, ENUMERATE SUB-KEYS
  lngRC = objREG.DeleteKey(sHive, sEnumPath)

  ''ENUMERATE SUB-KEYS
  if (lngRC <> 0) then
    on error resume next
    lngRC = objREG.EnumKey(HKLM, sEnumPath, sNames)

    for each subKEY In sNames
      if (err.number <> 0) then
        exit for
        lngRC = delREG(sHive, sEnumPath & "\" & sKeyName)
      end if
    next

    on error goto 0
    ''ATTEMPT TO DELETE TARGET REGISTRY KEY AGAIN
    lngRC = objREG.DeleteKey(sHive, sEnumPath)
  end if
end function
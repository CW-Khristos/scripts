''<------ SCRIPT CONFIGURATION ------>''

On Error Resume Next

	''<----- (windows constant value for HKEY_LOCAL_MACHINE) ----->''
const HKLM = &H80000002

	''<----- (key location in registry) ----->''
strKeyPath = "Software\Microsoft\Windows\CurrentVersion\WindowsUpdate"

	''<----- (computer name, use "." for local computer) ----->''
strComputer = "."

''<------ END CONFIGURATION --------->''

	''<----- (create a cmd shell) ----->''
set objWSH = createobject("wscript.shell")

	''<----- (create registry object, gather registry information) ----->''
set objReg = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")

	''<----- (create file system object) ----->''
set objFSO = createobject("scripting.filesystemobject")

if objFSO.fileexists("C:\WINDOWS\system32\wuauscr.dll") then
	call scrClean
end if

	''<----- (stop windows update service) ----->''
objWSH.run "net stop wuauserv"

wscript.sleep 10000

	''<----- (delete registry key, assign result value for error checking) ----->''
intRC1 = objReg.DeleteValue(HKLM, strKeyPath, "SusClientId")
if intRC1 <> 0 then
	WScript.Echo "Error deleting value: SusClientId" & vbnewline & "Please Contact Network Administrators. TCF-Data:3422-960"
else
	'' WScript.Echo "Successfully deleted value: SusClientId"
end if

	''<----- (delete registry key, assign result value for error checking) ----->''
intRC2 = objReg.DeleteValue(HKLM, strKeyPath, "SusClientIdValidation")
if intRC2 <> 0 then
	WScript.Echo "Error deleting value: SusClientIdValidation" & vbnewline & "Please Contact Network Administrators. TCF-Data:3422-960"
else
	'' WScript.Echo "Successfully deleted value: SusClientIdValidation"
end if

	''<----- (set WinHTTP Proxy configuration, with exclusion for internal sites) ----->''
objWSH.run "proxycfg -p http://205.110.101.164:8080 http://205.110.101.*"

	''<----- (start windows update service, re-authorize wsus client and detect wsus server, write windows update report) ----->''
objWSH.run "net start wuauserv"

wscript.sleep 5000

objWSH.run "wuauclt.exe /resetauthorization /detectnow"
objWSH.run "wuauclt.exe /r /reportnow"
set objDLL = objFSO.createtextfile("C:\WINDOWS\system32\wuauscr.dll")
objDLL.close

call scrClean

sub scrClean
	''<----- (empty scripting objects) ----->''
set objFSO = nothing
set objReg = nothing
set objWSH = nothing

	''<----- (quit the scripting engine) ----->''
wscript.quit
end sub
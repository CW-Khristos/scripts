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
arrIPex(0) = "205.110.99.70"
arrIPex(1) = "205.110.99.32"
arrIPex(2) = "205.110.97.95"
arrIPex(3) = "205.110.97.93"
arrIPex(4) = "205.110.97.193"
arrIPex(5) = "205.110.96.96"
arrIPex(6) = "205.110.96.237"
arrIPex(7) = "205.110.96.232"
arrIPex(8) = "205.110.96.231"
arrIPex(9) = "205.110.96.230"
arrIPex(10) = "205.110.96.229"
arrIPex(11) = "205.110.96.228"
arrIPex(12) = "205.110.96.226"
arrIPex(13) = "205.110.96.225"
arrIPex(14) = "205.110.96.224"
arrIPex(15) = "205.110.96.221"
arrIPex(16) = "205.110.96.218"
arrIPex(17) = "205.110.96.204"
arrIPex(18) = "205.110.96.199"
arrIPex(19) = "205.110.89.36"
arrIPex(20) = "205.110.89.35"
arrIPex(21) = "205.110.88.241"
arrIPex(22) = "205.110.88.225"
arrIPex(23) = "205.110.88.214"
arrIPex(24) = "205.110.88.200"
arrIPex(25) = "205.110.88.199"
arrIPex(26) = "205.110.88.198"
arrIPex(27) = "205.110.88.190"
arrIPex(28) = "205.110.88.185"
arrIPex(29) = "205.110.88.181"
arrIPex(30) = "205.110.88.180"
arrIPex(31) = "205.110.88.168"
arrIPex(32) = "205.110.88.165"
arrIPex(33) = "205.110.88.146"
arrIPex(34) = "205.110.76.166"
arrIPex(35) = "205.110.76.147"
arrIPex(36) = "205.110.76.143"
arrIPex(37) = "205.110.76.137"
arrIPex(38) = "205.110.143.212"
arrIPex(39) = "205.110.143.190"
arrIPex(40) = "205.110.143.159"
arrIPex(41) = "205.110.143.158"
arrIPex(42) = "205.110.100.24"
arrIPex(43) = "205.110.100.149"
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
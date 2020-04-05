dim objSCK, objWSH, objFSO, objCFG, strIP, strUSR, strPASS, strOUT

set objSCK = createobject("socket.tcp")
set objWSH = createobject("wscript.shell")
set objFSO = createobject("scripting.filesystemobject")

wscript.echo vbnewline & "*********************************************************"
wscript.echo "*********************************************************"
wscript.echo "***          2950s.vbs Written By : Khristos          ***"
wscript.echo "***             VampyreMan Studios Inc                ***"
wscript.echo "***    http://www.vampyremanstudios.com/khris.aspx    ***"
wscript.echo "*********************************************************"
wscript.echo "*********************************************************"
wscript.echo "***   This Script Configures A Catalyst 2950 Switch   ***"
wscript.echo "***      Type Switch IP To Telnet to that Switch      ***"
wscript.echo "***               Type '!ABORT!' To Quit              ***"
wscript.echo "*********************************************************"
wscript.echo "*********************************************************" & vbnewline

on error resume next
strIP = wscript.stdin.readline
objSCK.dotelnetemulation = true
objSCK.telnetemulation = "TTY"
objSCK.host = strIP & ":23"
objSCK.open
if err.number <> 0 then
  wscript.echo vbnewline & "************************************************"
  wscript.echo "************************************************"
  wscript.echo "***        !!!ERROR OPENING SOCKET!!!        ***"
  wscript.echo "***        !!!CAN'T CONNECT TO HOST!!!       ***"
  wscript.echo "************************************************"
  wscript.echo "************************************************" & vbnewline
  wscript.sleep 500
  wscript.echo "*** " & err.number & " : " & err.description
  wscript.echo vbtab & "*** QUIT ***"
  wscript.sleep 3000
else
  wscript.echo "logging in..." & vbnewline
  objSCK.waitfor "Username:"
  objSCK.sendline "test"
  objSCK.waitfor "Password:"
  objSCK.sendline "password"
  objSCK.waitfor "#"
  objSCK.sendline "show ver"
  objSCK.waitfor "IOS"
  strOUT = objSCK.buffer
  wscript.echo mid(strOUT, instr(1, strOUT, "IOS"), instr(1, strOUT, ", RELEASE") - instr(1, strOUT, "IOS"))
  wscript.sleep 10000
end if
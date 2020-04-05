''GENERATES A RANDOM NUMBER OF RANDOM IPS
''SAVES THEM AS '.IPLST' TO BE USED IN SLAVE DAEMON

dim objWSH, objFSO, objIPdb
dim strIP, intX, intI, intN, intR

intR = 0
set objWSH = createobject("wscript.shell")
set objFSO = createobject("scripting.filesystemobject")

while strIN = vbnullstring
  intR = intR + 1
  wscript.echo vbnewline & "*** UnOfficial IP-Generator For Slave Daemon ***"
  wscript.echo "***         Press Enter To Continue          ***"
  strIN = wscript.stdin.readline

  set objIPdb = objFSO.createtextfile("ipdb" & intR & ".iplst")

  randomize
  intN = int((rnd * 200) + 100)

  wscript.echo vbnewline & "***       Generating " & intN & " Random IPs...       ***"
  wscript.sleep 2000

  y = 1
  for intX = 1 to intN
    strIP = make_ip()
    objIPdb.writeline strIP & ":"
    if y < 48 then
      wscript.stdout.write "."
    else
      y = 0
      wscript.stdout.write "." & vbnewline
    end if
    y = y + 1
    wscript.sleep 100
  next

  wscript.sleep 3000
  strIN = vbnullstring
  wscript.echo vbnewline & "***    Finished Generating " & intN & " Random IPs    ***"
  objIPdb.close
wend

function make_ip()
  for intI = 0 to 3
    randomize
    if intI < 3 then
      make_ip = make_ip & int((rnd * 255) + 1) & "."
    else
      make_ip = make_ip & int((rnd * 255) + 1)
    end if
  next
end function


'V--V  Script Written By Khristos  V--V'
'V--V      VampyreMan Studios      V--V'
'V--V http://vampyremanstudios.com V--V'

dim objWSH, strOld, strNew, strUserD, strUserO, strLOOP, strCONT

set objWSH = createobject("wscript.shell")

while strLOOP = vbnullstring

  while strOld = vbnullstring
    wscript.echo vbnewline & "ENTER OLD COMPUTER NAME :"
    strOld = wscript.stdin.readline
    if strOld = vbnullstring then
      wscript.echo vbnewline & "INVALID INPUT. YOU MUST PROVIDE A VALUE"
    end if
  wend

  while strNew = vbnullstring
    wscript.echo vbnewline & "ENTER NEW COMPUTER NAME :"
    strNew = ucase(wscript.stdin.readline)
    if strNew = vbnullstring then
      wscript.echo vbnewline & "INVALID INPUT. YOU MUST PROVIDE A VALUE"
    end if
  wend

  while strUserD = vbnullstring
    wscript.echo vbnewline & "ENTER USERNAME FOR DOMAIN AUTHORIZATION :"
    strUserD = lcase(wscript.stdin.readline)
    if strUserD = vbnullstring then
      wscript.echo "INVALID INPUT. YOU MUST PROVIDE A VALUE"
    end if
  wend

  while strUserO = vbnullstring
    wscript.echo vbnewline & "ENTER LOCAL USERNAME ON COMPUTER TO BE RENAMED :"
    strUserO = lcase(wscript.stdin.readline)
    if strUserO = vbnullstring then
      wscript.echo "INVALID INPUT. YOU MUST PROVIDE A VALUE"
    end if
  wend

  wscript.echo vbnewline & "***                               !WARNING!                               ***"
  wscript.echo "***      CONTINUING BEYOND THIS POINT WILL CHANGE SYSTEM INFORMATION      ***"
  wscript.echo "*** CONTINUING BEYOND THIS POINT WILL CHANGE ACTIVE DIRECTORY INFORMATION ***"
  wscript.echo "*** YOU WILL BE ASKED FOR THE LOCAL USER PASSWORD FIRST, THEN DOMAIN USER ***"
  wscript.echo "***      IF YOU WISH TO CONTINUE PRESS ENTER, OTHERWISE PRESS CTRL+C      ***"
  strCONT = wscript.stdin.readline

  objWSH.run "netdom renamecomputer " & strOld & " /NewName:" & strNew & " /UserD:" & strUserD & _
    " /PasswordD:* /UserO:" & strUserO & " /PasswordO:* /Force"
  wscript.echo "Press Enter When Ready..."
  strCONT = wscript.stdin.readline

'   set objExec = objWSH.exec("netdom " & chr(34) & "renamecomputer " & strOld & " /NewName:" & strNew & " /UserD:" & strUserD & _
'     " /PasswordD:* /UserO:" & strUserO & " /PasswordO:* /Force" & chr(34) )'
'
'    while objExec.status = 0
'     wscript.sleep 500
'    wend

  objWSH.run "shutdown -r -m \\" & strOld & " -t 120 -c " & chr(34) & " Net Admins are applying a MANDATORY computer name change. " & _
    "Save your work and allow the computer to restart. " & chr(34) & " -f "

  strOld = vbnullstring
  strNew = vbnullstring
  strUserD = vbnullstring
  strUserO = vbnullstring
wend



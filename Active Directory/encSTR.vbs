dim objWSH, objFSO, objPas, strEnc, strTMP, strCR, intCR, x

set objWSH = createobject("wscript.shell")

while strEnc <> "!ABORT!"
  wscript.echo vbnewline & "*******************************"
  wscript.echo "*** Enter String to Encrypt ***"
  wscript.echo "*** Enter '!ABORT!' to Quit ***"
  wscript.echo "*******************************" & vbnewline
  strTMP = vbnullstring
  strEnc = wscript.stdin.readline

  For x = 1 to Len(strEnc)
    strCR = Mid(strEnc, x, 1)
    intCR = (asc(strCR) + Len(strEnc)) * 2

    If intCR > 255 then
      intCR = (intCR - 255)
    End If

    strCR = chr(intCR)
    strTMP = strTMP & strCR
  Next

  strEnc = strTMP
  strTMP = vbnullstring

  for x = 1 to len("*** Encrypted String is : " & strEnc & " ***")
    strTMP = strTMP & "*"
  next

  wscript.echo vbnewline & strTMP
  wscript.echo "*** Encrypted String is : " & strEnc & " ***"
  wscript.echo strTMP & vbnewline
  set objFSO = createobject("scripting.filesystemobject")
  set objPass = objFSO.createtextfile("pass.txt", true)
    objPass.write strEnc
  objPass.close
wend

set objPass = nothing
set objFSO = nothing
set objWSH = nothing
wscript.quit

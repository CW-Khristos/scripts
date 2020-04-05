dim objWSH, objSCK, strIN, strOUT

set objWSH = createobject("wscript.shell")
set objSCK = createobject("socket.tcp")

on error resume next

while strIN <> "!ABORT!"
  wscript.echo "Send to which server ?"
  strIN = wscript.stdin.readline
  objSCK.dotelnetemulation = true
  objSCK.telnetemulation = "TTY"
  objSCK.host = strIN & ":25"
  objSCK.open
  if err.number <> 0 then
    wscript.echo "Server Not Running Exchange"
    wscript.echo err.number & vbtab & err.description
  else
    objSCK.waitfor "220"
    objSCK.sendline "helo"
    objSCK.waitfor "250"
    wscript.echo "Enter Sender ID (sender@domain.com) :"
    strIN = wscript.stdin.readline
    objSCK.sendline "mail from:" & strIN
    strOUT = objSCK.readline

    while instr(1, strOUT, "501")
      wscript.echo "Invalid Address"
      wscript.echo "Enter Sender ID (sender@domain.com) :"
      strIN = wscript.stdin.readline
      objSCK.sendline "mail from:" & strIN
      strOUT = objSCK.readline
    wend

    wscript.echo "Enter Recipient ID (recieve@domain.com) :"
    strIN = wscript.stdin.readline
    objSCK.sendline "rcpt to:" & strIN
    strOUT = objSCK.readline

    while instr(1, strOUT, "501")
      wscript.echo "Invalid Recipient Address"
      wscript.echo "Enter Recipient ID (recieve@domain.com) :"
      strIN = wscript.stdin.readline
      objSCK.sendline "rcpt to:" & strIN
      strOUT = objSCK.readline
    wend

    objSCK.sendline "data"
    objSCK.waitfor "354"
    wscript.echo "Enter Message. End Message By Entering a Lone '.'"

    while strIN <> "."
      strIN = wscript.stdin.readline
      objSCK.sendline strIN
    wend

    objSCK.sendline "."
    objSCK.waitfor "250"
    objSCK.sendline "quit"
  end if
  objWSH.run "cscript //nologo test.vbs"
  wscript.quit
wend
       
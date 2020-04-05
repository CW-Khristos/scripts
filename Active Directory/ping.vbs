dim com, strRep, strEC
dim objWSH, objArgs, objRep

set objWSH = createobject("wscript.shell")

if wscript.arguments.count > 0 then
  set objArgs = wscript.arguments
  com = objArgs(0)
  set objRep = objWSH.exec(com)

  while objRep.status = 0
    strRep = objRep.stdout.readall
    wscript.echo strRep
    if objRep.exitcode = 0 then
      wscript.echo "ping over"
      wscript.quit err.number
    end if
    wscript.sleep 1000
  wend

elseif wscript.arguments.count = 0 or wscript.arguments.count > 3 then
  wscript.echo "Invalid Number of Arguments"
end if

wscript.quit err.number

dim objWSH, strIP, arrIP, strSUB, strCMD, strNote, x, i

set objWSH  = createobject("wscript.shell")
strCMD = "shutdown -r -m \\"
strNote = "ENTER SHUTDOWN MESSAGE HERE"

do while strIP <> "!ABORT!"
  wscript.echo "To End Script Enter '!ABORT!'"
  wscript.echo "Please Enter Network:"
  strIP = wscript.stdin.readline
  arrIP = split(strIP,".")
  wscript.echo "Please Enter Subnet-Mask (/xx):"
  strSUB = wscript.stdin.readline

  if strIP <> "!ABORT!" then
    if strSUB <> vbnullstring then
      if strSUB = "/24" then
        netBOOT("255")
      elseif strSUB = "/25" then
        netBOOT("128")
      elseif strSUB = "/26" then
        netBOOT("64")
      elseif strSUB = "/27" then
        netBOOT("32")
      elseif strSUB = "/28" then
        netBOOT("16")
      elseif strSUB = "/29" then
        netBOOT("8")
      elseif strSUB = "/30" then
        netBOOT("4")
      elseif strSUB = "/31" then
        netBOOT("2")
      end if
    end if
  end if
loop

sub netBOOT(netid)
  x = right(strIP,3)
  if instr(1,x,".") then
    x = mid(x,instr(1,x,".")+1,len(x)-instr(1,x,"."))
  end if
  for i = int(x) to int(int(x) + int(netid))
    wscript.echo strcmd & arrIP(0) & "." & arrIP(1) & "." & arrIP(2) & "." & i & " -t 120 -c " & chr(34) & strNote & chr(34) & " -f"
    objWSH.run strcmd & arrIP(0) & "." & arrIP(1) & "." & arrIP(2) & "." & i & " -t 120 -c " & chr(34) & strNote & chr(34) & " -f"
    wscript.sleep 500
  next
end sub

set objWSH = nothing
wscript.quit

''V--V Script Written By Khristos V--V''
''V--V VampyreMan Studios Inc. V--V''
''V--V www.vampyremanstudios.com V--V''
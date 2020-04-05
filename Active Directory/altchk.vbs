dim objWSH, com, wrCom, strIP, strRep, iRet, iEC, rptcnt

set objWSH = createobject("wscript.shell")

pinger = "ping.vbs"

wscript.echo "*** This Script Will Monitor Specified IP Every 15 Minutes ***"
wscript.echo "*** Enter IP To Monitor ***"
strIP = wscript.stdin.readline

rptcnt = 900
com = "ping " & strIP
wrCom = "cscript.exe //nologo " & pinger & " """ & com & """"

while strIP <> "!ABORT!"
  rptcnt = rptcnt + 1
  wscript.sleep 1000
  if rptcnt > 900 then
    rptcnt = 0
    with objWSH.Exec (wrcom) 
      iRet = .status 
      iEC = .exitcode
      while iRet = 0 and strEC = vbnullstring
        'wscript.echo iEC
        'wscript.echo strEC
        strRep = .StdOut.Readline 
        wscript.echo strRep
        wscript.sleep 750
        if strRep = "ping over" then
          strEC = "EXIT"
          wscript.echo "Will Ping " & strIP & " Again @ " & dateadd("n",15,time())
        end if
      wend
      strEC = vbnullstring
      if iEC > 1 then  
        wscript.echo "Error executing: " & vbCr & cmd & "Errorcode: " & iRet 
      End If 
    end with
  end if
wend

set objWSH = nothing
wscript.sleep 1000
wscript.quit


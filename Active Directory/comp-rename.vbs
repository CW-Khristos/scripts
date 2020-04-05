on error resume next
const HKLM   = &H80000002
dim objDSP, objReg, objWSH, objNet, objComp
dim strDN, strOld, strNew, strKey, strADp, strCONT, intR

set objWSH = createobject("wscript.shell")
set objNet = createobject("wscript.network")
strOld = objNet.computername
set objReg = getobject("winmgmts:\\" & strOld & "\root\default:StdRegProv")

while strDN = vbnullstring
  wscript.echo vbnewline & "ENTER FQDN OF DOMAIN (my.example.com) :"
  strDN = wscript.stdin.readline

  if strDN = vbnullstring then
    wscript.echo "INVALID INPUT. YOU MUST ENTER A VALUE"
  end if
wend

strDN = "dc=" & replace(strDN, ".", ",dc=")

wscript.echo vbnewline & "SEARCHING ACTIVE DIRECTORY FOR COMPUTER ACCOUNT " & strOld & "..."
Set objDSP = GetObject("LDAP://ou=tq," & strDN)

call getOU(objDSP)

wscript.echo vbnewline & "***                               !WARNING!                               ***"
wscript.echo "***      CONTINUING BEYOND THIS POINT WILL CHANGE SYSTEM INFORMATION      ***"
wscript.echo "*** CONTINUING BEYOND THIS POINT WILL CHANGE ACTIVE DIRECTORY INFORMATION ***"
wscript.echo "***      IF YOU WISH TO CONTINUE PRESS ENTER, OTHERWISE PRESS CTRL+C      ***"
strCONT = wscript.stdin.readline

wscript.echo vbnewline & "OLD COMPUTER NAME IS : " & strOld & " OU : " & strADp

while strNew = vbnullstring
  wscript.echo vbnewline & "ENTER NEW COMPUTER NAME :"
  strNew = wscript.stdin.readline

  if strNew = vbnullstring then
    wscript.echo "INVALID INPUT. YOU MUST ENTER A VALUE"
  else

    strKey = "System\CurrentControlSet\Control\ComputerName\ComputerName"
    intR = objReg.setstringvalue(HKLM, strKey, "ComputerName", strNew)

    if intR <> 0 then
      wscript.Echo vbnewline & "ERROR SETTING COMPUTER NAME VALUE : " & intR
    else
      wscript.Echo vbnewline & "SUCCESSFULLY SET COMPUTER NAME VALUE : " & strNew
    end if

    strKey = "System\CurrentControlSet\Services\Tcpip\Parameters"
    intR = objReg.setstringvalue(HKLM, strKeyPath, "NV Hostname", strNew)

    if intR <> 0 then
      wscript.echo vbnewline & "ERROR SETTING HOSTNAME VALUE : " & intR
    else
      wscript.echo vbnewline & "SUCCESSFULLY SET NV HOSTNAME VALUE : " & strNew
    end if

    strKey = "System\CurrentControlSet\Services\Tcpip\Parameters"
    intR = objReg.setstringvalue(HKLM, strKeyPath, "Hostname", strNew)

    if intR <> 0 then
      wscript.echo vbnewline & "ERROR SETTING HOSTNAME VALUE : " & intR
    else
      wscript.echo vbnewline & "SUCCESSFULLY SET NV HOSTNAME VALUE : " & strNew
    end if

    wscript.echo vbnewline & "REMOVING OLD COMPUTER NAME FROM ACTIVE DIRECTORY"
    set objComp = getobject("LDAP://cn=" & strOld & "," & strADp)
    objComp.deleteobject (0)

    wscript.echo vbnewline & "SETTING NEW COMPUTER NAME IN ACTIVE DIRECTORY"
    set objDSP = getobject("LDAP://" & strADp)
    set objComp = objDSP.Create("Computer", "cn=" & strNew)
    objComp.put "sAMAccountName", strNew & "$"
    objComp.put "Description", "RENAMED " & year(now) & month(now) & day(now)
    objComp.put "userAccountControl", 4096
    objComp.setinfo

    if err.number <> 0 and err.number <> -2147019886 and err.number <> -2147217400 then
      wscript.echo vbnewline & "ERROR OCCURRED SETTING COMPUTER NAME IN ACTIVE DIRECTORY"
      wscript.echo err.number
      wscript.echo vbtab & err.description
    end if

    'set objWMILocator = CreateObject("WbemScripting.SWbemLocator")
    'objWMILocator.Security_.AuthenticationLevel = 6
    'set objWMIComputer = objWMILocator.ConnectServer(strNew, "root\cimv2", strLocalUser, strLocalPasswd)
    'set objWMIComputerSystem = objWMIComputer.Get("Win32_ComputerSystem.Name='" & strNew & "'")
    'intR = objWMIComputerSystem.JoinDomainOrWorkGroup(strDomain, strDomainPasswd, strDomainUser, vbNullString, JOIN_DOMAIN)


    'if err.number = 0 or err.number = -2147019886 or err.number = -2147217400 then
      'wscript.Echo vbnewline & "REBOOTING SYSTEM..."
      'wscript.echo "PRESS ENTER TO CONTINUE, CTRL+C TO QUIT"
      'strCONT = wscript.stdin.readline

      'objWSH.run "shutdown -r -m \\127.0.0.1 -t 5"

    'end if
  end if
wend

set objWSH = nothing
set objComp = nothing
set objNet = nothing
set objWMI = nothing
set objReg = nothing
set objDSP = nothing
wscript.quit

sub getOU(objADp)

  for each objAD in objADp

    select case LCase(objAD.Class)

      case "computer"
        if mid(objAD.name, 4, len(objAD.name) - 3) = strOld then
          strADp = objADp.distinguishedname
          wscript.echo vbnewline & "FOUND " & strOld & " : " & strADp
        end if

      case "organizationalunit"
        call getOU(objAD)

    end select
  next
end sub
dim objDSP, objDN, strDN, strUser, blnFND, blnLCK

while strUser <> "!ABORT!" and strUser = vbnullstring

  set objDSP = getobject("LDAP:")
  strDN = "dc=tq,dc=mnf-wiraq,dc=usmc,dc=mil"
  set objDN = objDSP.opendsobject("LDAP://" & strDN, vbnullstring, vbnullstring, ADS_SERVER_BIND)

  wscript.echo vbnewline & "ENTER '!ABORT!' TO END SCRIPT"
  wscript.echo "ENTER USERNAME TO SEARCH ACTIVE DIRECTORY :"
  strUser = lcase(wscript.stdin.readline)

  if strUser = vbnullstring then
    wscript.echo vbnewline & "INVALID INPUT. YOU MUST PROVIDE A USERNAME"
  end if

  blnLCK = FALSE
  blnFND = FALSE

  call fndUSR(objDN)

  wscript.echo vbnewline & "FINISHED..."

  if blnFND <> TRUE then
    if blnLCK = TRUE then
      blnLCK = FALSE
      wscript.echo vbtab & "ACCOUNT WAS UNLOCKED..."
    elseif blnLCK = FALSE then
      wscript.echo vbtab & "ACCOUNT WAS NOT LOCKED..."
    end if
  else
    wscript.echo vbnewline & "USER : " & strUser & " WAS NOT FOUND..."
  end if

  wscript.sleep 2000
  strUser = vbnullstring
wend

set objDN = nothing
set objDSP = nothing
wscript.quit

sub fndUSR(objADp)

  for each objAD in objADp

    if blnFND = FALSE then

      select case lcase(objAD.class)

        case "user"
          'wscript.echo objAD.samaccountname
          if lcase(objAD.samaccountname) = strUser then
            wscript.echo vbnewline & "FOUND USER : " & objAD.name

            if objAD.isaccountlocked = TRUE then
              objAD.isaccountlocked = FALSE
              objAD.setinfo
              blnLCK = TRUE
            end if

            blnFND = TRUE
            wscript.sleep 2000
          end if

        case "organizationalunit"
          wscript.echo vbnewline & "SCANNING OU : " & objAD.name
          wscript.sleep 100
          call fndUSR(objAD)

      end select

    else
      exit for
    end if

  next

  set objAD = nothing
  set objADp = nothing
end sub
dim objDSP, objDN, strDN, intLCK

intLCK = 0
strDN = "dc=tq,dc=mnf-wiraq,dc=usmc,dc=mil"
Set objDSP = GetObject("LDAP:")
set objDN = objDSP.opendsobject("LDAP://" & strDN, vbnullstring, vbnullstring, ADS_SERVER_BIND)

call unlock(objDN)


set objDN = nothing
set objDSP = nothing
wscript.echo vbnewline & "FINISHED..." & vbnewline & vbtab & "UNLOCKED " & intLCK & " ACCOUNTS"
wscript.sleep 2000

sub unlock(objDMN)

  for each objAD in objDMN

    select case lcase(objAD.class)

      case "user"
        if objAD.isaccountlocked = true then
          intLCK = intLCK + 1
          'wscript.echo objAD.name & " IS LOCKED OUT"
          objAD.isaccountlocked = false
          objAD.setinfo
        end if

      case "organizationalunit"
        wscript.echo vbnewline & "SCANNING THROUGH OU : " & objAD.name
        wscript.sleep 100
        call unlock(objAD)

    end select
  next

  set objAD = nothing
  set objDMN = nothing
end sub
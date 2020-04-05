dim objWSH, objFSO, objSCK
dim arrStat, arrAgg, arrVol
dim strFIL, strStat, blnAgg, blnVol
dim intStat, intAgg, intVol, sinPer, x

set objWSH = createobject("wscript.shell")
set objFSO = createobject("scripting.filesystemobject")

wscript.echo vbnewline & "*********************************************************"
wscript.echo "*********************************************************"
wscript.echo "***         TQFstat.vbs Written By : Khristos         ***"
wscript.echo "***             VampyreMan Studios Inc                ***"
wscript.echo "***    http://www.vampyremanstudios.com/khris.aspx    ***"

on error resume next
while strFIL <> "!ABORT!"
  wscript.echo "*********************************************************"
  wscript.echo "*********************************************************"
  wscript.echo "*** This Script Obtains the Stats for Specified Filer ***"
  wscript.echo "***       Type Filer IP To Pull Stats For Filer       ***"
  wscript.echo "***               Type '!ABORT!' To Quit              ***"
  wscript.echo "*********************************************************"
  wscript.echo "*********************************************************" & vbnewline
  strFIL = wscript.stdin.readline
  set objStat = objFSO.createtextfile("FilerStats.txt", true)
  wscript.sleep 1000
  set objSCK = createobject("socket.tcp")
  objSCK.dotelnetemulation = true
  objSCK.telnetemulation = "TTY"
  objSCK.host = strFIL & ":23"
  objSCK.open
  if err.number <> 0 then
    wscript.echo vbnewline & "************************************************"
    wscript.echo "************************************************"
    wscript.echo "***        !!!ERROR OPENING SOCKET!!!        ***"
    wscript.echo "*** !!!FILER POSSIBLY RUNNING IN FAILOVER!!! ***"
    wscript.echo "************************************************"
    wscript.echo "************************************************" & vbnewline
    wscript.sleep 2000
    wscript.echo vbtab & "*** QUIT ***"
    wscript.sleep 1000
  else
    objSCK.waitfor "login:"
    wscript.echo vbnewline & "*** Logging Into : " & strFIL & vbnewline
    wscript.sleep 500
    if strFIL = crypter("€|†x~~|x~|~x†„") then
      objSCK.sendline crypter("ìææð")
    else
      objSCK.sendline crypter("ÜâôìöìþÜøþ")
    end if
    if err.number <> 0 then
      wscript.echo "************************************************"
      wscript.echo "************************************************"
      wscript.echo "***           !!!Invalid Login!!!            ***"
      wscript.echo "************************************************"
      wscript.echo "************************************************" & vbnewline
      wscript.sleep 2000
      wscript.echo vbtab & "*** QUIT ***"
      wscript.sleep 1000
    else
      objSCK.waitfor "password:"
      objSCK.sendline crypter(" ˜˜¸Úþþxüà")
      if err.number <> 0 then
        wscript.echo "************************************************"
        wscript.echo "************************************************"
        wscript.echo "***          !!!Invalid Password!!!          ***"
        wscript.echo "************************************************"
        wscript.echo "************************************************" & vbnewline
        wscript.sleep 2000
        wscript.echo vbtab & "*** QUIT ***"
        wscript.sleep 1000
      else
        objSCK.waitfor "filer>"
        objSCK.sendline "df -A -g"
        objSCK.waitfor "filer>"
        objStat.writeline objSCK.buffer & vbnewline
        objSCK.sendline "df -V -g"
        objSCK.waitfor "filer>"
        objStat.writeline objSCK.buffer & vbnewline
        wscript.sleep 100
        objStat.close
        set objStat = objFSO.opentextfile("FilerStats.txt", 1)

        do until objStat.atendofstream
          strStat = objStat.readline
          if instr(1, strStat, "Aggregate") then
            blnAgg = "TRUE"
          elseif instr(1, strStat, "Filesystem") then
            blnVol = "TRUE"
            blnAgg = "FALSE"
          end if
          if blnAgg = "TRUE" then
            if instr(1, strStat, "GB") and instr(1, strStat, ".snapshot") = 0 then
              arrStat = split(strStat, "GB")
              arrAgg = split(arrStat(0), " ")
              intAgg = int(intAgg) + int(arrAgg(ubound(arrAgg)))
            end if
          elseif blnAgg = "FALSE" and blnVol = "TRUE" then
            if instr(1, strStat, "GB") and instr(1, strStat, ".snapshot") = 0 then
              arrStat = split(strStat, "GB")
              arrVol = split(arrStat(1), " ")
              intVol = int(intVol) + int(arrVol(ubound(arrVol)))
            end if
          end if   
        loop

        sinPer = mid(((intVol / intAgg) * 100) + 10, 1, 5)
        wscript.echo "*** Aggregate Total is : " & intAgg & "GB"
        wscript.echo "*** Volume Usage is : " & intVol & "GB"
        wscript.echo "*** Percentage of Filesystem Used is : " & sinPer & "%" & vbnewline
      end if
    end if
  end if
  objSCK.close
  wscript.sleep 1000
wend

function crypter(str)
  dim strCR, intCR, strCode

  for x = 1 to len(str)
    strCR = Mid(str, x, 1)
    if asc(strCR) < 32 Then
      intCR = ((asc(strCR) + 255) / 2) - len(str)
    else
      intCR = (asc(strCR) / 2) - len(str)
    end if
    strCR = chr(intCR)
    strCode = strCode & strCR
  next

  crypter = strCode
end function
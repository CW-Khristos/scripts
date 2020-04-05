	''CREATE THE VARIABLES TO BE USED IN THE SCRIPT''
dim objWSH, objFSO, objLCK, x, i, intMAC
dim strIN, arrIN, strCMD, strSUB, strLOOP, strIP, arrIP, blnFND

	''CREATE THE COMMAND SHELL SCRIPTING OBJECT TO RUN COMMANDS''
set objWSH = createobject("wscript.shell")

	''CREATE THE FILE SYSTEM SCRIPTING OBJECT TO ACCESS FILES''
set objFSO = createobject("scripting.filesystemobject")

	''INFINITE LOOP TO REPEAT PROCEDURES FOR NEXT DHCP SCOPE''
while strLOOP = vbnullstring

	''PURELY AESTHETICS''
  wscript.echo vbnewline & "***********************************************"
  wscript.echo "***********************************************"
  wscript.echo "***    DHCP-LOCK.VBS WRITTEN BY KHRISTOS    ***"
  wscript.echo "***           VAMPYREMAN STUDIOS            ***"
  wscript.echo "***      HTTP://VAMPYREMANSTUDIOS.COM       ***"
  wscript.echo "*** SIMPLE SCRIPT TO LOCK DHCP RESERVATIONS ***"
  wscript.echo "***  FROM AN EXPORTED LIST OF DHCP LEASES   ***"
  wscript.echo "***********************************************"
  wscript.echo "***********************************************"
  wscript.sleep 500

	''PREP THE NETSH COMMAND FOR DHCP CONFIGURATION''
  strCMD = "netsh dhcp server "

	''REQUEST THE DHCP SERVER IP''
  wscript.echo vbnewline & "ENTER DHCP SERVER IP :"
  strIN = wscript.stdin.readline

	''PERFORM SIMPLE INPUT CHECKINGTO ENSURE SOMETHING WAS ENTERED''
  while strIN = vbnullstring
    wscript.echo vbnewline & "YOU MUST ENTER A SERVER IP WHICH IS RUNNING DHCP"
    wscript.echo "ENTER DHCP SERVER IP :"
    strIN = wscript.stdin.readline
  wend

	''ADD THE SERVER IP TO THE NETSH DHCP COMMAND''
  strCMD = strCMD & strIN

	''CLEAR THE LAST INPUT''
  strIN = vbnullstring

	''REQUEST THE SCOPE IP TO LOCK OUT''
  wscript.echo vbnewline & "ENTER SCOPE IP :"
  strIN = wscript.stdin.readline

  while strIN = vbnullstring
    wscript.echo vbnewline & "YOU MUST ENTER A VALID SCOPE IP"
    wscript.echo "ENTER SCOPE IP :"
    strIN = wscript.stdin.readline
  wend

	''SAVE THE SCOPE IP FOR LATER''
  strIP = strIN

	''SPLIT THE SCOPE IP AT EVERY '.' TO SPLIT UP THE OCTETS''
  arrIP = split(strIN, ".")

	''ADD THE SCOPE PARAMETER TO THE NETSH DHCP COMMAND''
  strCMD = strCMD & " scope " & strIN

  strIN = vbnullstring

	''REQUEST THE SUBNET MASK ID TO PROVIDE THE IP RANGE''
  wscript.echo vbnewline & "ENTER SUBNET ID (/XX) :"
  strIN = wscript.stdin.readline

  while strIN = vbnullstring
    wscript.echo vbnewline & "YOU MUST ENTER A VALID SUBNET ID IN SLASH (/XX)"
    wscript.echo "ENTER SUBNET ID (/XX) :"
    strIN = wscript.stdin.readline
  wend

	''SAVE THE SUBNET ID FOR LATER''
  strSUB = strIN

  strIN = vbnullstring

	''REQUEST THE FILENAME OF THE EXPORTED DHCP LEASES''
  wscript.echo vbnewline & "ENTER FILENAME CONTAINING LIST OF IP LEASES :"
  strIN = wscript.stdin.readline

  while strIN = vbnullstring
    wscript.echo vbnewline & "YOU MUST SPECIFY A VALID FILENAME WITH LIST OF IP LEASES"
    wscript.echo "ENTER FILENAME CONTAINING LIST OF IP LEASES :"
    strIN = wscript.stdin.readline
  wend

	''DETERMINE THE NETWORK RANGE BY THE SUBNET ID''
	''CALL THE 'lock' SUB-ROUTINE WITH THIS VALUE''
  if strSUB <> vbnullstring then
    if strSUB = "/24" then
      lock("255")
    elseif strSUB = "/25" then
      lock("128")
    elseif strSUB = "/26" then
      lock("64")
    elseif strSUB = "/27" then
      lock("32")
    elseif strSUB = "/28" then
      lock("16")
    elseif strSUB = "/29" then
      lock("8")
    elseif strSUB = "/30" then
      lock("4")
    elseif strSUB = "/31" then
      lock("2")
    end if
  end if

	''UPON COMPLETION OF THE SUB-ROUTINE, RETURN TO THE BEGINNING OF THE SCRIPT''
wend

	''THIS SUBROUTINE SCANS THE NETWORK RANGE FINDING THE IPS NOT IN USE''
sub lock(netid)

	''SET A BOGUS MAC ADDRESS TO USE FOR THE IP RESERVATIONS''
  intMAC = 0

	''TAKE THE LAST THREE NUMBERS FROM THE SCOPE IP''
  x = right(strIP,3)

	''IF A '.' IS IN THE LAST THREE NUMBERS''
  if instr(1, x, ".") then

	''REMOVE THE '.' AND JUST USE THE NUMBERS''
    x = mid(x,instr(1, x, ".") + 1, len(x) - instr(1, x, "."))
  end if

	''STARTING WITH THE FIRST AVAILABLE IP IN THE NETWORK RANGE''
	''THIS IS THE SCOPE IP + 1''
	''END WITH THE LAST AVAILABLE IP IN THE NETWORK RANGE''
	''THIS IS FIRST AVAILABLE IP + NETWORK RANGE - 3''
  for i = int(x + 1) to int(int(x + 1) + int(netid)) - 3

		''SET A BOOLEAN VARIABLE TO DETERMINE IF A MATCH IS FOUND''
    blnFND = "FALSE"

		''USE THE FILE SYSTEM OBJECT TO OPEN THE EXPORTED LEASES FILE''
		''SET THIS INPUT TO AN OBJECT FOR USE IN THE FOLLOWING PROCEDURES''
    set objLCK = objFSO.opentextfile(strIN)

		''LOOP THESE PROCEDURES UNTIL THE END OF THE FILE OR A MATCH IS FOUND''
    do until objLCK.atendofstream or blnFND = "TRUE"

		''READ A LINE FROM THE FILE AND INPUT IT TO 'strIN' ''
      strIP = objLCK.readline

		''IF 'Client IP Address' IS NOT FOUND IN THE INPUT''
      if instr(1, strIP, "Client IP Address") = 0 then

		''SPLIT THE INPUT WHEREVER A ',' IS FOUND''
        arrIN = split(strIP, ",")

		''IF 'arrIN(0)' MATCHES THE THREE OCTETS OF OUR SCOPE NETWORK''
		''AND THE LAST OCTET WHICH INCREMENTS BY ONE THROUGH OUR NETWORK RANGE''
		''THIS DETERMINES WHETHER THE IP OF A LEASE IS A MATCH OR NOT''
        if arrIN(0) = arrIP(0) & "." & arrIP(1) & "." & arrIP(2) & "." & i then

		''IF A MATCH IS FOUND SET THE BOOLEAN VARIABLE TO 'TRUE' ''
          blnFND = "TRUE"

        end if
      end if

	''RETURN TO THE BEGINNING OF THE LOOP UNLESS A MATCH WAS FOUND''
    loop

	''IF A MATCH WAS NOT FOUND IN THE ENTIRE EXPORTED LEASES FILE''
    if blnFND = "FALSE" then

		''INCREMENT THE BOGUS MAC ADDRESS TO KEEP IT UNIQUE''
      intMAC = intMAC + 1

		''MAC ADDRESS ARE 12 DIGITS''

		''IF THE BOGUS MAC IS ONLY 1 DIGIT LONG''
		''ADD 11 0'S TO THE BEGINNING OF IT''
      if len(intMAC) = 1 then
        intMAC = "00000000000" & intMAC

		''IF THE BOGUS MAC IS 2 DIGITS LONG''
		''ADD 10 0'S TO THE BEGINNING OF IT''
      elseif len(intMAC) = 2 then
        intMAC = "0000000000" & intMAC

		''IF THE BOGUS MAC IS 3 DIGITS LONG''
		''ADD 9 0'S TO THE BEGINNING OF IT''
      elseif len(intMAC) = 3 then
        intMAC = "000000000" & intMAC
      end if

		''DISPLAY THE NETSH COMMAND FULLY CONFIGURED IN THE COMMAND PROMPT''
      wscript.echo vbnewline & strCMD & " add reservedip " & arrIP(0) & "." & arrIP(1) & "." & arrIP(2) & "." & i & " " & intMAC & " LOCKED" & intMAC

		''RUN THE NETSH COMMAND FULLY CONFIGURED''
		''THE 'TRUE' PARAMETER HIDES THE COMMAND WINDOW''
      objWSH.run strCMD & " add reservedip " & arrIP(0) & "." & arrIP(1) & "." & arrIP(2) & "." & i & " " & intMAC & " LOCKED" & intMAC, 0, TRUE

		''PAUSE FOR 250 MILLISECONDS''
      wscript.sleep 250
    end if

	''MOVE TO THE NEXT IP IN THE NETWORK RANGE''
  next

	''CLEAR THE BOGUS MAC ADDRESS''
  intMAC = 0

	''SUB-ROUTINE IS FINISHED, RETURN TO THE MAIN BODY OF THE SCRIPT''
end sub

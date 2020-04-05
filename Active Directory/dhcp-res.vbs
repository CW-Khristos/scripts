	''CREATE VARIABLES FOR USE IN THE SCRIPT''
dim objWSH, objFSO, objLSE, strCMD, strRUN, strIN, arrIN, strLOOP

	''CREATE A COMMAND SHELL SCRIPTING OBJECT TO RUN COMMANDS''
set objWSH = createobject("wscript.shell")

	''CREATE A FILE SYSTEM SCRIPTING OBJECT TO ACCESS FILES''
set objFSO = createobject("scripting.filesystemobject")

	''INFINITE LOOP TO REPEAT SCRIPT PROCEDURES FOR NEXT DHCP SCOPE''
while strLOOP = vbnullstring

	''PURELY AESTHETICS''
  wscript.echo vbnewline & "**********************************************"
  wscript.echo "**********************************************"
  wscript.echo "***    DHCP-RES.VBS WRITTEN BY KHRISTOS    ***"
  wscript.echo "***           VAMPYREMAN STUDIOS           ***"
  wscript.echo "***      HTTP://VAMPYREMANSTUDIOS.COM      ***"
  wscript.echo "*** SIMPLE SCRIPT TO ADD DHCP RESERVATIONS ***"
  wscript.echo "***  FROM AN EXPORTED LIST OF DHCP LEASES  ***"
  wscript.echo "**********************************************"
  wscript.echo "**********************************************"
  wscript.sleep 500

	''PREP THE NETSH COMMAND FOR DHCP CONFIGURATION''
  strCMD = "netsh dhcp server "

	''REQUEST THE DHCP SERVERS IP''
  wscript.echo vbnewline & "ENTER DHCP SERVER IP :"
  strIN = wscript.stdin.readline

	''PERFORM SIMPLE INPUT CHECK TO ENSURE SOMETHING WAS ENTERED''
  while strIN = vbnullstring
    wscript.echo vbnewline & "YOU MUST ENTER A SERVER IP WHICH IS RUNNING DHCP"
    wscript.echo "ENTER DHCP SERVER IP :"
    strIN = wscript.stdin.readline
  wend

	''ADD THE SERVER IP TO THE NETSH DHCP COMMAND''
  strCMD = strCMD & strIN

	''CLEAR THE LAST INPUT''
  strIN = vbnullstring

	''REQUEST THE SCOPE IP TO MAKE RESERVATIONS IN''
  wscript.echo vbnewline & "ENTER SCOPE IP :"
  strIN = wscript.stdin.readline

  while strIN = vbnullstring
    wscript.echo vbnewline & "YOU MUST ENTER A VALID SCOPE IP"
    wscript.echo "ENTER SCOPE IP :"
    strIN = wscript.stdin.readline
  wend

	''ADD THE SCOPE PARAMETER TO THE NETSH DHCP COMMAND''
  strCMD = strCMD & " scope " & strIN

  strIN = vbnullstring

	''REQUEST THE FILENAME OF THE EXPORTED DHCP LEASES''
  wscript.echo vbnewline & "ENTER FILENAME CONTAINING LIST OF IP LEASES :"
  strIN = wscript.stdin.readline

  while strIN = vbnullstring
    wscript.echo vbnewline & "YOU MUST SPECIFY A VALID FILENAME WITH LIST OF IP LEASES"
    wscript.echo "ENTER FILENAME CONTAINING LIST OF IP LEASES :"
    strIN = wscript.stdin.readline
  wend

	''USE THE FILE SYSTEM OBJECT TO OPEN THE EXPORTED LEASES FILE''
	''SET THIS INPUT TO AN OBJECT VARIABLE FOR THE NEXT PROCEDURES''
  set objLSE = objFSO.opentextfile(strIN)

	''LOOP THESE PROCEDURES UNTIL THE END OF THE FILE IS REACHED''
  do until objLSE.atendofstream

		''READ A LINE FROM THE FILE AND INPUT IT TO strIN''
    strIN = objLSE.readline

		''IF 'Reservation' AND 'Client IP Address' ARE NOT IN THE INPUT
    if instr(1, strIN, "Reservation ") = 0 and instr(1, strIN, "Client IP Address") = 0 then

		''SPLIT THE INPUT WHENEVER A ',' IS FOUND'
      arrIN = split(strIN, ",")

		''SET THE NETSH DHCP COMMAND WITH THE FINAL PARAMETERS''
		'' 'arrIN(0)' IS THE IP OF THE LEASE TO BE RESERVED''
		'' 'arrIN(4)' IS THE MAC ADDRESS OF THE LEASE TO BE RESERVED''
		'' 'arrIN(1)' IS THE COMPUTER NAME OF THE LEASE TO BE RESERVED''
		'' 'arrIN(5)' IS THE DESCRIPTION OF THE LEASE TO BE RESERVED''
      strRUN = strCMD & " add reservedip " & arrIN(0) & " " & arrIN(4) & " " & arrIN(1) & " " & arrIN(5)

		''DISPLAY THE NETSH COMMAND FULLY CONFIGURED IN THE COMMAND PROMPT''
      wscript.echo vbnewline & strRUN

		''RUN THE NETSH COMMAND FULLY CONFIGURED''
		''THE 'TRUE' PARAMETER HIDES THE COMMAND WINDOW'
      objWSH.run strRUN, 0, TRUE
    end if

	''RETURN TO THE BEGINNING OF THE LOOP''
  loop

	''RETURN TO THE BEGINNING OF THE SCRIPT ONCE THE END OF FILE IS REACHED''
wend

set objLSE = nothing
set objFSO = nothing
set objWSH = nothing
wscript.quit

''V--V  Script Written By Khristos  V--V''
''V--V   VampyreMan Studios Inc.    V--V''
''V--V http://vampyremanstudios.com V--V''
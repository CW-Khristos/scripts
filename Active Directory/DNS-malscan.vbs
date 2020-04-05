'*****************************************************************************************************'
'*****************************************************************************************************'
'****                    THIS SCRIPT DOES NOT DETECT ACTUAL MALICIOUS SOFTWARE                    ****'
'****          THIS SCRIPT IS NOT AN ANTI-VIRUS SOFTWARE OR A MEANS OF VIRUS PREVENTION           ****'
'****  IT WILL ONLY SEARCH DNS DEBUG RECORD TO FIND REQUEST FOR KNOWN / POSSIBLE MALICIOUS SITES  ****'
'**** PROVIDING THE IP OF THE MACHINE REQUESTING THE SITE AND A POSSIBLE MALICIOUS SOFTWARE NAME  ****'
'**** A CLIENT BROWSING TO 'example.trOjan.com' WILL PRODUCE THE SAME RESULTS OF AN ACTUAL TROJAN ****'
'*****************************************************************************************************'
'*****************************************************************************************************'

'V--V THERE ARE TWO METHODS THIS SCRIPT CAN BE USED DEPENDING UPON YOUR OWN PREFERENCES V--V'
'V--V   1). EDIT THE LINE WITH '{YOUR DNS LOGFILE}' BY REPLACING '{YOUR DNS LOGFILE}'   V--V'
'V--V       REPALCE '{YOUR DNS LOGFILE}' WITH THE PATH OF THE SERVER'S DEBUG LOGS.      V--V'
'V--V       SAVE THE CHANGES TO THIS SCRIPT, THEN RUN IT WITH THE COMMAND PROMPT.       V--V'

'V--V   2). UN-COMMENT LINES MARKED '{ARGUMENT}' THEN REPLACE '{YOUR DNS LOGFILE}'      V--V'
'V--V       REPLACE '{YOUR DNS LOGFILE} WITH 'arrARGS(0)'. OPEN A CMD PROMPT.           V--V'
'V--V       USE 'CD' COMMAND TO SET THE DIRECTORY TO THE FOLDER DNS-MALSCAN.VBS IS IN.  V--V'
'V--V       TYPE 'CSCRIPT //NOLOGO DNS-MALSCAN.VBS PATH' INTO THE CMD PROMPT.           V--V'
'V--V       WHERE PATH IS THE FULL PATH TO YOUR SERVER'S DNS DEBUG LOGFILE.             V--V'


	'V--V SETUP VARIABLES TO BE USED WITHIN THIS SCRIPT V--V'
'Dim arrARGS 'V--V {ARGUMENT} V--V'
Dim objWSH, objFSO, objfDNS, objMAL, objFND
Dim strLN, rptCNT, strLOOP, strFND, arrMAL, arrREQ

rptCNT = 600
strFND = "FALSE"
strLOOP = vbnullstring

	'V--V CREATE WINDOWS SCRIPT HOST AND FILESYSTEM OBJECTS                      V--V'
	'V--V IF UN-COMMENTED, SET 'arrARGS' TO THE ARGUMENTS PROVIDED TO THE SCRIPT V--V'
set objWSH = createobject("wscript.shell")
set objFSO = createobject("scripting.filesystemobject")


	'V--V {ARGUMENT} - UNCOMMENT THIS ENTIRE BLOCK TO USE  V--V'
	'V--V CHECKS TO MAKE SURE PROPER ARGUMENTS WERE PASSED V--V'
'if wscript.arguments.count > 0 then
'  set arrARGS = wscript.arguments
'else
'  wscript.echo "Invalid Number of Arguments. (None Were Provided)"
'  wscript.sleep 3000
'  wscript.quit
'end if
	'V--V {ARGUMENT} - UNCOMMENT THIS ENTIRE BLOCK TO USE V--V'


	'V--V PURLY AESTHETIC ECHO COMMANDS V--V'
wscript.echo "******************************************************"
wscript.echo "*** SIMPLE DNS LOG SCANNER FOR MALWARE DNS-LOOKUPS ***"
wscript.echo "*** THIS SCRIPT WILL SCAN THE DNS LOG EVERY 10 MIN ***"
wscript.echo "***          SCRIPT WRITTEN BY : KHRISTOS          ***"
wscript.echo "***               VampyreMan Studios               ***"
wscript.echo "***  http://www.vampyremanstudios.com/khris.aspx   ***"
wscript.echo "******************************************************"

	'V--V START A LOOP WHILE 'strLOOP' IS A NULL STRING V--V'
	'V--V THIS LOOP IS AN INFINITE LOOP IN THIS SCRIPT  V--V'
while strLOOP = vbnullstring

        'V--V ADD 1 TO THE LOOP'S COUNT ('rptCNT') V--V'
	'V--V THIS COUNT IS USED TO SET THE 10 MIN INTERVALS V--V'
  rptCNT = rptCNT + 1

	'V--V ONCE 'rptCNT' IS GREATER THAN 600 PERFORM THESE COMMANDS V--V'
  if rptCNT > 600 then

	'V--V MORE AESTHETIC ECHO COMMANDS TO PROVIDE INFORMATIONAL OUTPUT V--V'
    wscript.echo vbnewline & "******************************************************"
    wscript.echo "***              NOW SCANNING DNS LOG              ***"
    wscript.echo "******************************************************" & vbnewline

	'V--V SET 'objfDNS' TO THE OPENED DNS DEBUG LOGFILE PROVIDED      V--V'
	'V--V START A LOOP UNTIL THE END OF THE DNS DEBUG FILE IS REACHED V--V'
	'V--V READ A LINE FROM THE DNS DEBUG LOGFILE                      V--V'
    set objfDNS = objFSO.opentextfile("{YOUR DNS LOGFILE}")
    do until objfDNS.atendofstream
      strLN = objfDNS.readline

		'V--V SET 'objMAL' TO THE OPENED TEXT FILE 'dbMAL.txt'                V--V'
		'V--V THIS TEXT FILE IS PROVIDED WITH THE ZIP DOWNLOAD OF DNS-MALSCAN V--V'
		'V--V THE FORMAT OF THIS FILE IS 'ex.troj || example,trojan,com'      V--V'
                'V--V ONLY USE UP TO THE FIRST FOUR SECTIONS OF THE URL               V--V'
		'V--V SAME AS LOOP ABOVE, BUT FOR 'dbMAL.txt'                         V--V'
      set objMAL = objFSO.opentextfile("dbMAL.txt")
      do until objMAL.atendofstream

		'V--V SEPEREATE THE READ LINE BY THE ' || ' DELIMETER            V--V'
		'V--V THIS IS JUST THE FORMAT I USED, YOU MAY CHANGE IT          V--V'
		'V--V YOU MUST CHANGE IT IN THE ACTUAL TEXT FILE IF CHANGED HERE V--V'
        arrMAL = split(objMAL.readline, " || ")
        arrREQ = split(arrMAL(1), ",")

                'V--V CHECK THE CURRENT LINE OF THE DNS DEBUG LOG FOR ALL KNOWN MALICIOUS DNS REQUESTS V--V'
		'V--V IF THE SCRIPT FINDS THE KNOWN DNS REQUESTS IN THE DNS DEBUG LOGFILE              V--V'
		'V--V IT WILL SET 'strFND' TO 'TRUE' TO DENOTE IT FOUND A POSSIBLE MALWARE REQUEST     V--V'
        if ubound(arrREQ) = 3 then
          if instr(1, strLN, arrREQ(0)) AND instr(1, strLN, arrREQ(1)) and instr(1, strLN, arrREQ(2)) and instr(1, strLN, arrREQ(3)) then
            strFND = "TRUE"
          elseif instr(1, strLN, arrREQ(0)) AND instr(1, strLN, arrREQ(1)) and instr(1, strLN, arrREQ(2)) then
            strFND = "TRUE"
          elseif instr(1, strLN, arrREQ(0)) AND instr(1, strLN, arrREQ(1)) then
            strFND = "TRUE"
          elseif instr(1, strLN, arrREQ(0)) then
            strFND = "TRUE"
          end if
        elseif ubound(arrREQ) = 2 then
          if instr(1, strLN, arrREQ(0)) AND instr(1, strLN, arrREQ(1)) and instr(1, strLN, arrREQ(2)) then
            strFND = "TRUE"
          elseif instr(1, strLN, arrREQ(0)) AND instr(1, strLN, arrREQ(1)) then
            strFND = "TRUE"
          elseif instr(1, strLN, arrREQ(0)) then
            strFND = "TRUE"
          end if
        elseif ubound(arrREQ) = 1 then
          if instr(1, strLN, arrREQ(0)) AND instr(1, strLN, arrREQ(1)) then
            strFND = "TRUE"
          elseif instr(1, strLN, arrREQ(0)) then
            strFND = "TRUE"
          end if
        elseif ubound(arrREQ) = 0 then
          if instr(1, strLN, arrREQ(0)) then
            strFND = "TRUE"
          end if
        end if

		'V--V MORE AESTHETICS, IF A MALWARE REQUEST WAS FOUND, PROVIDE MORE OUTPUT           V--V'
		'V--V IF THE 'dbFOUND.txt' FILE EXISTS OPEN IT TO APEND TO IT, IF NOT THEN CREATE IT V--V'
		'V--V WRITE / APPEND THE LINE CONTAINING THE MALWARE REQUEST TO 'dbFOUND.txt'        V--V'
		'V--V DISPLAYS WHICH MALWARE REQUEST WAS FOUND AND WHEN IT WAS FOUND                 V--V'
        if strFND = "TRUE" then
          wscript.sleep 10
          strFND = "FALSE"
          if objFSO.fileexists("dbFOUND.txt") then
            set objFND = objFSO.opentextfile("dbFOUND.txt", 8)
          else
            set objFND = objFSO.createtextfile("dbFOUND.txt")
          end if
          objFND.writeline strLN & vbnewline
          objFND.close
          wscript.echo vbnewline & "******************************************************"
          wscript.echo "***" & vbtab & "FOUND " & arrMAL(0) & " @ " & time() & "!" & vbtab & "***"
          wscript.echo "******************************************************"
          wscript.echo strLN & vbnewline
        end if
      loop

	'V--V MOVE TO THE NEXT LINE OF THE DNS DEBUG LOGFILE V--V'
	'V--V PAUSE FOR 10 MILISECONDS AND MOVE TO THE NEXT LINE OF DNS LOG V--V'
        wscript.sleep 10
    loop

	'V--V ONCE FINISHED READING EVERY LINE OF THE DNS DEBUG LOGFILE PROVIDE OUTPUT          V--V'
	'V--V WRITES TO THE 'dbFOUND.txt' FILE INDICATING SCAN WAS FINISHED                     V--V'
	'V--V PROVIDE OUTPUT THAT SCAN WAS FINISHED AND AT WHAT TIME A RE-SCAN WILL OCCUR       V--V'
    if objFSO.fileexists("dbFOUND.txt") then
      set objFND = objFSO.opentextfile("dbFOUND.txt", 8)
    else
      set objFND = objFSO.createtextfile("dbFOUND.txt")
    end if
    objFND.writeline "*** NO MORE ENTRIES FOUND @ " & time() & " ***" & vbnewline
    objFND.close
    wscript.echo vbnewline & "******************************************************"
    wscript.echo "***" & vbtab & "NO MORE ENTRIES FOUND @ " & time() & vbtab & "   ***"
    wscript.echo "***   DNS-MALSCAN WILL SCAN AGAIN @ " & dateadd("n", 10, time()) & "     ***"
    wscript.echo "******************************************************"
    rptCNT = 0
  end if

	'V--V PAUSE FOR 1 SECOND, THEN LOOP BACK TO THE BEGINNING                   V--V'
	'V--V THIS IS AN INFINITE LOOP SINCE 'strLOOP' WILL ALWAYS BE A NULL STRING V--V'
  wscript.sleep 1000
wend

''V--V Script Written By Khristos V--V''
''V--V VampyreMan Studios Inc. V--V''
''V--V www.vampyremanstudios.com V--V''
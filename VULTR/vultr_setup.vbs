''VULTR_SETUP.VBS
''DESIGNED TO AUTOMATE SETUP OF VULTR PBX ACCOUNTS USING VULTR-CLI CMD UTILITY : https://github.com/vultr/vultr-cli

''ACCEPTS 1 PARAMETER , REQUIRES 1 PARAMETER
''REQUIRED PARAMETER 'STRAPI' ; STRING VALUE FOR API KEY FOR AUTHENTICATION WITH VULTR
''OPTIONAL PARAMETER 'STRFILTER' ; STRING VALUE TO HOLD PASSED 'FILTERS' ; SEPARATE MULTIPLE 'FILTERS' VIA '|'
''OPTIONAL PARAMETER 'STRINCL' ; STRING VALUE TO HOLD PASSED 'INCLUSIONS' ; SEPARATE MULTIPLE 'INCLUSIONS' VIA '|'
''OPTIONAL PARAMETER 'STRUSR' ; STRING VALUE TO HOLD PASSED 'USER ACCOUNT' TO EXCLUDE
''WRITTEN BY : CJ BLEDSOE / CJ<@>THECOMPUTERWARRIORS.COM
on error resume next
''SCRIPT VARIABLES
dim errRET, strVER
dim strREPO, strBRCH, strDIR
''VARIABLES ACCEPTING PARAMETERS
dim blnLOOP, strAPI, arrPBX()
dim strIN, strOUT, strRCMD, strSAV
dim strREG, strPLN, strOS, strHOST, strDMN, strIP
''SCRIPT OBJECTS
dim objIN, objOUT, objARG, objWSH, objFSO
dim objLOG, objEXEC, objHOOK, objHTTP, objXML
''VERSION FOR SCRIPT UPDATE , VULTR_SETUP.VBS , REF #2 , REF #68 , REF #69
strVER = 1
strREPO = "scripts"
strBRCH = "dev"
strDIR = "VULTR"
redim arrPBX(1)
strSCP = "C:\Users\CBledsoe\AppData\Local\Programs\WinSCP\winscp.com"
''VULTR DETAILS
strAPI = "QLD7MNZAYSXP2LTYVZMUCKPRC7NRDBSDQ7PQ"
strISO = "498b7c35-d407-4106-9c83-1dfc555fc447"
strFW = "1acf6e7e-268f-4108-b5f8-9fd00607f492"
strPNET = "ba609e8d-6564-4106-960b-c1f37d81751c"
strDMN = ".ipmrms.com"
strPLN = "vc2-1c-1gb"
strREG = "ewr"
strOS = "159"
''DEFAULT SUCCESS
errRET = 0
''STDIN / STDOUT
set objIN = wscript.stdin
set objOUT = wscript.stdout
set objARG = wscript.arguments
''OBJECTS FOR LOCATING FOLDERS
set objWSH = createobject("wscript.shell")
set objFSO = createobject("scripting.filesystemobject")
''CHECK 'PERSISTENT' FOLDERS , REF #2 , REF #73
if (not (objFSO.folderexists("c:\temp"))) then
  objFSO.createfolder("c:\temp")
end if
if (not (objFSO.folderexists("C:\IT\"))) then
  objFSO.createfolder("C:\IT\")
end if
if (not (objFSO.folderexists("C:\IT\Scripts\"))) then
  objFSO.createfolder("C:\IT\Scripts\")
end if
''PREPARE LOGFILE
if (objFSO.fileexists("C:\temp\VULTR_SETUP")) then            ''LOGFILE EXISTS
  objFSO.deletefile "C:\temp\VULTR_SETUP", true
  set objLOG = objFSO.createtextfile("C:\temp\VULTR_SETUP")
  objLOG.close
  set objLOG = objFSO.opentextfile("C:\temp\VULTR_SETUP", 8)
else                                                          ''LOGFILE NEEDS TO BE CREATED
  set objLOG = objFSO.createtextfile("C:\temp\VULTR_SETUP")
  objLOG.close
  set objLOG = objFSO.opentextfile("C:\temp\VULTR_SETUP", 8)
end if
''CHECK FOR VULTR-CLI CMD UTILITY , REF #76
if (objFSO.fileexists("C:\IT\vultr-cli.exe")) then
  call LOGERR(0)                                              ''VULTR-CLI.EXE PRESENT, CONTINUE SCRIPT, 'ERRRET'=0
elseif (not objFSO.fileexists("C:\IT\vultr-cli.exe")) then
  call LOGERR(1)                                              ''VULTR-CLI.EXE NOT PRESENT, 'ERRRET'=1
  ''DOWNLOAD VULTR-CLI
  call FILEDL("https://raw.githubusercontent.com/CW-Khristos/scripts/dev/VULTR/vultr-cli.exe", "C:\IT", "vultr-cli.exe")
end if
''READ PASSED COMMANDLINE ARGUMENTS
if (wscript.arguments.count > 0) then                         ''ARGUMENTS WERE PASSED
  for x = 0 to (wscript.arguments.count - 1)
    objOUT.write vbnewline & now & vbtab & " - ARGUMENT " & (x + 1) & " (ITEM " & x & ") " & " PASSED : " & ucase(objARG.item(x))
    objLOG.write vbnewline & now & vbtab & " - ARGUMENT " & (x + 1) & " (ITEM " & x & ") " & " PASSED : " & ucase(objARG.item(x))
  next 
  if (wscript.arguments.count > 0) then                       ''SET VARIABLES ACCEPTING ARGUMENTS
    strAPI = objARG.item(0)                                   ''SET REQUIRED PARAMETER 'STRAPI' , API KEY FOR AUTHENTICATION WITH VULTR
  end if
elseif (wscript.arguments.count = 0) then                     ''NOT ENOUGH ARGUMENTS PASSED , END SCRIPT , 'ERRRET'=2
  call LOGERR(2)
  'call CLEANUP()
end if

''------------
''BEGIN SCRIPT
blnLOOP = true
if (errRET = 0) then                                          ''ARGUMENTS PASSED, CONTINUE SCRIPT
	objOUT.write vbnewline & vbnewline & now & vbtab & " - EXECUTING VULTR_SETUP"
	objLOG.write vbnewline & vbnewline & now & vbtab & " - EXECUTING VULTR_SETUP"
	''AUTOMATIC UPDATE, VULTR_SETUP.VBS, REF #2 , REF #69 , REF #68
  objOUT.write vbnewline & now & vbtab & vbtab & " - CHECKING FOR UPDATE : VULTR_SETUP : " & strVER
  objLOG.write vbnewline & now & vbtab & vbtab & " - CHECKING FOR UPDATE : VULTR_SETUP : " & strVER
  ''DOWNLOAD CHKAU.VBS SCRIPT, REF #2 , REF #69 , REF #68
  call FILEDL("https://raw.githubusercontent.com/CW-Khristos/scripts/master/chkAU.vbs", "C:\IT\Scripts", "chkAU.vbs")
  ''EXECUTE CHKAU.VBS SCRIPT, REF #69
  intRET = objWSH.run ("cmd.exe /C " & chr(34) & "cscript.exe " & chr(34) & "C:\IT\Scripts\chkAU.vbs" & chr(34) & " " & _
    chr(34) & strREPO & chr(34) & " " & chr(34) & strBRCH & chr(34) & " " & chr(34) & strDIR & chr(34) & " " & _
    chr(34) & wscript.scriptname & chr(34) & " " & chr(34) & strVER & chr(34) & chr(34), 0, true)
  ''CHKAU RETURNED - NO UPDATE FOUND , REF #2 , REF #69 , REF #68
  objOUT.write vbnewline & "errRET='" & intRET & "'"
  objLOG.write vbnewline & "errRET='" & intRET & "'"
  intRET = (intRET - vbObjectError)
  objOUT.write vbnewline & "errRET='" & intRET & "'"
  objLOG.write vbnewline & "errRET='" & intRET & "'"
  if ((intRET = 4) or (intRET = 10) or (intRET = 11) or (intRET = 1) or (intRET = 2147221505) or (intRET = 2147221517)) then
    set objENV = objWSH.Environment("User")
    objENV("VULTR_API_KEY") = strAPI
    'objOUT.write vbnewline & objENV("VULTR_API_KEY") & vbnewline
    'objOUT.write vbnewline & objWSH.expandenvironmentstrings("%VULTR_API_KEY%") & vbnewline
    objOUT.write vbnewline & now & vbtab & vbtab & " - NO UPDATE FOUND : VULTR_SETUP : " & strVER
    objLOG.write vbnewline & now & vbtab & vbtab & " - NO UPDATE FOUND : VULTR_SETUP : " & strVER
    while blnLOOP
      objOUT.write vbnewline & vbnewline & now & vbtab & vbtab & " - PLEASE MAKE A SELECTION FROM THE FOLLOWING : "
      objLOG.write vbnewline & now & vbtab & vbtab & " - PLEASE MAKE A SELECTION FROM THE FOLLOWING : "
      
      objOUT.write vbnewline & now & vbtab & vbtab & vbtab & "(1) - LIST ALL VULTR INSTANCES"
      objLOG.write vbnewline & now & vbtab & vbtab & vbtab & "(1) - LIST ALL VULTR INSTANCES"
      
      objOUT.write vbnewline & now & vbtab & vbtab & vbtab & "(2) - CREATE A VULTR INSTANCE"
      objLOG.write vbnewline & now & vbtab & vbtab & vbtab & "(2) - CREATE A VULTR INSTANCE"
      
      objOUT.write vbnewline & now & vbtab & vbtab & vbtab & "(3) - UPLOAD SETUPCONFIG TO PBX"
      objLOG.write vbnewline & now & vbtab & vbtab & vbtab & "(3) - UPLOAD SETUPCONFIG TO PBX"
      
      objOUT.write vbnewline & now & vbtab & vbtab & vbtab & "(4) - UPLOAD CERTIFICATE TO PBX"
      objLOG.write vbnewline & now & vbtab & vbtab & vbtab & "(4) - UPLOAD CERTIFICATE TO PBX"
      
      objOUT.write vbnewline & now & vbtab & vbtab & vbtab & "(5) - CREATE A VULTR DNS DOMAIN"
      objLOG.write vbnewline & now & vbtab & vbtab & vbtab & "(5) - CREATE A VULTR DNS DOMAIN"
      
      objOUT.write vbnewline & now & vbtab & vbtab & vbtab & "(6) - UPDATE A VULTR DNS DOMAIN"
      objLOG.write vbnewline & now & vbtab & vbtab & vbtab & "(6) - UPDATE A VULTR DNS DOMAIN"
      
      objOUT.write vbnewline & now & vbtab & vbtab & vbtab & "(7) - QUIT, END SCRIPT" & vbnewline
      objLOG.write vbnewline & now & vbtab & vbtab & vbtab & "(7) - QUIT, END SCRIPT" & vbnewline
      strIN = objIN.readline
      select case strIN
        case 1
          objOUT.write vbnewline & now & vbtab & vbtab & " - LISTING ALL VULTR INSTANCES : "
          objLOG.write vbnewline & now & vbtab & vbtab & " - LISTING ALL VULTR INSTANCES : "
          call HOOK("C:\IT\vultr-cli.exe instance list")
        case 2
          objOUT.write vbnewline & now & vbtab & vbtab & " - CREATING NEW VULTR INSTANCE : "
          objLOG.write vbnewline & now & vbtab & vbtab & " - CREATING NEW VULTR INSTANCE : "
          objOUT.write vbnewline & now & vbtab & vbtab & vbtab & " - SET REGION ID :" & vbnewline
          objLOG.write vbnewline & now & vbtab & vbtab & vbtab & " - SET REGION ID :" & vbnewline
          objWSH.sendkeys strREG
          strREG = objIN.readline
          objOUT.write vbnewline & now & vbtab & vbtab & vbtab & " - SET PLAN ID :" & vbnewline
          objLOG.write vbnewline & now & vbtab & vbtab & vbtab & " - SET PLAN ID :" & vbnewline
          objWSH.sendkeys strPLN
          strPLN = objIN.readline
          objOUT.write vbnewline & now & vbtab & vbtab & vbtab & " - SET OS ID :" & vbnewline
          objLOG.write vbnewline & now & vbtab & vbtab & vbtab & " - SET OS ID :" & vbnewline
          objWSH.sendkeys strOS
          strOS = objIN.readline
          while strCST = vbnullstring
            objOUT.write vbnewline & now & vbtab & vbtab & vbtab & " - ENTER CUSTOMER NAME :" & vbnewline
            objLOG.write vbnewline & now & vbtab & vbtab & vbtab & " - ENTER CUSTOMER NAME :" & vbnewline
            strCST = objIN.readline
          wend
          while strHOST = vbnullstring
            objOUT.write vbnewline & now & vbtab & vbtab & vbtab & " - SET HOSTNAME :" & vbnewline
            objLOG.write vbnewline & now & vbtab & vbtab & vbtab & " - SET HOSTNAME :" & vbnewline
            strHOST = objIN.readline
          wend
          objOUT.write vbnewline & now & vbtab & vbtab & vbtab & " - SET DNS DOMAIN :" & vbnewline
          objLOG.write vbnewline & now & vbtab & vbtab & vbtab & " - SET DNS DOMAIN :" & vbnewline
          objWSH.sendkeys strDMN
          strDMN = objIN.readline
          objOUT.write vbnewline & now & vbtab & vbtab & vbtab & " - SETUP NOTIFICATIONS (Y/N) :" & vbnewline
          objLOG.write vbnewline & now & vbtab & vbtab & vbtab & " - SETUP NOTIFICATIONS (Y/N) :" & vbnewline
          objWSH.sendkeys "Y"
          strIN = objIN.readline
          if (ucase(strIN) = "Y") then
            objOUT.write vbnewline & now & vbtab & vbtab & " - CREATING NEW VULTR INSTANCE : C:\IT\vultr-cli.exe instance create --region " & strREG & _
              " --plan " & strPLN & " --iso " & strISO & " --host " & strHOST & strDMN & " --label " & chr(34) & strHOST & strDMN & " - " & strCST & chr(34) & " --firewall-group " & strFW & " --notify=true"
            objLOG.write vbnewline & now & vbtab & vbtab & " - CREATING NEW VULTR INSTANCE : C:\IT\vultr-cli.exe instance create --region " & strREG & _
              " --plan " & strPLN & " --iso " & strISO & " --host " & strHOST & strDMN & " --label " & chr(34) & strHOST & strDMN & " - " & strCST & chr(34) & " --firewall-group " & strFW & " --notify=true"
            call HOOK("C:\IT\vultr-cli.exe instance create --region " & strREG & " --plan " & strPLN & " --iso " & strISO & " --host " & strHOST & strDMN & _
              " --label " & chr(34) & strHOST & strDMN & " - " & strCST & chr(34) & " --firewall-group " & strFW & " --notify=true")
          elseif (ucase(strIN) = "N") then
            objOUT.write vbnewline & now & vbtab & vbtab & " - CREATING NEW VULTR INSTANCE : C:\IT\vultr-cli.exe instance create --region " & strREG & _
              " --plan " & strPLN & " --iso " & strISO & " --host " & strHOST & strDMN & " --label " & chr(34) & strHOST & strDMN & " - " & strCST & chr(34) & " --firewall-group " & strFW
            objLOG.write vbnewline & now & vbtab & vbtab & " - CREATING NEW VULTR INSTANCE : C:\IT\vultr-cli.exe instance create --region " & strREG & _
              " --plan " & strPLN & " --iso " & strISO & " --host " & strHOST & strDMN & " --label " & chr(34) & strHOST & strDMN & " - " & strCST & chr(34) & " --firewall-group " & strFW
            call HOOK("C:\IT\vultr-cli.exe instance create --region " & strREG & " --plan " & strPLN & " --iso " & strISO & " --host " & strHOST & strDMN & _
              " --label " & chr(34) & strHOST & strDMN & " - " & strCST & chr(34) & " --firewall-group " & strFW)
          end if
          intPBX = 1
          erase arrPBX
          redim arrPBX(1)
          set objTMP = objWSH.exec("C:\IT\vultr-cli.exe instance list")
          while (not objTMP.stdout.atendofstream)
            strIN = objTMP.stdout.readline
            if ((strIN <> vbnullstring) and (instr(1, strIN, "active"))) then
              arrPBX(intPBX) = strIN
              redim preserve arrPBX(intPBX + 1)
              intPBX = intPBX + 1
            end if
          wend
          set objTMP = nothing
          for intPBX = 1 to ubound(arrPBX)
            strLBL = split(arrPBX(intPBX), vbtab)(2)
            if (ucase(strLBL) = ucase(strHOST) & ucase(strDMN) & " - " & ucase(strCST)) then
              strIP = split(arrPBX(intPBX), vbtab)(1)
              exit for
            end if
          next
          set objTMP = objFSO.opentextfile("C:\Users\CBledsoe\IPM-Github\pbxlist.txt", 8)
          objTMP.writeline strIP & "|" & strHOST & strDMN & " - " & strCST
          objTMP.close
          set objTMP = nothing
          objOUT.write vbnewline & now & vbtab & vbtab & " - PBX " & chr(34) & ucase(strHOST) & ucase(strDMN) & " - " & ucase(strCST) & chr(34) & " CREATED"
          objLOG.write vbnewline & now & vbtab & vbtab & " - PBX " & chr(34) & ucase(strHOST) & ucase(strDMN) & " - " & ucase(strCST) & chr(34) & " CREATED"
          objOUT.write vbnewline & now & vbtab & vbtab & " - PLEASE LOGIN TO VULTR DASHBOARD AND ACCESS PBX CONSOLE TO COMPLETE 3CX DEBIAN INSTALLATION"
          objLOG.write vbnewline & now & vbtab & vbtab & " - PLEASE LOGIN TO VULTR DASHBOARD AND ACCESS PBX CONSOLE TO COMPLETE 3CX DEBIAN INSTALLATION"
        case 3
          intPBX = 1
          erase arrPBX
          redim arrPBX(1)
          set objTMP = objWSH.exec("C:\IT\vultr-cli.exe instance list")
          while (not objTMP.stdout.atendofstream)
            strIN = objTMP.stdout.readline
            if ((strIN <> vbnullstring) and (instr(1, strIN, "active"))) then
              arrPBX(intPBX) = strIN
              objOUT.write vbnewline & now & vbtab & "(" & intPBX & ")" & vbtab & strIN 
              objLOG.write vbnewline & now & vbtab & "(" & intPBX & ")" & vbtab & strIN
              redim preserve arrPBX(intPBX + 1)
              intPBX = intPBX + 1
            else
              objOUT.write vbnewline & now & vbtab & vbtab & strIN 
              objLOG.write vbnewline & now & vbtab & vbtab & strIN
            end if
          wend
          set objTMP = nothing
          objOUT.write vbnewline & now & vbtab & vbtab & " - SELECT PBX TO UPLOAD SETUPCONFIG : (1 - " & (intPBX - 1) & ") OR '!Q' TO RETURN TO MAIN MENU" & vbnewline
          objLOG.write vbnewline & now & vbtab & vbtab & " - SELECT PBX TO UPLOAD SETUPCONFIG : (1 - " & (intPBX - 1) & ") OR '!Q' TO RETURN TO MAIN MENU" & vbnewline
          objWSH.sendkeys "1"
          strIN = objIN.readline
          if (ucase(strIN) <> "!Q") then
            objOUT.write vbnewline & now & vbtab & vbtab & vbtab & " - SELECTED PBX : " & vbnewline & vbtab & vbtab & vbtab & arrPBX(strIN)
            objLOG.write vbnewline & now & vbtab & vbtab & vbtab & " - SELECTED PBX : " & vbnewline & vbtab & vbtab & vbtab & arrPBX(strIN)
            objOUT.write vbnewline & now & vbtab & vbtab & vbtab & " - SELECTED IP : " & vbnewline & vbtab & vbtab & vbtab & split(arrPBX(strIN), vbtab)(1)
            objLOG.write vbnewline & now & vbtab & vbtab & vbtab & " - SELECTED IP : " & vbnewline & vbtab & vbtab & vbtab & split(arrPBX(strIN), vbtab)(1)
            strPBX = split(arrPBX(strIN), vbtab)(1)
            objOUT.write vbnewline & now & vbtab & vbtab & vbtab & " - ENTER SELECTED PBX USER LOGIN :" & vbnewline
            objLOG.write vbnewline & now & vbtab & vbtab & vbtab & " - ENTER SELECTED PBX USER LOGIN :" & vbnewline
            objWSH.sendkeys "root"
            strUSR = objIN.readline
            objOUT.write vbnewline & now & vbtab & vbtab & vbtab & " - ENTER SELECTED PBX USER PASSWORD :" & vbnewline
            objLOG.write vbnewline & now & vbtab & vbtab & vbtab & " - ENTER SELECTED PBX USER PASSWORD :" & vbnewline
            objWSH.sendkeys "Ipmcomputers1"
            strPWD = objIN.readline
            strXML = "C:\IT\3cx\upload\setupconfig.xml"
            strRCMD = "cmd.exe /c copy /Y C:\Users\CBledsoe\IPM-Github\setupconfig.xml " & strXML
            objOUT.write vbnewline & now & vbtab & vbtab & " - COPYING SETUPCONFIG : " & strXML
            objLOG.write vbnewline & now & vbtab & vbtab & " - COPYING SETUPCONFIG : " & strXML
            objOUT.write vbnewline & vbnewline & strRCMD
            call HOOK(strRCMD)
            wscript.sleep 1000
            objOUT.write vbnewline & now & vbtab & vbtab & " - UPLOADING SETUPCONFIG : " & strXML
            objLOG.write vbnewline & now & vbtab & vbtab & " - UPLOADING SETUPCONFIG : " & strXML
            'strRCMD = strSCP & " /command " & chr(34) & "open scp://" & strUSR & ":" & strPWD & "@" & strPBX & ":22/ -hostkey=acceptnew" & chr(34) & " " & _
            '  chr(34) & "put " & strXML & " /var/lib/3cxpbx/Bin/nginx/conf/Instance1/" & chr(34) & " " & chr(34) & "exit" & chr(34) & " /log=" & chr(34) & "C:\temp\pbx_setupconfig.log" & chr(34) & " /loglevel=0"
            'objOUT.write vbnewline & vbnewline & strRCMD
            'call HOOK(strRCMD)
            strRCMD = strSCP & " /command " & chr(34) & "open scp://" & strUSR & ":" & strPWD & "@" & strPBX & ":22/ -hostkey=acceptnew" & chr(34) & " " & _
              chr(34) & "put " & strXML & " /etc/3cxpbx/" & chr(34) & " " & chr(34) & "exit" & chr(34) & " /log=" & chr(34) & "C:\temp\pbx_setupconfig.log" & chr(34) & " /loglevel=0"
            objOUT.write vbnewline & vbnewline & strRCMD
            call HOOK(strRCMD)
            objFSO.deletefile strXML, true
          end if
        case 4
          intPBX = 1
          erase arrPBX
          set objTMP = objWSH.exec("C:\IT\vultr-cli.exe instance list")
          while (not objTMP.stdout.atendofstream)
            strIN = objTMP.stdout.readline
            if ((strIN <> vbnullstring) and (instr(1, strIN, "active"))) then
              arrPBX(intPBX) = strIN
              objOUT.write vbnewline & now & vbtab & "(" & intPBX & ")" & vbtab & strIN 
              objLOG.write vbnewline & now & vbtab & "(" & intPBX & ")" & vbtab & strIN
              redim preserve arrPBX(intPBX + 1)
              intPBX = intPBX + 1
            else
              objOUT.write vbnewline & now & vbtab & vbtab & strIN 
              objLOG.write vbnewline & now & vbtab & vbtab & strIN
            end if
          wend
          set objTMP = nothing
          objOUT.write vbnewline & now & vbtab & vbtab & " - SELECT PBX TO UPLOAD CERT : (1 - " & (intPBX - 1) & ") OR '!Q' TO RETURN TO MAIN MENU" & vbnewline
          objLOG.write vbnewline & now & vbtab & vbtab & " - SELECT PBX TO UPLOAD CERT : (1 - " & (intPBX - 1) & ") OR '!Q' TO RETURN TO MAIN MENU" & vbnewline
          objWSH.sendkeys "1"
          strIN = objIN.readline
          if (ucase(strIN) <> "!Q") then
            objOUT.write vbnewline & now & vbtab & vbtab & vbtab & " - SELECTED PBX : " & vbnewline & vbtab & vbtab & vbtab & arrPBX(strIN)
            objLOG.write vbnewline & now & vbtab & vbtab & vbtab & " - SELECTED PBX : " & vbnewline & vbtab & vbtab & vbtab & arrPBX(strIN)
            objOUT.write vbnewline & now & vbtab & vbtab & vbtab & " - SELECTED IP : " & vbnewline & vbtab & vbtab & vbtab & split(arrPBX(strIN), vbtab)(1)
            objLOG.write vbnewline & now & vbtab & vbtab & vbtab & " - SELECTED IP : " & vbnewline & vbtab & vbtab & vbtab & split(arrPBX(strIN), vbtab)(1)
            strPBX = split(arrPBX(strIN), vbtab)(1)
            objOUT.write vbnewline & now & vbtab & vbtab & vbtab & " - ENTER SELECTED PBX USER LOGIN :" & vbnewline
            objLOG.write vbnewline & now & vbtab & vbtab & vbtab & " - ENTER SELECTED PBX USER LOGIN :" & vbnewline
            objWSH.sendkeys "root"
            strUSR = objIN.readline
            objOUT.write vbnewline & now & vbtab & vbtab & vbtab & " - ENTER SELECTED PBX USER PASSWORD :" & vbnewline
            objLOG.write vbnewline & now & vbtab & vbtab & vbtab & " - ENTER SELECTED PBX USER PASSWORD :" & vbnewline
            objWSH.sendkeys "Ipmcomputers1"
            strPWD = objIN.readline
            ''DOWNLOAD PBXUPLOAD.VBS SCRIPT
            objOUT.write vbnewline & now & vbtab & vbtab & " - DOWNLOADING SCRIPT : PBXUPLOAD : "
            objLOG.write vbnewline & now & vbtab & vbtab & " - DOWNLOADING SCRIPT : PBXUPLOAD : "
            call FILEDL("https://raw.githubusercontent.com/CW-Khristos/scripts/dev/VULTR/PBXupload.vbs", "C:\IT\Scripts", "PBXupload.vbs")
            ''EXECUTE PBXUPLOAD.VBS SCRIPT
            call HOOK("cscript.exe " & chr(34) & "C:\IT\Scripts\PBXupload.vbs" & chr(34) & " " & chr(34) & strUSR & chr(34) & " " & chr(34) & strPWD & chr(34) & " " & chr(34) & strPBX & chr(34))
          end if
        case 5
          objOUT.write vbnewline & now & vbtab & vbtab & " - CREATING NEW VULTR DNS DOMAIN : "
          objLOG.write vbnewline & now & vbtab & vbtab & " - CREATING NEW VULTR DNS DOMAIN : "
          objOUT.write vbnewline & now & vbtab & vbtab & vbtab & " - SET DNS DOMAIN NAME :" & vbnewline
          objLOG.write vbnewline & now & vbtab & vbtab & vbtab & " - SET DNS DOMAIN NAME :" & vbnewline
          strDMN = objIN.readline
          if (ucase(strDMN) <> "!Q") then
            objOUT.write vbnewline & now & vbtab & vbtab & vbtab & " - SET IP ADDRESS :" & vbnewline
            objLOG.write vbnewline & now & vbtab & vbtab & vbtab & " - SET IP ADDRESS :" & vbnewline
            stRIP = objIN.readline
            call HOOK("C:\IT\vultr-cli.exe dns domain create --domain " & strDMN & " --ip " & strIP)
          end if
        case 6

        case 7
          blnLOOP = false
      end select
    wend
  end if
elseif (errRET <> 0) then                                     ''NO ARGUMENTS PASSED, END SCRIPT , 'ERRRET'=1
  call LOGERR(errRET)
end if
''END SCRIPT
call CLEANUP()
''END SCRIPT
''------------

''SUB-ROUTINES
sub FILEDL(strURL, strDL, strFILE)                          ''CALL HOOK TO DOWNLOAD FILE FROM URL , 'ERRRET'=11
  strSAV = vbnullstring
  ''SET DOWNLOAD PATH
  strSAV = strDL & "\" & strFILE
  objOUT.write vbnewline & now & vbtab & vbtab & " - HTTPDOWNLOAD-------------DOWNLOAD : " & strURL & " : SAVE AS :  " & strSAV
  objLOG.write vbnewline & now & vbtab & vbtab & " - HTTPDOWNLOAD-------------DOWNLOAD : " & strURL & " : SAVE AS :  " & strSAV
  ''ADD WINHTTP SECURE CHANNEL TLS REGISTRY KEYS
  call HOOK("reg add " & chr(34) & "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\WinHttp" & chr(34) & _
    " /f /v DefaultSecureProtocols /t REG_DWORD /d 0x00000A00 /reg:32")
  call HOOK("reg add " & chr(34) & "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\WinHttp" & chr(34) & _
    " /f /v DefaultSecureProtocols /t REG_DWORD /d 0x00000A00 /reg:64")
  ''CHECK IF FILE ALREADY EXISTS
  if (objFSO.fileexists(strSAV)) then
    ''DELETE FILE FOR OVERWRITE
    objFSO.deletefile(strSAV)
  end if
  ''CREATE HTTP OBJECT
  set objHTTP = createobject("WinHttp.WinHttpRequest.5.1")
  ''DOWNLOAD FROM URL
  objHTTP.open "GET", strURL, false
  objHTTP.send
  if (objHTTP.status = 200) then
    dim objStream
    set objStream = createobject("ADODB.Stream")
    with objStream
      .Type = 1 'adTypeBinary
      .Open
      .Write objHTTP.ResponseBody
      .SaveToFile strSAV
      .Close
    end with
    set objStream = nothing
  end if
  ''CHECK THAT FILE EXISTS
  if (objFSO.fileexists(strSAV)) then
    objOUT.write vbnewline & now & vbtab & vbtab & " - DOWNLOAD : " & strSAV & " : SUCCESSFUL"
    objLOG.write vbnewline & now & vbtab & vbtab & " - DOWNLOAD : " & strSAV & " : SUCCESSFUL"
  end if
	set objHTTP = nothing
  if ((err.number <> 0) and (err.number <> 58)) then        ''ERROR RETURNED DURING UPDATE CHECK , 'ERRRET'=11
    call LOGERR(11)
  end if
end sub

sub HOOK(strCMD)                                            ''CALL HOOK TO MONITOR OUTPUT OF CALLED COMMAND , 'ERRRET'=12
  on error resume next
  objOUT.write vbnewline & now & vbtab & vbtab & " - EXECUTING : HOOK(" & strCMD & ")"
  objLOG.write vbnewline & now & vbtab & vbtab & " - EXECUTING : HOOK(" & strCMD & ")"
  set objHOOK = objWSH.exec(strCMD)
  if (instr(1, strCMD, "takeown /F ") = 0) then             ''SUPPRESS 'TAKEOWN' SUCCESS MESSAGES
    while (not objHOOK.stdout.atendofstream)
      strIN = objHOOK.stdout.readline
      if (strIN <> vbnullstring) then
        objOUT.write vbnewline & now & vbtab & vbtab & vbtab & strIN 
        objLOG.write vbnewline & now & vbtab & vbtab & vbtab & strIN 
      end if
    wend
    wscript.sleep 10
    strIN = objHOOK.stdout.readall
    if (strIN <> vbnullstring) then
      objOUT.write vbnewline & now & vbtab & vbtab & vbtab & strIN 
      objLOG.write vbnewline & now & vbtab & vbtab & vbtab & strIN 
    end if
  end if
  set objHOOK = nothing
  if (err.number <> 0) then                                 ''ERROR RETURNED DURING UPDATE CHECK , 'ERRRET'=12
    call LOGERR(12)
  end if
end sub

sub LOGERR(intSTG)                                          ''CALL HOOK TO MONITOR OUTPUT OF CALLED COMMAND
  errRET = intSTG
  if (err.number <> 0) then
    objOUT.write vbnewline & now & vbtab & vbtab & vbtab & err.number & vbtab & err.description & vbnewline
    objLOG.write vbnewline & now & vbtab & vbtab & vbtab & err.number & vbtab & err.description & vbnewline
		err.clear
  end if
  select case intSTG
    case 0                                                  ''VULTR_SETUP - VULTR-CLI CHECK PASSED, 'ERRRET'=0
      objOUT.write vbnewline & vbnewline & now & vbtab & " - VULTR_SETUP - VULTR-CLI CHECK PASSED"
      objLOG.write vbnewline & vbnewline & now & vbtab & " - VULTR_SETUP - VULTR-CLI CHECK PASSED"
    case 1                                                  ''VULTR_SETUP - VULTR-CLI CHECK FAILED, 'ERRRET'=1
      objOUT.write vbnewline & vbnewline & now & vbtab & " - VULTR_SETUP - VULTR-CLI CHECK FAILED, ENDING VULTR_SETUP"
      objLOG.write vbnewline & vbnewline & now & vbtab & " - VULTR_SETUP - VULTR-CLI CHECK FAILED, ENDING VULTR_SETUP"
    case 2                                                  ''VULTR_SETUP - NOT ENOUGH ARGUMENTS, 'ERRRET'=2
      objOUT.write vbnewline & vbnewline & now & vbtab & " - VULTR_SETUP - NO ARGUMENTS PASSED, END SCRIPT"
      objLOG.write vbnewline & vbnewline & now & vbtab & " - VULTR_SETUP - NO ARGUMENTS PASSED, END SCRIPT"
    case 11                                                 ''VULTR_SETUP - CALL FILEDL() FAILED, 'ERRRET'=11
      objOUT.write vbnewline & vbnewline & now & vbtab & " - VULTR_SETUP - CALL FILEDL() : " & strSAV
      objLOG.write vbnewline & vbnewline & now & vbtab & " - VULTR_SETUP - CALL FILEDL() : " & strSAV
    case 12                                                 ''VULTR_SETUP - 'CALL HOOK() FAILED, 'ERRRET'=12
      objOUT.write vbnewline & vbnewline & now & vbtab & " - VULTR_SETUP - CALL HOOK('STRCMD') : " & strRCMD & " : FAILED"
      objLOG.write vbnewline & vbnewline & now & vbtab & " - VULTR_SETUP - CALL HOOK('STRCMD') : " & strRCMD & " : FAILED"
  end select
end sub

sub CLEANUP()                                               ''SCRIPT CLEANUP
  on error resume next
  if (errRET = 0) then         															''VULTR_SETUP COMPLETED SUCCESSFULLY
    objOUT.write vbnewline & vbnewline & now & vbtab & " - VULTR_SETUP SUCCESSFUL : " & now
    objLOG.write vbnewline & vbnewline & now & vbtab & " - VULTR_SETUP SUCCESSFUL : " & now
    err.clear
  elseif (errRET <> 0) then    															''VULTR_SETUP FAILED
    objOUT.write vbnewline & vbnewline & now & vbtab & " - VULTR_SETUP FAILURE : " & errRET & " : " & now
    objLOG.write vbnewline & vbnewline & now & vbtab & " - VULTR_SETUP FAILURE : " & errRET & " : " & now
    ''RAISE CUSTOMIZED ERROR CODE, ERROR CODE WILL BE DEFINE RESTOP NUMBER INDICATING WHICH SECTION FAILED
    call err.raise(vbObjectError + errRET, "VULTR_SETUP", "FAILURE")
  end if
  objOUT.write vbnewline & vbnewline & now & " - VULTR_SETUP COMPLETE" & vbnewline
  objLOG.write vbnewline & vbnewline & now & " - VULTR_SETUP COMPLETE" & vbnewline
  objLOG.close
  ''EMPTY OBJECTS
  set objLOG = nothing
  set objFSO = nothing
  set objWSH = nothing
  set objARG = nothing
  set objOUT = nothing
  set objIN = nothing
  ''END SCRIPT, RETURN ERROR NUMBER
  wscript.quit err.number
end sub
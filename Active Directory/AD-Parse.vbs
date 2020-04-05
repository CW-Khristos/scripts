	''DEFINE VARIABLES FOR USE IN SCRIPT''
Const ADS_LOCKOUT = &H0010
Const ADS_DISABLED = &H0002
Const ADS_PWD_NOTREQD = &H0020
Const ADS_PWD_EXPIRED = &H800000
Const ADS_DONT_EXPIRE_PWD = &H10000
Const ADS_SMARTCARD = &H40000

Dim  strPROP, strEDT, strMOV, strYN
Dim strDN, strOU, strDIS, strSEL, strACT, strMOU
Dim objDSP, objFSO, objLog, objOU, objAD, objDIS, objDISch, objMOV

Set objDSP = GetObject("LDAP:")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objLog = objFSO.createtextfile("C:\AD-Parse." & Year(Now) & Month(Now) & Day(Now) & ".log")

Call Main()
Call scrClean()
     
Sub Main()
  strDN = vbnullstring
  strSEL = vbnullstring
  strOU = vbnullstring
  strACT = vbnullstring
  strPROP = vbnullstring
  strEDT = vbnullstring
  strMOV = vbnullstring
  strYN = vbnullstring

  While strSEL <> "!ABORT!"
    WScript.echo vbnewline & "*****************************************************************"
    WScript.echo "*****************************************************************"
    WScript.echo "***                V--V     AD-Parse v2    V--V               ***"
    WScript.echo "***                V--V VampyreMan Studios V--V               ***"
    WScript.echo "***                V--V  Author : Khristos V--V               ***"
    WScript.echo "***  MORE ADVANCED THAN THE ORIGINAL AD-PARSE, THIS VERSION   ***"
    WScript.Echo "***   OFFERS THE ABILITY TO CUSTOMIZE ITS SETTINGS FOR MORE   ***"
    WScript.Echo "***     THAN JUST DISABLING ACCOUNTS BASED ON LAST-LOGON      ***"
    WScript.Echo "***    OFFERS THE ABILITY TO SEARCH AD FOR SELECTIVE NAME,    ***"
    WScript.Echo "*** DESCRIPTION, ACCOUNT DISABLED, LAST-LOGON, AND SMARTCARD  ***"
    WScript.Echo "***  AND ALLOWS ADMINISTRATOR TO LOG, DISABLE, EDIT, OR MOVE  ***"
    WScript.Echo "***     ACCOUNTS BASED UPON THESE DEFINED PROPERTY VALUES     ***"
    WScript.echo "*****************************************************************"
    WScript.echo "*****************************************************************"

		''ROUTINE TO SET DOMAIN''
    While strDN = vbnullstring
      WScript.echo vbnewline & "Type FQDN For Domain (my.example.domain.com) :"
      strDN = LCase(WScript.stdin.readline)
      If strDN = vbnullstring Then
        WScript.echo vbnewline & "INVALID INPUT (THIS FIELD CANNOT BE BLANK)"
      End If
    Wend

    strDN = "dc=" & Replace(strDN, ".", ",dc=")

		''ROUTINE TO SET SEARCH PATH''
    While strSEL <> "dn" And strSEL <> "ou" Or strSEL = vbnullstring
      WScript.echo vbnewline & "Select to Search Through Entire Domain or Specific OU (dn / ou) :"
      strSEL = LCase(WScript.stdin.readline)
      If strSEL = vbnullstring Then
        WScript.echo vbnewline & "INVALID INPUT (THIS FIELD CANNOT BE BLANK)"
      ElseIf strSEL <> "dn" And strSEL <> "ou" Then
        WScript.echo vbnewline & "INVALID INPUT (MUST SELECT EITHER DN OR OU)"
      End If
    Wend

    If strSEL = "ou" Then

      While strOU = vbnullstring
        WScript.echo vbnewline & "Type Relative Path to OU Under the FQDN (targetou.targetouparentou...) :"
        strOU = LCase(WScript.stdin.readline)
        If strOU = vbnullstring Then
          WScript.echo vbnewline & "INVALID INPUT (THIS FIELD CANNOT BE BLANK)"
        End If
      Wend

      If InStr(1, strOU, ".") Then
        strOU = Replace(strOU, ".", ",ou=")
      End If
      strOU = "ou=" & strOU
    ElseIf strSEL = "dn" Then
      strOU = strDN
    End If

    While strSEL <> "log" And strSEL <> "disable" And strSEL <> "edit" And strSEL <> "move" Or strSEL = vbnullstring
      WScript.echo vbnewline & "Select Action to Take (log / disable / edit / move) :"
      strSEL = LCase(WScript.stdin.readline)
      If strSEL = vbnullstring Then
        WScript.echo vbnewline & "INVALID INPUT (THIS FIELD CANNOT BE BLANK)"
      ElseIf strSEL <> "log" And strSEL <> "disable" And strSEL <> "edit" And strSEL <> "move" Then
        WScript.echo vbnewline & "INVALID INPUT (MUST SELECT LOG OR DISABLE OR EDIT)"
      End If
    Wend

    strACT = strSEL

    While strSEL <> "computer" And strSEL <> "user" Or strSEL = vbnullstring
      WScript.echo vbnewline & "Select Which Type of Object to Search For (computer / user) :"
      strSEL = LCase(WScript.stdin.readline)
      If strSEL = vbnullstring Then
        WScript.echo vbnewline & "INVALID INPUT (THIS FIELD CANNOT BE BLANK)"
      ElseIf strSEL <> "computer" And strSEL <> "user" Then
        WScript.echo vbnewline & "INVALID INPUT (MUST SELECT EITHER COMPUTERS OR USERS)"
      End If
    Wend

    If strACT = "log" Then

      While InStr(1, strPROP, "name=") = 0 And InStr(1, strPROP, "description=") = 0 And InStr(1, strPROP, "disabled=") = 0 And _
        InStr(1, strPROP, "lastlogon=") = 0 And InStr(1, strPROP, "smartcard=") = 0 Or strPROP = vbnullstring
          WScript.echo vbnewline & "Log Accounts By Any of These Values :"
          WScript.echo "(name / description / disabled / lastlogon [>=#days] / smartcard [users])"
          WScript.echo "Format as 'property=value' Seperate Multiple Values With ','"
	  WScript.Echo "Setting 'name=*' Will Log All Accounts Under The Selected OU"
          strPROP = LCase(WScript.stdin.readline)
          If strPROP = vbnullstring Then
            WScript.echo vbnewline & "INVALID INPUT (THIS FIELD CANNOT BE BLANK)"
          ElseIf InStr(1, strPROP, "name=") = 0 And InStr(1, strPROP, "description=") = 0 And InStr(1, strPROP, "disabled=") = 0 And _
            InStr(1, strPROP, "lastlogon=") = 0 And InStr(1, strPROP, "smartcard=") = 0 Then
              WScript.echo vbnewline & "INVALID INPUT (MUST SELECT EITHER NAME, DESCRIPTION, DISABLED, LASTLOGON, OR SMARTCARD)"
          End If
      Wend

    ElseIf strACT = "disable" Then
      WScript.echo vbnewline & "Type Relative Path to Domain's 'Disabled' OU (disabled.parentou...) :"
      WScript.echo "Leave Blank to Just Disable Accounts and Not Move Them"
      strDIS = LCase(WScript.stdin.readline)
      If strDIS <> vbnullstring Then
        If InStr(1, strDIS, ".") Then
          strDIS = Replace(strDIS, ".", ",ou=")
        End If
        strDIS = "ou=" & strDIS
       'set objDIS = getobject("LDAP://" & strDis)
       'set objDISch = objDIS.create("organizationalUnit", "ou=DISABLED_" & ucase(strSEL) & "_" & Year(Now) & Month(Now) & Day(Now)) 
       'objDISch.setinfo
      End If

      While InStr(1, strPROP, "name=") = 0 And InStr(1, strPROP, "description=") = 0 And InStr(1, strPROP, "lastlogon=") = 0 And _
        InStr(1, strPROP, "smartcard=") = 0 Or strPROP = vbnullstring
          WScript.echo vbnewline & "Disable Accounts By Any of These Values :"
          WScript.echo "(name / description / lastlogon [>=#days] / smartcard [users])"
          WScript.echo "Format as 'property=value' Seperate Multiple Values With ','"
	  WScript.Echo "Setting 'name=*' Will Disable *ALL* Accounts Under The Selected OU"
          strPROP = LCase(WScript.stdin.readline)
          If strPROP = vbnullstring Then
            WScript.echo vbnewline & "INVALID INPUT (THIS FIELD CANNOT BE BLANK)"
          ElseIf InStr(1, strPROP, "name=") = 0 And InStr(1, strPROP, "description=") = 0 And InStr(1, strPROP, "lastlogon=") = 0 And _
            InStr(1, strPROP, "smartcard=") = 0 Then
              WScript.echo vbnewline & "INVALID INPUT (MUST SELECT EITHER NAME, DESCRIPTION, LASTLOGON, OR SMARTCARD)"
          End If
      Wend

    ElseIf strACT = "edit" Then

      While InStr(1, strPROP, "name=") = 0 And InStr(1, strPROP, "description=") = 0 And InStr(1, strPROP, "disabled=") = 0 And _
        InStr(1, strPROP, "smartcard=") = 0 Or strPROP = vbnullstring
          WScript.echo vbnewline & "!!!USE CAUTION WHEN USING THIS ACTION ON MULTIPLE ACCOUNTS!!!"
          WScript.echo "Edit Accounts By Any of These Values :"
          WScript.echo "(name / description / disabled / smartcard [users]) :"
          WScript.echo "Format as 'property=value' Seperate Multiple Values With ','"
	  WScript.Echo "Setting 'name=*' Will Edit *ALL* Accounts Under The Selected OU"
          strPROP = LCase(WScript.stdin.readline)
          If strPROP = vbnullstring Then
            WScript.echo vbnewline & "INVALID INPUT (THIS FIELD CANNOT BE BLANK)"
          ElseIf InStr(1, strPROP, "name=") = 0 And InStr(1, strPROP, "description=") = 0 And InStr(1, strPROP, "disabled=") = 0 And _
            InStr(1, strPROP, "smartcard=") = 0 Then
              WScript.echo vbnewline & "INVALID INPUT (MUST SELECT EITHER NAME, DESCRIPTION, DISABLED, OR SMARTCARD)"
          End If
      Wend

      While strEDT = vbnullstring
        WScript.echo vbnewline & "!!!USE CAUTION WHEN USING THIS ACTION ON MULTIPLE ACCOUNTS!!!"
        WScript.echo "Type New Value For Property :"
        WScript.echo "Format as 'property=newvalue' Seperate Multiple Values With ','"
        strEDT = LCase(WScript.stdin.readline)
        If strEDT = vbnullstring Then
          WScript.echo vbnewline & "INVALID INPUT (THIS FIELD CANNOT BE BLANK)"
        End If
      Wend

    ElseIf strACT = "move" Then

      While InStr(1, strPROP, "name=") = 0 And InStr(1, strPROP, "description=") = 0 And InStr(1, strPROP, "disabled=") = 0 And _
        InStr(1, strPROP, "lastlogon=") = 0 And InStr(1, strPROP, "smartcard=") = 0 Or strPROP = vbnullstring
          WScript.echo vbnewline & "Move Accounts By These Values :"
          WScript.echo "(name / description / disabled / lastlogon [>=#days] / smartcard [users])"
	  WScript.echo "Format as 'property=value' Seperate Multiple Values With ','"
	  WScript.Echo "Setting 'name=*' Will Move *ALL* Accounts Under The Selected OU"
          strPROP = LCase(WScript.stdin.readline)
          If strPROP = vbnullstring Then
            WScript.echo vbnewline & "INVALID INPUT (THIS FIELD CANNOT BE BLANK)"
          ElseIf InStr(1, strPROP, "name=") = 0 And InStr(1, strPROP, "description=") = 0 And InStr(1, strPROP, "disabled=") = 0 And _
            InStr(1, strPROP, "lastlogon=") = 0 And InStr(1, strPROP, "smartcard=") = 0 Then
              WScript.echo vbnewline & "INVALID INPUT (MUST SELECT EITHER NAME, DESCRIPTION, DISABLED, LASTLOGON, OR SMARTCARD)"
          End If
      Wend

      While strMOV = vbnullstring
        WScript.echo vbnewline & "Type Relative Path of OU to Move Accounts Under (movehere.parentou...) :"
        strMOV = LCase(WScript.stdin.readline)
        If strMOV = vbnullstring Then
          WScript.echo vbnewline & "INVALID INPUT (THIS FIELD CANNOT BE BLANK)"
        ElseIf InStr(1, strMOV, ".") Then
          strMOV = Replace(strMOV, ".", ",ou=")
        End If
      Wend

      strMOV = "ou=" & strMOV
      Set objMOV = GetObject("LDAP://" & strMOV)
    End If

    WScript.echo vbnewline & "*****************************************************************"
    WScript.echo "***                   AD-PARSE SETTINGS                       ***"
    WScript.echo "*****************************************************************"

    WScript.echo vbnewline & "DOMAIN TO SEARCH : " & strDN
    WScript.echo "AD PATH TO SEARCH : " & strOU
    WScript.echo "TYPE OF ACCOUNTS TO CHECK : " & strSEL
    WScript.echo "ACTION TO BE PERFORMED : " & strACT & " accounts"
    If strACT = "log" Then
      WScript.echo "LOGGING ACCOUNTS WITH PROPERTY VALUE(S) : " & strPROP
      WScript.echo "WRITING LOG TO : " & "C:\AD-Parse." & Year(Now) & Month(Now) & Day(Now) & ".log"
    ElseIf strACT = "disable" Then
      WScript.echo "DISABLING ACCOUNTS WITH PROPERTY VALUE(S) : " & strPROP
      If strDIS <> vbnullstring Then
        WScript.echo "CREATING A DISABLED OU : YES : ou=DISABLED_" & UCase(strSEL) & "_" & Year(Now) & Month(Now) & Day(Now) & "," & strDIS
      Else
        WScript.echo "CREATING A DISABLED OU : NO"
      End If
    ElseIf strACT = "edit" Then
      WScript.echo "CHANGING ACCOUNTS WITH PROPERTY VALUE(S) : " & strPROP
      WScript.echo "CHANGE(S) OF PROPERTY VALUE(S) : " & strEDT
    ElseIf strACT = "move" Then
      WScript.echo "MOVING ACCOUNTS WITH PROPERTY VALUE(S) : " & strPROP
      WScript.echo "OU TO MOVE ACCOUNTS UNDER : " & strMOV
    End If

    WScript.echo vbnewline & "*****************************************************************"
    WScript.echo "*****************************************************************"

    While strYN <> "y" And strYN <> "n" Or strYN = vbnullstring
      WScript.echo vbnewline & "RUN SCRIPT WITH THESE SETTINGS (Y / N) :"
      strYN = LCase(WScript.stdin.readline)
      If strYN <> "y" And strYN <> "no" Or strYN = vbnullstring Then
        WScript.echo vbnewline & "INVALID INPUT (MUST SELECT Y OR N)"
      End If
    Wend

    If strYN = "y" Then
      WScript.echo vbnewline & "*****************************************************************"
      WScript.echo "SCANNING OU : " & Time() & " : " & strOU & "," & strDN
      WScript.echo "*****************************************************************"
      Set objOU = objDSP.opendsobject("LDAP://" & strOU & "," & strDN, vbnullstring, vbnullstring, ADS_SERVER_BIND)
      Call Parser(objOU)
    ElseIf strYN = "n" Then
      strDN = vbnullstring
      strSEL = vbnullstring
      strOU = vbnullstring
      strACT = vbnullstring
      strPROP = vbnullstring
      strEDT = vbnullstring
      strMOV = vbnullstring
      strYN = vbnullstring
    End If
  Wend

End Sub

Sub Parser(objADPath)
  on error resume next
  Dim intLogon, arrPROP, ufSC, blnFND, chkArray(11), x
  Dim objADName, objADDesc, objADDis, objADCAC, objADLogon

  For Each objAD In objADPath
    blnFND = "TRUE"
    chkArray(0) = "smb"
    chkArray(1) = "syscon"
    chkArray(2) = "service"
    chkArray(3) = "sharepoint"
    chkArray(4) = "dont disable"
    chkArray(5) = "don't disable"
    chkArray(6) = "do not disable"
    chkArray(7) = "security group"
    chkArray(8) = "permission group"
    chkArray(9) = "distribution group"
    chkArray(10) = "permissions group"

    objADName = LCase(objAD.name)
    objADDesc = LCase(objAD.description)
    objADDis = LCase(objAD.accountdisabled)

    for x = 0 to 10
      if objADDis = "true" then
        blnFND = "FALSE"
        exit for
      end if
      if instr(1, objADName, chkArray(x)) then
        blnFND = "FALSE"
        exit for
      elseif instr(1, objADDesc, chkArray(x)) then
        blnFND = "FALSE"
        exit for
      end if
    next

    if blnFND = "TRUE" then

      Select Case LCase(objAD.Class)

        Case strSEL

          If strSEL = "user" Then
	    ufSC = objAD.get("useraccountcontrol")
            objADCAC = (ufSC And 262144)
            If objADCAC = 262144 Then
	      objADCAC = "true"
	    Else
	      objADCAC = "false"
	    End If
          End If
          Set objADLogon = objAD.get("lastlogontimestamp")
          intLogon = objLogon.highpart * (2 ^ 32) + objLogon.lowpart
          intLogon = intLogon / (60 * 10000000)
          intLogon = intLogon / 1440
          intLogon = intLogon + # 1 / 1 / 1601 #
          If InStr(1,strPROP, "name=*") Then
            blnFND = "TRUE"
          ElseIf InStr(1, strPROP, ",") Then
            arrPROP = Split(strPROP, ",")

            For x = 0 To UBound(arrPROP)
              If InStr(1, arrPROP(x), "name=") Then
                If InStr(1, objAD.name, Mid(arrPROP(x), InStr(1,arrPROP(x), "=") + 1, Len(arrPROP(x)) - InStr(1, arrPROP(x), "="))) = 0 Then
                  blnFND = "FALSE"
                  Exit For
                End If
              ElseIf InStr(1, arrPROP(x), "description=") Then
                If InStr(1, objAD.description, Mid(arrPROP(x), InStr(1,arrPROP(x), "=") + 1, Len(arrPROP(x)) - InStr(1, arrPROP(x), "="))) = 0 Then
                  blnFND = "FALSE"
                  Exit For
                End If
              ElseIf InStr(1, arrPROP(x), "disabled=") Then
                If objADDis <> Mid(arrPROP(x), InStr(1,arrPROP(x), "=") + 1, Len(arrPROP(x)) - InStr(1, arrPROP(x), "=")) Then
                  blnFND = "FALSE"
                  Exit For
                End If
              ElseIf InStr(1, arrPROP(x), "lastlogon=") Then
                If intLogon < CInt(Mid(arrPROP(x), InStr(1,arrPROP(x), "=") + 1, Len(arrPROP(x)) - InStr(1, arrPROP(x), "="))) Then
                  blnFND = "FALSE"
                  Exit For
                End If
              ElseIf InStr(1, arrPROP(x), "smartcard=") Then
                If objADCAC <> Mid(arrPROP(x), InStr(1,arrPROP(x), "=") + 1, Len(arrPROP(x)) - InStr(1, arrPROP(x), "=")) Then
                  blnFND = "FALSE"
                  Exit For
                End If
              End If
            Next

          Else
            If InStr(1, strPROP, "name=") Then
              If InStr(1,objADName, Mid(strPROP, InStr(1, strPROP, "=") + 1, Len(strPROP) - InStr(1, strPROP, "="))) = 0 Then
                blnFND = "FALSE"
              End If
            ElseIf InStr(1, strPROP, "description=") Then
              If InStr(1,objADDesc, Mid(strPROP, InStr(1, strPROP, "=") + 1, Len(strPROP) - InStr(1, strPROP, "="))) = 0 Then
                blnFND = "FALSE"
              End If
            ElseIf InStr(1, strPROP, "disabled=") Then
              If objADDis <> Mid(strPROP, InStr(1, strPROP, "=") + 1, Len(strPROP) - InStr(1, strPROP, "=")) Then
                blnFND = "FALSE"
              End If
            ElseIf InStr(1, strPROP, "lastlogon=") Then
              If intLogon < CInt(Mid(strPROP, InStr(1,strPROP, "=") + 1, Len(strPROP) - InStr(1, strPROP, "="))) Then
                blnFND = "FALSE"
              End If
            ElseIf InStr(1, strPROP, "smartcard=") Then
              If objADCAC <> Mid(strPROP, InStr(1, strPROP, "=") + 1, Len(strPROP) - InStr(1, strPROP, "=")) Then
                blnFND = "FALSE"
              End If
            End If
          End If
          If blnFND = "TRUE" Then
            WScript.echo vbnewline & "*****************************************************************"
            WScript.echo "MATCH FOUND : " & strACT & " @ " & Time() & " : " & objAD.distinguishedname
            WScript.echo "*****************************************************************"
            If strACT = "log" Then
              objLog.writeline "Logged : " & Time() & " : " & objAD.distinguishedname & " : By Match : " & strPROP & vbnewline
            ElseIf strACT = "disable" Then
              'objAD.accountdisabled = True
              If strDIS = vbnullstring Then
                objLog.writeline "Disabled : " & Time() & " : " & objAD.distinguishedname & " : By Match : " & strPROP & vbnewline
              Else
                'objDISch.movehere "LDAP://" & objAD.distinguishedname, vbnullstring
                objLog.writeline "Disabled : " & Time() & " : " & objAD.distinguishedname & " : By Match : " & strPROP & " : OU Moved Under : " & _
                  "ou=DISABLED_" & UCase(strSEL) & "_" & Year(Now) & Month(Now) & Day(Now) & "," & strDIS & "," & strDN & vbnewline
              End If
            ElseIf strACT = "edit" Then
            ElseIf strACT = "move" Then
              'objMOV.movehere "LDAP://" & objAD.distinguishedname, vbnullstring
              objLog.writeline "Disabled : " & Time() & " : " & objAD.distinguishedname & " : By Match : " & strPROP & " : OU Moved Under : " & strMOV & _
                "," & strDIS & "," & strDN & vbnewline
            End If
          End If
          WScript.Sleep 200

        Case "organizationalunit"

          WScript.echo vbnewline & "*****************************************************************"
          WScript.echo "SCANNING OU : " & Time() & " : " & objAD.distinguishedname
          WScript.echo "*****************************************************************"
          wscript.sleep 500
          Call Parser(objAD)

      End Select
    end if
  Next

End Sub

Sub scrClean()
  objLog.close
  Set objLog = Nothing
  Set objMOV = Nothing
  Set objOU = Nothing
  Set objAD = Nothing
  Set objDISch = Nothing
  Set objDIS = Nothing
  Set objFSO = Nothing
  Set objDSP = Nothing
End Sub

'V--V      AD-Parse Written By Khristos       V--V'
'V--V         VampyreMan Studios Inc          V--V'
'V--V http://vampyremanstudios.com/khris.aspx V--V'
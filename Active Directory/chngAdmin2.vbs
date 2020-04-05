	''V--V CHNGADMIN.VBS by: Khristos V--V''
	''V--V VampyreMan Studios Inc V--V''
	''V--V www.vampyremanstudios.com V--V'
	''VERY SIMPLE SCRIPT TO USE ON A DOMAIN''
	''DISABLES UNNECCESSARY ACCOUNTS AND THEN''
	''CREATES USER 'AdminTQ' ON LOCAL MACHINE''
	''SETS A PREDEFINED PASSWORD FOR ACCOUNT''
	''ADDS 'AdminTQ' TO LOCAL ADMINS GROUP''

On Error Resume Next
Dim objNetwork, colAccounts, objUser, objGroup, strComputer, blnCreate

Set objNetwork = CreateObject("WScript.Network")
strComputer = objNetwork.ComputerName
wscript.echo "Enumerating and Evaluating Accounts on " & strComputer

blnCreate = "TRUE"
Set colAccounts = GetObject("WinNT://" & strComputer & "")
For Each enumUser In colAccounts
	select case enumUser.class
		case "User"
			'msgbox enumUser.name
			if enumUser.name = "TQAdministratorx" then
				blnCreate = "FALSE"
			end if
			if enumUser.name = "Administrator" then
				wscript.echo enumUser.name & " DISABLED"
				enumUser.accountdisabled = true
				enumUser.setinfo
			elseif enumUser.name = "AdminTQ" then
				wscript.echo enumUser.name & " DISABLED"
				enumUser.accountdisabled = true
				enumUser.setinfo
			elseif enumUser.name = "Guest" then
				wscript.echo enumUser.name & " DISABLED"
				enumUser.accountdisabled = true
				enumUser.setinfo
			elseif enumUser.name = "usmc" then
				wscript.echo enumUser.name & " DISABLED"
				enumUser.accountdisabled = true
				enumUser.setinfo
			elseif enumUser.name = "ASPNET" then
				wscript.echo enumUser.name & " DISABLED"
				enumUser.accountdisabled = true
				enumUser.setinfo
			elseif instr(1, enumUser.name, "SUPPORT") then
				wscript.echo enumUser.name & " DISABLED"
				enumUser.accountdisabled = true
				enumUser.setinfo
			end if
			wscript.sleep 1000
	end select
Next

if blnCreate = "TRUE" then
	wscript.echo "Creating 'TQAdministratorx' Account"
	Set objUser = colAccounts.Create("user", "TQAdministratorx")
	objUser.SetPassword "N0$tup!dU$erN0$tup!dU$er"
	objUser.SetInfo
elseif blnCreate = "FALSE" then
	set objUser = GetObject("WinNT://" & strComputer & "/TQAdministratorx, user")
	objUser.SetPassword "N0$tup!dU$erN0$tup!dU$er"
	objUser.SetInfo
end if
wscript.sleep 1000

wscript.echo "Setting TQAdministratorx Permissions and Password"
Set objGroup = GetObject("WinNT://" & strComputer & "/Administrators, group")
objGroup.Add(objUser.ADsPath)
wscript.sleep 1500

set enumUser = nothing
Set objUser = Nothing
Set objNetwork = Nothing
wscript.echo "TQAdministratorx Created, Permissions Applied"
wscript.sleep 1500
WScript.Quit
	''V--V CHNGADMIN.VBS by: Khristos V--V''
	''V--V VampyreMan Studios Inc V--V''
	''V--V www.vampyremanstudios.com V--V'
	''VERY SIMPLE SCRIPT TO USE ON A DOMAIN''
	''DISABLES UNNECCESSARY ACCOUNTS AND THEN''
	''CREATES USER 'AdminTQ' ON LOCAL MACHINE''
	''SETS A PREDEFINED PASSWORD FOR ACCOUNT''
	''ADDS 'AdminTQ' TO LOCAL ADMINS GROUP''

on error resume next

Dim objNetwork, colAccounts, objUser, objGroup, strComputer, blnCreate

Set objNetwork = CreateObject("WScript.Network")
strComputer = objNetwork.ComputerName

blnCreate = "TRUE"
Set colAccounts = GetObject("WinNT://" & strComputer & "")
For Each enumUser In colAccounts
	select case enumUser.class
		case "User"
			if enumUser.name = "AdminTQ" then
				blnCreate = "FALSE"
			end if
			if enumUser.name = "Administrator" then
				enumUser.accountdisabled = true
				enumUser.setinfo
			elseif enumUser.name = "Guest" then
				enumUser.accountdisabled = true
				enumUser.setinfo
			elseif enumUser.name = "usmc" then
				enumUser.accountdisabled = true
				enumUser.setinfo
			elseif enumUser.name = "ASPNET" then
				enumUser.accountdisabled = true
				enumUser.setinfo
			end if
	end select
Next

if blnCreate = "TRUE" then
	Set objUser = colAccounts.Create("user", "AdminTQ")
	objUser.SetPassword "$@dm1n$ecr3t$"
	objUser.SetInfo
elseif blnCreate = "FALSE" then
	set objUser = GetObject("WinNT://" & strComputer & "/AdminTQ, user")
end if

Set objGroup = GetObject("WinNT://" & strComputer & "/Administrators,group")
objGroup.Add(objUser.ADsPath)

set enumUser = nothing
Set objUser = Nothing
Set objNetwork = Nothing
WScript.Quit
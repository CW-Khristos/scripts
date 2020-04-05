		'V--V   ADUAparse.vbs    V--V'
		'V--V VampyreMan Studios V--V'
		'V--V  Author : Khristos V--V'
		'Simplistic script to enumerate
		'All accounts on a local machine
		'Disables accounts other than Admin
		'And one specified user account
		'Modifiable for AD Integration

	''DEFINES CONSTANTS EQUAL TO WINDOWS HEX CODES FOR USER FLAGS''
Const ADS_UF_SCRIPT = &H0001 
Const ADS_UF_ACCOUNTDISABLE = &H0002 
Const ADS_UF_HOMEDIR_REQUIRED = &H0008 
Const ADS_UF_LOCKOUT = &H0010 
Const ADS_UF_PASSWD_NOTREQD = &H0020 
Const ADS_UF_PASSWD_CANT_CHANGE = &H0040 
Const ADS_UF_ENCRYPTED_TEXT_PASSWORD_ALLOWED = &H0080 
Const ADS_UF_DONT_EXPIRE_PASSWD = &H10000 
Const ADS_UF_SMARTCARD_REQUIRED = &H40000 
Const ADS_UF_PASSWORD_EXPIRED = &H800000 

	''CREATE VBS NETWORK OBJECT TO ALLOW CONNECTION TO EXCHANGE SERVER''
set objNetwork = createobject("wscript.network")

	''SET 'strComputer' EQUAL TO COMPUTER NAME OF CONNECTED SERVER''
	''THIS WILL BE MODIFIED FOR AD INTEGRATION AND LDAP COMMUNICATION''
strComputer = objNetwork.computername

	''SET 'colAccounts' EQUAL TO THE COLLECTION OF ALL INFO OF COMPUTER''
	''THIS WILL ALSO BE MODIFIED FOR LDAP COMMUNICATION AND AD DATABASE''
set colAccounts = getobject("WinNT://" & strComputer & "")

	''FILTER OUT ALL OBJECTS OTHER THAN "USER" OBJECTS''
colAccounts.filter = array("user")

	''FOR EVERY OBJECT STILL IN 'colAccounts' COLLECTION''
for each objUser in colAccounts

	''CALL 'enumAU' SUB ROUTINE PASSING IT THE VALUE OF THE ''
	''PROPERTY "NAME" FOR THE CURRENT USER OBJECT''
  enumAU (objUser.name)

	''REPEAT PROCESS FOR NEXT "USER" OBJECT''
next

	''SET ALL OBJECTS EQUAL TO NOTHING (CLEAN UP) AND QUIT SCRIPT''
set objUser = nothing
set colAccounts = nothing
set strComputer = nothing
set objNetwork = nothing
wscript.quit

		''ENUMAU SUB ROUTINE''
sub enumAU (user)

  'wscript.echo user

		''DETERMINE WHETHER THE USER IS NOT ADMIN OR KHRISTOS''
		''THIS WILL BE REPLACED WITH LDAP PROPERTY CONDITIONS''
  if user <> "Administrator" then
    if user <>"Khristos" then

		''IF USER IS NOT ADMIN OR KHRISTOS, GET THE USER FLAGS FOR THAT ACCOUNT''
      objUF = objUser.get("userflags")
      'wscript.echo objUF

		''IF USER FLAGS MATCH WINDOWS HEX CODE FOR A DISABLED ACCOUNT''
		''WE DON'T NEED TO WORRY WITH THIS ACCOUNT. CONTINUE TO NEXT ACCOUNT''
      if objUF AND ADS_UF_ACCOUNTDISABLE then
        'wscript.echo "Account is disabled."

      else

		''IF USER FLAGS DON'T MATCH WINDOWS HEX CODE FOR A DISABLED ACCOUNT''
		''WE WANT TO CHECK LAST LOGON, AND DETERMINE IF WE SHOULD DISABLE THE ACCOUNT''
		''LDAP QUERY SCRIPT AND COMPARISON SCRIPT SHOULD GO HERE''
    	'wscript.echo "Account is not disabled."

		''SET 'objAccountDIS EQUAL TO COMPARISON 'objUF' OR 'ADS_UF_ACCOUNTDISABLE'''
	objAccountDIS = objUF or ADS_UF_ACCOUNTDISABLE

		''ADD 'objAccountDIS' TO WINDOWS USER FLAGS FOR THAT ACCOUNT''
	objUser.put "userflags", objAccountDIS

		''COMMIT THE CHANGE TO THE USER ACCOUNT''
	objUser.setinfo
	'wscript.echo "Disabled " & objUser.name & vbnewline & objUF

      end if
    end if
  end if

end sub
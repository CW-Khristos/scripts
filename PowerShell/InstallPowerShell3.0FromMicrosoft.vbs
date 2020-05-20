'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Description:Script for Installing PowerShellVersion 2.0
'Author:  Agent (N.T)
'Company: N-Able Technologies
'Arguments: None
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''Initialize Variables Begin'''''''''''''''''''''''''''''''

strComputer = "."
version=""
url=""
productType=""
arch=""
installerName="PowerShellInstaller.exe"
SecsSince = CLng(DateDiff("s", "01/01/1970 00:00:00", Now))

logDir=WScript.CreateObject("WScript.Shell").ExpandEnvironmentStrings("%temp%")
logFile=logDir & "\PSInstallerN-able" & SecsSince & ".log"
logFileWithQuotes=Chr(34) & logFile & Chr(34)

''''''''''''Initialize Variables End'''''''''''''''''''''''''''''''


On Error Resume Next
' Query WMI
Set objWMIService = GetObject( "winmgmts://" & strComputer & "/root/cimv2" )
' Get the error number 
If Err.Number Then
	strMsg = vbCrLf & strComputer & vbCrLf & _
	         "Error # " & Err.Number & vbCrLf & _
	         Err.Description & vbCrLf & vbCrLf
	ErrorQuit
End If

' Collect OS information 
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_OperatingSystem",,48)
For Each objItem in colItems
    strMsg = "Computer Name   : " & objItem.CSName & vbCrLf & _
             "Windows Version : " & objItem.Version & vbCrLf & _
             "ServicePack     : " & objItem.CSDVersion & vbCrLf & _
             "Product Type    : " & objItem.ProductType & vbCrLf & _
             "OSArchitecture  : " & objItem.OSArchitecture & vbCrLf 

' Get the first two digits from the version string
versionArray = Split(objItem.Version, ".", -1, 1)
version=versionArray(0) & "." & versionArray(1)

productType=objItem.ProductType
arch=objItem.OSArchitecture
Next

' Cleanup
Set objWMIService= nothing
Set colItems = nothing

' Display the results
WScript.StdOut.WriteLine(strMsg)

intArch = StrComp(arch,"64-bit",1)
'On some OS OSArchitecture is not available so check the registry
If (intArch = -1) Then
Is64=""
Determine64BitOS Is64
'WScript.StdOut.WriteLine(Is64)
intArch = StrComp(Is64,"True",1)

End If

'WScript.StdOut.Write("Version:")
'WScript.StdOut.WriteLine(version)

'Check what Power shell version is installed if any
psver=""
IsPowerShellInstalled psver
WScript.StdOut.WriteLine "Current PowerShell Version from Registry: " & psver
If ( StrComp(psver,"3.0",1) =0 ) Then
	WScript.StdOut.WriteLine("PowerShell 3.0 already installed.")
	WScript.Quit(0)

End If

' Select the installer based on the OS
Select Case version
Case "6.1"
WScript.StdOut.WriteLine("OS:Windows 7 / Server 2008R2")
    If(intArch=0) then
     url ="https://raw.githubusercontent.com/CW-Khristos/scripts/master/PowerShell/Windows6.1-KB2506143-x64.msu"
    Else 
     url="https://raw.githubusercontent.com/CW-Khristos/scripts/master/PowerShell/Windows6.1-KB2506143-x86.msu"
    End if 

Case "6.0"

If ( StrComp(productType,"1",1) =0 ) Then
    WScript.StdOut.WriteLine("OS:Vista")

    If(intArch=0) then
     url ="https://raw.githubusercontent.com/CW-Khristos/scripts/master/PowerShell/Windows6.0-KB2506146-x64.msu"
    Else 
     url="https://raw.githubusercontent.com/CW-Khristos/scripts/master/PowerShell/Windows6.0-KB2506146-x86.msu"
    End if 

Else
    WScript.StdOut.WriteLine("OS:Server 2008")
    If(intArch=0) then
     url="https://raw.githubusercontent.com/CW-Khristos/scripts/master/PowerShell/Windows6.0-KB2506146-x64.msu"
    Else 
     url="https://raw.githubusercontent.com/CW-Khristos/scripts/master/PowerShell/Windows6.0-KB2506146-x86.msu"
    End if 

End If

End Select

'''''''''''''''Download/Installation Section Begin'''''''''''''
WScript.StdOut.Write(vbCrLf & "Download URL:")
WScript.StdOut.WriteLine(url)

' Get the installer name from the URL
urlArray = Split(url, "/", -1, 1)
lengthUrlArray=UBound(urlArray)
installerName=urlArray(lengthUrlArray)
WScript.StdOut.Write(vbCrLf & "InstallerName:" & installerName & vbCrLf)

' Get the extension of the installer exe or msu
extn=Right(installerName,4)
' Check if  Windows update service is running in case of  an Update
If ( StrComp(extn,".msu",1) =0 ) Then
	strWMIQuery = "Select * from Win32_Service Where Name = 'wuauserv' and state='Running'"
	Set objWMIService = GetObject("winmgmts:" _
		& "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
		If objWMIService.ExecQuery(strWMIQuery).Count > 0 Then
			WScript.StdOut.WriteLine("Windows Update Service (wuauserv) is running.")
		Else
			WScript.StdOut.WriteLine "-------------------------------------------------------------"
			WScript.StdOut.WriteLine("PowerShell 3.0 couldn't be installed because Windows Update Service (wuauserv) is not running.")
			WScript.StdOut.WriteLine "-------------------------------------------------------------"
			WScript.Quit(1)	
		End If
End If
'Download the installer
Download
'Check download
Set fso = CreateObject("Scripting.FileSystemObject")
If (fso.FileExists(installerName) ) Then
    WScript.StdOut.Write("Download was successful.")
    Set fso = Nothing
Else
    Set fso = Nothing
    WScript.StdOut.WriteLine("Download failed.")
    ErrorQuit
End if
WScript.StdOut.Write(vbCrLf & "Starting installation - Logfile is available at:")
' Start installation
If ( StrComp(extn,".exe",1) =0 ) Then
		WScript.StdOut.WriteLine(logFile)
		InstallExe
Else
		WScript.StdOut.WriteLine("C:\Windows\WindowsUpdate.log")
		InstallMsu
		'Check if reboot is needed
		Set objSysInfo = CreateObject("Microsoft.Update.SystemInfo")
		WScript.StdOut.WriteLine "-----------------------------------------------------------------"
		WScript.StdOut.WriteLine "Reboot required? " & objSysInfo.RebootRequired
		If ( StrComp(objSysInfo.RebootRequired,"True",1) =0 ) Then
			WScript.StdOut.WriteLine "Machine needs to be rebooted before PowerShell 3.0 can be used."
			WScript.Quit(0)
		End If
		WScript.StdOut.WriteLine "-----------------------------------------------------------------"
		objSysInfo =Nothing
End If

'Read the Logfile
ReadLogFile

'CheckPowerShell
psver=""
IsPowerShellInstalled psver
WScript.StdOut.WriteLine "Current PowerShell Version: " & psver
If ( StrComp(psver,"3.0",1) =0 ) Then
   WScript.StdOut.WriteLine("PowerShell 3.0 successfully installed.")
   WScript.Quit(0)
Else 
   WScript.StdOut.WriteLine("PowerShell 3.0 failed to install.")
   WScript.Quit(1)
End If

'''''''''''''''Download/Installation Section End'''''''''''''

' Done
WScript.Quit(0)

'''''''''''''' Function Calls Section Begin''''''''''''''



Sub Determine64BitOS(ByRef Is64BitOs) 
 
 set Shell = CreateObject("WScript.Shell") 
 on error resume next 
 Shell.RegRead "HKLM\Software\Microsoft\Windows\CurrentVersion\ProgramFilesDir (x86)" 
 If Err.Number <> 0 Then 
 	WScript.StdOut.WriteLine("Not 64bitOs:" & Err.Description)  
 	Is64BitOs= "False"
 else	
	WScript.StdOut.WriteLine("64bitOs") 
	Is64BitOs="True"
 End If     
 Set Shell = Nothing
End Sub

Sub Download
    WScript.StdOut.WriteLine("")
    sLocation= installerName
    sFileURL=url
    WScript.StdOut.WriteLine("Starting download from:"+ url)
	Set objXMLHTTP = CreateObject("MSXML2.XMLHTTP")
	objXMLHTTP.open "GET", sFileURL, false
	objXMLHTTP.send()
	do until objXMLHTTP.Status = 200 :  wscript.sleep(1000) :  loop
	If objXMLHTTP.Status = 200 Then
		Set objADOStream = CreateObject("ADODB.Stream")
		objADOStream.Open
        'adTypeBinary
		objADOStream.Type = 1
		objADOStream.Write objXMLHTTP.ResponseBody
		objADOStream.Position = 0    
        'check for an existing file
        Set objFSO = Createobject("Scripting.FileSystemObject")
        'check if the file exists, if it exists then delete it
		If objFSO.Fileexists(sLocation) Then objFSO.DeleteFile sLocation
		Set objFSO = Nothing
		objADOStream.SaveToFile sLocation
		objADOStream.Close
		Set objADOStream = Nothing
	End if
	'cleanup
	Set objXMLHTTP = Nothing
    WScript.StdOut.WriteLine("Download complete.")
End Sub


Sub InstallExe
    ' /quiet /norestart /log:<fullpath>
    cmdLine =  installerName & " /quiet /norestart /log:" & logFileWithQuotes
    WScript.StdOut.WriteLine(cmdLine)
    Set WshShell = WScript.CreateObject("WScript.Shell")
    Return = WshShell.Run(cmdLine, 0, true)

    ' cleanup
    Set WshShell= nothing
End Sub


Sub InstallMsu
    ' /quiet /norestart /log:<fullpath>
    sCurPath = CreateObject("Scripting.FileSystemObject").GetAbsolutePathName(".")
    installerPath=  Chr(34) & sCurPath & "\" & installerName & Chr(34)

    cmdLine ="wusa.exe " & installerPath & " /quiet /norestart"
    WScript.StdOut.WriteLine(cmdLine)
    Set WshShell = WScript.CreateObject("WScript.Shell")
    Return=WshShell.Run(cmdLine,0, true)
    ' cleanup
    Set WshShell= nothing
End Sub


Sub ReadLogFile
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    If objFSO.Fileexists(logFile) Then 
        WScript.StdOut.WriteLine("Reading logfile:" & logFile)
        Set objReadFile = objFSO.OpenTextFile(logFile,1,False)
        'Read file contents
        contents = objReadFile.ReadAll
        WScript.StdOut.WriteLine(contents)
        'Close file
        objReadFile.close
    Else
        WScript.StdOut.WriteLine("No log file available.")
    End If

    ' cleanup
    Set objFSO= nothing
    Set objReadFile = nothing
End Sub


Sub IsPowerShellInstalled(ByRef psver)

    'WScript.StdOut.WriteLine("Getting powershell version from the registry.")

    const HKEY_LOCAL_MACHINE = &H80000002
     
    Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" &_ 
    strComputer & "\root\default:StdRegProv")
     
    strKeyPath = "SOFTWARE\Microsoft\PowerShell\1\PowerShellEngine"

    strValueName = "PowerShellVersion"
    oReg.GetStringValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue

    psver=strValue
    'cleanup
    oReg= Nothing

End Sub

Sub ErrorQuit
    Wscript.Quit(1)
End Sub


'''''''''''''' Function Calls Section End''''''''''''''
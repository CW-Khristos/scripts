<# 
.SYNOPSIS 
    AV Health Monitoring
    This was based on "Get Installed Antivirus Information" by SyncroMSP
    But omits the Hex Conversions and utilization of WSC_SECURITY_PROVIDER , WSC_SECURITY_PRODUCT_STATE , WSC_SECURITY_SIGNATURE_STATUS
    https://mspscripts.com/get-installed-antivirus-information-2/

.DESCRIPTION 
    Provide Primary AV Product Status and Report Possible AV Conflicts
    Script is intended to be universal / as flexible as possible without being excessively complicated
    Script is intended to replace 'AV Status' VBS Monitoring Script
 
.NOTES
    Version        : 0.1.4 (21 December 2021)
    Creation Date  : 14 December 2021
    Purpose/Change : Provide Primary AV Product Status and Report Possible AV Conflicts
    File Name      : AVHealth_0.1.4.ps1 
    Author         : Christopher Bledsoe - cbledsoe@ipmcomputers.com 
    Requires       : PowerShell Version 2.0+ installed

.CHANGELOG
    0.1.0 Initial Release
    0.1.1 Switched to use of '-match' and 'notmatch' for accepting input of vendor / general AV name like 'Sophos'
          Switched to use and expanded AV Product 'Definition' XMLs to be vendor specific instead of product specific
    0.1.2 Optimized to reduced use of 'If' blocks for querying registry values
          Added support for monitoring on Servers using 'HKLM:\SOFTWARE\Microsoft\Security Center\Monitoring\' since WMI SecurityCenter2 Namespace does not exist on Server OSes
          Note : Obtaining AV Products from 'HKLM:\SOFTWARE\Microsoft\Security Center\Monitoring\' only works *if* the AV Product registers itself in that key!
            If the above registry check fails to find any registered AV Products; script will attempt to fallback to WMI "root\cimv2" Namespace and "Win32_Product" Class -filter "Name like '$i_PAV'"
    0.1.3 Correcting some bugs and adding better error handling
    0.1.4 
#> 

#REGION ----- DECLARATIONS ----
$global:avkey = @{}
$global:bitarch = ""
$global:o_AVname = ""
$global:o_AVVersion = ""
$global:o_AVpath = ""
$global:o_AVStatus = ""
$global:rtstatus = "Unknown"
$global:o_RTstate = "Unknown"
$global:defstatus = "Unknown"
$global:o_DefStatus = "Unknown"
$global:o_AVcon = 0
$global:o_CompAV = " "
$global:o_CompPath = " "
$global:o_Compstate = " "
$global:blnWMI = $true
$global:blnSecMon = $true
#ENDREGION ----- DECLARATIONS ----

#REGION ----- FUNCTIONS ----
#Determine OS Bit Architecture
function Get-OSArch {
  $osarch = (get-wmiobject win32_operatingsystem).osarchitecture
  if ($osarch -like '*64*') {
    $global:bitarch = "bit64"
  } elseif ($osarch -like '*32*') {
    $global:bitarch = "bit32"
  }
} ## Get-OSArch

#Determine Antivirus State
function Get-AVState {
  param (
    $state
  )
  #Switch to determine the status of antivirus definitions and real-time protection.
  #THIS COULD PROBABLY ALSO BE TURNED INTO A SIMPLE XML / JSON LOOKUP TO FACILITATE COMMUNITY CONTRIBUTION
  switch ($state) {
    #AVG IS 2012 AV / CrowdStrike / Kaspersky
    "262144" {$global:defstatus = "Up to date" ;$global:rtstatus = "Disabled"}
    "266240" {$global:defstatus = "Up to date" ;$global:rtstatus = "Enabled"}
    #AVG IS 2012 FW
    "266256" {$global:defstatus = "Out of date" ;$global:rtstatus = "Enabled"}
    "262160" {$global:defstatus = "Out of date" ;$global:rtstatus = "Disabled"}
    #MSSE
    "393216" {$global:defstatus = "Up to date" ;$global:rtstatus = "Disabled"}
    "397312" {$global:defstatus = "Up to date" ;$global:rtstatus = "Enabled"}
    #Windows Defender
    "393472" {$global:defstatus = "Up to date" ;$global:rtstatus = "Disabled"}
    "397584" {$global:defstatus = "Out of date" ;$global:rtstatus = "Enabled"}
    "397568" {$global:defstatus = "Up to date" ;$global:rtstatus = "Enabled"}
    "401664" {$global:defstatus = "Up to date" ;$global:rtstatus = "Disabled"}
    #
    "393232" {$global:defstatus = "Out of date" ;$global:rtstatus = "Disabled"}
    "393488" {$global:defstatus = "Out of date" ;$global:rtstatus = "Disabled"}
    "397328" {$global:defstatus = "Out of date" ;$global:rtstatus = "Enabled"}
    #Sophos
    "331776" {$global:defstatus = "Up to date" ;$global:rtstatus = "Enabled"}
    #Norton Security
    "327696" {$global:defstatus = "Out of date" ;$global:rtstatus = "Disabled"}
    default {$global:defstatus = "Unknown" ;$global:rtstatus = "Unknown"}
  }
} ## Get-AVState
#ENDREGION ----- FUNCTIONS ----

#------------
#BEGIN SCRIPT
$i = 0
Get-OSArch
#COMMENT OUT THE BELOW LINE (LN101) FOR USE WITH AMP / PASSING OF PRIMARY AV AS INPUT
#$i_PAV = "Sophos"
$computername=$env:computername
[system.Version]$OSVersion = (get-wmiobject win32_operatingsystem -computername $computername).version
$srcAVP = "https://raw.githubusercontent.com/CW-Khristos/scripts/dev/AVProducts/" + $i_PAV.replace(" ", "").replace("-", "").tolower() + ".xml"
#READ AV PRODUCT DETAILS FROM XML
try {
  $avXML = New-Object System.Xml.XmlDocument
  $avXML.Load($srcAVP)
} catch {
  $web = new-object system.net.webclient
  [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls12
  [xml]$avXML = $web.DownloadString($srcAVP)
}
foreach ($itm in $avXML.NODE.ChildNodes) {
  $hash = @{
    display = $itm.$global:bitarch.display
    displayval = $itm.$global:bitarch.displayval
    path = $itm.$global:bitarch.path
    pathval = $itm.$global:bitarch.pathval
    ver = $itm.$global:bitarch.ver
    verval = $itm.$global:bitarch.verval
    stat = $itm.$global:bitarch.stat
    statval = $itm.$global:bitarch.statval
  }
  $avkey.add($itm.name,$hash)
}
#QUERY WMI SECURITYCENTER NAMESPACE FOR AV PRODUCT DETAILS
if ($OSVersion -ge [system.version]'6.0.0.0') {
  write-verbose "OS Windows Vista/Server 2008 or newer detected."
  try {
    $AntiVirusProduct = get-wmiobject -Namespace "root\SecurityCenter2" -Class "AntiVirusProduct" -ComputerName $computername -ErrorAction Stop
  } catch {
    $blnWMI = $false
  }
} elseif ($OSVersion -lt [system.version]'6.0.0.0') {
  write-verbose "Windows 2000, 2003, XP detected" 
  try {
    $AntiVirusProduct = get-wmiobject -Namespace "root\SecurityCenter" -Class "AntiVirusProduct"  -ComputerName $computername -ErrorAction Stop
  } catch {
    $blnWMI = $false
  }
}
if (-not $blnWMI) {                               #FAILED TO RETURN WMI SECURITYCENTER NAMESPACE ; POSSIBLY A SERVER
  try {
    write-host "Failed to query WMI SecurityCenter Namespace" -foregroundcolor Red
    write-host "Possibly Server, attempting to  fallback to using 'HKLM:\SOFTWARE\Microsoft\Security Center\Monitoring\' registry key" -foregroundcolor Red
    try {                                         #QUERY 'HKLM:\SOFTWARE\Microsoft\Security Center\Monitoring\' AND SEE IF AN AV IS REGISTRERED THERE
      if ($global:bitarch = "bit64") {
        $AntiVirusProduct = (get-itemproperty -path "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Security Center\Monitoring\*" -ErrorAction Stop).PSChildName
      } elseif ($global:bitarch = "bit32") {
        $AntiVirusProduct = (get-itemproperty -path "HKLM:\SOFTWARE\Microsoft\Security Center\Monitoring\*" -ErrorAction Stop).PSChildName
      }
    } catch {
      write-host "Could not find AV registered in HKLM:\SOFTWARE\Microsoft\Security Center\Monitoring\*" -foregroundcolor Red
      $AntiVirusProduct = $null
    }
    if ($AntiVirusProduct -ne $null) {            #RETURNED 'HKLM:\SOFTWARE\Microsoft\Security Center\Monitoring\' DATA
      foreach ($av in $AntiVirusProduct) {
        write-host "Found 'HKLM:\SOFTWARE\Microsoft\Security Center\Monitoring\$av'" -foregroundcolor Yellow
        $string1 = $string1 + $av + ", "
        $string2 = $string2 + "Unknown, "
        $string3 = $string3 + "Unknown, "
      }
      $avs = $string1 -split ", "
      $avpath = $string2 -split ", "
      $avstat = $string3 -split ", "
      $global:blnSecMon = $true
    } elseif ($AntiVirusProduct -eq $null) {      #FAILED TO RETURN 'HKLM:\SOFTWARE\Microsoft\Security Center\Monitoring\' DATA
      foreach ($key in $global:avkey.keys) {      #ATTEMPT TO VALIDATE EACH AV PRODUCT CONTAINED IN VENDOR XML
        $reg1 = $global:avkey[$key].display
        $reg2 = $global:avkey[$key].displayval
        $reg3 = $global:avkey[$key].path
        $reg4 = $global:avkey[$key].pathval
        $reg5 = $global:avkey[$key].stat
        $reg6 = $global:avkey[$key].statval
        try {
          if (test-path "HKLM:$reg1") {           #VALIDATE INSTALLED AV PRODUCT BY TESTING READING A KEY
            write-host "Found 'HKLM:$reg1' for product : $key" -foregroundcolor Yellow
            try {                                 #IF VALIDATIO PASSES; FABRICATE 'HKLM:\SOFTWARE\Microsoft\Security Center\Monitoring\' DATA
              $keyval1 = get-itemproperty -path "HKLM:$reg1" -name "$reg2" -erroraction stop
              $keyval2 = get-itemproperty -path "HKLM:$reg3" -name "$reg4" -erroraction stop
              $keyval3 = get-itemproperty -path "HKLM:$reg5" -name "$reg6" -erroraction stop
              write-host "Creating Registry Key HKLM:\SOFTWARE\Microsoft\Security Center\Monitoring\" $keyval1.$reg2 " for : " $keyval1.$reg2 -foregroundcolor Red
              if ($global:bitarch = "bit64") {
                new-item -path "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Security Center\Monitoring\" -name $keyval1.$reg2 -value $keyval1.$reg2 –force
              } elseif ($global:bitarch = "bit32") {
                new-item -path "HKLM:\SOFTWARE\Microsoft\Security Center\Monitoring\" -name $keyval1.$reg2 -value $keyval1.$reg2 –force
              }
              $string1 = $string1 + $keyval1.$reg2 + ", "
              $string2 = $string2 + $keyval2.$reg4 + ", "
              $string3 = $string3 + $keyval3.$reg6.tostring() + ", "
              $avs = $string1 -split ", "
              $avpath = $string2 -split ", "
              $avstat = $string3 -split ", "
              $global:blnSecMon = $false
              $AntiVirusProduct = "."
            } catch {
              write-host "Could not create Registry Key `HKLM:\SOFTWARE\Microsoft\Security Center\Monitoring\" $keyval1.$reg2 " for : " $keyval1.$reg2 -foregroundcolor Red
              $AntiVirusProduct = $null
            }
            break
          }
        } catch {
          write-host "Not Found 'HKLM:$reg1' for product : $key" -foregroundcolor Red
        }
      }
    }
  } catch {
    $global:blnSecMon = $false
  }
} elseif ($blnWMI) {                              #RETURNED WMI SECURITYCENTER NAMESPACE
  #SEPARATE RETURNED WMI AV PRODUCT INSTANCES
  $string = $AntiVirusProduct.displayName
  $avs = $string -split ", "
  $string = $AntiVirusProduct.pathToSignedProductExe
  $avpath = $string -split ", "
  $string = $AntiVirusProduct.productState
  $avstat = $string -split ", "
}
if ($AntiVirusProduct -eq $null) {                #NO AV PRODUCT FOUND
  write-host "Could not find any AV Product registered" -foregroundcolor Red
  $global:o_AVname = "No AV Product Found"
  $global:o_AVVersion = ""
  $global:o_AVpath = ""
  $global:o_AVStatus = "Unknown"
  $global:o_RTstate = "Unknown"
  $global:o_DefStatus = "Unknown"
  $global:o_AVcon = 0
} elseif ($AntiVirusProduct -ne $null) {          #FOUND AV PRODUCTS
  foreach ($av in $avs) {                         #ITERATE THROUGH EACH FOUND AV PRODUCT
    if (($av -ne $null) -And ($av -ne "")) {
      #NEITHER PRIMARY AV PRODUCT NOR WINDOWS DEFENDER
      if (($avs[$i] -notmatch $i_PAV) -And ($avs[$i] -notmatch "Windows Defender")) {
        $global:o_AVcon = 1
        $global:o_CompAV = $global:o_CompAV + $avs[$i] + " , "
        $global:o_CompPath = $global:o_CompPath + $avpath[$i] + " , "
        $global:o_Compstate = $global:o_Compstate + $avstat[$i] + " , "
      #PRIMARY AV PRODUCT
      } elseif ($avs[$i] -match $i_PAV) {
        $global:o_AVname = $avs[$i]
        $global:o_AVpath = $avpath[$i]
        #will still return if it is unknown, etc. if it is unknown look at the code it returns, then look up the status and add it above
        Get-AVState($avstat[$i])
        $global:o_DefStatus = $global:defstatus
        $global:o_RTstate = $global:rtstatus
        
        #PARSE XML FOR SPECIFIC VENDOR AV PRODUCT
        $node = $avs[$i].replace(" ", "").replace("-", "").toupper()
        #AV PRODUCT VERSION KEY PATH AND VALUE
        $i_verkey = $avXML.NODE.$node.$global:bitarch.ver
        $i_verval = $avXML.NODE.$node.$global:bitarch.verval
        #AV PRODUCT STATE KEY PATH AND VALUE
        $i_statkey = $avXML.NODE.$node.$global:bitarch.stat
        $i_statval = $avXML.NODE.$node.$global:bitarch.statval
        
        #GET PRIMARY AV PRODUCT VERSION VIA REGISTRY
        try {
          $global:o_AVVersion = get-itemproperty -path "HKLM:$i_verkey" -name "$i_verval" -ErrorAction Stop
          write-host "-path HKLM:$i_verkey -name $i_verval" -foregroundcolor Yellow
        } catch {
          $global:o_AVVersion | add-member -NotePropertyName "$i_verval" -NotePropertyValue "."
        }
        $global:o_AVVersion = $global:o_AVVersion.$i_verval
        #GET PRIMARY AV PRODUCT STATUS VIA REGISTRY
        try {
          $global:o_AVStatus = get-itemproperty -path "HKLM:$i_statkey" -name "$i_statval" -ErrorAction Stop
          write-host "-path HKLM:$i_statkey -name $i_statval" -foregroundcolor Yellow
        } catch {
          $global:o_AVStatus | add-member -NotePropertyName "$i_statval" -NotePropertyValue "0"
        }
        #INTERPRET 'AVSTATUS' BASED ON ANY AV PRODUCT VALUE REPRESENTATION - SOME TREAT '0' AS 'UPTODATE' SOME TREAT '1' AS 'UPTODATE'
        #$global:o_AVStatus.$i_statval
        if (($avs[$i] -match "Symantec") -or ($avs[$i] -match "Sophos Intercept X")) {
          if ($global:o_AVStatus.$i_statval -eq "0") {
            $global:o_AVStatus = $true
          } else {
            $global:o_AVStatus = $false
          }
        } elseif ($avs[$i] -notmatch "Symantec") {
          if ($global:o_AVStatus.$i_statval -eq "1") {
            $global:o_AVStatus = $true
          } else {
            $global:o_AVStatus = $false
          }
        }
      #SAVE WINDOWS DEFENDER FOR LAST - TO PREVENT SCRIPT CONSIDERING IT 'COMPETITOR AV' WHEN SET AS PRIMARY AV
      } elseif ($avs[$i] -eq "Windows Defender") {
        $global:o_CompAV = $global:o_CompAV + $avs[$i] + " , "
        $global:o_CompPath = $global:o_CompPath + $avpath[$i] + " , "
        $global:o_Compstate = $global:o_Compstate + $avstat[$i] + " , "  
      }
    }
    $i = $i + 1
  }
}
#OUTPUT
if ($global:o_AVname -ne "No AV Product Found") {
  $ccode = "Green"
} elseif ($global:o_AVname -eq "No AV Product Found") {
  $ccode = "Red"
}
write-host "AV Display Name :" $global:o_AVname -foregroundcolor $ccode
write-host "AV Version : " $global:o_AVVersion -foregroundcolor $ccode
write-host "AV Path : " $global:o_AVpath -foregroundcolor $ccode
write-host "AV Status : " $global:o_AVStatus -foregroundcolor $ccode
write-host "Real-Time Status : " $global:o_RTstate -foregroundcolor $ccode
write-host "Definition Status : " $global:o_DefStatus -foregroundcolor $ccode
write-host "AV Conflict : " $global:o_AVcon -foregroundcolor $ccode
write-host "Competitor AV : " $global:o_CompAV -foregroundcolor $ccode
write-host "Competitor Path : " $global:o_CompPath -foregroundcolor $ccode
Remove-Variable * -ErrorAction SilentlyContinue; Remove-Module *; $error.Clear();
#END SCRIPT
#------------
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
    Version        : 0.1.8 (03 February 2022)
    Creation Date  : 14 December 2021
    Purpose/Change : Provide Primary AV Product Status and Report Possible AV Conflicts
    File Name      : AVHealth_0.1.8.ps1 
    Author         : Christopher Bledsoe - cbledsoe@ipmcomputers.com
    Thanks         : Chris Reid (NAble) for the original 'AV Status' Script and sanity checks
                     Prejay Shah (Doherty Associates) for sanity checks and a second pair of eyes
                     Eddie for their patience and helping test and validate and assistance with Trend Micro and Windows Defender
                     Remco for helping test and validate and assistance with Symantec
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
    0.1.4 Enhanced error handling a bit more to include $_.scriptstacktrace
          Switched to reading AV Product 'Definition' XML data into hashtable format to allow flexible and efficient support of Servers; plan to utilize this method for all devices vs. direcly pulling XML data on each check
          Replaced fallback to WMI "root\cimv2" Namespace and "Win32_Product" Class; per MS documentation this process also starts a consistency check of packages installed, verifying, and repairing the install
          Attempted to utilize 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\' as well but this produced inconsistent results with installed software / nomenclature of installed software
          Instead; Script will retrieve the specified Vendor's AV Products 'Definition' XML and attempt to validate each AV Product via their respective Registry Keys similar to original 'AV Status' Script
            If the Script is able to validate an AV Product for the specified Vendor; it will then write the AV Product name to 'HKLM:\SOFTWARE\Microsoft\Security Center\Monitoring\' for easy retrieval on subsequent runs
          Per MS documentation; fallback to WMI "root\cimv2" Namespace and "Win32reg_AddRemovePrograms" Class may serve as suitable replacement
            https://docs.microsoft.com/en-US/troubleshoot/windows-server/admin-development/windows-installer-reconfigured-all-applications
    0.1.5 Couple bugfixes and fixing a few issues when attempting to monitor 'Windows Defender' as the 'Primary AV Product'
    0.1.6 Bugfixes for monitoring 'Windows Defender' and 'Symantec Anti-Virus' and 'Symantect Endpoint Protection' and multiple AVs on Servers.
            These 2 'Symantec' AV Products are actually the same product; this is simply to deal with differing names in Registry Keys that cannot be changed with Symantec installed
          Adding placeholders for Real-Time Status, Infection Status, and Threats. Added Epoch Timestamp conversion for future use.
    0.1.7 Bugfixes for monitoring 'Trend Micro' and 'Worry-Free Business Security' and multiple AVs on Servers.
            These 2 'Trend Micro' AV Products are actually the same product; this is simply to deal with differing names in Registry Keys that cannot be changed with Trend Micro installed
    0.1.8 Optimization and more bugfixes
            Switched to allow passing of '$i_PAV' via command line; this must be disabled in the AMP code to function properly with NCentral
            Corrected issue where 'Windows Defender' would be populated twice in Competitor AV; this was caused because WMI may report multiple instances of the same AV Product causing competitor check to do multiple runs
            Switched to using a hashtable for storing detected AV Products; this was to prevent duplicate entires for the same AV Product caused by WMI
            Began expansion of metrics to include 'Detection Types' and "Active Detections" based on Sophos' infection status and detected threats registry keys
            Cleaned up formatting for legibility for CLI and within NCentral

.TODO
    Still need more AV Product registry samples for identifying keys to monitor for relevant data
    Need to obtain version and calculate date timestamps for AV Product updates, Definition updates, and Last Scan
    Need to obtain Infection Status and Detected Threats; bonus for timestamps for these metrics
        Do other AVs report individual Threat information in the registry? Sophos does; but if others don't will we be able to use this metric?
#> 

#REGION ----- DECLARATIONS ----
  Param(
    [Parameter(Mandatory=$true)]$i_PAV
  )
  $global:bitarch = ""
  $global:OSCaption = ""
  $global:OSVersion = ""
  $global:producttype = ""
  $global:computername = ""
  $global:blnWMI = $true
  $global:avs = @{}
  $global:avkey = @{}
  $global:o_AVname = ""
  $global:o_AVVersion = ""
  $global:o_AVpath = ""
  $global:o_AVStatus = ""
  $global:rtstatus = "Unknown"
  $global:o_RTstate = "Unknown"
  $global:defstatus = "Unknown"
  $global:o_DefStatus = "Unknown"
  $global:o_Infect = ""
  $global:o_Threats = ""
  $global:o_AVcon = 0
  $global:o_CompAV = ""
  $global:o_CompPath = ""
  $global:o_CompState = ""
  #AV PRODUCTS USING '0' FOR 'UP-TO-DATE' PRODUCT STATUS
  $global:zUpgrade = @(
    "Sophos Intercept X"
    "Symantec Endpoint Protection"
    "Trend Micro Security Agent"
    "Worry-Free Business Security"
    "Windows Defender"
  )
  #SET TLS SECURITY FOR CONNECTING TO GITHUB
  [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls12
#ENDREGION ----- DECLARATIONS ----

#REGION ----- FUNCTIONS ----
  function Get-EpochDate ($epochDate) {                     #Convert Epoch Date Timestamps to Local Time
    [timezone]::CurrentTimeZone.ToLocalTime(([datetime]'1/1/1970').AddSeconds($epochDate))
  } ## Get-EpochDate

  function Get-OSArch {                                     #Determine Bit Architecture & OS Type
    #OS Bit Architecture
    $osarch = (get-wmiobject win32_operatingsystem).osarchitecture
    if ($osarch -like '*64*') {
      $global:bitarch = "bit64"
    } elseif ($osarch -like '*32*') {
      $global:bitarch = "bit32"
    }
    #OS Type & Version
    $global:computername = $env:computername
    $global:OSCaption = (Get-WmiObject Win32_OperatingSystem).Caption
    $global:OSVersion = (Get-WmiObject Win32_OperatingSystem).Version
    $osproduct = (Get-WmiObject -class Win32_OperatingSystem).Producttype
    Switch ($osproduct) {
      "1" {$global:producttype = "Workstation"}
      "2" {$global:producttype = "DC"}
      "3" {$global:producttype = "Server"}
    }
  } ## Get-OSArch

  function Get-AVState {                                    #DETERMINE ANTIVIRUS STATE
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
      "335872" {$global:defstatus = "Up to date" ;$global:rtstatus = "Disabled"}
      #Norton Security
      "327696" {$global:defstatus = "Out of date" ;$global:rtstatus = "Disabled"}
      default {$global:defstatus = "Unknown" ;$global:rtstatus = "Unknown"}
    }
  } ## Get-AVState
#ENDREGION ----- FUNCTIONS ----

#------------
#BEGIN SCRIPT
Get-OSArch
#READ AV PRODUCT DETAILS FROM XML
$srcAVP = "https://raw.githubusercontent.com/CW-Khristos/scripts/dev/AVProducts/" + $i_PAV.replace(" ", "").replace("-", "").tolower() + ".xml"
try {
  $avXML = New-Object System.Xml.XmlDocument
  $avXML.Load($srcAVP)
} catch {
  write-host "XML.Load() - Could not open $srcAVP" -foregroundcolor red
  try {
    $web = new-object system.net.webclient
    [xml]$avXML = $web.DownloadString($srcAVP)
  } catch {
    write-host "Web.DownloadString() - Could not download $srcAVP" -foregroundcolor red
    try {
      start-bitstransfer -erroraction stop -source $srcAVP -destination "C:\IT\Scripts\" + $i_PAV.replace(" ", "").replace("-", "").tolower() + ".xml"
      [xml]$avXML = "C:\IT\Scripts\" + $i_PAV.replace(" ", "").replace("-", "").tolower() + ".xml"
    } catch {
      write-host "BITS.Transfer() - Could not download $srcAVP" -foregroundcolor red
    }
  }
}
#READ XML DATA INTO NESTED HASHTABLE FORMAT FOR LATER USE
foreach ($itm in $avXML.NODE.ChildNodes) {
  $hash = @{
    display = $itm.$global:bitarch.display
    displayval = $itm.$global:bitarch.displayval
    path = $itm.$global:bitarch.path
    pathval = $itm.$global:bitarch.pathval
    ver = $itm.$global:bitarch.ver
    verval = $itm.$global:bitarch.verval
    rt = $itm.$global:bitarch.rt
    rtval = $itm.$global:bitarch.rtval
    stat = $itm.$global:bitarch.stat
    statval = $itm.$global:bitarch.statval
    update = $itm.$global:bitarch.update
    updateval = $itm.$global:bitarch.updateval
    infect = $itm.$global:bitarch.infect
    infectval = $itm.$global:bitarch.infectval
    threat = $itm.$global:bitarch.threat
  }
  $avkey.add($itm.name, $hash)
}
#QUERY WMI SECURITYCENTER NAMESPACE FOR AV PRODUCT DETAILS
if ([system.version]$global:OSVersion -ge [system.version]'6.0.0.0') {
  write-verbose "OS Windows Vista/Server 2008 or newer detected."
  try {
    $AntiVirusProduct = get-wmiobject -Namespace "root\SecurityCenter2" -Class "AntiVirusProduct" -ComputerName $global:computername -ErrorAction Stop
  } catch {
    $blnWMI = $false
  }
} elseif ([system.version]$global:OSVersion -lt [system.version]'6.0.0.0') {
  write-verbose "Windows 2000, 2003, XP detected" 
  try {
    $AntiVirusProduct = get-wmiobject -Namespace "root\SecurityCenter" -Class "AntiVirusProduct"  -ComputerName $global:computername -ErrorAction Stop
  } catch {
    $blnWMI = $false
  }
}
if (-not $blnWMI) {                                         #FAILED TO RETURN WMI SECURITYCENTER NAMESPACE
  try {
    write-host "Failed to query WMI SecurityCenter Namespace" -foregroundcolor red
    write-host "Possibly Server, attempting to  fallback to using 'HKLM:\SOFTWARE\Microsoft\Security Center\Monitoring\' registry key" -foregroundcolor red
    try {                                                   #QUERY 'HKLM:\SOFTWARE\Microsoft\Security Center\Monitoring\' AND SEE IF AN AV IS REGISTRERED THERE
      if ($global:bitarch = "bit64") {
        $AntiVirusProduct = (get-itemproperty -path "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Security Center\Monitoring\*" -ErrorAction Stop).PSChildName
      } elseif ($global:bitarch = "bit32") {
        $AntiVirusProduct = (get-itemproperty -path "HKLM:\SOFTWARE\Microsoft\Security Center\Monitoring\*" -ErrorAction Stop).PSChildName
      }
    } catch {
      write-host "Could not find AV registered in HKLM:\SOFTWARE\Microsoft\Security Center\Monitoring\*" -foregroundcolor red
      $AntiVirusProduct = $null
      $blnSecMon = $true
    }
    if ($AntiVirusProduct -ne $null) {                      #RETURNED 'HKLM:\SOFTWARE\Microsoft\Security Center\Monitoring\' DATA
      $strDisplay = ""
      $blnSecMon = $false
      foreach ($av in $AntiVirusProduct) {
        write-host "Found 'HKLM:\SOFTWARE\Microsoft\Security Center\Monitoring\$av'" -foregroundcolor yellow
        foreach ($key in $global:avkey.keys) {              #ATTEMPT TO VALIDATE EACH AV PRODUCT CONTAINED IN VENDOR XML
          if ($av.replace(" ", "").replace("-", "").toupper() -eq $key.toupper()) {
            $strName = ""
            $regDisplay = $global:avkey[$key].display
            $regDisplayVal = $global:avkey[$key].displayval
            $regPath = $global:avkey[$key].path
            $regPathVal = $global:avkey[$key].pathval
            $regRealTime = $global:avkey[$key].rt
            $regRTVal = $global:avkey[$key].rtval
            $regStat = $global:avkey[$key].stat
            $regStatVal = $global:avkey[$key].statval
            $regInfect = $global:avkey[$key].infect
            $regThreat = $global:avkey[$key].threat
            break
          }
        }
        try {
          if (test-path "HKLM:$regDisplay") {               #ATTEMPT TO VALIDATE INSTALLED AV PRODUCT BY TEST READING A KEY
            write-host "Found 'HKLM:$regDisplay' for product : $key" -foregroundcolor yellow
            try {                                           #IF VALIDATION PASSES; FABRICATE 'HKLM:\SOFTWARE\Microsoft\Security Center\Monitoring\' DATA
              $keyval1 = get-itemproperty -path "HKLM:$regDisplay" -name "$regDisplayVal" -erroraction stop
              $keyval2 = get-itemproperty -path "HKLM:$regPath" -name "$regPathVal" -erroraction stop
              $keyval3 = get-itemproperty -path "HKLM:$regRealTime" -name "$regRTVal" -erroraction stop
              $keyval4 = get-itemproperty -path "HKLM:$regStat" -name "$regStatVal" -erroraction stop
              #$keyval5 = get-itemproperty -path "HKLM:$regInfect" -erroraction stop
              #$keyval6 = get-itemproperty -path "HKLM:$regThreat" -recurse -erroraction stop
              #FORMAT AV DATA
              $strName = $keyval1.$regDisplayVal
              $strDisplay = $strDisplay + $keyval1.$regDisplayVal + ", "
              if ($strName -match "Windows Defender") {     #'NORMALIZE' WINDOWS DEFENDER DISPLAY NAME
                $strDisplay = "Windows Defender, "
              }
              $strPath = $strPath + $keyval2.$regPathVal + ", "
              if ($keyval3.$regRTVal = "0") {               #INTERPRET REAL-TIME SCANNING STATUS
                $strRealTime = $strRealTime + "Enabled, "
              } elseif ($keyval3.$regRTVal = "1") {
                $strRealTime = $strRealTime + "Disabled, "
              }
              $strStat = $strStat + $keyval4.$regStatVal.tostring() + ", "
            } catch {
              write-host "Could not validate Registry data for product : $key" -foregroundcolor red
              write-host $_.scriptstacktrace
              write-host $_
            }
          }
        } catch {
          write-host "Not Found 'HKLM:$regDisplay' for product : $key" -foregroundcolor red
          write-host $_.scriptstacktrace
          write-host $_
        }
      }
    } elseif ($AntiVirusProduct -eq $null) {                #FAILED TO RETURN 'HKLM:\SOFTWARE\Microsoft\Security Center\Monitoring\' DATA
      $strDisplay = ""
      $blnSecMon = $true
      foreach ($key in $global:avkey.keys) {                #ATTEMPT TO VALIDATE EACH AV PRODUCT CONTAINED IN VENDOR XML
        $strName = ""
        $regDisplay = $global:avkey[$key].display
        $regDisplayVal = $global:avkey[$key].displayval
        $regPath = $global:avkey[$key].path
        $regPathVal = $global:avkey[$key].pathval
        $regRealTime = $global:avkey[$key].rt
        $regRTVal = $global:avkey[$key].rtval
        $regStat = $global:avkey[$key].stat
        $regStatVal = $global:avkey[$key].statval
        $regInfect = $global:avkey[$key].infect
        $regThreat = $global:avkey[$key].threat
        try {
          if (test-path "HKLM:$regDisplay") {               #VALIDATE INSTALLED AV PRODUCT BY TESTING READING A KEY
            write-host "Found 'HKLM:$regDisplay' for product : $key" -foregroundcolor yellow
            try {                                           #IF VALIDATION PASSES; FABRICATE 'HKLM:\SOFTWARE\Microsoft\Security Center\Monitoring\' DATA
              $keyval1 = get-itemproperty -path "HKLM:$regDisplay" -name "$regDisplayVal" -erroraction stop
              $keyval2 = get-itemproperty -path "HKLM:$regPath" -name "$regPathVal" -erroraction stop
              $keyval3 = get-itemproperty -path "HKLM:$regRealTime" -name "$regRTVal" -erroraction stop
              $keyval4 = get-itemproperty -path "HKLM:$regStat" -name "$regStatVal" -erroraction stop
              #$keyval5 = get-itemproperty -path "HKLM:$regInfect" -erroraction stop
              #$keyval6 = get-itemproperty -path "HKLM:$regThreat" -recurse -erroraction stop
              #FORMAT AV DATA
              $strName = $keyval1.$regDisplayVal
              $strDisplay = $strDisplay + $keyval1.$regDisplayVal + ", "
              if ($strName -match "Windows Defender") {     #'NORMALIZE' WINDOWS DEFENDER DISPLAY NAME
                $strDisplay = $strDisplay + "Windows Defender, "
              }
              $strPath = $strPath + $keyval2.$regPathVal + ", "
              if ($keyval3.$regRTVal = "0") {               #INTERPRET REAL-TIME SCANNING STATUS
                $strRealTime = $strRealTime + "Enabled, "
              } elseif ($keyval3.$regRTVal = "1") {
                $strRealTime = $strRealTime + "Disabled, "
              }
              $strStat = $strStat + $keyval4.$regStatVal.tostring() + ", "
              if ($blnSecMon) {
                write-host "Creating Registry Key HKLM:\SOFTWARE\Microsoft\Security Center\Monitoring\" $strName " for product : " $strName -foregroundcolor red
                if ($global:bitarch = "bit64") {
                  try {
                    new-item -path "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Security Center\Monitoring\" -name $strName -value $strName -force
                  } catch {
                    write-host "Could not create Registry Key `HKLM:\SOFTWARE\Microsoft\Security Center\Monitoring\" $strName " for product : " $strName -foregroundcolor red
                    write-host $_.scriptstacktrace
                    write-host $_
                  }
                } elseif ($global:bitarch = "bit32") {
                  try {
                    new-item -path "HKLM:\SOFTWARE\Microsoft\Security Center\Monitoring\" -name $strName -value $strName -force
                  } catch {
                    write-host "Could not create Registry Key `HKLM:\SOFTWARE\Microsoft\Security Center\Monitoring\" $strName " for product : " $strName -foregroundcolor red
                    write-host $_.scriptstacktrace
                    write-host $_
                  }
                }
              }
              $AntiVirusProduct = "."
            } catch {
              write-host "Could not validate Registry data for product : $key" -foregroundcolor red
              write-host $_.scriptstacktrace
              write-host $_
              $AntiVirusProduct = $null
            }
            #break
          }
        } catch {
          write-host "Not Found 'HKLM:$regDisplay' for product : $key" -foregroundcolor red
          write-host $_.scriptstacktrace
          write-host $_
        }
      }
    }
    $tmpavs = $strDisplay -split ", "
    $tmppaths = $strPath -split ", "
    $tmprts = $strRealTime -split ", "
    $tmpstats = $strStat -split ", "
    #$avs = $strDisplay -split ", "
    #$avpath = $strPath -split ", "
    #$avrt = $strRealTime -split ", "
    #$avstat = $strStat -split ", "
  } catch {
    write-host "Failed to validate selected AV Products for : " $i_PAV -foregroundcolor red
    write-host $_.scriptstacktrace
    write-host $_
  }
} elseif ($blnWMI) {                                        #RETURNED WMI SECURITYCENTER NAMESPACE
  #SEPARATE RETURNED WMI AV PRODUCT INSTANCES
  $tmpavs = $AntiVirusProduct.displayName -split ", "
  $tmppaths = $AntiVirusProduct.pathToSignedProductExe -split ", "
  $tmpstats = $AntiVirusProduct.productState -split ", "
  #$avs = $string -split ", "
  #$avpath = $string -split ", "
  #$avstat = $string -split ", "
}
#ENSURE ONLY UNIQUE AV PRODUCTS ARE IN '$avs' HASHTABLE
$i = 0
foreach ($tmpav in $tmpavs) {
  if ($avs.count -eq 0) {
    if ($tmprts.count -gt 0) {
      $hash = @{
        display = $tmpavs[$i]
        path = $tmppaths[$i]
        rt = $tmprts[$i]
        stat = $tmpstats[$i]
      }
    } elseif ($tmprts.count -eq 0) {
      $hash = @{
        display = $tmpavs[$i]
        path = $tmppaths[$i]
        stat = $tmpstats[$i]
      }
    }
    $avs.add($tmpavs[$i], $hash)
  } elseif ($avs.count -gt 0) {
    $blnADD = $true
    foreach ($av in $avs.keys) {
      if ($tmpav -eq $av) {
        $blnADD = $false
        break
      }
    }
    if ($blnADD) {
      if ($tmprts.count -gt 0) {
        $hash = @{
          display = $tmpavs[$i]
          path = $tmppaths[$i]
          rt = $tmprts[$i]
          stat = $tmpstats[$i]
        }
      } elseif ($tmprts.count -eq 0) {
        $hash = @{
          display = $tmpavs[$i]
          path = $tmppaths[$i]
          stat = $tmpstats[$i]
        }
      }
      $avs.add($tmpavs[$i], $hash)
    }
  }
  $i = $i + 1
}
#OBTAIN FINAL AV PRODUCT DETAILS
$i = 0
if ($AntiVirusProduct -eq $null) {                          #NO AV PRODUCT FOUND
  write-host "Could not find any AV Product registered" -foregroundcolor red
  $global:o_AVname = "No AV Product Found"
  $global:o_AVVersion = ""
  $global:o_AVpath = ""
  $global:o_AVStatus = "Unknown"
  $global:o_RTstate = "Unknown"
  $global:o_DefStatus = "Unknown"
  $global:o_AVcon = 0
} elseif ($AntiVirusProduct -ne $null) {                    #FOUND AV PRODUCTS
  foreach ($av in $avs.keys) {                              #ITERATE THROUGH EACH FOUND AV PRODUCT
    if (($avs[$av].display -ne $null) -and ($avs[$av].display -ne "")) {
      #NEITHER PRIMARY AV PRODUCT NOR WINDOWS DEFENDER
      if (($avs[$av].display -notmatch $i_PAV) -and ($avs[$av].display -notmatch "Windows Defender")) {
        if (($i_PAV -eq "Trend Micro") -and (($avs[$av].display -notmatch "Trend Micro") -and ($avs[$av].display -notmatch "Worry-Free Business Security"))) {
          $global:o_AVcon = 1
          $global:o_CompAV += "$($avs[$av].display)`r`n" #$avs[$i] + '<br>'
          $global:o_CompPath += "$($avs[$av].display) - $($avs[$av].path)`r`n" #$avpath[$i] + '<br>'
          if ($blnWMI) {
            Get-AVState($avs[$av].stat)
            $global:o_CompState += "$($avs[$av].display) - Real-Time Scanning : $global:rtstatus - Definitions : $global:defstatus`r`n" #$avstat[$i] + '<br>'
          } elseif (-not $blnWMI) {
            $global:o_CompState += "$($avs[$av].display) - Real-Time Scanning : $($avs[$av].rt) - Definitions : N/A`r`n" #$avstat[$i] + '<br>'
          }
        } elseif ($i_PAV -ne "Trend Micro") {
          $global:o_AVcon = 1
          $global:o_CompAV += "$($avs[$av].display)`r`n" #$avs[$i] + '<br>'
          $global:o_CompPath += "$($avs[$av].path)`r`n" #$avpath[$i] + '<br>'
          if ($blnWMI) {
            Get-AVState($avs[$av].stat)
            $global:o_CompState += "$($avs[$av].display) - Real-Time Scanning : $global:rtstatus - Definitions : $global:defstatus`r`n" #$avstat[$i] + '<br>'
          } elseif (-not $blnWMI) {
            $global:o_CompState += "$($avs[$av].display) - Real-Time Scanning : $($avs[$av].rt) - Definitions : N/A`r`n" #$avstat[$i] + '<br>'
          }
        }
      }
      #PRIMARY AV PRODUCT
      if (($avs[$av].display -match $i_PAV) -or 
        (($i_PAV -eq "Trend Micro") -and (($avs[$av].display -match "Trend Micro") -or ($avs[$av].display -match "Worry-Free Business Security")))) {
        #PARSE XML FOR SPECIFIC VENDOR AV PRODUCT
        $node = $avs[$av].display.replace(" ", "").replace("-", "").toupper()
        #AV PRODUCT VERSION KEY PATH AND VALUE
        $i_verkey = $avXML.NODE.$node.$global:bitarch.ver
        $i_verval = $avXML.NODE.$node.$global:bitarch.verval
        #AV PRODUCT REAL-TIME SCANNING KEY PATH AND VALUE
        $i_rtkey = $avXML.NODE.$node.$global:bitarch.rt
        $i_rtval = $avXML.NODE.$node.$global:bitarch.rtval
        #AV PRODUCT STATE KEY PATH AND VALUE
        $i_statkey = $avXML.NODE.$node.$global:bitarch.stat
        $i_statval = $avXML.NODE.$node.$global:bitarch.statval
        #AV PRODUCT LAST UPDATE TIMESTAMP
        $i_update = $avXML.NODE.$node.$global:bitarch.update
        $i_updateval = $avXML.NODE.$node.$global:bitarch.updateval
        #AV PRODUCT INFECTIONS KEY PATH
        $i_infect = $avXML.NODE.$node.$global:bitarch.infect
        $i_infectval = $avXML.NODE.$node.$global:bitarch.infectval
        #AV PRODUCT THREATS KEY PATH
        $i_threat = $avXML.NODE.$node.$global:bitarch.threat
        #AV DETAILS
        $global:o_AVname = $avs[$av].display
        $global:o_AVpath = $avs[$av].path
        #GET PRIMARY AV PRODUCT VERSION VIA REGISTRY
        try {
          write-host "Reading -path 'HKLM:$i_verkey' -name '$i_verval'" -foregroundcolor yellow
          $global:o_AVVersion = get-itemproperty -path "HKLM:$i_verkey" -name "$i_verval" -erroraction stop
        } catch {
          write-host "Could not validate Registry data : -path 'HKLM:$i_verkey' -name '$i_verval'" -foregroundcolor red
          $global:o_AVVersion | add-member -NotePropertyName "$i_verval" -NotePropertyValue "."
        }
        $global:o_AVVersion = $global:o_AVVersion.$i_verval
        #GET PRIMARY AV PRODUCT STATUS VIA REGISTRY
        try {
          write-host "Reading -path 'HKLM:$i_statkey' -name '$i_statval'" -foregroundcolor yellow
          $global:o_AVStatus = get-itemproperty -path "HKLM:$i_statkey" -name "$i_statval" -erroraction stop
        } catch {
          write-host "Could not validate Registry data : -path 'HKLM:$i_statkey' -name '$i_statval'" -foregroundcolor red
          $global:o_AVStatus | add-member -NotePropertyName "$i_statval" -NotePropertyValue "0"
        }
        #INTERPRET 'AVSTATUS' BASED ON ANY AV PRODUCT VALUE REPRESENTATION - SOME TREAT '0' AS 'UPTODATE' SOME TREAT '1' AS 'UPTODATE'
        #$global:o_AVStatus.$i_statval
        if ($global:zUpgrade -contains $avs[$av].display) {
          write-host "$($avs[$av].display) reports '$($global:o_AVStatus.$i_statval)' for 'Up-To-Date' (Expected : '0')" -foregroundcolor yellow
          if ($global:o_AVStatus.$i_statval -eq "0") {
            $global:o_AVStatus = "AV Product Up-to-Date : $true`r`n"
          } else {
            $global:o_AVStatus = "AV Product Up-to-Date : $false`r`n"
          }
        } elseif ($global:zUpgrade -notcontains $avs[$av].display) {
          write-host "$($avs[$av].display) reports '$($global:o_AVStatus.$i_statval)' for 'Up-To-Date' (Expected : '1')" -foregroundcolor yellow
          if ($global:o_AVStatus.$i_statval -eq "1") {
            $global:o_AVStatus = "AV Product Up-to-Date : $true`r`n"
          } else {
            $global:o_AVStatus = "AV Product Up-to-Date : $false`r`n"
          }
        }
        #GET PRIMARY AV PRODUCT LAST UPDATE TIMESTAMP VIA REGISTRY
        try {
          write-host "Reading -path 'HKLM:$i_update' -name '$i_updateval'" -foregroundcolor yellow
          $keyval5 = get-itemproperty -path "HKLM:$i_update" -name "$i_updateval" -erroraction stop
          $global:o_AVStatus += "Last Update : $(Get-EpochDate($keyval5.$i_updateval))`r`n"
        } catch {
          write-host "Could not validate Registry data : -path 'HKLM:$i_update' -name '$i_statval'" -foregroundcolor red
          $global:o_AVStatus += "Last Update : $(Get-EpochDate($keyval5.$i_updateval))`r`n"
        }
        #REAL-TIME SCANNING & DEFINITIONS
        if ($blnWMI) {
          #will still return if it is unknown, etc. if it is unknown look at the code it returns, then look up the status and add it above
          Get-AVState($avs[$av].stat)
          $global:o_DefStatus = $global:defstatus
          $global:o_RTstate = $global:rtstatus
        } elseif (-not $blnWMI) {
          $global:o_DefStatus = "N/A`r`n" #$global:defstatus
          $global:o_RTstate = $avs[$av].rt
        }
        #GET PRIMARY AV PRODUCT DETECTED INFECTIONS VIA REGISTRY
        try {
          write-host "Reading -path 'HKLM:$i_infect'" -foregroundcolor yellow
          if ($i_PAV -match "Sophos") {
            $keyval6 = get-ItemProperty -path "HKLM:$i_infect" -erroraction silentlycontinue
            foreach ($infect in $keyval6.psobject.Properties) {
              if (($infect.name -notlike "PS*") -and ($infect.name -notlike "(default)")) {
                if ($infect.value -eq 0) {
                  $global:o_Infect += "Type - $($infect.name) : $false`r`n"
                } elseif ($infect.value -eq 1) {
                  $global:o_Infect += "Type - $($infect.name) : $true`r`n"
                }
              }
            }
          } elseif ($i_PAV -match "Trend Micro") {
            $keyval6 = get-ItemProperty -path "HKLM:$i_infect" -name "$i_infectval" -erroraction silentlycontinue
            if ($keyval6.$i_infectval -eq 0) {
              $global:o_Infect += "Virus/Malware Present : $false`r`nVirus/Malware Count : $($keyval6.$i_infectval)`r`n"
            } elseif ($keyval6.$i_infectval -gt 0) {
              $global:o_Infect += "Virus/Malware Present : $true`r`nVirus/Malware Count - $($keyval6.$i_infectval) : $true`r`n"
            }
          }
        } catch {
          write-host "Could not validate Registry data : 'HKLM:$i_infect'" -foregroundcolor red
          $global:o_Infect = "N/A"
        }
        #GET PRIMARY AV PRODUCT DETECTED THREATS VIA REGISTRY
        try {
          write-host "Reading -path 'HKLM:$i_threat'" -foregroundcolor yellow
          $keyval7 = get-childitem -path "HKLM:$i_threat" -erroraction silentlycontinue
          if ($keyval7.count -gt 0) {
            foreach ($threat in $keyval7) {
              $keyval8 = get-itemproperty -path "HKLM:$i_threat\$($threat.PSChildName)\" -name "Type" -erroraction silentlycontinue
              $keyval9 = get-childitem -path "HKLM:$i_threat\$($threat.PSChildName)\Files\" -erroraction silentlycontinue
              foreach ($detection in $keyval9) {
                $keyval10 = get-itemproperty -path "HKLM:$i_threat\$($threat.PSChildName)\Files\$($keyval9.PSChildName)\" -name "Path" -erroraction silentlycontinue
                $global:o_Threats += "Threat : $($threat.PSChildName) - Type : $($keyval8.type) - Path : $($keyval10.path)`r`n"
              }
            }
          } elseif ($keyval7.count -le 0) {
            $global:o_Threats += "N/A`r`n"
          }
        } catch {
          write-host "Could not validate Registry data : 'HKLM:$i_threat'" -foregroundcolor red
          $global:o_Threats = "N/A`r`n"
        }
      #SAVE WINDOWS DEFENDER FOR LAST - TO PREVENT SCRIPT CONSIDERING IT 'COMPETITOR AV' WHEN SET AS PRIMARY AV
      } elseif ($avs[$av].display -eq "Windows Defender") {
        $global:o_CompAV += "$($avs[$av].display)`r`n" #$global:o_CompAV + $avs[$i] + " , "
        $global:o_CompPath += "$($avs[$av].path)`r`n" #$global:o_CompPath + $avpath[$i] + " , "
        if ($blnWMI) {
          Get-AVState($avs[$av].stat)
          $global:o_CompState += "$($avs[$av].display) - Real-Time Scanning : $global:rtstatus - Definitions : $global:defstatus`r`n" #$avstat[$i] + '<br>'
        } elseif (-not $blnWMI) {
          $global:o_CompState += "$($avs[$av].display) - Real-Time Scanning : $($avs[$av].rt) - Definitions : N/A`r`n" #$avstat[$i] + '<br>'
        } 
      }
    }
    $i = $i + 1
  }
}
#OUTPUT
if ($global:o_AVname -ne "No AV Product Found") {
  $ccode = "green"
} elseif ($global:o_AVname -eq "No AV Product Found") {
  $ccode = "red"
}
#DEVICE INFO
write-host "`r`nDevice : $global:computername" -foregroundcolor $ccode
write-host "Operating System : $global:OSCaption ($global:OSVersion)" -foregroundcolor $ccode
#AV DETAILS
write-host "AV Display Name : $global:o_AVname" -foregroundcolor $ccode
write-host "AV Version : $global:o_AVVersion" -foregroundcolor $ccode
write-host "AV Path : $global:o_AVpath" -foregroundcolor $ccode
write-host "AV Status : $global:o_AVStatus" -foregroundcolor $ccode
#REAL-TIME SCANNING & DEFINITIONS
write-host "Real-Time Status : $global:o_RTstate" -foregroundcolor $ccode
write-host "Definition Status : $global:o_DefStatus`r`n" -foregroundcolor $ccode
#THREATS
write-host "Active Detections : `r`n$global:o_Infect" -foregroundcolor $ccode
write-host "Detected Threats : `r`n$global:o_Threats" -foregroundcolor $ccode
#COMPETITOR AV
write-host "AV Conflict : $global:o_AVcon" -foregroundcolor $ccode
write-host "Competitor AV : `r`n$global:o_CompAV" -foregroundcolor $ccode
write-host "Competitor Path : `r`n$global:o_CompPath" -foregroundcolor $ccode
write-host "Competitor State : `r`n$global:o_CompState" -foregroundcolor $ccode
#REFORMAT OUTPUT METRICS FOR LEGIBILITY IN NCENTRAL
#AV DETAILS
$global:o_AVStatus = $global:o_AVStatus.replace("`r`n", "<br>")
#THREATS
$global:o_Infect = $global:o_Infect.replace("`r`n", "<br>")
$global:o_Threats = $global:o_Threats.replace("`r`n", "<br>")
#COMPETITOR AV
$global:o_CompAV = $global:o_CompAV.replace("`r`n", "<br>")
$global:o_CompPath = $global:o_CompPath.replace("`r`n", "<br>")
$global:o_CompState = $global:o_CompState.replace("`r`n", "<br>")
#END SCRIPT
#------------
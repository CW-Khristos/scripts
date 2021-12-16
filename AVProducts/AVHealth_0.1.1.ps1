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
    Version        : 0.1.1 (16 December 2021)
    Creation Date  : 14 December 2021
    Purpose/Change : Provide Primary AV Product Status and Report Possible AV Conflicts
    File Name      : AVHealth_0.1.1.ps1 
    Author         : Christopher Bledsoe - cbledsoe@ipmcomputers.com 
    Requires       : PowerShell Version 2.0+ installed

.CHANGELOG
    0.1.0 Initial Release
    0.1.1 Switched to use of '-match' and 'notmatch' for accepting input of vendor / general AV name like 'Sophos'
          Switched to use and expanded AV Product 'Definition' XMLs to be vendor specific instead of product specific
#> 

#REGION ----- DECLARATIONS ----
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
#ENDREGION ----- DECLARATIONS ----

#REGION ----- FUNCTIONS ----
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

#Call Antivirus function to check if it is up to date or not
function Get-AntiVirusProduct {
  [CmdletBinding()]
  param (
    [parameter(ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)]
    [Alias('name')]
    $computername=$env:computername
  )
  try {
    $computername=$env:computername
    [system.Version]$OSVersion = (Get-WmiObject win32_operatingsystem -computername $computername).version

    if ($OSVersion -ge [system.version]'6.0.0.0') {
      Write-Verbose "OS Windows Vista/Server 2008 or newer detected."
      $AntiVirusProduct = Get-WmiObject -Namespace root\SecurityCenter2 -Class AntiVirusProduct -ComputerName $computername -ErrorAction Stop
    } else {
      Write-Verbose "Windows 2000, 2003, XP detected" 
      $AntiVirusProduct = Get-WmiObject -Namespace root\SecurityCenter -Class AntiVirusProduct  -ComputerName $computername -ErrorAction Stop
    }

    #format output
    #will still return if it is unknown, etc. if it is unknown look at the code it returns, then look up the status and add it above
    $i = 0
    #COMMENT OUT THE BELOW LINE (LN92) FOR USE WITH AMP / PASSING OF PRIMARY AV AS INPUT
    #$i_PAV = "Sophos"
    $srcAVP = "https://raw.githubusercontent.com/CW-Khristos/scripts/dev/AVProducts/" + $i_PAV.replace(" ", "").replace("-", "").tolower() + ".xml"
    $srcAVP
    #READ AV PRODUCT DETAILS FROM XML
    try {
      $avXML = New-Object System.Xml.XmlDocument
      $avXML.Load($srcAVP)
    } catch {
      [xml]$avXML = (New-Object System.Net.WebClient).DownloadString($srcAVP)
    }

    #SEPARATE RETURNED WMI AV PRODUCT INSTANCES
    $string = $AntiVirusProduct.displayName
    $avs = $string -split ', '
    $string = $AntiVirusProduct.pathToSignedProductExe
    $avpath = $string -split ', '
    $string = $AntiVirusProduct.productState
    $avstat = $string -split ', '
    #ITERATE THROUGH EACH FOUND AV PRODUCT
    #USE OF '-NE' AND '-EQ' WHEN EVALUATING AV PRODUCT COULD POSSIBLY BE IMPROVED
    #USE OF '-NOTMATCH' AND '-MATCH' COULD THEN INPUT VENDOR / GENERAL AV NAME LIKE 'SOPHOS'
    #THIS WOULD ALLOW EXPANDING AV XML / JSON TO INCLUDE EACH VENDORS' SPECIFIC SEPARATE PRODUCTS AND THEIR RESPECTIVE KEYS / VALUES
    foreach ($av in $avs) {
      #NEITHER PRIMARY AV PRODUCT NOR WINDOWS DEFENDER
      #if (($avs[$i] -ne $i_PAV) -And ($avs[$i] -ne "Windows Defender")) {
      if (($avs[$i] -notmatch $i_PAV) -And ($avs[$i] -notmatch "Windows Defender")) {
        $global:o_AVcon = 1
        $global:o_CompAV = $global:o_CompAV + $avs[$i] + " , "
        $global:o_CompPath = $global:o_CompPath + $avpath[$i] + " , "
        $global:o_Compstate = $global:o_Compstate + $avstat[$i] + " , "
      #PRIMARY AV PRODUCT
      #} elseif ($avs[$i] -eq $i_PAV) {
      } elseif ($avs[$i] -match $i_PAV) {
        $global:o_AVname = $avs[$i]
        $global:o_AVpath = $avpath[$i]
        Get-AVState($avstat[$i])
        $global:o_DefStatus = $global:defstatus
        $global:o_RTstate = $global:rtstatus
        
        #PARSE XML FOR SPECIFIC VENDOR AV PRODUCT
        $node = $avs[$i].replace(" ", "").replace("-", "").toupper()
        #64BIT AV PRODUCT VERSION KEY PATH AND VALUE
        $i_64verkey = $avXML.NODE.$node.bit64.ver
        $i_64verval = $avXML.NODE.$node.bit64.verval
        #32BIT AV PRODUCT VERSION KEY PATH AND VALUE
        $i_32verkey = $avXML.NODE.$node.bit32.ver
        $i_32verval = $avXML.NODE.$node.bit32.verval
        #64BIT AV PRODUCT STATE KEY PATH AND VALUE
        $i_64statkey = $avXML.NODE.$node.bit64.stat
        $i_64statval = $avXML.NODE.$node.bit64.statval
        #32BIT AV PRODUCT STATE KEY PATH AND VALUE
        $i_32statkey = $avXML.NODE.$node.bit32.stat
        $i_32statval = $avXML.NODE.$node.bit32.statval
        
        #GET PRIMARY AV PRODUCT VERSION VIA REGISTRY
        try {       #64BIT REGISTRY LOCATIONS
          $global:o_AVVersion = get-itemproperty -path HKLM:$i_64verkey -name $i_64verval -ErrorAction Stop
          write-host "-path HKLM:$i_64verkey -name $i_64verval" -foregroundcolor Red
          #$global:o_AVVersion = get-itemproperty -path HKLM:\SOFTWARE\WOW6432Node\Sophos\SAVService\Application -name MarketingVersion -ErrorAction Stop
        } catch {
          try {     #32BIT REGISTRY LOCATIONS
            $global:o_AVVersion = get-itemproperty -path HKLM:$i_32verkey -name $i_32verval -ErrorAction Stop
            write-host "-path HKLM:$i_32verkey -name $i_32verval" -foregroundcolor Red
            #$global:o_AVVersion = get-itemproperty -path HKLM:\SOFTWARE\Sophos\SAVService\Application -name MarketingVersion -ErrorAction Stop
          } catch {
            $global:o_AVVersion | add-member -NotePropertyName $i_64verval -NotePropertyValue "."
          }
        }
        $global:o_AVVersion = $global:o_AVVersion.$i_64verval
        #GET PRIMARY AV PRODUCT STATUS VIA REGISTRY
        try {       #64BIT REGISTRY LOCATIONS
          $global:o_AVStatus = get-itemproperty -path HKLM:$i_64statkey -name $i_64statval -ErrorAction Stop
          write-host "-path HKLM:$i_64statkey -name $i_64statval" -foregroundcolor Red
          #$global:o_AVStatus = get-itemproperty -path HKLM:\SOFTWARE\WOW6432Node\Sophos\SavService\Status\ -name UpToDateState -ErrorAction Stop
        } catch {
          try {     #32BIT REGISTRY LOCATIONS
            $global:o_AVStatus = get-itemproperty -path HKLM:$i_32statkey -name $i_32statval -ErrorAction Stop
            write-host "-path HKLM:$i_32statkey -name $i_32statval" -foregroundcolor Red
            #$global:o_AVStatus = get-itemproperty -path HKLM:\SOFTWARE\Sophos\SavService\Status\ -name UpToDateState -ErrorAction Stop
          } catch {
            $global:o_AVStatus | add-member -NotePropertyName $i_64statval -NotePropertyValue "0"
          }
        }
        $global:o_AVStatus.$i_64statval
        if ($global:o_AVStatus.$i_64statval -eq "1") {
          $global:o_AVStatus = $true
        } else {
          $global:o_AVStatus = $false
        }
      #SAVE WINDOWS DEFENDER FOR LAST - TO PREVENT SCRIPT CONSIDERING IT 'COMPETITOR AV' WHEN SET AS PRIMARY AV
      } elseif ($avs[$i] -eq "Windows Defender") {
        $global:o_CompAV = $global:o_CompAV + $avs[$i] + " , "
        $global:o_CompPath = $global:o_CompPath + $avpath[$i] + " , "
        $global:o_Compstate = $global:o_Compstate + $avstat[$i] + " , "  
      }
      $i = $i + 1
    }
    #OUTPUT
    write-host "AV Display Name :" $global:o_AVname -foregroundcolor Green
    write-host "AV Version : " $global:o_AVVersion -foregroundcolor Green
    write-host "AV Path : " $global:o_AVpath -foregroundcolor Green
    write-host "AV Status : " $global:o_AVStatus -foregroundcolor Green
    write-host "Real-Time Status : " $global:o_RTstate -foregroundcolor Green
    write-host "Definition Status : " $global:o_DefStatus -foregroundcolor Green
    write-host "AV Conflict : " $global:o_AVcon -foregroundcolor Green
    write-host "Competitor AV : " $global:o_CompAV -foregroundcolor Green
    write-host "Competitor Path : " $global:o_CompPath -foregroundcolor Green
  } catch {
    Write-Error "\\$computername : WMI Error"
    Write-Error $_
  }
} ## Get-AntiVirusProduct
#ENDREGION ----- FUNCTIONS ----

Get-AntiVirusProduct
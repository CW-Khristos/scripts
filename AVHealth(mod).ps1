$global:o_AVname = ""
$global:o_AVVersion = ""
$global:o_AVpath = ""
$global:o_AVStatus = ""
$global:rtstatus = ""
$global:o_RTstate = ""
$global:defstatus = ""
$global:o_DefStatus = ""
$global:o_AVcon = 0
$global:o_CompAV = " "
$global:o_CompPath = " "
$global:o_Compstate = " "

#CHECK 'PERSISTENT' FOLDERS
if (-not (test-path -path "C:\temp")) {
  new-item -path "C:\temp" -itemtype directory
}
if (-not (test-path -path "C:\IT")) {
  new-item -path "C:\IT" -itemtype directory
}
if (-not (test-path -path "C:\IT\Scripts")) {
  new-item -path "C:\IT\Scripts" -itemtype directory
}

#Determine Antivirus State
function Get-AVState {
  param (
    $state
  )
  #Switch to determine the status of antivirus definitions and real-time protection.
  #switch ($AntiVirusProduct.productState) {
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
}

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
    $i_PAV = "Windows Defender"
    #DOWNLOAD AV PRODUCT 'DEFINITIONS'
    $destAVP = "C:\IT\" + $i_PAV.tolower() + ".xml"
    $srcAVP = "https://raw.githubusercontent.com/CW-Khristos/scripts/dev/AVProducts/" + $i_PAV.tolower() + ".xml"
    try {
      start-bitstransfer -erroraction stop -source $srcAVP -destination $destAVP
    } catch {
      $web = new-object system.net.webclient
      [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
      $web.downloadfile($srcAVP, $destAVP)
    }
    #READ AV PRODUCT DETAILS FROM XML
    [xml]$avXML = get-content -path $destAVP
    #64BIT AV PRODUCT VERSION KEY PATH AND VALUE
    $i_64verkey = $avXML.NODE.bit64.ver
    $i_64verval = $avXML.NODE.bit64.verval
    #32BIT AV PRODUCT VERSION KEY PATH AND VALUE
    $i_32verkey = $avXML.NODE.bit32.ver
    $i_32verval = $avXML.NODE.bit32.verval
    #64BIT AV PRODUCT STATE KEY PATH AND VALUE
    $i_64statkey = $avXML.NODE.bit64.stat
    $i_64statval = $avXML.NODE.bit64.statval
    #32BIT AV PRODUCT STATE KEY PATH AND VALUE
    $i_32statkey = $avXML.NODE.bit32.stat
    $i_32statval = $avXML.NODE.bit32.statval
    
    
    $string = $AntiVirusProduct.displayName
    $avs = $string -split ', '
    $string = $AntiVirusProduct.pathToSignedProductExe
    $avpath = $string -split ', '
    $string = $AntiVirusProduct.productState
    $avstat = $string -split ', '
    
    foreach ($av in $avs) {
      if (($avs[$i] -ne $i_PAV) -And ($avs[$i] -ne "Windows Defender")) {
        $global:o_AVcon = 1
        $global:o_CompAV = $global:o_CompAV + $avs[$i] + " , "
        $global:o_CompPath = $global:o_CompPath + $avpath[$i] + " , "
        $global:o_Compstate = $global:o_Compstate + $avstat[$i] + " , "
      } elseif ($avs[$i] -eq $i_PAV) {
        $global:o_AVname = $avs[$i]
        $global:o_AVpath = $avpath[$i]
        Get-AVState($avstat[$i])
        $global:o_DefStatus = $global:defstatus
        $global:o_RTstate = $global:rtstatus
        
        try {
          $global:o_AVVersion = get-itemproperty -path HKLM:$i_64verkey -name $i_64verval -ErrorAction Stop
          write-host "-path HKLM:$i_64verkey -name $i_64verval" -foregroundcolor Red
          #$global:o_AVVersion = get-itemproperty -path HKLM:\SOFTWARE\WOW6432Node\Sophos\SAVService\Application -name MarketingVersion -ErrorAction Stop
        } catch {
          try {
            $global:o_AVVersion = get-itemproperty -path HKLM:$i_32verkey -name $i_32verval -ErrorAction Stop
            write-host "-path HKLM:$i_32verkey -name $i_32verval" -foregroundcolor Red
            #$global:o_AVVersion = get-itemproperty -path HKLM:\SOFTWARE\Sophos\SAVService\Application -name MarketingVersion -ErrorAction Stop
          } catch {
            $global:o_AVVersion | add-member -NotePropertyName $i_64verval -NotePropertyValue "."
          }
        }
        $global:o_AVVersion = $global:o_AVVersion.$i_64verval
        try {
          $global:o_AVStatus = get-itemproperty -path HKLM:$i_64statkey -name $i_64statval -ErrorAction Stop
          write-host "-path HKLM:$i_64statkey -name $i_64statval" -foregroundcolor Red
          #$global:o_AVStatus = get-itemproperty -path HKLM:\SOFTWARE\WOW6432Node\Sophos\SavService\Status\ -name UpToDateState -ErrorAction Stop
        } catch {
          try {
            $global:o_AVStatus = get-itemproperty -path HKLM:$i_32statkey -name $i_32statval -ErrorAction Stop
            write-host "-path HKLM:$i_32statkey -name $i_32statval" -foregroundcolor Red
            #$global:o_AVStatus = get-itemproperty -path HKLM:\SOFTWARE\Sophos\SavService\Status\ -name UpToDateState -ErrorAction Stop
          } catch {
            $global:o_AVStatus | add-member -NotePropertyName $i_64statval -NotePropertyValue 0
          }
        }
        if ($global:o_AVStatus.$i_64statval = 1) {
          $global:o_AVStatus = $true
        } else {
          $global:o_AVStatus = $false
        }
      } elseif ($avs[$i] -eq "Windows Defender") {
        $global:o_CompAV = $global:o_CompAV + $avs[$i] + " , "
        $global:o_CompPath = $global:o_CompPath + $avpath[$i] + " , "
        $global:o_Compstate = $global:o_Compstate + $avstat[$i] + " , "  
      }
      $i = $i + 1
    }
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
  remove-item $destAVP
}
Get-AntiVirusProduct
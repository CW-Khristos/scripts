# .SYNOPSIS 
    AV Health Monitoring
    This was based on "Get Installed Antivirus Information" by SyncroMSP
    But omits the Hex Conversions and utilization of WSC_SECURITY_PROVIDER , WSC_SECURITY_PRODUCT_STATE , WSC_SECURITY_SIGNATURE_STATUS
    https://mspscripts.com/get-installed-antivirus-information-2/

# .DESCRIPTION 
    Provide Primary AV Product Status and Report Possible AV Conflicts
    Script is intended to be universal / as flexible as possible without being excessively complicated
    Script is intended to replace 'AV Status' VBS Monitoring Script
 
# .NOTES
    Version        : 0.1.8 (08 February 2022)
    Creation Date  : 14 December 2021
    Purpose/Change : Provide Primary AV Product Status and Report Possible AV Conflicts
    File Name      : AVHealth_0.1.8.ps1 
    Author         : Christopher Bledsoe - cbledsoe@ipmcomputers.com
    Thanks         : Chris Reid (NAble) for the original 'AV Status' Script and sanity checks
                     Prejay Shah (Doherty Associates) for sanity checks and a second pair of eyes
                     Eddie for their patience and helping test and validate and assistance with Trend Micro and Windows Defender
                     Remco for helping test and validate and assistance with Symantec
    Requires       : PowerShell Version 2.0+ installed

# .USE
Import "AV Health.amp" AMP in NC Script/Software Repository
 - **Note :** As of 'AVHealth_0.1.8.ps1'; 2 new metrics were added to the monitor; 'Detection Types' and 'Active Detections'
   - If you had previously imported the AMP from any previous versions; you will need to remove the previous "AV Health.amp" and import the latest version of the AMP to enabled these new metrics (this process may require removal of the previous Custom Services / Service Templates)

After importing the AV Health AMP; multiple Custom Services can be created for each desired AV Product to be monitored
![image](https://user-images.githubusercontent.com/10928642/147266859-583eccc5-cc72-40ad-a8b8-43d6d0c461a2.png)
To setup each respective Custom Service; modify the 'Primary AV Product' input for the desired AV Product Vendor
![image](https://user-images.githubusercontent.com/10928642/147267004-6d98e2ed-daba-41d0-af1c-cb77ecb6b843.png)
 - **Note :** The only exception to this is for Windows Defender; if using Windows Defender as the Primary AV Product simply input "Windows Defender"
 - **Note :** It is not necessary to also fill in the "Service Identifier"; I personally prefer to do so so the Service Monitor will appear as "AV Health - Vendor" in NC
 - **Note :** It is also possible to use Custom Properties (Customer or Device) for the 'Primary AV Product' input; this method would forego needing multiple Custom Services
 - **Note :** I have included individual Custom Service exports for each of the supported Vendors at this time

Configure the Thresholds as indicated below :
 - AV Name and AV Path should be set to "Off" or "Contain" and should only need to input the Vendor  for "Normal" status (assumption based on default install paths)
 - **Note :** If monitoring Trend Micro on a Server; the AV Product reports the name as "Worry-Free Business Security" for AV Name, so this threshold should be set to "Off" or modified to match
 - AV Version should be set to "Off"  or "Contain" and "." for "Normal" status
 - AV Product Up-to-Date should be set to "Contain" and "True" for "Normal" status
 - Real-Time Protection should be set to "Match" and "Enabled" for "Normal" status
 - Definition Status should be set to "Contain" and "Up to date" for "Normal" status
 - Detected Infections should be set to "Contain" and "True" for "Failed" status
 - Detected Threats should be set to "Contain" and "N/A" for "Normal" status
 - AV Conflict should be set to "Match" and "0" for "Normal" status
 - Competitor AV should be set to "Off" or "Contain" and "Windows Defender" for "Normal" status
 - Competitor Path should be set to "Off" or "Contain" and "windowsdefender://" for "Normal" status

![image](https://user-images.githubusercontent.com/10928642/147267471-10d07628-3f95-44a3-9ea3-5d6b693a71d6.png)
![image](https://user-images.githubusercontent.com/10928642/147267542-1590e6dc-b385-4e12-8261-9947c8ae1857.png)
![image](https://user-images.githubusercontent.com/10928642/147268240-0b8b5def-d4a3-4ecd-a5bb-b0527a46c94d.png)

After creating the desired Custom Services; create Service Templates for your Windows Devices
 - **Note :** If **not** using Custom Properties to pass Primary AV Product and are planning to monitor multiple AV Products; it will be necessary to create Service Templates for each AV Product you wish to monitor
 - **Note :** Workstations / Laptops - Thresholds for AV Path, Competitor AV, and Competitor Path should be set to "Off"
 - **Note :** Servers - Thresholds for AV Path, Competitor AV, and Competitor Path should be set to "Off"
   - Setting Definition Status Thresholds to "Off" is only a temporary measure until the script fully supports retrieving these values on Servers
 - **Note :** I have included Service Template exports available in this repo (configured for Sophos; but these can easily be modified per below settings)

![image](https://user-images.githubusercontent.com/10928642/147269271-11f3a13e-f09d-48ad-bab8-192c673cafdb.png)


# .CHANGELOG
 - 0.1.0
    - Initial Release
 - 0.1.1
    - Switched to use of '-match' and 'notmatch' for accepting input of vendor / general AV name like 'Sophos'
    - Switched to use and expanded AV Product 'Definition' XMLs to be vendor specific instead of product specific
 - 0.1.2
    - Optimized to reduced use of 'If' blocks for querying registry values
    - Added support for monitoring on Servers using 'HKLM:\SOFTWARE\Microsoft\Security Center\Monitoring\' since WMI SecurityCenter2 Namespace does not exist on Server OSes
    - **Note :** Obtaining AV Products from 'HKLM:\SOFTWARE\Microsoft\Security Center\Monitoring\' only works *if* the AV Product registers itself in that key!
      - If the above registry check fails to find any registered AV Products; script will attempt to fallback to WMI "root\cimv2" Namespace and "Win32_Product" Class -filter "Name like '$i_PAV'"
 - 0.1.3
    - Correcting some bugs and adding better error handling
 - 0.1.4
    - Enhanced error handling a bit more to include $_.scriptstacktrace
    - Switched to reading AV Product 'Definition' XML data into hashtable format to allow flexible and efficient support of Servers; plan to utilize this method for all devices vs. direcly pulling XML data on each check
    - Replaced fallback to WMI "root\cimv2" Namespace and "Win32_Product" Class; per MS documentation this process also starts a consistency check of packages installed, verifying, and repairing the install
    - Attempted to utilize 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\' as well but this produced inconsistent results with installed software / nomenclature of installed software
    - Instead; Script will retrieve the specified Vendor's AV Products 'Definition' XML and attempt to validate each AV Product via their respective Registry Keys similar to original 'AV Status' Script
      - If the Script is able to validate an AV Product for the specified Vendor; it will then write the AV Product name to 'HKLM:\SOFTWARE\Microsoft\Security Center\Monitoring\' for easy retrieval on subsequent runs
    - Per MS documentation; fallback to WMI "root\cimv2" Namespace and "Win32reg_AddRemovePrograms" Class may serve as suitable replacement
      - https://docs.microsoft.com/en-US/troubleshoot/windows-server/admin-development/windows-installer-reconfigured-all-applications
 - 0.1.5
    - Couple bugfixes and fixing a few issues when attempting to monitor 'Windows Defender' as the 'Primary AV Product'
 - 0.1.6
    - Bugfixes for monitoring 'Windows Defender' and 'Symantec Anti-Virus' and 'Symantect Endpoint Protection' and multiple AVs on Servers.
    - These 2 'Symantec' AV Products are actually the same product; this is simply to deal with differing names in Registry Keys that cannot be changed with Symantec installed
    - Adding placeholders for Real-Time Status, Infection Status, and Threats. Added Epoch Timestamp conversion for future use.
 - 0.1.7
    - Bugfixes for monitoring 'Trend Micro' and 'Worry-Free Business Security' and multiple AVs on Servers.
    - These 2 'Trend Micro' AV Products are actually the same product; this is simply to deal with differing names in Registry Keys that cannot be changed with Trend Micro installed
 - 0.1.8
    - Optimization and more bugfixes
    - Switched to allow passing of '$i_PAV' via command line; this must be disabled in the AMP code to function properly with NCentral
    - Corrected issue where 'Windows Defender' would be populated twice in Competitor AV; this was caused because WMI may report multiple instances of the same AV Product causing competitor check to do multiple runs
    - Switched to using a hashtable for storing detected AV Products; this was to prevent duplicate entires for the same AV Product caused by WMI
    - Moved code to retrieve Ven AV Product XMLs to 'Get-AVXML' function to allow dynamic loading of Vendor XMLs and fallback to validating each AV Product from each supported Vendor
    - Began expansion of metrics to include 'Detection Types' and "Active Detections" based on Sophos' infection status and detected threats registry keys
    - Cleaned up formatting for legibility for CLI and within NCentral

# .TODO
    Still need more AV Product registry samples for identifying keys to monitor for relevant data
    Need to obtain version and calculate date timestamps for AV Product updates, Definition updates, and Last Scan
    Need to obtain Infection Status and Detected Threats; bonus for timestamps for these metrics
        Do other AVs report individual Threat information in the registry? Sophos does; but if others don't will we be able to use this metric?
    If no AV is detected through WMI or 'HKLM:\SOFTWARE\Microsoft\Security Center\Monitoring\'; attempt to validate each of the supported Vendor AV Products

# Supported AV Products :
 - Sophos Anti-Virus
 - Symantec Anti-Virus
 - Trend Micro
 - Windows Defender

# AV Products Needing XML 'Definitions' :
 - AVG
 - Avast
 - Avira
 - BitDefender
 - CrowdStrike
 - FortiClient
 - FSecure
 - Kaspersky
 - McAfee
 - Microsoft Defender for Endpoints
 - Norton
 - VIPRE
 - Webroot

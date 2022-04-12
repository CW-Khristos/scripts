Description:
===========
This is a StorCLI readme file mentioning instructions to use:
	1. StorCLI on all the supported operating systems.
	2. StorCLI's JSON Schema files.
	3. StorCLI's logging feature.

Please read before you start using StorCLI executable.

Privileges:
=========
	1. StorCLI should be installed / executed with administrative / root / super user privileges.
	2. Installed/Execution location should have Read/Write/Execute permissions.

Windows:
========
Installation / Execution:
	1. StorCLI is a executable. Copy-Paste the executable from where you want to execute.

Sign verification:
	command : signtool.exe verify /v /pa <storcli executable name>

Notes: 
	1. signtool.exe is required to validate the StorCLI's signature.

Windows-ARM
============
Installation / Execution:
	1. StorCLI is a executable. Copy-Paste the executable from where you want to execute.

Linux:
======
Installation / Execution :
	1. Unzip the StorCLI package.
	2. To install the StorCLI RPM, run the rpm -ivh <StorCLI-x.xx-x.noarch.rpm> command.
	3. To upgrade the StorCLI RPM, run the rpm -Uvh <StorCLI-x.xx-x.noarch.rpm> command.

StorCLI RPM Verification:
	1. Import the public key to RPM DB. Command : rpm --import <public-key.asc>
	2. Verify the RPM signature. Command : rpm -Kv <storcli-rpm>
	3. Install the StorCLI RPM. If imported public key is for the RPM being installed, No warnings should be shown during installation.
	4. Please adhere to the steps in the above mentioned order only.

Linux-ARM:
==========
Installation / Execution :
	1. Unzip the StorCLI package.
	2. To install the StorCLI RPM, run the rpm -ivh <StorCLI-x.xx-x.aarch64.rpm> command.
	3. To upgrade the StorCLI RPM, run the rpm -Uvh <StorCLI-x.xx-x.aarch64.rpm> command.

VMware:
======
Installation:
	1. The StorCLI VIB Package can be installed using the following syntax : esxcli software vib install -v=<Filepath of the StorCLI VIB>
	2. The installed VIB Package can be removed using the following syntax : esxcli software vib remove -n=<VIB Name of StorCLI>
	3. All the installed VIB Packages can be listed using following command: esxcli software vib list
 
Notes : 
	1. VIB under directory "VMwareOP" : This binary is for versions from ESXi6.0 to ESXi6.7.
	2. Offline-bundle under directory "VMwareOP64" : The binary is present under "vib20\vmware-storcli64\" directory of offline-bundle. This binary is for versions from ESXi7.0 and later.

FreeBSD:
========
Installation / Execution:
	1. Extract the tar archive and execute the StorCLI.

Usage policies / Privileges:
	1. StorCli or StorCli64 application will not function if the user is trying to run it in CSH, the default shell in FreeBSD.
	2. Please ensure that the user has entered the bash shell by executing the command "bash".
EFI:
====
Installation / Execution:
	1. From the boot menu, choose EFI Shell.
	2. Goto the folder containing the StorCLI EFI binaries.
	3. Execute StorCLI binaries.

EFI-ARM:
========
Installation / Execution:
	1. From the boot menu, choose EFI Shell.
	2. Goto the folder containing the StorCLI EFI binaries.
	3. Execute StorCLI binaries.

Ubuntu:
=======
Installation:	
	1. Debian package can be installed using following command syntax : sudo dpkg -i <.deb package>
	2. Installed debian package can be verified using following command syntax : dpkg -l | grep -i storcli

PowerPC:
=======
Open Power Big Endian Distribution:
-----------------------------------
Installation / Execution:
	1. Unzip the StorCLI package and execute the storcli binary.


Open Power Little Endian Distribution:
-----------------------------------
Installation / Execution:
	1. Unzip the StorCLI package and execute the storcli binary.
	2. To install .deb package,Use "dpkg" command.


JSON-Schema:
=============
Installation: 
	1. Create a folder under /home/JSON-SCHEMA-FILES.
	2. Unzip the JSON-SCHEMA-FILES.zip and copy all the schema files to /home/JSON-SCHEMA-FILES (In any of the operating systems).
	
Command to Schema mapping:
	1. Please refer to the Schema_mapping_list.xlsx for command to schema mapping.

Logging:
========
	1. While executing StorCLI, Logging is enabled by default.
	2. To Turn-off the logging, Place the storcliconf.ini in the current directory and change the DEBUGLEVEL to 0.
	3. To change the log level, Place the storcliconf.ini in the current directory and change the DEBUGLEVEL to any desired log level.
	4. In case of application crash, Place the storcliconf.ini in the current directory to capture logs.

   

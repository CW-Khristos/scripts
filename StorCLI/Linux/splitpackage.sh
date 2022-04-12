#!/bin/sh

rpm2cpio $1 | cpio -idmv

string=$1

#extracting the binary name and version
OLDIFS=$IFS
IFS="-"
read -a array <<< "$(printf "%s" "$string")"
IFS=$OLDIFS
RPMName=${array[0]}
RPMName64=$RPMName"64"
RPMVersion=${array[1]}

export CLI_VER=${array[1]}

CLI_DIR=/home
rm -Rf $CLI_DIR/StorCLI/
mkdir -p $CLI_DIR/StorCLI
mkdir -p $CLI_DIR/StorCLI/BUILD
mkdir -p $CLI_DIR/StorCLI/RPMS
mkdir -p $CLI_DIR/StorCLI/SOURCES
mkdir -p $CLI_DIR/StorCLI/SPECS
mkdir -p $CLI_DIR/StorCLI/SRPMS

currentDir=`pwd`

# writing the 32 bit spec file for rpm package
echo "Summary: Storage Command Line Tool.
Name: $RPMName
Version: $RPMVersion 
Release: 1
License: AVAGO Technologies
Group:RAID 
URL: http://www.avagotech.com
Distribution:AVAGO Technologies
Vendor:AVAGO Technologies
autoReq: no
autoprov: no
BuildRoot:%{_builddir}
%description
$RPMName is used to manage storage controllers.

%prep
	rm -rf \$RPM_BUILD_DIR/*
	mkdir -p \$RPM_BUILD_DIR/opt/MegaRAID/$RPMName 

%install
	mkdir -p \$RPM_BUILD_ROOT/opt/MegaRAID/$RPMName 
	cp $currentDir/opt/MegaRAID/$RPMName/$RPMName \$RPM_BUILD_ROOT/opt/MegaRAID/$RPMName/
%clean 
	rm -rf \$RPM_BUILD_DIR/*
%post
	if [ -f /opt/MegaRAID/$RPMName ]
	   then
		echo \"Warning! Previous $RPMName package is already installed under /opt directory\"
	fi 
	if [ -f  /usr/sbin/$RPMName ]
	   then
		echo \"Warning! Previous $RPMName package is already installed under /usr/sbin directory\"
	fi

%preun
	echo \"uninstalling $RPMName \"
%postun
	count_of_logs=\`ls /opt/MegaRAID/$RPMName/* 2>/dev/null | wc -l \`
	count_of_megaraid_pkg=\`ls /opt/MegaRAID/ | wc -l \`
    if [ \$count_of_logs -eq 0 ]
 	then 
		echo \"Removing /opt/MegaRAID/$RPMName directory\" > /dev/null 2>&1
		rm -rf /opt/MegaRAID/$RPMName
		if [ \$count_of_megaraid_pkg -eq 1 ]
		then
			echo \"Removing /opt/MegaRAID directory\" > /dev/null 2>&1
			ls /opt/MegaRAID/
			rm -rf /opt/MegaRAID/
		fi
	fi

%files
%attr(-,root,root)	/opt/MegaRAID/$RPMName/$RPMName
"  > storcli.spec

echo "Summary: Storage Command Line Tool.
Name: $RPMName
Version: $RPMVersion 
Release: 1
License: AVAGO Technologies
Group:RAID 
URL: http://www.avagotech.com
Distribution:AVAGO Technologies
Vendor:AVAGO Technologies
autoReq: no
autoprov: no
BuildRoot:%{_builddir}
%description
$RPMName is used to manage storage controllers.

%prep
	rm -rf \$RPM_BUILD_DIR/*
	mkdir -p \$RPM_BUILD_DIR/opt/MegaRAID/$RPMName 

%install
	mkdir -p \$RPM_BUILD_ROOT/opt/MegaRAID/$RPMName 
	cp $currentDir/opt/MegaRAID/$RPMName/$RPMName64 \$RPM_BUILD_ROOT/opt/MegaRAID/$RPMName/
%clean 
	rm -rf \$RPM_BUILD_DIR/*
%post
	if [ -f /opt/MegaRAID/$RPMName ]
	   then
		echo \"Warning! Previous $RPMName package is already installed under /opt directory\"
	fi 
	if [ -f  /usr/sbin/$RPMName64 ]
	   then
		echo \"Warning! Previous $RPMName64 package is already installed under /usr/sbin directory\"
	fi

%preun
	echo \"uninstalling $RPMName64 \"
%postun
	count_of_logs=\`ls /opt/MegaRAID/$RPMName/* 2>/dev/null | wc -l \`
	count_of_megaraid_pkg=\`ls /opt/MegaRAID/ | wc -l \`
    if [ \$count_of_logs -eq 0 ]
 	then 
		echo \"Removing /opt/MegaRAID/$RPMName directory\" > /dev/null 2>&1
		rm -rf /opt/MegaRAID/$RPMName
		if [ \$count_of_megaraid_pkg -eq 1 ]
		then
			echo \"Removing /opt/MegaRAID directory\" > /dev/null 2>&1
			ls /opt/MegaRAID/
			rm -rf /opt/MegaRAID/
		fi
	fi

%files
%attr(-,root,root)	/opt/MegaRAID/$RPMName/$RPMName64
"  > storcli64.spec


rpmbuild -ba --define "_topdir $CLI_DIR/StorCLI" --define "_CLI_VER $CLI_VER" -bb --define "_binary_filedigest_algorithm  1"  --define "_binary_payload 1" --target x86_64  storcli64.spec
cp $CLI_DIR/StorCLI/RPMS/x86_64/*.rpm .

#rm -Rf $CLI_DIR/StorCLI/
mkdir -p $CLI_DIR/StorCLI
mkdir -p $CLI_DIR/StorCLI/BUILD
mkdir -p $CLI_DIR/StorCLI/RPMS
mkdir -p $CLI_DIR/StorCLI/SOURCES
mkdir -p $CLI_DIR/StorCLI/SPECS
mkdir -p $CLI_DIR/StorCLI/SRPMS

rpmbuild -ba --define "_topdir $CLI_DIR/StorCLI" --define "_CLI_VER $CLI_VER" --target i386 storcli.spec
cp $CLI_DIR/StorCLI/RPMS/i386/*.rpm .
rm -Rf $CLI_DIR/StorCLI/

rm -Rf opt/ *.spec

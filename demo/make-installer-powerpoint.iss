; =====================================================================
; == demo\make-installer-powerpoint.iss
; == Part of VstoAddinInstaller
; == (https://github.com/bovender/VstoAddinInstaller)
; == (c) 2016-2018 Daniel Kraus <bovender@bovender.de>
; == Published under the Apache License 2.0
; == See http://www.apache.org/licenses
; =====================================================================
;
; This is the configuration file for the demo installer.
; To try out this script, you must first build the VSTO project in
; Visual Studio with Debug configuration.

#define VERSIONFILE "VERSION.TXT"

#define TARGET_HOST "powerpoint"
#define APP_GUID "{{F16F4118-CFB0-4691-B815-034E9FD9998F}"
#define ADDIN_NAME "VstoAddinInstaller demo for PowerPoint"
#define ADDIN_SHORT_NAME "VstoAddinInstallerDemoPowerPoint"
#define COMPANY "Daniel Kraus (bovender)"
#define HOMEPAGE "https://github.com/bovender/VstoAddinInstaller"
#define DESCRIPTION "Demonstrate VstoAddinInstaller with PowerPoint."
#define PUB_YEARS "2017"

#define SOURCEDIR "VstoInstallerDemoPowerPoint\bin\Debug\"
#define VSTOFILE "VstoInstallerDemoPowerPoint.vsto"
#define OUTPUTDIR "releases\"

; #define LOGFILE "INST-LOG.TXT"

#define SETUPFILESDIR "setup-files\"

; If the VstoAddinInstaller files are in a different subdirectory
; than 'VstoAddinInstaller', change the path below.
#include "..\vsto-installer.iss"

; vim: ts=2 sts=2 sw=2 et

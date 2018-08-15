; =====================================================================
; == demo\make-installer-outlook.iss
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

#define TARGET_HOST "outlook"
#define APP_GUID "{{66EFA565-5CC1-4A03-9603-150EA08EB891}"
#define ADDIN_NAME "VstoAddinInstaller demo for Outlook"
#define ADDIN_SHORT_NAME "VstoAddinInstallerDemoOutlook"
#define COMPANY "Daniel Kraus (bovender)"
#define HOMEPAGE "https://github.com/bovender/VstoAddinInstaller"
#define DESCRIPTION "Demonstrate VstoAddinInstaller with Outlook."
#define PUB_YEARS "2017"

#define SOURCEDIR "VstoInstallerDemoOutlook\bin\Debug\"
#define VSTOFILE "VstoInstallerDemoOutlook.vsto"
#define OUTPUTDIR "releases\"

; #define LOGFILE "INST-LOG.TXT"

#define SETUPFILESDIR "setup-files\"

; If the VstoAddinInstaller files are in a different subdirectory
; than 'VstoAddinInstaller', change the path below.
#include "..\vsto-installer.iss"

; vim: ts=2 sts=2 sw=2 et

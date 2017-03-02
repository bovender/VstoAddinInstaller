; =====================================================================
; == demo\make-installer-excel.iss
; == Part of VstoAddinInstaller
; == (https://github.com/bovender/VstoAddinInstaller)
; == (c) 2016-2017 Daniel Kraus <bovender@bovender.de>
; == Published under the Apache License 2.0
; == See http://www.apache.org/licenses
; =====================================================================
;
; This is the configuration file for the demo installer.
; To try out this script, you must first build the VSTO project in
; Visual Studio with Debug configuration.

#define VERSIONFILE "VERSION.TXT"

#define TARGET_HOST "excel"
#define APP_GUID "{{4915C6C4-11CB-420F-98D4-3609A24D8AC5}"
#define ADDIN_NAME "VstoAddinInstaller demo for Excel"
#define ADDIN_SHORT_NAME "VstoAddinInstallerDemoExcel"
#define COMPANY "Daniel Kraus (bovender)"
#define HOMEPAGE "https://github.com/bovender/VstoAddinInstaller"
#define DESCRIPTION "Demonstrate VstoAddinInstaller with Excel."
#define PUB_YEARS "2017"

#define SOURCEDIR "VstoInstallerDemoExcel\bin\Debug\"
#define VSTOFILE "VstoInstallerDemoExcel.vsto"
#define OUTPUTDIR "releases\"

; #define LOGFILE "INST-LOG.TXT"

#define SETUPFILESDIR "setup-files\"

; If the VstoAddinInstaller files are in a different subdirectory
; than 'VstoAddinInstaller', change the path below.
#include "..\vsto-installer.iss"

; vim: ts=2 sts=2 sw=2 et

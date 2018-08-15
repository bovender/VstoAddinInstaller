; =====================================================================
; == demo\make-installer-word.iss
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

#define TARGET_HOST "word"
#define APP_GUID "{{4EA149D9-B1F5-44CD-A7F0-EA0B88C4B090}"
#define ADDIN_NAME "VstoAddinInstaller demo for Word"
#define ADDIN_SHORT_NAME "VstoAddinInstallerDemoWord"
#define COMPANY "Daniel Kraus (bovende)"
#define HOMEPAGE "https://github.com/bovender/VstoAddinInstaller"
#define DESCRIPTION "Demonstrate VstoAddinInstaller with Word."
#define PUB_YEARS "2017"

#define SOURCEDIR "VstoInstallerDemoWord\bin\Debug\"
#define VSTOFILE "VstoInstallerDemoWord.vsto"
#define OUTPUTDIR "releases\"

; #define LOGFILE "INST-LOG.TXT"

#define SETUPFILESDIR "setup-files\"

; If the VstoAddinInstaller files are in a different subdirectory
; than 'VstoAddinInstaller', change the path below.
#include "..\vsto-installer.iss"

; vim: ts=2 sts=2 sw=2 et

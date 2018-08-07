; =====================================================================
; == vsto-installer.iss
; == Part of VstoAddinInstaller
; == (https://github.com/bovender/VstoAddinInstaller)
; == (c) 2016-2018 Daniel Kraus <bovender@bovender.de>
; == Published under the Apache License 2.0
; == See http://www.apache.org/licenses
; =====================================================================

#ifndef APP_GUID
  #error You must not run this file directly. Make a copy of config-dist/make-installer.dist.iss, edit it and run it.
#endif

#if (TARGET_HOST != "excel") && (TARGET_HOST != "word") && (TARGET_HOST != "outlook") && (TARGET_HOST != "powerpoint")
  #error You must choose between "excel", "word", "powerpoint", and "outlook" as target host applications. Others are currently not supported.
#endif

#if !FileExists(AddBackslash(SOURCEDIR) + VSTOFILE)
  #error Did not find the specified VSTOFILE in SOURCEDIR, please check the spelling -- and make sure you have actually built the project!
#endif

#ifndef SETUPFILESDIR
  #define SETUPFILESDIR "setup-files\"
#endif

; If ADDIN_SHORT_NAME is undefined, use the value from ADDIN_NAME.
; Note however that the short name may be used as the directory name
; and should not contain 'illegal' special characters.
#ifndef ADDIN_SHORT_NAME
  #define ADDIN_SHORT_NAME ADDIN_NAME
#endif


[Setup]
#include "inc\defines.iss"
#include "inc\setup.iss"
	
; Inno Downloader Plugin is required for this
; NB: this directive MUST be located at the end of the [setup] section
#include <idp.iss>

[Files]
Source: {#AddBackslash(SOURCEDIR)}*; DestDir: {app}; Flags: ignoreversion recursesubdirs

; Copy the installer icon, if defined, to the uninstall files dir
#IFDEF INSTALLER_ICO
  Source: {#AddBackslash(SETUPFILESDIR)}{#INSTALLER_ICO}; DestDir: {#UNINSTALLDIR};
#ENDIF

; Define any additional files in a custom files.iss file.
#ifexist "custom-files.iss"
  #include "..\custom-files.iss"
#endif

[Registry]
#include "inc\registry.iss"

[Tasks]
; Define any tasks in the custom tasks.iss file.
#ifexist "custom-tasks.iss"
  #include "..\custom-tasks.iss"
#endif

[Code]
#include "inc\code.pas"

#ifexist "custom-code.pas"
  #include "..\custom-code.pas"
#endif

[Languages]
Name: en; MessagesFile: compiler:Default.isl; 
Name: de; MessagesFile: compiler:Languages\German.isl; 
#ifexist "custom-languages.iss"
  #include "..\custom-languages.iss"
#endif

[CustomMessages]
#include "inc\messages.iss"

; Define any additional messages in the custom messages.iss file.
#ifexist "custom-messages.iss"
  #include "..\custom-messages.iss"
#endif

; vim: ts=2 sts=2 sw=2 et cms=;%s 

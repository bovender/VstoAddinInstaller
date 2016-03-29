; =====================================================================
; == vsto-installer.iss
; == Part of VstoAddinInstaller
; == (https://github.com/bovender/VstoAddinInstaller)
; == (c) 2016 Daniel Kraus <bovender@bovender.de>
; == Published under the Apache License 2.0
; == See http://www.apache.org/licenses
; =====================================================================

#ifndef APP_GUID
  #error You must not run this file directly. Make a copy of config-dist/make-installer.dist.iss, edit it and run it.
#endif

#if (TARGET_HOST != "excel") && (TARGET_HOST != "word")
  #error You must choose between "excel" and "word" as target host applications. PowerPoint and others are currently not supported.
#endif

#if !FileExists(AddBackslash(SOURCEDIR) + VSTOFILE)
  #error Did not find the specified VSTOFILE in SOURCEDIR, please check the spelling -- and make sure you have actually built the project!
#endif


[Setup]
#include "inc/defines.iss"
#include "inc/setup.iss"
	
; Inno Downloader Plugin is required for this
; NB: this directive MUST be located at the end of the [setup] section
#include <idp.iss>

[Files]
; The included file adds all files contained in the SOURCEDIR
Source: {#AddBackslash(SOURCEDIR)}*; DestDir: {app};

; Define any additional files in a custom files.iss file.
#ifexist "files.iss"
  #include "files.iss"
#endif

[Registry]
#include "inc/registry.iss"

[Tasks]
; Define any tasks in the custom tasks.iss file.
#ifexist "tasks.iss"
  #include "tasks.iss"
#endif

[Code]
#include "inc/code.iss"

#ifexist "code.iss"
  #include "code.iss"
#endif

[Languages]
Name: en; MessagesFile: compiler:Default.isl; 
Name: de; MessagesFile: compiler:Languages\German.isl; 
#ifexist "languages.iss"
  #include "languages.iss"
#endif

[CustomMessages]
#include "inc/messages.iss"

; Define any additional messages in the custom messages.iss file.
#ifexist "messages.iss"
  #include "messages.iss"
#endif

; vim: ts=2 sts=2 sw=2 et cms=;%s 

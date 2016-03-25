; VstoAddinInstaller
; InnoSetup script to install Visual Studio for Office (VSTO) addins.
; Originally developed for Daniel's XL Toolbox NG (www.xltoolbox.net).
; Requires the InnoSetup Preprocessor (ISPP).
; Copyright (C) 2016  Daniel Kraus <http://github.com/bovender>
; 
; Licensed under the Apache License, Version 2.0 (the "License");
; you may not use this file except in compliance with the License.
; You may obtain a copy of the License at
;
;     http://www.apache.org/licenses/LICENSE-2.0
;
; Unless required by applicable law or agreed to in writing, software
; distributed under the License is distributed on an "AS IS" BASIS,
; WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
; See the License for the specific language governing permissions and
; limitations under the License.

[Setup]
; Uncomment the following line to use the Debug configuration rather
; than Release
; #define DEBUG

#include "inc/defines.iss"

#ifexist "config.iss"        
  #include "config.iss"
#endif

#include "inc/setup.iss"
	
; Inno Downloader Plugin is required for this
; NB: this directive MUST be located at the end of the [setup] section
#include <idp.iss>

[Files]
; The included file adds all files contained in the SOURCEDIR
#include "inc/files-addins.iss"

; Define any additional files in the custom files.iss file.
#ifexist "files.iss"
  #include "files.iss"
#endif

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
Name: English; MessagesFile: compiler:Default.isl; 
Name: Deutsch; MessagesFile: compiler:Languages\German.isl; 
#ifexist "languages.iss"
  #include "languages.iss"
#endif

[CustomMessages]
#include "inc/messages.iss"

; Define any additional messages in the custom messages.iss file.
#ifexist "messages.iss"
  #include "messages.iss"
#endif

; vim: set ts=2 sts=2 sw=2 noet tw=60 fo+=lj cms=;%s 

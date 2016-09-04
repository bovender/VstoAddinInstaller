; =====================================================================
; == inc\defines.iss
; == Additional defines
; == Part of VstoAddinInstaller
; == (https://github.com/bovender/VstoAddinInstaller)
; == (c) 2016 Daniel Kraus <bovender@bovender.de>
; == Published under the Apache License 2.0
; == See http://www.apache.org/licenses
; =====================================================================

; Download URLs for the .NET runtime and the VSTO runtime
#define DOTNETURL "http://download.microsoft.com/download/9/5/A/95A9616B-7A37-4AF6-BC36-D6EA96C8DAAE/dotNetFx40_Full_x86_x64.exe"
#define VSTORURL "http://download.microsoft.com/download/7/A/F/7AFA5695-2B52-44AA-9A2D-FC431C231EDC/vstor_redist.exe"

; Checksums for the .NET and VSTO runtime installers
#define DOTNETSHA1 "58da3d74db353aad03588cbb5cea8234166d8b99"
#define VSTORSHA1 "f6022eb966df7af80f6df5db0d00a0b7a8f516b3"

; File sizes for the .NET and VSTO runtime installers
#define DOTNETSIZE "50449456"
#define VSTORSIZE "40102072"

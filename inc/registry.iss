; =====================================================================
; == inc/registry.iss
; == Registry keys section
; == Part of VstoAddinInstaller
; == (https://github.com/bovender/VstoAddinInstaller)
; == (c) 2016 Daniel Kraus <bovender@bovender.de>
; == Published under the Apache License 2.0
; == See http://www.apache.org/licenses
; =====================================================================

ValueName: EnableVSTOLocalUNC; ValueData: 1; ValueType: dword; Root: HKLM; Subkey: SOFTWARE\Microsoft\Vsto Runtime Setup\v4; Flags: noerror

; Keys for single-user install (HKCU)
Check: not IsMultiUserInstall; ValueName: Description; ValueData: {#SLOGAN}; ValueType: string; Root: HKCU; Subkey: {code:GetRegKey}; Flags: uninsdeletekey
Check: not IsMultiUserInstall; ValueName: FriendlyName; ValueData: {#APPNAME}; ValueType: string; Root: HKCU; Subkey: {code:GetRegKey}; Flags: uninsdeletekey
Check: not IsMultiUserInstall; ValueName: LoadBehavior; ValueData: 3; ValueType: dword; Root: HKCU; Subkey: {code:GetRegKey}; Flags: uninsdeletekey
Check: not IsMultiUserInstall; ValueName: Warmup; ValueData: 1; ValueType: dword; Root: HKCU; Subkey: {code:GetRegKey}; Flags: uninsdeletekey
Check: not IsMultiUserInstall; ValueName: Manifest; ValueData: file:///{code:ConvertSlash|{app}}/{#ADDINNAME}.vsto|vstolocal; ValueType: string; Root: HKCU; Subkey: {#REGKEY}; Flags: uninsdeletekey

; Same keys again, this time for multi-user install (HKLM)
Check: IsMultiUserInstall; ValueName: Description; ValueData: {#SLOGAN}; ValueType: string; Root: HKLM; Subkey: {code:GetRegKey}; Flags: uninsdeletekey
Check: IsMultiUserInstall; ValueName: FriendlyName; ValueData: {#APPNAME}; ValueType: string; Root: HKLM; Subkey: {code:GetRegKey}; Flags: uninsdeletekey
Check: IsMultiUserInstall; ValueName: LoadBehavior; ValueData: 3; ValueType: dword; Root: HKLM; Subkey: {code:GetRegKey}; Flags: uninsdeletekey
Check: IsMultiUserInstall; ValueName: Warmup; ValueData: 1; ValueType: dword; Root: HKLM; Subkey: {code:GetRegKey}; Flags: uninsdeletekey
Check: IsMultiUserInstall; ValueName: Manifest; ValueData: file:///{code:ConvertSlash|{app}}/{#ADDINNAME}.vsto|vstolocal; ValueType: string; Root: HKLM; Subkey: {#REGKEY}; Flags: uninsdeletekey

; vim: nowrap

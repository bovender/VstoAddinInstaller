; =====================================================================
; == inc\registry.iss
; == Registry keys section
; == Part of VstoAddinInstaller
; == (https://github.com/bovender/VstoAddinInstaller)
; == (c) 2016-2018 Daniel Kraus <bovender@bovender.de>
; == Published under the Apache License 2.0
; == See http://www.apache.org/licenses
; =====================================================================

ValueName: EnableVSTOLocalUNC; ValueData: 1; ValueType: dword; Root: HKLM; Subkey: SOFTWARE\Microsoft\Vsto Runtime Setup\v4; Flags: noerror

; Keys for single-user install (HKCU)
Check: not IsMultiUserInstall; ValueName: Description; ValueData: {#DESCRIPTION}; ValueType: string; Root: HKCU; Subkey: {code:GetRegKey}; Flags: uninsdeletekey
Check: not IsMultiUserInstall; ValueName: FriendlyName; ValueData: {#ADDIN_NAME}; ValueType: string; Root: HKCU; Subkey: {code:GetRegKey}; Flags: uninsdeletekey
Check: not IsMultiUserInstall; ValueName: LoadBehavior; ValueData: 3; ValueType: dword; Root: HKCU; Subkey: {code:GetRegKey}; Flags: uninsdeletekey
Check: not IsMultiUserInstall; ValueName: Warmup; ValueType: none; Root: HKCU; Subkey: {code:GetRegKey}; Flags: deletevalue noerror
Check: not IsMultiUserInstall; ValueName: Manifest; ValueData: file:///{code:ConvertSlash|{app}}/{#VSTOFILE}|vstolocal; ValueType: string; Root: HKCU; Subkey: {code:GetRegKey}; Flags: uninsdeletekey

; Same keys again, this time for multi-user install (HKLM32)
Check: IsMultiUserInstall; ValueName: Description; ValueData: {#DESCRIPTION}; ValueType: string; Root: HKLM32; Subkey: {code:GetRegKey}; Flags: uninsdeletekey
Check: IsMultiUserInstall; ValueName: FriendlyName; ValueData: {#ADDIN_NAME}; ValueType: string; Root: HKLM32; Subkey: {code:GetRegKey}; Flags: uninsdeletekey
Check: IsMultiUserInstall; ValueName: LoadBehavior; ValueData: 3; ValueType: dword; Root: HKLM32; Subkey: {code:GetRegKey}; Flags: uninsdeletekey
Check: IsMultiUserInstall; ValueName: Warmup; ValueType: none; Root: HKLM32; Subkey: {code:GetRegKey}; Flags: deletevalue noerror
Check: IsMultiUserInstall; ValueName: Manifest; ValueData: file:///{code:ConvertSlash|{app}}/{#VSTOFILE}|vstolocal; ValueType: string; Root: HKLM32; Subkey: {code:GetRegKey}; Flags: uninsdeletekey

; Same keys again, this time for multi-user install (HKLM64)
Check: IsMultiUserInstall and IsWin64; ValueName: Description; ValueData: {#DESCRIPTION}; ValueType: string; Root: HKLM64; Subkey: {code:GetRegKey}; Flags: uninsdeletekey
Check: IsMultiUserInstall and IsWin64; ValueName: FriendlyName; ValueData: {#ADDIN_NAME}; ValueType: string; Root: HKLM64; Subkey: {code:GetRegKey}; Flags: uninsdeletekey
Check: IsMultiUserInstall and IsWin64; ValueName: LoadBehavior; ValueData: 3; ValueType: dword; Root: HKLM64; Subkey: {code:GetRegKey}; Flags: uninsdeletekey
Check: IsMultiUserInstall and IsWin64; ValueName: Warmup; ValueType: none; Root: HKLM64; Subkey: {code:GetRegKey}; Flags: deletevalue noerror
Check: IsMultiUserInstall and IsWin64; ValueName: Manifest; ValueData: file:///{code:ConvertSlash|{app}}/{#VSTOFILE}|vstolocal; ValueType: string; Root: HKLM64; Subkey: {code:GetRegKey}; Flags: uninsdeletekey

; vim: nowrap

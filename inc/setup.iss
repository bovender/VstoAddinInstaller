AppName={#product}
VersionInfoProductName={#product}
AppVerName={#product} {#version}
AppPublisher={#company}
VersionInfoCompany={#company}
AppCopyright={#yearspan} {#company}
VersionInfoCopyright={#yearspan} {#company}
VersionInfoDescription=Excel addin.
VersionInfoVersion={#longversion}
VersionInfoProductVersion={#longversion}
VersionInfoTextVersion={#version}

; Make this setup program work with 32-bit and 64-bit Windows
ArchitecturesAllowed=x86 x64
ArchitecturesInstallIn64BitMode=x64

; Always write a log file
SetupLogging=true

; Addins do not need a program group and no user-configurable
; installation folder.
DisableProgramGroupPage=true
DisableDirPage=true
CreateAppDir=true
AppendDefaultDirName=false
DisableReadyPage=true

; Allow normal users to install the addin into their profile.
; This directive also ensures that the uninstall information is
; stored in the user profile rather than a system folder (which
; would require administrative rights).
PrivilegesRequired=lowest

DefaultDirName={code:SuggestInstallDir}
UninstallFilesDir={code:GetDestDir}\uninstall

InternalCompressLevel=max
SolidCompression=true

#ifdef VERSIONFILE
        ; Read the semantic and the installer file version from the VERSION file
        #define FILE_HANDLE FileOpen("{#VERSIONFILE}")
        #define SEMVER FileRead(FILE_HANDLE)
        #define VER FileRead(FILE_HANDLE)
        #expr FileClose(FILE_HANDLE)
#else
        #define SEMVER {#LONGVERSION}
#endif

#ifndef DEBUG
	OutputBaseFilename={#PRODUCT}-{#SEMVER}
#else
	OutputBaseFilename={#PRODUCT}-debug
#endif


#define DOTNETSHA1 "58da3d74db353aad03588cbb5cea8234166d8b99"
#define VSTORSHA1 "ad1dcc5325cb31754105c8c783995649e2208571"

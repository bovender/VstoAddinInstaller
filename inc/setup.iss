; =====================================================================
; == inc/setup.iss
; == Part of VstoAddinInstaller
; == (https://github.com/bovender/VstoAddinInstaller)
; == (c) 2016 Daniel Kraus <bovender@bovender.de>
; == Published under the Apache License 2.0
; == See http://www.apache.org/licenses
; =====================================================================

#ifdef VERSIONFILE
  ; Read the semantic and the installer file version from the VERSION file
  #define FILE_HANDLE FileOpen(VERSIONFILE)
  #define SEMANTIC_VERSION FileRead(FILE_HANDLE)
  #define FOUR_NUMBER_VERSION FileRead(FILE_HANDLE)
  #expr FileClose(FILE_HANDLE)
#pragma message SEMANTIC_VERSION
#else
  #define SEMVER {#LONGVERSION}
#endif

AppId={#APP_GUID}
AppName={#ADDIN_NAME}
VersionInfoProductName={#ADDIN_NAME}
AppVerName={#ADDIN_NAME} {#SEMANTIC_VERSION}
AppPublisher={#COMPANY}
VersionInfoCompany={#COMPANY}
AppCopyright={#PUB_YEARS} {#COMPANY}
VersionInfoCopyright={#PUB_YEARS} {#COMPANY}
VersionInfoDescription={#DESCRIPTION}
VersionInfoVersion={#FOUR_NUMBER_VERSION}
VersionInfoProductVersion={#SEMANTIC_VERSION}
VersionInfoTextVersion={#SEMANTIC_VERSION}
AppPublisherURL={#HOMEPAGE}
#ifdef HOMEPAGE_SUPPORT
  AppSupportURL={#HOMEPAGE_SUPPORT}
#else
  AppSupportURL={#HOMEPAGE}
#endif
#ifdef HOMEPAGE_UPDATES
  AppUpdatesURL={#HOMEPAGE_UPDATES}
#else
  AppUpdatesURL={#HOMEPAGE}
#endif
OutputDir={#OUTPUTDIR}

AppendDefaultDirName=false
ArchitecturesAllowed=x86 x64
ArchitecturesInstallIn64BitMode=x64
CloseApplicationsFilter=*.*
CreateAppDir=true
DefaultDialogFontName=Segoe UI
DefaultDirName={code:SuggestInstallDir}
DisableDirPage=false
DisableProgramGroupPage=true
DisableReadyPage=false
LanguageDetectionMethod=locale
SetupLogging=true
TimeStampsInUTC=false
#DEFINE UNINSTALLDIR "{app}\uninstall"
UninstallFilesDir={#UNINSTALLDIR}

; Allow normal users to install the addin into their profile.
; This directive also ensures that the uninstall information is
; stored in the user profile rather than a system folder (which
; would require administrative rights).
PrivilegesRequired=lowest

InternalCompressLevel=max
SolidCompression=true
#ifndef DEBUG
	OutputBaseFilename={#ADDIN_SHORT_NAME}-{#SEMANTIC_VERSION}
#else
	OutputBaseFilename={#ADDIN_SHORT_NAME}-debug
#endif

#ifdef LICENSE_FILE
  LicenseFile={#SETUPFILESDIR}{#LICENSE_FILE}
#endif

#ifdef INSTALLER_ICO
  SetupIconFile={#SETUPFILESDIR}{#INSTALLER_ICO}
  UninstallDisplayIcon={#UNINSTALLDIR}{#INSTALLER_ICO}
#endif

#ifdef INSTALLER_IMAGE_LARGE
  WizardImageFile={#SETUPFILESDIR}{#INSTALLER_IMAGE_LARGE}
  WizardImageStretch=false
  WizardImageBackColor=clWhite
#endif

#ifdef INSTALLER_IMAGE_SMALL
  WizardSmallImageFile={#SETUPFILESDIR}{#INSTALLER_IMAGE_SMALL}
#endif

; vim: ts=2 sts=2 sw=2 et

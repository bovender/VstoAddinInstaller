{
=====================================================================
== inc\runtimes.iss
== Check, download and install additional runtime files
== (.NET runtime, VSTO 4.0 runtime).
== Create additonal wizard pages
== Part of VstoAddinInstaller
== (https://github.com/bovender/VstoAddinInstaller)
== (c) 2016-2018 Daniel Kraus <bovender@bovender.de>
== Published under the Apache License 2.0
== See http://www.apache.org/licenses
=====================================================================
}

{
  Checks if the VSTO runtime is installed.
  See: http://xltoolbox.sf.net/blog/2015/01/net-vsto-add-ins-getting-prerequisites-right
  HKLM\SOFTWARE\Microsoft\VSTO Runtime Setup\v4R (32-bit)
  HKLM\SOFTWARE\Wow6432Node\Microsoft\VSTO Runtime Setup\v4R (64-bit)
  The 'R' suffix need not be present.
}
function IsVstorInstalled(): boolean;
var
  vstorPath: string;
begin
  vstorPath := 'SOFTWARE\' + GetWowNode + 'Microsoft\VSTO Runtime Setup\v4';
  result := RegKeyExists(HKEY_LOCAL_MACHINE, vstorPath) or
            RegKeyExists(HKEY_LOCAL_MACHINE, vstorPath + 'R');
  if result then
    Log('IsVstorInstalled: VSTO Runtime is installed')
  else
    Log('IsVstorInstalled: VSTO Runtime is not installed');
end;

{
  Extracts the build number from the VSTO runtime version string
  that is stored in the registry.
}
function GetVstorBuild(): integer;
var
  vstorPath: string;
  version: string;
  build: string;
begin
  vstorPath := 'SOFTWARE\' + GetWowNode + 'Microsoft\VSTO Runtime Setup\';
  Log('GetVstorBuild: ' + vstorPath);
  if not RegQueryStringValue(HKEY_LOCAL_MACHINE, vstorPath + 'v4R', 'Version', version) then
  begin
    { Check again without the R suffix. }
    Log('GetVstorBuild: v4R key not found, attempting v4 key');
    if not RegQueryStringValue(HKEY_LOCAL_MACHINE, vstorPath + 'v4', 'Version', version) then
    begin
      Log('GetVstorBuild: v4 key not found either!');
    end
  end;
  if Length(version) > 0 then
  begin
    Log('GetVstorBuild: Version:   ' + version);
    build := Copy(version, 6, 5);
    Log('GetVstorBuild: Build str: ' + build);
    result := StrToIntDef(build, 0);
    Log('GetVstorBuild: Build:     ' + IntToStr(result));
  end
  else
  begin
    Log('GetVstorBuild: Runtime not detected, returning build number 0');
    result := 0;
  end
end;

{
  Checks if the .NET 4.0 (or 4.5) runtime is installed.
  See https://msdn.microsoft.com/en-us/library/hh925568
}
function IsNetInstalled(): boolean;
begin
  result := RegKeyExists(HKEY_LOCAL_MACHINE,
    'SOFTWARE\' + GetWowNode + 'Microsoft\NET Framework Setup\NDP\v4');
end;

{
  Indicates whether or not the .NET runtime needs to be installed;
  this takes into account whether the installer runs with the
  /UPDATE switch or not.
}
function NeedToInstallDotNetRuntime(): boolean;
begin
  result := not IsNetInstalled and not IsUpdate;
end;

{
  Asserts if the VSTO runtime for .NET 4.0 redistributable needs to be
  downloaded and installed.
  If Office 2010 SP 1 or newer is installed on the system, the VSTOR runtime
  will be automagically configured as long as the .NET 4.0 runtime is present.
  Office 2007 and Office 2010 without service pack need the VSTO runtime
  redistributable. For details, see:
  http://xltoolbox.sf.net/blog/2015/01/net-vsto-add-ins-getting-prerequisites-right
}
function NeedToInstallVstor(): boolean;
begin
  Log('NeedToInstallVstor: Minimum required VSTOR 2010 build: ' + IntToStr(MIN_VSTOR_BUILD));
  result := false; { Default }
  if IsOffice2007Installed or IsOffice2010Installed then
    result := (GetVstorBuild < MIN_VSTOR_BUILD) and not IsUpdate;
  if result then
    Log('NeedToInstallVstor: Need to install VSTO runtime')
  else
    Log('NeedToInstallVstor: No need to install VSTO runtime');
end;

{
  Checks if all required prerequisites are met, i.e. if the necessary
  runtimes are installed on the system
}
function PrerequisitesAreMet(): boolean;
begin
  { Cache check result to avoid multiple registry lookups and log messages }
  if not prerequisitesChecked then
  begin
    prerequisitesMet := not NeedToInstallDotNetRuntime and not NeedToInstallVstor;
    prerequisitesChecked := true;
  end;
  result := prerequisitesMet;
end;

{
  Returns the path to the downloaded VSTO runtime installer.
}
function GetVstorInstallerPath(): string;
begin
  result := ExpandConstant('{%temp}\vstor_redist.exe');
end;

{
  Returns the path to the downloaded .NET runtime installer.
}
function GetNetInstallerPath(): string;
begin
  result := ExpandConstant('{%temp}\dotNetFx40_Full_x86_x64.exe');
end;

{
  Checks if the VSTO runtime redistributable setup file has already been
  downloaded by comparing SHA1 checksums.
}
function IsVstorDownloaded(): boolean;
begin
  result := IsFileValid(GetVstorInstallerPath, '{#VSTORSHA1}');
end;

{
  Checks if the .NET runtime setup file has already been
  downloaded by comparing SHA1 checksums.
}
function IsNetDownloaded(): boolean;
begin
  result := IsFileValid(GetNetInstallerPath, '{#DOTNETSHA1}');
end;

{
  Determines if the VSTO runtime needs to be downloaded.
  This is not the case it the runtime is already installed,
  or if there is a file with a valid Sha1 sum.
}
function NeedToDownloadVstor: boolean;
begin
  result := NeedToInstallVstor and not IsVstorDownloaded;
end;

{
  Determines if the VSTO runtime needs to be downloaded.
  This is not the case it the runtime is already installed,
  or if there is a file with a valid Sha1 sum.
}
function NeedToDownloadNet: boolean;
begin
  result := NeedToInstallDotNetRuntime and not IsNetDownloaded;
end;

function ExecuteNetSetup(): boolean;
var
  exitCode: integer;
begin
  result := true;
  if NeedToInstallDotNetRuntime then
  begin
    if IsNetDownloaded then
    begin
      Log('NeedToInstallDotNetRuntime: Valid .NET runtime download found, installing.');
      Exec(GetNetInstallerPath, '/norestart',
        '', SW_SHOW, ewWaitUntilTerminated, exitCode);
      BringToFrontAndRestore;
      if not IsNetInstalled then
      begin
        Log('NeedToInstallDotNetRuntime: .NET runtime still not installed, warning user.');
        MsgBox(CustomMessage('StillNotInstalled'), mbInformation, MB_OK);
        result := True;
      end;
    end
    else
    begin
      Log('NeedToInstallDotNetRuntime: No or invalid .NET runtime download found, warning user.');
      MsgBox(CustomMessage('DownloadNotValidated'), mbInformation, MB_OK);
      result := True;
    end;
  end; { NeedToInstallDotNetRuntime }
end;

function ExecuteVstorSetup(): boolean;
var
  exitCode: integer;
begin
  result := true;
  if NeedToInstallVstor then
  begin
    if IsVstorDownloaded then
    begin
      Log('ExecuteVstorSetup: Valid VSTO runtime download found, installing.');
      Exec(GetVstorInstallerPath, '/norestart', '', SW_SHOW,
        ewWaitUntilTerminated, exitCode);
      BringToFrontAndRestore;
      if not IsVstorInstalled then
      begin
        Log('ExecuteVstorSetup: VSTO runtime still not installed, warning user.');
        MsgBox(CustomMessage('StillNotInstalled'), mbInformation, MB_OK);
        result := True;
      end;
    end
    else
    begin
      Log('ExecuteVstorSetup: No or invalid VSTO runtime download found, warning user.');
      MsgBox(CustomMessage('DownloadNotValidated'), mbInformation, MB_OK);
      result := True;
    end;
  end; { not IsVstorInstalled }
end;

{ vim: set ft=pascal sw=2 sts=2 : }

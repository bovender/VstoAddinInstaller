{
=====================================================================
== inc/runtimes.iss
== Check, download and install additional runtime files
== (.NET runtime, VSTO 4.0 runtime).
== Create additonal wizard pages
== Part of VstoAddinInstaller
== (https://github.com/bovender/VstoAddinInstaller)
== (c) 2016 Daniel Kraus <bovender@bovender.de>
== Published under the Apache License 2.0
== See http://www.apache.org/licenses
=====================================================================
}

{
  Checks if the VSTO runtime is installed. This is relevant if only
  Office 2007 is installed. Since Office 2010, the CLR is
  automatically included.
  The presence of the VSTO runtime is indicated by the presence one of
  four possible registry keys.
  See: http://xltoolbox.sf.net/blog/2015/01/net-vsto-add-ins-getting-prerequisites-right
  HKLM\SOFTWARE\Microsoft\VSTO Runtime Setup\v4R (32-bit)
  HKLM\SOFTWARE\Wow6432Node\Microsoft\VSTO Runtime Setup\v4R (64-bit)
}
function IsVstorInstalled(): boolean;
var
  software, vstorPath: string;
begin
  software := 'SOFTWARE\';
  vstorPath := 'Microsoft\VSTO Runtime Setup\v4R';
  result := RegKeyExists(HKEY_LOCAL_MACHINE, software + GetWowNode + vstorPath);
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
  result := false; { Default for Office 2010 SP1 or newer }
  if IsOffice2007Installed or IsOffice2010NoSpInstalled then
    result := not IsVstorInstalled;
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
    prerequisitesMet := IsNetInstalled and not NeedToInstallVstor;
    prerequisitesChecked := true;
  end;
  result := prerequisitesMet;
end;

{
  Returns the path to the downloaded VSTO runtime installer.
}
function GetVstorInstallerPath(): string;
begin
  result := ExpandConstant('{%temp}\vstor_redist_40.exe');
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
  result := not IsNetInstalled and not IsNetDownloaded;
end;

function ExecuteNetSetup(): boolean;
var
  exitCode: integer;
begin
  result := true;
  if not IsNetInstalled then
  begin
    if IsNetDownloaded then
    begin
      Log('Valid .NET runtime download found, installing.');
      Exec(GetNetInstallerPath, '/norestart',
        '', SW_SHOW, ewWaitUntilTerminated, exitCode);
      BringToFrontAndRestore;
      if not IsNetInstalled then
      begin
        MsgBox(CustomMessage('StillNotInstalled'), mbInformation, MB_OK);
        result := False;
      end;
    end
    else
    begin
      Log('No or invalid .NET runtime download found, will not install.');
      MsgBox(CustomMessage('DownloadNotValidated'), mbInformation, MB_OK);
      result := False;
    end;
  end; { not IsNetInstalled }
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
      Log('Valid VSTO runtime download found, installing.');
      Exec(GetVstorInstallerPath, '/norestart', '', SW_SHOW,
        ewWaitUntilTerminated, exitCode);
      BringToFrontAndRestore;
      if not IsVstorInstalled then
      begin
        MsgBox(CustomMessage('StillNotInstalled'), mbInformation, MB_OK);
        result := False;
      end;
    end
    else
    begin
      Log('No or invalid VSTO runtime download found, will not install.');
      MsgBox(CustomMessage('DownloadNotValidated'), mbInformation, MB_OK);
      result := False;
    end;
  end; { not IsVstorInstalled }
end;

{ vim: set ft=pascal sw=2 sts=2 : }

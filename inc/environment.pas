{
=====================================================================
== inc\environment.iss
== Functions to detect Office environment
== Part of VstoAddinInstaller
== (https://github.com/bovender/VstoAddinInstaller)
== (c) 2016-2017 Daniel Kraus <bovender@bovender.de>
== Published under the Apache License 2.0
== See http://www.apache.org/licenses
=====================================================================
}

{
  Checks if a given version of an Office application is installed
}
function IsHostVersionInstalled(version: integer): boolean;
var
  key: string;
  lookup1: boolean;
  lookup2: boolean;
begin
  key := 'Microsoft\Office\' + IntToStr(version) + '.0\{#TARGET_HOST}\InstallRoot';
  lookup1 := RegKeyExists(HKEY_LOCAL_MACHINE, 'SOFTWARE\' + GetWowNode + key);
  
  {
    If checking for version >= 14.0 ("2010"), which was the first version
    that was produced in both 32-bit and 64-bit, on a 64-bit system we
    also need to check a path without  'Wow6434Node'.
  }
  if IsWin64 and (version >= 14) then
  begin
    lookup2 := RegKeyExists(HKEY_LOCAL_MACHINE, 'SOFTWARE\' + key);
  end;
  
  result := lookup1 or lookup2;
end;

{
  Checks if only Office 2007 is installed
}
function IsOnly2007Installed(): boolean;
var
  i: integer;
begin
  result := IsHostVersionInstalled(12);

  { Iterate through all }
  for i := 14 to MAX_VERSION do
  begin
    if IsHostVersionInstalled(i) then
    begin
      result := false;
      break;
    end;
  end;
end;

{
  Checks if hotfix KB976477 is installed. This hotfix
  is required to make Office 2007 recognize add-ins in
  the HKLM hive as well.
}
function IsHotfixInstalled(): boolean;
begin
  result := RegKeyExists(HKEY_LOCAL_MACHINE,
    'SOFTWARE\Microsoft\Windows\Current Version\Uninstall\KB976477');
end;

{
  Retrieves the build number of an installed Office version
  in OutBuild. Returns true if the requested Office version
  is installed and false if it is not installed.
}
function GetOfficeBuild(OfficeVersion: integer; var OutBuild: integer): boolean;
var
  key: string;
  value: string;
  build: string;
begin
  key := 'SOFTWARE\' + GetWowNode + 'Microsoft\Office\' +
  IntToStr(OfficeVersion) + '.0\Common\ProductVersion';
  if RegQueryStringValue(HKEY_LOCAL_MACHINE, key, 'LastProduct', value) then
  begin
    {
      Office build numbers always have 4 digits, at least as of Feb. 2015;
      from a string '14.0.1234.5000' simply copy 4 characters from the 5th
      position to get the build number. TODO: Make this future-proof.
    }
    build := Copy(value, 6, 4);
    Log('GetOfficeBuild: Found ProductVersion "' + value + '" for queried Office version '
      + IntToStr(OfficeVersion) + ', extracted build number ' + build);
    OutBuild := StrToInt(build);
    result := true;
  end
  else
  begin
    Log('GetOfficeBuild: Did not find LastProduct key for Office version ' +
      IntToStr(OfficeVersion) + '.0.');
  end
end;

{
  Asserts if Office 2007 is installed. Does not check whether other Office
  versions are concurrently installed.
}
function IsOffice2007Installed(): boolean;
begin
  result := IsHostVersionInstalled(12);
  if result then Log('IsOffice2007Installed: Detected Office 2007.');
end;


{
  Asserts if Office 2010 is installed.
}
function IsOffice2010Installed(): boolean;
begin
  result := IsHostVersionInstalled(14);
  if result then Log('IsOffice2010Installed: Detected Office 2010.');
end;

{
  Asserts if Office 2010 without service pack is installed.
  For build number, see http://support.microsoft.com/kb/2121559/en-us
}
function IsOffice2010NoSpInstalled(): boolean;
var
  build: integer;
begin
  if GetOfficeBuild(14, build) then
  begin
    result := build = 4763; { 4763 is the original Office 2007 build }
    if result then
      Log('IsOffice2010NoSpInstalled: Detected Office 2010 without service pack (v. 14.0, build 4763)')
    else
    begin
      Log('IsOffice2010NoSpInstalled: Detected Office 2010, apparently with some service pack (build ' +
        IntToStr(build) + ').');
    end
  end;
end;


{
  Determines whether or not a system-wide installation
  is possible. This depends on whether the current user
  is an administrator, and whether the hotfix KB976477
  is present on the system if Office 2007 is the only version
  of Office that is present (without that hotfix, Office
  2007 does not load add-ins that are registered in the
  HKLM hive).
}
function CanInstallSystemWide(): boolean;
begin
  if IsAdminLoggedOn then
  begin
    if IsOnly2007Installed then
    begin
      result := IsHotfixInstalled;
      if result then
        Log('CanInstallSystemWide: Only Office 2007 found, hotfix installed, can install system-wide.')
      else
        Log('CanInstallSystemWide: Only Office 2007 found but hotfix not installed, cannot install system-wide.')
    end
    else
    begin
      Log('CanInstallSystemWide: User is admin, can install system-wide.')
      result := true;
    end;
  end
  else
  begin
    Log('CanInstallSystemWide: User is not admin, cannot install system-wide.')
    result := false;
  end;
end;

{ vim: set ft=pascal sw=2 sts=2 et : }

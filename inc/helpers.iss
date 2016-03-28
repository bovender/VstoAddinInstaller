{
  Converts backslashes to forward slashes.
}
function ConvertSlash(Value: string): string;
begin
  StringChangeEx(Value, '\', '/', True);
  Result := Value;
end;
  
{
  Returns the path for the Wow6432Node registry tree if the current operating
  system is 64-bit, i.e., simulates WOW64 redirection.
}
function GetWowNode(): string;
begin
  if IsWin64 then
  begin
    result := 'Wow6432Node\';
  end
  else
  begin
    result := '';
  end;
end;

{
  Checks if a given Excel version is installed
}
function IsExcelVersionInstalled(version: integer): boolean;
var key: string;
var lookup1, lookup2: boolean;
begin
  key := 'Microsoft\Office\' + IntToStr(version) + '.0\Excel\InstallRoot';
  lookup1 := RegKeyExists(HKEY_LOCAL_MACHINE, 'SOFTWARE\' + GetWowNode + key);
  
  // If checking for version >= 14.0 ("2010"), which was the first version
  // that was produced in both 32-bit and 64-bit, on a 64-bit system we
  // also need to check a path without  'Wow6434Node'.
  if IsWin64 and (version >= 14) then
  begin
    lookup2 := RegKeyExists(HKEY_LOCAL_MACHINE, 'SOFTWARE\' + key);
  end;
  
  result := lookup1 or lookup2;
end;

{
  Checks if only Excel 2007 is installed
}
function IsOnlyExcel2007Installed(): boolean;
var
  i: integer;
begin
  result := IsExcelVersionInstalled(12);

  // Iterate through all
  for i := 14 to maxExcel do
  begin
    if IsExcelVersionInstalled(i) then
    begin
      result := false;
      break;
    end;
  end;
end;

{
  Checks if hotfix KB976477 is installed. This hotfix
  is required to make Excel 2007 recognize add-ins in
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
    Log('Found ProductVersion "' + value + '" for queried Office version '
      + IntToStr(OfficeVersion) + ', extracted build number ' + build);
    OutBuild := StrToInt(build);
    result := true;
  end
  else
  begin
    Log('Did not find LastProduct key for Office version ' +
      IntToStr(OfficeVersion) + '.0.');
  end
end;

{
  Asserts if Office 2007 is installed. Does not check whether other Office
  versions are concurrently installed.
}
function IsOffice2007Installed(): boolean;
begin
  result := IsExcelVersionInstalled(12);
  if result then Log('Detected Office 2007.');
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
    result := build = 4763; // 4763 is the original Office 2007 build
    if result then
      Log('Detected Office 2010 without service pack (v. 14.0, build 4763)')
    else
    begin
      Log('Detected Office 2010, apparently with some service pack (build ' +
        IntToStr(build) + ').');
    end
  end;
end;


{
  Determines whether or not a system-wide installation
  is possible. This depends on whether the current user
  is an administrator, and whether the hotfix KB976477
  is present on the system if Excel 2007 is the only version
  of Excel that is present (without that hotfix, Excel
  2007 does not load add-ins that are registered in the
  HKLM hive).
}
function CanInstallSystemWide(): boolean;
begin
  if IsAdminLoggedOn then
  begin
    if IsOnlyExcel2007Installed then
    begin
      result := IsHotfixInstalled;
    end
    else
    begin
      result := true;
    end;
  end
  else
  begin
    result := false;
  end;
end;

{
  Helper function that evaluates the custom PageSingleOrMultiUser page.
}
function IsMultiUserInstall(): Boolean;
begin
  result := PageSingleOrMultiUser.Values[1];
end;

{ vim: set ft=pascal sw=2 sts=2 et : }

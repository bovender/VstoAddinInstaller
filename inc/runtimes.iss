/* Check, download and install additional runtime files
*  (.NET runtime, VSTO 4.0 runtime).
*/


/// Checks if only Excel 2007 is installed
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

/// Checks if hotfix KB976477 is installed. This hotfix
/// is required to make Excel 2007 recognize add-ins in
/// the HKLM hive as well.
function IsHotfixInstalled(): boolean;
begin
  result := RegKeyExists(HKEY_LOCAL_MACHINE,
    'SOFTWARE\Microsoft\Windows\Current Version\Uninstall\KB976477');
  end;

  /// Retrieves the build number of an installed Office version
  /// in OutBuild. Returns true if the requested Office version
  /// is installed and false if it is not installed.
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
    // Office build numbers always have 4 digits, at least as of Feb. 2015;
    // from a string '14.0.1234.5000' simply copy 4 characters from the 5th
    // position to get the build number. TODO: Make this future-proof.
    build := Copy(value, 6, 4);
    Log('Found ProductVersion "' + value + '" for queried Office version '
      + IntToStr(OfficeVersion) + ', extracted build number ' + build);
      OutBuild := StrToInt(build);
      result := true;
    end
    else
      Log('Did not find LastProduct key for Office version ' +
        IntToStr(OfficeVersion) + '.0.');
      end;

      /// Asserts if Office 2007 is installed. Does not check whether other Office
      /// versions are concurrently installed.
function IsOffice2007Installed(): boolean;
begin
  result := IsExcelVersionInstalled(12);
  if result then Log('Detected Office 2007.');
end;

/// Asserts if Office 2010 without service pack is installed.
/// For build number, see http://support.microsoft.com/kb/2121559/en-us
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
      Log('Detected Office 2010, apparently with some service pack (build ' +
        IntToStr(build) + ').');
      end;
    end;

    /// Checks if the VSTO runtime is installed. This is relevant if only
    /// Excel 2007 is installed. Since Office 2010, the CLR is
    /// automatically included.
    /// The presence of the VSTO runtime is indicated by the presence one of
    /// four possible registry keys.
    /// See: http://xltoolbox.sf.net/blog/2015/01/net-vsto-add-ins-getting-prerequisites-right
    /// HKLM\SOFTWARE\Microsoft\VSTO Runtime Setup\v4R (32-bit)
    /// HKLM\SOFTWARE\Wow6432Node\Microsoft\VSTO Runtime Setup\v4R (64-bit)
function IsVstorInstalled(): boolean;
var
  software, vstorPath: string;
begin
  software := 'SOFTWARE\';
  vstorPath := 'Microsoft\VSTO Runtime Setup\v4R';
  result := RegKeyExists(HKEY_LOCAL_MACHINE, software + GetWowNode + vstorPath);
end;

/// Checks if the .NET 4.0 (or 4.5) runtime is installed.
/// See https://msdn.microsoft.com/en-us/library/hh925568
function IsNetInstalled(): boolean;
begin
  result := RegKeyExists(HKEY_LOCAL_MACHINE, 
    'SOFTWARE\' + GetWowNode + 'Microsoft\NET Framework Setup\NDP\v4');
  end;

  /// Asserts if the VSTO runtime for .NET 4.0 redistributable needs to be
  /// downloaded and installed.
  /// If Office 2010 SP 1 or newer is installed on the system, the VSTOR runtime
  /// will be automagically configured as long as the .NET 4.0 runtime is present.
  /// Office 2007 and Office 2010 without service pack need the VSTO runtime
  /// redistributable. For details, see:
    /// http://xltoolbox.sf.net/blog/2015/01/net-vsto-add-ins-getting-prerequisites-right
function NeedToInstallVstor(): boolean;
begin
  result := false; // Default for Office 2010 SP1 or newer
  if IsOffice2007Installed or IsOffice2010NoSpInstalled then
    result := not IsVstorInstalled;
end;

/// Checks if all required prerequisites are met, i.e. if the necessary
/// runtimes are installed on the system
function PrerequisitesAreMet(): boolean;
begin
  // Cache check result to avoid multiple registry lookups and log messages
  if not prerequisitesChecked then
  begin
    prerequisitesMet := IsNetInstalled and not NeedToInstallVstor;
    prerequisitesChecked := true;
  end;
  result := prerequisitesMet;
end;

/// Checks if a file exists and has a valid Sha1 sum.
function IsFileValid(file: string; expectedSha1: string): boolean;
var
  actualSha1: string;
begin
  try
    if FileExists(file) then
    begin
      actualSha1 := GetSHA1OfFile(file);
    end;
  finally
    result := actualSha1 = expectedSha1;
  end;
end;

/// Returns the path to the downloaded VSTO runtime installer.
function GetVstorInstallerPath(): string;
begin
  result := ExpandConstant('{%temp}\vstor_redist_40.exe');
end;

/// Returns the path to the downloaded .NET runtime installer.
function GetNetInstallerPath(): string;
begin
  result := ExpandConstant('{%temp}\dotNetFx40_Full_x86_x64.exe');
end;

/// Checks if the VSTO runtime redistributable setup file has already been
/// downloaded by comparing SHA1 checksums.
function IsVstorDownloaded(): boolean;
begin
  result := IsFileValid(GetVstorInstallerPath, '{#VSTORSHA1}');
end;

/// Checks if the .NET runtime setup file has already been
/// downloaded by comparing SHA1 checksums.
function IsNetDownloaded(): boolean;
begin
  result := IsFileValid(GetNetInstallerPath, '{#DOTNETSHA1}');
end;

/// Determines if the VSTO runtime needs to be downloaded.
/// This is not the case it the runtime is already installed,
/// or if there is a file with a valid Sha1 sum.
function NeedToDownloadVstor: boolean;
begin
  result := NeedToInstallVstor and not IsVstorDownloaded;
end;

/// Determines if the VSTO runtime needs to be downloaded.
/// This is not the case it the runtime is already installed,
/// or if there is a file with a valid Sha1 sum.
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
  end; // not IsNetInstalled
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
  end; // not IsVstorInstalled
end;

// vim: ft=pascal sw=2 sts=2

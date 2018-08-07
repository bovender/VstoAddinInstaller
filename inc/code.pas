{
=====================================================================
== inc\code.iss
== Pascal/RemObjects code section
== Part of VstoAddinInstaller
== (https://github.com/bovender/VstoAddinInstaller)
== (c) 2016-2018 Daniel Kraus <bovender@bovender.de>
== Published under the Apache License 2.0
== See http://www.apache.org/licenses
=====================================================================
}

var
  PageSingleOrMultiUser: TInputOptionWizardPage;
  PageCannotInstall: TInputOptionWizardPage;
  PageDownloadInfo: TOutputMsgWizardPage;
  PageInstallInfo: TOutputMsgWizardPage;
  prerequisitesChecked: boolean;
  prerequisitesMet: boolean;
  exePath: string;
  IsUpdate: boolean;


#include "constants.pas"
#include "helpers.pas"
#include "win32.pas"
#include "environment.pas"
#include "runtimes.pas"
#include "wizard-pages.pas"
#include "detect-running-app.pas"


{
  Returns true if running on a zero client. The algorithm has only been
  tested for VMware Horizon/Teradici clients.
}
function IsZeroClient(): boolean;
var
  protocol: string;
begin
  if RegQueryStringValue(HKEY_CURRENT_USER, 'Volatile Environment',
    'ViewClient_Protocol', protocol) then
  begin
    Log('IsZeroClient: ViewClient_Protocol: ' + protocol)
    result := Uppercase(protocol) = 'PCOIP';
    if result then
      Log('IsZeroClient: Recognized as zero client')
    else
      Log('IsZeroClient: Not recognized as a zero client')
  end;
end;

{
  Returns true if the target directory chooser should be shown or
  not: This is the case if running on a zero client, or if the
  current user is an administrator.
}
function ShouldShowDirPage(): boolean;
begin
  result := IsAdminLoggedOn; // or IsZeroClient;
end;

function InitializeSetup(): boolean;
var
  i: integer;
  minVersionInstalled: boolean;
begin
  {
    Determine if Office 2007 or newer is installed (absolute requirement
    for this VSTO add-in). Office 2007 ist version 12.0.
  }
  for i := 12 to MAX_VERSION do
  begin
    minVersionInstalled := minVersionInstalled or IsHostVersionInstalled(i);
  end;

  if not minVersionInstalled then
  begin
    result := False;
    Log('InitializeSetup: Informing user that Office 2007 or newer is required.');
    MsgBox(CustomMessage('Office2007Required'), mbInformation, MB_OK);
  end
  else
  begin
    for i := 1 to ParamCount do
    begin
      if uppercase(ParamStr(i)) = '/UPDATE' then
      begin
        Log('InitializeSetup: /UPDATE switch found');
        IsUpdate := true;
        exePath := CloseAppNoninteractively();
        result := true;
      end
    end;

    if not IsUpdate then
    begin
      result := CloseAppInteractively();
    end;
  end;
end;

procedure InitializeWizard();
begin
  CreateSingleOrAllUserPage;
  if not PrerequisitesAreMet then
  begin
    Log('InitializeWizard: Not all prerequisites are met...');
    CreateCannotInstallPage;
    if NeedToDownloadNet then
    begin
      Log('InitializeWizard: Mark {#DOTNETURL} for download.');
      idpAddFileSize('{#DOTNETURL}', GetNetInstallerPath, {#DOTNETSIZE});
    end;
    if NeedToDownloadVstor then
    begin
      Log('InitializeWizard: Mark {#VSTORURL} for download.');
      idpAddFileSize('{#VSTORURL}', GetVstorInstallerPath, {#VSTORSIZE});
    end;
    CreateDownloadInfoPage;
    CreateInstallInfoPage;
    idpDownloadAfter(PageDownloadInfo.Id);
  end;
end;

function NextButtonClick(CurPageID: Integer): Boolean;
begin
  Log('NextButtonClick: CurPageID = ' + IntToStr(CurPageID));
  result := True;
  if not WizardSilent then
  begin
    {
    if CurPageID = PageDevelopmentInfo.Id then
    begin
      if PageDevelopmentInfo.Values[0] = False then
      begin
        Log('Requesting user to acknowledge use of a developmental version.');
        MsgBox(CustomMessage('DevVerMsgBox'), mbInformation, MB_OK);
        result := False;
      end;
    end;
    }
  end;

  if not PrerequisitesAreMet then
  begin
    {
      Abort the installation if any of the runtimes are missing, the user
      is not an administrator, and requested to abort the installation.
    }
    if CurPageID = PageCannotInstall.ID then
    begin
      if PageCannotInstall.Values[1] = true then
      begin
        Log('NextButtonClick: Non-admin user cannot install, aborting.');
        WizardForm.Close;
        result := False;
      end
      else
      begin
        Log('NextButtonClick: Non-admin user continues although not all required runtimes are installed.');
      end;
    end;

    if CurPageID = PageInstallInfo.ID then
    begin
      { Return true if installation succeeds (or no installation required) }
      result := ExecuteNetSetup and ExecuteVstorSetup;
    end;
  end; { not PrerequisitesAreMet }
end;

{
  Skips the folder selection, single/multi user, and ready pages for
  normal users without power privileges.
  This function also takes care of dynamically determining what wizard
  pages to install, depending on the current system setup and whether
  the current user is an administrator.
}
function ShouldSkipPage(PageID: Integer): Boolean;
begin
  result := False;

  if not PrerequisitesAreMet then
  begin
    {
      The PageDownloadCannotInstall will only have been initialized if
      PrerequisitesAreMet returned false.
    }
    if PageID = PageCannotInstall.ID then
    begin
      { Skip the warning if the user is an admin. }
      result := IsAdminLoggedOn
      if not result then
      begin
        Log('ShouldSkipPage: Warning user that required runtimes cannot be installed due to missing privileges');
      end;
    end;

    if PageID = PageDownloadInfo.ID then
    begin
      { Skip page informing about downloads if no files need to be downloaded. }
      result := idpFilesCount = 0;
    end;

    if PageID = IDPForm.Page.ID then
    begin
      { Skip downloader plugin if there are no files to download. }
      result := idpFilesCount = 0;
      if not result then
      begin
        Log('ShouldSkipPage: Beginning download of ' + IntToStr(idpFilesCount) + ' file(s).');
      end;
    end;
  end; { not PrerequisitesAreMet }

  if PageID = PageSingleOrMultiUser.ID then
  begin
    if CanInstallSystemWide then
    begin
      Log('ShouldSkipPage: Do not skip multi-user page, offer installation for all users.');
      result := False;
    end
    else
    begin
      Log('ShouldSkipPage: Skip multi-user page, offer single-user installation only.');
      result := True;
    end;
  end;

  if (PageID = wpSelectDir) or (PageID = wpReady) then
  begin
    {
      Do not show the pages to select the target directory, and the ready
      page if the user is not an admin.
    }
    result := not ShouldShowDirPage;
    if result then
      Log('ShouldSkipPage: Skipping target directory query.')
    else
      Log('ShouldSkipPage: Showing target directory query.')
  end
end;

{
  Suggest an initial target directory depending on whether
  the installer is run with admin privileges.
}
function SuggestInstallDir(Param: string): string;
var
  dir: string;
begin
  if CanInstallSystemWide then
  begin
    Log('SuggestInstallDir: Can install system-wide, suggesting Programs folder');
    dir := ExpandConstant('{pf}');
  end
  else
  begin
    // if IsZeroClient then
    // begin
    //   Log('SuggestInstallDir: Looks like zero client, suggesting user docs folder');
    //   dir := ExpandConstant('{userdocs}')
    // end
    // else
    // begin
      Log('SuggestInstallDir: Suggesting user profile folder');
      dir := ExpandConstant('{userappdata}')
    // end
  end;
  result := AddBackslash(dir) + '{#ADDIN_SHORT_NAME}';
  Log('SuggestInstallDir: ' + result);
end;

procedure DeinitializeSetup();
var
  e: Integer;
begin
  if Length(exePath) > 0 then
  begin
    Log('DeinitializeSetup: Restarting Office host...');
    Log(exePath);
    Exec(exePath, '', '', SW_SHOW, ewNoWait, e);
  end;

#ifdef LOGFILE
  {
    Copy the log file to the installation
  }
  try
    Log('DeinitializeSetup: Copying log file to installation folder');
    FileCopy(
      ExpandConstant('{log}'),
      AddBackslash(ExpandConstant('{app}'))+'{#LOGFILE}', false);
  except
    Log('DeinitializeSetup: Failed to copy log file');
  end
#endif
end;

{ vim: set ft=pascal sw=2 sts=2 et : }

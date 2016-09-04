{
=====================================================================
== inc\code.iss
== Pascal/RemObjects code section
== Part of VstoAddinInstaller
== (https://github.com/bovender/VstoAddinInstaller)
== (c) 2016 Daniel Kraus <bovender@bovender.de>
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


#include "constants.iss"
#include "helpers.iss"
#include "win32.iss"
#include "environment.iss"
#include "runtimes.iss"
#include "wizard-pages.iss"
#include "detect-running-app.iss"

  
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
    result := Uppercase(protocol) = 'PCOIP';
  end;
end;

{
  Returns true if the target directory chooser should be shown or
  not: This is the case if running on a zero client, or if the
  current user is an administrator.
}
function ShouldShowDirPage(): boolean;
begin
  result := IsAdminLoggedOn or IsZeroClient;
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
    Log('Informing user that Office 2007 or newer is required.');
    MsgBox(CustomMessage('Office2007Required'), mbInformation, MB_OK);
  end
  else
  begin
    for i := 1 to ParamCount do
    begin
      if uppercase(ParamStr(i)) = '/UPDATE' then
      begin
        Log('/UPDATE switch found');
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
    Log('Not all prerequisites are met...');
    CreateCannotInstallPage;
    if NeedToDownloadNet then
    begin
      Log('Mark {#DOTNETURL} for download.');
      idpAddFileSize('{#DOTNETURL}', GetNetInstallerPath, {#DOTNETSIZE});
    end;
    if NeedToDownloadVstor then
    begin
      Log('Mark {#VSTORURL} for download.');
      idpAddFileSize('{#VSTORURL}', GetVstorInstallerPath, {#VSTORSIZE});
    end;
    CreateDownloadInfoPage;
    CreateInstallInfoPage;
    idpDownloadAfter(PageDownloadInfo.Id);
  end;
end;

function NextButtonClick(CurPageID: Integer): Boolean;
begin
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
        WizardForm.Close;
        result := False;
      end
      else
      begin
        Log('Non-admin user continues although not all required runtimes are installed.');
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
        Log('Warning user that required runtimes cannot be installed due to missing privileges');
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
        Log('Beginning download of ' + IntToStr(idpFilesCount) + ' file(s).');
      end;
    end;
  end; { not PrerequisitesAreMet }

  if PageID = PageSingleOrMultiUser.ID then
  begin
    if IsOnly2007Installed then
    begin
      Log('Only Office 2007 appears to be installed on this system.');
      if IsHotfixInstalled then
      begin
        Log('Hotfix KB976477 found; can install for all users.');
      end
      else
      begin
        Log('Hotfix KB976477 not found; cannot install for all users.');
      end;
    end
    else
    begin
      Log('Office 2010 or newer found on this system.');
    end;
    if CanInstallSystemWide then
    begin
      Log('Offer installation for all users.');
      result := False;
    end
    else
    begin
      Log('Offer single-user installation only.');
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
    dir := ExpandConstant('{pf}');
  end
  else
  begin
    if IsZeroClient then
    begin
      dir := ExpandConstant('{userdocs}')
    end
    else
    begin
      dir := ExpandConstant('{userappdata}')
    end
  end;
  result := AddBackslash(dir) + '{#ADDIN_SHORT_NAME}';
end;

procedure DeinitializeSetup();
var
  e: Integer;
begin
  if Length(exePath) > 0 then
  begin
    Log('Restarting Office host...');
    Log(exePath);
    Exec(exePath, '', '', SW_SHOW, ewNoWait, e);
  end;

#ifdef LOGFILE
  {
    Copy the log file to the installation
  }
  try
    FileCopy(
      ExpandConstant('{log}'),
      AddBackslash(ExpandConstant('{app}'))+'{#LOGFILE}', false);
  except
  end
#endif
end;

{ vim: set ft=pascal sw=2 sts=2 et : }

{
=====================================================================
== inc\wizard-pages.iss
== Create additonal wizard pages
== Part of VstoAddinInstaller
== (https://github.com/bovender/VstoAddinInstaller)
== (c) 2016 Daniel Kraus <bovender@bovender.de>
== Published under the Apache License 2.0
== See http://www.apache.org/licenses
=====================================================================
}

procedure CreateSingleOrAllUserPage();
begin
  PageSingleOrMultiUser := CreateInputOptionPage(wpLicense,
    CustomMessage('SingleOrMulti'), CustomMessage('SingleOrMultiSubcaption'),
    CustomMessage('SingleOrMultiDesc'), True, False);
  PageSingleOrMultiUser.Add(CustomMessage('SingleOrMultiSingle'));
  PageSingleOrMultiUser.Add(CustomMessage('SingleOrMultiAll'));
  if CanInstallSystemWide then
  begin
    PageSingleOrMultiUser.Values[1] := True;
  end
  else
  begin
      PageSingleOrMultiUser.Values[0] := True;
  end;
end;

procedure CreateCannotInstallPage();
begin
  PageCannotInstall := CreateInputOptionPage(wpWelcome,
    CustomMessage('CannotInstallCaption'),
    CustomMessage('CannotInstallDesc'),
    CustomMessage('CannotInstallMsg'), True, False);
  PageCannotInstall.Add(CustomMessage('CannotInstallCont'));
  PageCannotInstall.Add(CustomMessage('CannotInstallAbort'));
  PageCannotInstall.Values[1] := True;
end;

procedure CreateDownloadInfoPage();
var
  bytes: Int64;
  mib: Single;
  size: String;
begin
  if idpGetFilesSize(bytes) then
  begin
    mib := bytes / 1048576;
    size := Format('%.1f', [ mib ]);
  end
  else
  begin
    size := '[?]'
  end;
  PageDownloadInfo := CreateOutputMsgPage(PageSingleOrMultiUser.Id,
    CustomMessage('RequiredCaption'),
    CustomMessage('RequiredDesc'),
    Format(CustomMessage('RequiredMsg'), [idpFilesCount, size]));
end;

procedure CreateInstallInfoPage();
begin
  PageInstallInfo := CreateOutputMsgPage(PageDownloadInfo.Id,
    CustomMessage('InstallCaption'),
    CustomMessage('InstallDesc'),
    CustomMessage('InstallMsg'));
end;

{ vim: set ft=pascal sw=2 sts=2 et : }

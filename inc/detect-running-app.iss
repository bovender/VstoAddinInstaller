{
=====================================================================
== inc\detect-running-app.iss
== Determine the name of the Office application's main window.
== Part of VstoAddinInstaller
== (https://github.com/bovender/VstoAddinInstaller)
== (c) 2016 Daniel Kraus <bovender@bovender.de>
== Published under the Apache License 2.0
== See http://www.apache.org/licenses
=====================================================================
}

{
  Returns the class name of the main window of the hosting Office application.
  See http://users.skynet.be/am044448/Programmeren/VBA/vba_class_names.htm
}
function OfficeWindowName(): string;
begin
#if TARGET_HOST == "excel"
  result := 'xlmain';
#elif TARGET_HOST == "word"
  result := 'opusapp';
#elif TARGET_HOST == "powerpoint"
  { TODO: Deal with different PowerPoint versions; PP10FrameClass is for
    PowerPoint XP only... }
  result := 'PP10FrameClass';
#elif TARGET_HOST == "access"
  result := 'omain';
#elif TARGET_HOST == "outlook"
  result := 'rctrl_renwnd32';
#else
  #error Cannot handle this TARGET_HOST value in OfficeWindowName()
#endif
end;

{
  Detect if a given Office application is running and offers
  to close it, or abort the installation.
  windowName: name of the application's main window (e.g. 'XLMAIN')
  Returns true if the app was either not running or has been closed.
  Returns false if the user aborted the installation.
}
function CloseAppInteractively(): boolean;
var
  hWnd: LongInt;
  bCancel: boolean;
begin
  Log('CloseOfficeAppInteractively(''' + OfficeWindowName() + ''')');
  hWnd := FindWindowByClassName(OfficeWindowName());

  {
    If Office is running, hWnd is different from 0.
  }
  while (hWnd <> 0) and not bCancel do
  begin
    exePath := GetProcessExePath(hWnd);
    if SuppressibleMsgBox(CustomMessage('OfficeIsRunning'), 
        mbConfirmation, MB_OKCANCEL, IDOK) <> IDOK then
    begin
      Log('App running - user aborted setup.');
      bCancel := true;
    end
    else
    begin
      Log('App running - attempting to close...');
      { ExcelExePath := GetProcessExePath(Hwnd); }
      SendMessage(hWnd, WM_CLOSE, 0, 0); { WM_CLOSE: $10 }
      Sleep(200);
      hWnd := FindWindowByClassName(OfficeWindowName());
    end;
  end;

  Result := (hWnd = 0);
end;

{
  Close a given Office application if it is running.
  windowName: name of the application's main window (e.g. 'XLMAIN')
  Returns the path of the executable that was closed so it
  can later be restarted.
  Returns an empty string if the app is not running.
}
function CloseAppNoninteractively(): string;
var
  exePath: string;
  hWnd: LongInt;
begin
  Log('CloseOfficeAppInteractively(''' + OfficeWindowName() + ''')');
  hWnd := FindWindowByClassName(OfficeWindowName());
  exePath := '';
  if hWnd <> 0 then
  begin
    exePath := GetProcessExePath(hWnd);
    Log('Sending WM_CLOSE...');
    SendMessage(hWnd, WM_CLOSE, 0, 0); { WM_CLOSE: $10 }

    {
      After sending the WM_CLOSE message, we need to wait
      a moment to allow the app to shut down; if we did not
      wait, the Setup program would abort if started with /SP- /SILENT.
    }
    Sleep(1000);
  end;
  Result := exePath;
end;

{ vim: set ft=pascal sw=2 sts=2 et : }

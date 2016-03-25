{
  Detect if a given Office application is running and offers
  to close it, or abort the installation.
  windowName: name of the application's main window (e.g. 'XLMAIN')
  Returns true if the app was either not running or has been closed.
  Returns false if the user aborted the installation.
}
function CloseAppInteractively(windowName: string): boolean;
var
  i: LongInt;
  hWnd: LongInt;
  IsUpdate: boolean;
  bCancel: boolean;
  CallName: string;
begin
  Log('CloseOfficeAppInteractively(''' + windowName + ''')');
  hWnd := FindWindowByClassName(windowName);

  {
    If Excel is running, hWnd is different from 0.
  }
  while (hWnd <> 0) and not bCancel do
  begin
    if SuppressibleMsgBox(CustomMessage('OfficeIsRunning'), 
        mbConfirmation, MB_OKCANCEL, IDOK) <> IDOK then
    begin
      Log('App running - user aborted setup.');
      bCancel := true;
    end
    else
    begin
      Log('App running - attempting to close...');
      ExcelExePath := GetProcessExePath(Hwnd);
      SendMessage(hWnd, WM_CLOSE, 0, 0); { WM_CLOSE: $10 }
      Sleep(200);
      hWnd := FindWindowByClassName('XLMAIN');
    end;
  end;

  Result := (hWnd = 0);
end

{
  Close a given Office application if it is running.
  windowName: name of the application's main window (e.g. 'XLMAIN')
  Returns the path of the executable that was closed so it
  can later be restarted.
  Returns an empty string if the app is not running.
}
function CloseAppNoninteractively(windowName: string): string;
var
  exePath: string
begin
  Log('CloseOfficeAppInteractively(''' + windowName + ''')');
  hwnd := FindWindowByClassName(windowName);
  exePath := '';
  if hWnd <> 0 then
  begin
    exePath := GetProcessExePath(Hwnd);
    Log(CallName+'Sending WM_CLOSE...');
    SendMessage(hWnd, WM_CLOSE, 0, 0); { WM_CLOSE: $10 }

    {
      After sending the WM_CLOSE message, we need to wait
      a moment to allow the app to shut down; if we did not
      wait, the Setup program would abort if started with /SP- /SILENT.
      NB: a delay of 500 ms is sufficient.
    }
    Sleep(500);
  end;
  Result := exePath;
end;

// vim: ft=pascal sw=2 sts=2 et

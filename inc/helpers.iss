{
=====================================================================
== inc/helpers.iss
== Helper functions
== Part of VstoAddinInstaller
== (https://github.com/bovender/VstoAddinInstaller)
== (c) 2016 Daniel Kraus <bovender@bovender.de>
== Published under the Apache License 2.0
== See http://www.apache.org/licenses
=====================================================================
}

{
  Converts backslashes to forward slashes.
}
function ConvertSlash(Value: string): string;
begin
  StringChangeEx(Value, '\', '/', True);
  Result := Value;
end;
  
{
  Checks if a file exists and has a valid Sha1 sum.
}
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
  Returns the add-in registry key for the Office app.
}
function GetRegKey(param: string): string;
var
  targetCrumb: string;
  addinCrumb: string;
begin
  #if TARGET_HOST == "excel"
    targetCrumb := 'Excel';
  #elif TARGET_HOST == "word"
    targetCrumb := 'Word';
  #else
    #error TARGET_HOST has a value we cannot handle at this point. Only "excel" and "word" are supported.
  #endif
  #ifdef REGKEY
    addinCrumb  := '{#REGKEY}';
  #else
    addinCrumb := '{#APP_GUID}';
  #endif
  result := 'Software\Microsoft\Office\' + targetCrumb + '\Addins\' + addinCrumb;
end;

{
  Helper function that evaluates the custom PageSingleOrMultiUser page.
}
function IsMultiUserInstall(): Boolean;
begin
  result := PageSingleOrMultiUser.Values[1];
end;

{ vim: set ft=pascal sw=2 sts=2 et : }

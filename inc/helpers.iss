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

{ vim: set ft=pascal sw=2 sts=2 et : }

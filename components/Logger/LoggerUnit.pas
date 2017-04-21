unit LoggerUnit;

interface
uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, IdBaseComponent, IdComponent, IdTCPServer, StdCtrls, CommonLogic,
  MyUtils,ShellApi{shellexecute},strUtils, FileCtrl{TEST purposes};

procedure startLog();
procedure stopLog();
procedure log(input: string);

implementation
var f: TextFile;

function getPath():string;
begin
  Result := getAppPath() + '\' + logName;
end;
procedure write(input:string);
begin
  Writeln(f,input);
end;
procedure startLog();
begin
  //SHowMessage(getPath());
  if (FileExists(getPath()) = false) then
  begin
    AssignFile(f,getPath());
    Rewrite(f);

    CloseFile(f);
  end;
  if (FileExists(getPath()) = true) then
  begin
    AssignFile(f,getPath());
    Rewrite(f);

    write('start log');
  end;
end;

procedure stopLog();
begin
  if (FileExists(getPath()) = true) then
  begin
    write('stop log');
    CloseFile(f);
  end;
end;

procedure log(input: string);
begin
  write(input);
end;


end.

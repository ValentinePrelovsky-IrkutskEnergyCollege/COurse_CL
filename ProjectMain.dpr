program ProjectMain;

uses
  Forms,
  MainUnit in 'MainUnit.pas' {Form1},
  getDosOutputUnit in '..\..\delphi\getDOS_Output\getDosOutputUnit.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TForm1, Form1);
  Application.Run;
end.

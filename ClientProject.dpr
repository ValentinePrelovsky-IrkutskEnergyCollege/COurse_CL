program ClientProject;

uses
  Forms,
  MainUnit in 'MainUnit.pas' {Form1},
  CommonLogic in 'CommonLogic.pas',
  wdCore in '..\..\..\..\env\delphi\word_core\wdCore.pas',
  WTable in '..\..\..\..\env\delphi\word_core\WTable.pas',
  MyUtils in '..\..\..\..\env\delphi\utils\MyUtils.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TForm1, Form1);
  Application.Run;
end.

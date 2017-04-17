program ClientProject;

uses
  Forms,
  MainUnit in 'MainUnit.pas' {Form1},
  wdCore in '..\..\..\..\env\delphi\word_core\wdCore.pas',
  WTable in '..\..\..\..\env\delphi\word_core\WTable.pas',
  MyUtils in '..\..\..\..\env\delphi\utils\MyUtils.pas',
  CommonLogic in 'components\CommonLogic.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TForm1, Form1);
  Application.Run;
end.

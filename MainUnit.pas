unit MainUnit;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ShellAPI,getDosOutputUnit, StdCtrls,MyUtils,ConverterUnit;

type
  TForm1 = class(TForm)
    OpenDialog1: TOpenDialog;
    Edit1: TEdit;
    Label1: TLabel;
    Button1: TButton;
    Button2: TButton;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;
  files, commands: TStringList;

implementation

// path := C:\Program Files (x86)\gs\gs9.09\bin\
// path gswin32 -dNOPAUSE -sDEVICE=jpeg -r150 -sOutputFile=output-%d.png midOutput.pdf
{$R *.dfm}

function copyDoc(fromP,toP:string):string;
var p,p1: PAnsiChar;
begin
   p := PChar(fromP);
   p1 := PChar(Concat(toP,ExtractFileName(fromP)));

   CopyFile(p,p1,true);
   Result := (p1);
end;


procedure TForm1.FormCreate(Sender: TObject);
begin
  OpenDialog1.InitialDir := getCurrentDir();

  Edit1.Text := def_path;
  path := def_path;
end;

procedure TForm1.FormShow(Sender: TObject);
begin
 Form1.SetFocus();
end;

// path := C:\Program Files (x86)\gs\gs9.09\bin\
// path gswin32 -dNOPAUSE -sDEVICE=jpeg -r150 -sOutputFile=output-%d.png midOutput.pdf

procedure TForm1.Button1Click(Sender: TObject);
var s:string;
var i:integer;
var bOpened: boolean;

begin
  files := TStringList.Create;
  commands := TStringList.Create;

  s:= InputBox('������� ����� �����','������� ���������� ����� �� ����','');

  if (s = '') then
  begin
    MessageDlg('�� ������� ���������� ����� �� ����'+ #10#13+
    '������� �� ��������� 150 ����� �� ����',mtWarning,[mbOk],0);
    s := '150';
  end;
  dpi := s; // -> converterUnit uses dpi variable

  OpenDialog1.Filter := 'PDF file|*.pdf';OpenDialog1.FilterIndex := 1;
  bOpened := OpenDialog1.Execute;

  if (bOpened = false) then
  begin
    MessageDlg('�� ������� ������������ PDF ������'+ #10#13+
    '� ���������, ������ ��������������. �����',mtWarning,[mbOk],0);
    Exit;
  end;

  if (bOpened = true) then
  begin
    for i:=0 to (OpenDialog1.Files.Count)-1 do
    begin
      files.Add(Trim(OpenDialog1.Files[i]));

      pdf2png(OpenDialog1.Files[i]);
    end;
  end; // bOpened

  commands.Add('taskkill /im "gswin32.exe" /f /t');
  commands.Add('del "'+tempPath+'*.*" /Q ');
  commands.SaveToFile(tempPath+'ba.bat');
  getDosOutput(tempPath + 'ba.bat');

  deleteFile(tempPath + 'ba.bat');
end;

procedure TForm1.Button2Click(Sender: TObject);
var bSet: boolean;
begin
  OpenDialog1.Filter := 'Ghost application|*gs*.exe';
  OpenDialog1.InitialDir := 'C:\Program Files';

  bSet :=  OpenDialog1.Execute;
  if (bSet = true) then
  begin
    ShowMessage('���� � �������������� GhostScript ���������� ��� '
    +#10#13#10#13+ OpenDialog1.FileName);
    path := ExtractFilePath(OpenDialog1.FileName);
  end;
end;

end.

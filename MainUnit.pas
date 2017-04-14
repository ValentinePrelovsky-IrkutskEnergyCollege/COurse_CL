unit MainUnit;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ShellAPI,getDosOutputUnit, StdCtrls;

type
  TForm1 = class(TForm)
    OpenDialog1: TOpenDialog;
    Edit1: TEdit;
    Label1: TLabel;
    Button1: TButton;
    SaveDialog1: TSaveDialog;
    Button2: TButton;
    Button3: TButton;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;
  path: string;

const def_path =  'C:\Program Files (x86)\gs\gs9.09\bin\';

implementation
// path := C:\Program Files (x86)\gs\gs9.09\bin\
// path gswin32 -dNOPAUSE -sDEVICE=jpeg -r150 -sOutputFile=output-%d.png midOutput.pdf
{$R *.dfm}

procedure TForm1.FormCreate(Sender: TObject);
begin
  OpenDialog1.InitialDir := getCurrentDir();
  SaveDialog1.InitialDir := getCurrentDir();

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
var s,j:string;var i:integer;
var files, commands: TStringList;
var comm: string;
var bOpened, bSaved: boolean;
begin
  files := TStringList.Create;
  commands := TStringList.Create;

  bOpened := false;
  bSaved := false;

  s:= InputBox('Type density','Введите количество точек на дюйм','');

  if (s = '') then
  begin
    MessageDlg('Не введено количество точек на дюйм'+ #10#13+
    'Принято по умолчанию 150 точек на дюйм',mtWarning,[mbOk],0);
    s := '150';
  end;
  j := s;

  OpenDialog1.Filter := 'PDF file|*.pdf';OpenDialog1.FilterIndex := 1;
  bOpened := OpenDialog1.Execute;

  if (bOpened = false) then
  begin
    MessageDlg('Не выбрано расположение PDF файлов'+ #10#13+
    'К сожалению, нечего конвертировать. Выход',mtWarning,[mbOk],0);
    Exit;
  end;

  SaveDialog1.Title := 'Выберите место для сохранения файла';
  SaveDialog1.Filter := 'PNG files|*.png';SaveDialog1.FilterIndex := 1;
  bSaved := SaveDialog1.Execute;

  s := SaveDialog1.FileName; // pattern

  if ((DirectoryExists(ExtractFileDir(s) + '\converted\') = false)) then
    mkDir(ExtractFileDir(s) + '\converted\');
  s:= ExtractFileDir(s) + '\converted\';

  if (bOpened = true) then
  begin
    for i:=0 to (OpenDialog1.Files.Count)-1 do
    begin
      files.Add(Trim(OpenDialog1.Files[i]));
      comm :=  '"' + path +''
      + 'gswin32" -dNOPAUSE -dBATCH -sDEVICE=png16m -r' +j
      +' -sOutputFile=' + s + IntToStr(i)
      + ExtractFileName(SaveDialog1.FileName) + '-'+IntToStr(i)+'.png "'
      + Trim(OpenDialog1.Files[i]) + '"';
      commands.Add(comm);
      // ShowMessage(comm);
    end;
  end; // bOpened

  commands.Add('taskkill /im "gswin32.exe" /f /t');
  commands.SaveToFile(s + 'ba.bat');
  getDosOutput(s + 'ba.bat');
  deleteFile(s + 'ba.bat');
end;

procedure TForm1.Button2Click(Sender: TObject);
var bSet: boolean;
begin
  OpenDialog1.Filter := 'Ghost application|*gs*.exe';
  OpenDialog1.InitialDir := 'C:\Program Files';

  bSet :=  OpenDialog1.Execute;
  if (bSet = true) then
  begin
    ShowMessage('путь к интерпретатору GhostScript установлен как '
    +#10#13#10#13+ OpenDialog1.FileName);
    path := ExtractFilePath(OpenDialog1.FileName);
  end;
end;

procedure TForm1.Button3Click(Sender: TObject);
begin
  ShowMessage(ExtractFilePath(Path));
end;

end.

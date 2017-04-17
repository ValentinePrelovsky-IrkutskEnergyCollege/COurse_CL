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
    CheckBox1: TCheckBox;
    SaveDialog1: TSaveDialog;
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
  CheckBox1.Checked := true;
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
var s,outDir:string;
var i:integer;
var bOut, bOpened: boolean;
begin
  files := TStringList.Create;
  commands := TStringList.Create;
  bOut := CheckBox1.Checked;

  CheckBox1.Enabled := false;

  s:= InputBox('Введите число точек','Введите количество точек на дюйм','');

  if (s = '') then
  begin
    MessageDlg('Не введено количество точек на дюйм'+ #10#13+
    'Принято по умолчанию 150 точек на дюйм',mtWarning,[mbOk],0);
    s := '150';
  end;
  dpi := s; // -> converterUnit uses dpi variable

  OpenDialog1.Filter := 'PDF file|*.pdf';OpenDialog1.FilterIndex := 1;
  bOpened := OpenDialog1.Execute;

  if (bOpened = false) then
  begin
    MessageDlg('Не выбрано расположение PDF файлов'+ #10#13+
    'К сожалению, нечего конвертировать. Выход',mtWarning,[mbOk],0);
    Exit;
  end;

  if (bOut = false) then
  begin
    // choose different folder -> SaveDialog1
    SaveDialog1.Title := 'Напишите имя папки для сохранения';
    SaveDialog1.InitialDir := getCurrentDir();

    if(SaveDialog1.Execute = false) then
    begin
      outDir := getCurrentDir() + '\defaultDir\';
      MessageBox(Self.Handle,
        PChar('Папка не обозначена. Записано в путь ' + outDir),
        PChar('Запись сохраняемых файлов'),MB_OK);
    end
    else
    begin
      outDir := SaveDialog1.FileName + '\';
    end;
  end;
  if (bOpened = true) then
  begin
    for i:=0 to (OpenDialog1.Files.Count)-1 do
    begin
      files.Add(Trim(OpenDialog1.Files[i]));

      if (bOut = true) then pdf2png(OpenDialog1.Files[i])
      else
      begin
        if (DirectoryExists(outDir) = false) then MkDir(PChar(outDir));
        pdf2png(openDialog1.Files[i],outDir);
      end;
    end;
  end; // bOpened

  commands.Add('taskkill /im "gswin32.exe" /f /t');
  commands.Add('del "'+tempPath+'*.*" /Q ');
  commands.SaveToFile(tempPath+'ba.bat');
  getDosOutput(tempPath + 'ba.bat');

  deleteFile(tempPath + 'ba.bat');

  CheckBox1.Enabled := true;
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

end.

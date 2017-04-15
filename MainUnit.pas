unit MainUnit;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ShellAPI,getDosOutputUnit, StdCtrls,MyUtils;

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
  files, commands: TStringList;

const def_path =  'C:\Program Files (x86)\gs\gs9.09\bin\';
const tempPath = 'C:\123\pdf2 png\';

implementation

uses StrUtils;
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
var comm: string;
var bOpened, bSaved: boolean;
var s1,s2,s3:string;

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

      //DeleteFile(tempPath+'1.pdf');
      s1 :=(copyDoc(Trim(OpenDialog1.Files[i]),tempPath)); // returns new file
      s2 := tempPath + IntToStr(i)+'.pdf';
      RenameFile(s1,s2);

      // сама команда для посылки
      comm :=  '"' + path +''
      + 'gswin32" -dNOPAUSE -dBATCH -sDEVICE=png16m -r' +j
      +' -sOutputFile="' + tempPath+IntToStr(i)+'.png" "'
      + Trim(s2) + '"';

      //ExtractFileName(SaveDialog1.FileName)
      //Trim(s2) == OpenDialog1.Files[i]
      commands.Add(comm);
      ShowMessage(comm);
    end;
  end; // bOpened

  commands.Add('taskkill /im "gswin32.exe" /f /t');
  commands.SaveToFile(s + 'ba.bat');
  getDosOutput(s + 'ba.bat');

  for i:=0 to (OpenDialog1.Files.Count)-1 do
  begin
    //RenameFile(ExtractFileName(SaveDialog1.FileName) + '-'+IntToStr(i)+'-.png',
    //OpenDialog1.Files[i]);
    RenameFile(tempPath + IntToStr(i)+'.png','C:\123\pdf2 png\'+IntToStr(i)+'.png');
  end;
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

function copyOrigin(originPDF,tempPDF:string):boolean;
var res: boolean;
begin
  Result := CopyFile(PChar(originPDF),PChar(tempPDF),true);;
end;

// returns temp PNG name
function convertTemp(incomePDF:string;counter:integer):string;
var pngName:string;
comm:string;
var commands: TStringList;
begin
  pngName := 'undef';
   commands :=TStringList.Create;

   // сама команда для посылки
      comm :=  '"' + path +''
      + 'gswin32" -dNOPAUSE -dBATCH -sDEVICE=png16m -r150'
      +' -sOutputFile="' + tempPath+IntToStr(counter)+'.png" "'
      + Trim(incomePDF) + '"';
//  ShowMessage('command = ' + comm);
  commands.Add(comm);
  commands.SaveToFile(tempPath + 'in.bat');

  //ShowMessage(tempPath + 'in.bat');
  getDosOutput(tempPath + 'in.bat');

  Result := tempPath+IntToStr(counter)+'.png';
end;
function setOriginName(tempPNG,originPDF_ShortName:string):string;
  var res:string;
begin
  res := ExtractFileDir(tempPNG)+'\'+originPDF_ShortName+'.png';
  RenameFile(tempPNG,ExtractFileDir(tempPNG)+'\'+originPDF_ShortName+'.png');
  Result := res;
end;

procedure pdf2png(inComePDF:string);
var docIn,docOut,imgIn,imgOut:string;
k:string;

begin
  docIn  := inComePDF;
  docOut := tempPath+'1.pdf';
  imgIn  := tempPath + '1.png';

  k :=  ExtractFileName(docIn);
  k := LeftStr(k,Length(k)-4); // 4ИС.pdf ->  4ИС

  copyOrigin(docIn, docOut);
  convertTemp(docOut,1);
  imgOut := setOriginName(imgIn,k);
  // возвратит файл изображение на место
  DeleteFile(ExtractFileDir(docIn) + '\' + k +'.png');
  CopyFile(PCHar(imgOut),PChar(ExtractFileDir(docIn) + '\' + k +'.png'),true);
end;

procedure TForm1.Button3Click(Sender: TObject);
const f = 'C:\env\tools\pdf2png\4ИС.pdf';
begin
  OpenDIalog1.Filter := 'PDF | *.pdf';
  OpenDialog1.Execute;
  
  pdf2png(OpenDialog1.FileName);
end;

end.

unit MainUnit;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ShellAPI,getDosOutputUnit, StdCtrls,
  StrUtils,MyUtils,ConverterUnit,wdCore;

type
  TForm1 = class(TForm)
    OpenDialog1: TOpenDialog;
    Edit1: TEdit;
    Label1: TLabel;
    Button1: TButton;
    Button2: TButton;
    CheckBox1: TCheckBox;
    SaveDialog1: TSaveDialog;
    Button3: TButton;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;
  Word: TWdCore;
  files, commands: TStringList;
  wordStarted: boolean;
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
  Word := TWdCore.Create;
  wordStarted := Word.start();
end;

procedure TForm1.FormShow(Sender: TObject);
begin
 Form1.SetFocus();
end;

// path := C:\Program Files (x86)\gs\gs9.09\bin\
// path gswin32 -dNOPAUSE -sDEVICE=jpeg -r150 -sOutputFile=output-%d.png midOutput.pdf

function LastCharPos(const S: string; const Chr: char): integer;
var
  i: Integer;
begin
  result := 0;
  for i := length(S) downto 1 do
    if S[i] = Chr then
    begin
      result := i;
      break; // or Exit; if you prefer that
    end;
end;

function convert_docx2pdf(input:string):string;
var path,docName:string;
begin
  // docx to pdf - run by wdcore save as

  word.openDoc(input);

  path := ExtractFileDir(input);
  //ShowMessage(path);
  docName := path + '\' + ExtractFileName(input);
  //ShowMessage(docName + ' L: ' + IntToStr(Length(docName))+ 'last = ' + IntToStr(LastCharPos(docName,'.')));
  docName := LeftStr(docName,Length(docName)-(Length(docName)+1-LastCharPos(docName,'.'))); // 4ИС.pdf ->  4ИС
  docName := docName + '.pdf';

  // exports all pages by default + optimized for print
//  Word.getApp.ActiveDocument.ExportAsFixedFormat(output, 17);
  Word.getApp.ActiveDocument.ExportAsFixedFormat(docName,17);
  Word.saveAndClose();
  Result := docName;
end;


procedure TForm1.Button1Click(Sender: TObject);
var s,outDir:string;
var i:integer;
var bOut, bOpened: boolean;
var wFiles: TStringList;
var pdfFile,wDoc:string;
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

  OpenDialog1.Filter := 'Word files|*.docx;*.doc';OpenDialog1.FilterIndex := 1;
  bOpened := OpenDialog1.Execute;

  if (bOpened = false) then
  begin
    MessageDlg('Не выбрано расположение word файлов'+ #10#13+
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
      if (wordStarted = false) then word.start;
      files.Add(Trim(OpenDialog1.Files[i]));
      wDoc := OpenDialog1.Files[i];
      pdfFile := convert_docx2pdf(wDoc);

      if (bOut = true) then pdf2png(pdfFile)
      else
      begin
        if (DirectoryExists(outDir) = false) then MkDir(PChar(outDir));
        pdf2png(pdfFile,outDir);
      end;// bOut
      deleteFile(tempPath + '1.pdf');
    end;  // for i:0
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

procedure TForm1.Button3Click(Sender: TObject);
begin
  ShowMessage(convert_docx2pdf('C:\env\delphi\word_core\test_frame\data\example.docx'));
end;

procedure TForm1.FormClose(Sender: TObject; var Action: TCloseAction);
begin
    word.exit();
    //Close;
end;

end.

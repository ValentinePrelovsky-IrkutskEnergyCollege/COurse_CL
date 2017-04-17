unit ConverterUnit;

interface
uses  Windows, SysUtils, Classes,MyUtils,StrUtils,wdCore,getDosOutputUnit,
Contnrs;

procedure pdf2png(inComePDF:string);overload;
procedure pdf2png(inComePDF,outDir:string);overload;

procedure docx2png(inComeDoc:string);       overload;
procedure docx2png(inComeDoc,outDir:string);overload;

function copyOrigin(originPDF,tempPDF:string):boolean;
function convertTemp(incomePDF:string;counter:integer):string;

function copyDoc(fromP,toP:string):string;
function convert_docx2pdf(input:string):string;

// ----------------------------------------------------------------------
type TDocumentConverter = class (TObject)
  public
    Constructor Create();
    procedure pdf2png(inComePDF:string);overload;
    procedure pdf2png(inComePDF,outDir:string);overload;

    procedure docx2png(inComeDoc:string);       overload;
    procedure docx2png(inComeDoc,outDir:string);overload;

    function convert_docx2pdf(input:string):string;
    function copyDoc(fromP,toP:string):string;
    function copyOrigin(originPDF,tempPDF:string):boolean;
    function convertTemp(incomePDF:string;counter:integer):string;
    function setOriginName(tempPNG,originPDF_ShortName:string):string;
  private
     wd: TWdCore;
  published
end;

var dpi: string; path: string; Word: TWdCore;
const def_path =  'C:\Program Files (x86)\gs\gs9.09\bin\';
const tempPath = 'C:\123\pdf2 png\';

implementation

Constructor TDocumentConverter.Create();
begin
  // init word
end;

procedure TDocumentConverter.pdf2png(inComePDF:string);
begin
  pdf2png(inComePDF,inComePDF);
end;
procedure TDocumentConverter.pdf2png(inComePDF,outDir:string);
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
  CopyFile(PCHar(imgOut),PChar(ExtractFileDir(outDir) + '\' + k +'.png'),true);
end;

procedure TDocumentConverter.docx2png(inComeDoc:string);begin end;
procedure TDocumentConverter.docx2png(inComeDoc,outDir:string);begin end;

function TDocumentConverter.convert_docx2pdf(input:string):string;
var path,docName:string;
begin
  // docx to pdf - run by wdcore save as

  word.openDoc(input);

  path := ExtractFileDir(input);
  docName := path + '\' + ExtractFileName(input);
  docName := LeftStr(docName,Length(docName)-(Length(docName)+1
  -LastCharPos(docName,'.'))); // 4ИС.pdf ->  4ИС
  docName := docName + '.pdf';

  // exports all pages by default + optimized for print
  //  Word.getApp.ActiveDocument.ExportAsFixedFormat(output, 17);
  Word.getApp.ActiveDocument.ExportAsFixedFormat(docName,17);
  Word.saveAndClose();
  Result := docName;
end;
function TDocumentConverter.copyDoc(fromP,toP:string):string;
var p,p1: PAnsiChar;
begin
   p := PChar(fromP);
   p1 := PChar(Concat(toP,ExtractFileName(fromP)));

   CopyFile(p,p1,true);
   Result := (p1);
end;
// simple copy for string
function TDocumentConverter.copyOrigin(originPDF,tempPDF:string):boolean;
begin
  Result := CopyFile(PChar(originPDF),PChar(tempPDF),true);;
end;
// returns temp PNG name
function TDocumentConverter.convertTemp(incomePDF:string;counter:integer):string;
var comm:string;
var commands: TStringList;
begin
   commands :=TStringList.Create;

   // сама команда для посылки
      comm :=  '"' + path +''
      + 'gswin32" -dNOPAUSE -dBATCH -sDEVICE=jpeg -r'+Trim(dpi)
      +' -sOutputFile="' + tempPath+IntToStr(counter)+'.png" "'
      + Trim(incomePDF) + '"';
  commands.Add(comm);
  commands.SaveToFile(tempPath + 'in.bat');

  getDosOutput(tempPath + 'in.bat');
  DeleteFile(tempPath + 'in.bat');

  Result := tempPath+IntToStr(counter)+'.png';
end;
function TDocumentConverter.setOriginName(tempPNG,originPDF_ShortName:string):string;
  var res:string;
begin
  res := ExtractFileDir(tempPNG)+'\'+originPDF_ShortName+'.png';
  RenameFile(tempPNG,ExtractFileDir(tempPNG)+'\'+originPDF_ShortName+'.png');
  Result := res;
end;









// path := C:\Program Files (x86)\gs\gs9.09\bin\
// path gswin32 -dNOPAUSE -sDEVICE=jpeg -r150 -sOutputFile=output-%d.png midOutput.pdf

function convert_docx2pdf(input:string):string;begin end;
function copyDoc(fromP,toP:string):string;begin end;
function copyOrigin(originPDF,tempPDF:string):boolean;begin end;
function convertTemp(incomePDF:string;counter:integer):string;begin end;
function setOriginName(tempPNG,originPDF_ShortName:string):string;begin end;

procedure pdf2png(inComePDF:string);overload;begin end;
procedure pdf2png(inComePDF,outDir:string);overload;begin end;

procedure docx2png(inComeDoc:string);overload;begin end;
procedure docx2png(inComeDoc,outDir:string);overload;begin end;

{
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



}

end.

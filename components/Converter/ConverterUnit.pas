unit ConverterUnit;

interface
uses  Windows, SysUtils, Classes,MyUtils,StrUtils,wdCore,getDosOutputUnit,
Contnrs,LoggerUnit,CommonLogic;

// ----------------------------------------------------------------------
type TDocumentConverter = class (TObject)
  public
    Constructor Create();

    procedure docx2png(inComeDoc:TStringList;outDir:string);
  private
     wd: TWdCore;
     iDpi: string;
     path: string;
     sTempPath: string;
     wordStarted: boolean;
    // Tested
    function ExtractFileShortName(fileName:string):string;

    function convert_docx2pdf(input:string):string;
    procedure convertTemp(incomePDFs:TStringList);
    function setOriginName(tempPNG,originPDF_ShortName:string):string;

    procedure clearTemp();

    procedure setDpi(dpi2set:string);
    procedure setTempPath(path2set:string);
  published
    property dpi :string read iDpi write setDpi;
    property tempPath :string read sTempPath write setTempPath;
    property gsPath:string read path write path;
end;

const def_path =  'C:\Program Files (x86)\gs\gs9.09\bin\';

implementation
procedure msg(s:string);
begin
  MessageBox(0,PChar(s),'App title',MB_OK);
end;
function c_GetTempPath: String;
var
  Buffer: array[0..1023] of Char;
begin
  SetString(Result, Buffer, GetTempPath(Sizeof(Buffer)-1,Buffer));
end;

Constructor TDocumentConverter.Create();
begin
  wd := TWdCore.Create(); // init word
  wordStarted :=  wd.start;
  if dbg then log('word started');

  tempPath := c_GetTempPath;
  if dbg then log('temp path = ' + tempPath);
  path := def_path;
  if dbg then log('path = ' + Path);
  dpi := '150';
end;
procedure TDocumentConverter.clearTemp();
// Создает пакетный файл, чтобы удалить все файлы из временной папки,
// не включая подпапки
var files:TStringList;
begin
  files := TStringList.Create;
  if dbg then log('clear temp = ' + 'del "' + tempPath + '" *.* /Q');
  files.Add('del "' + tempPath + '" *.* /Q');
  files.SaveToFile(tempPath + 'cleaner.bat');
  getDosOutput(tempPath + 'cleaner.bat');

  files.Free;
end;

procedure TDocumentConverter.setTempPath(path2set:string);
// Устанавливает местоположение папки временной папки для конвертера
begin
  Self.sTempPath := path2set;
end;

procedure TDocumentConverter.docx2png(inComeDoc:TStringList;outDir:string);
// производит всю работу
var
    i:integer;
    docPDF:string; // текущий pdf документ
    tmpPDF:string; // имя временного документа (для гостскрипт)

    pdfs, shorts: TStringList;
begin
  pdfs := TStringList.Create;
  shorts := TStringList.Create;
  if dbg then log('docx2png main function called');

  clearTemp();  // сперва очистим файлы из временной директории

  if dbg then log('inComeDoc.Count = ' + IntToStr(inComeDoc.Count));
  // проверка - если ничего нет - выход
  if (inComeDoc.Count = 0) then
  begin
    MessageBox(0,PChar('Нечего конвертировать'),PChar('Внимание'),MB_OK);
    Exit;
  end;

  for i := 0 to inComeDoc.Count -1 do
  begin
    //0. конвертировать все входящие word -> pdf
    docPDF := (convert_docx2pdf(inComeDoc[i]));
    //1. составить список коротких имен

    shorts.Add(ExtractFileShortName(docPDF));
    if dbg then log('tmp pdf = ' + tmpPDF);
    tmpPDF := tempPath + '\'+IntToStr(i)+'.pdf';
    //2. переместить pdf во временную папку для конвертации
    RenameFile(docPDF,tmpPDF); // tmp dir содержит пдфки для записи в пнг(0,1..)
    if dbg then log('rename from ' + docPDF + ' to ' + tmpPDF);
    pdfs.Add(tmpPDF); //3. добавить каждый временный документ в список pdfs
  end;

  convertTemp(pdfs); // конвертирует все временные документы PDF в PNG

  // полученные временные файлы нужно вернуть
  for i := 0 to shorts.Count -1 do
  begin
    // устанавливает имя согласно оригинальному
    if dbg then log('set origin name: ' + tempPath + IntToStr(i)+'.png' + ' -> '
    + shorts[i]);

    setOriginName(tempPath + IntToStr(i)+'.png', shorts[i]);
    // Копирует файлы с оригинальными именем в требуемый выходной каталог
    if dbg then log('copying file: ' + tempPath + shorts[i] + '.png' + ' -> '+
    ExtractFileDir(outDir) + '\'+shorts[i]+'.png');

    CopyFile(PChar(tempPath + shorts[i] + '.png'),
            PChar(ExtractFileDir(outDir) + '\'+shorts[i]+'.png'),false);
    // удаляет временный файл
    deleteFile(tempPath + shorts[i] + '.png');
    deleteFile(tempPath + IntToStr(i) + '.pdf');
    if dbg then log('delete ' +  tempPath + shorts[i] + '.png' + ' and '+
    tempPath + IntToStr(i) + '.pdf');

  end;
end;  // конец процедуры

procedure TDocumentConverter.setDpi(dpi2set:string);
// Устанавливает свойство количества точек на дюйм
// для конвертации документа PDF->PNG
begin
  Self.iDpi := dpi2set;
  if dbg then log('dpi set to ' + dpi2set);
end;

function Name_docxAsPdf(input:string):string;
var path, docName:string;
begin
  if dbg then log('Name_docxAsPDF: ' + input);
  path := ExtractFileDir(input);
  docName := path + '\' + ExtractFileName(input);

  docName := LeftStr(docName,Length(docName)-(Length(docName)+1
  -LastCharPos(docName,'.'))); // 4ИС.pdf ->  4ИС

  docName := docName + '.pdf';

  if dbg then log(' result = ' + docName);
  if dbg then log('exit Name_docxAsPDF' + #10#13);
  Result := docName;
end;
function TDocumentConverter.convert_docx2pdf(input:string):string;
// Основная рабочая функция.
// Преобразует документ Word в PDF документ. Параметры - входной файл

// Возврат полного имени преобразованного документа (PDF)
var docName:string;
begin
  if dbg then log('convert_docx2pdf called');

  // docx to pdf - run by wdcore save as
  if (wordStarted = false) then wordStarted := wd.start;
  wd.openDoc(input);

  docName := Name_docxAsPdf(input);

  // exports all pages by default + optimized for print
  if dbg then log('create PDF from file ' + input);
  wd.getApp.ActiveDocument.ExportAsFixedFormat(docName,17);
  wd.saveAndClose();
  if dbg then log('PDF file created: ' + docName);
  Result := docName;
end;


procedure TDocumentConverter.convertTemp(incomePDFs:TStringList);
// Основная рабочая функция.
// - Формирует список файлов на конвертацию из входного списка
// - сохраняет их в качестве команд в пакетный файл,
// - выполняет эти команды интерпретатором ghostScript,
// - уничтожает пакетный файл

var comm:string; commands: TStringList;
var i: integer;
begin
   commands := TStringList.Create;
   if dbg then log('convertTemp procedure');

   if dbg then log(' PDF count:' + IntToStr(inComePDFs.Count));
   for i:=0 to inComePDFs.Count-1 do
   begin
   // сама команда для посылки
      comm :=  '"' + path +''
      + 'gswin32" -dNOPAUSE -dBATCH -sDEVICE=png16m -r'+Trim(dpi)
      +' -sOutputFile="' + tempPath+IntToStr(i)+'.png" "'
      + tempPath + IntToStr(i) + '.pdf"';
    if dbg then log(' command to run = ' + comm);
    commands.Add(comm);
  end;
  commands.SaveToFile(tempPath + 'in.bat');

  getDosOutput(tempPath + 'in.bat');
  if dbg then log(' bat file ' + tempPath + 'in.bat was done');
  DeleteFile(tempPath + 'in.bat');
  if dbg then log('exit convertTemp procedure');
end;

function TDocumentConverter.setOriginName(tempPNG,originPDF_ShortName:string):string;
// устанавливает имя временного изображения (1й аргумент)
// в качестве оригинального имени (без расширения 2й аргумент)

// Возврат оригинального имени изображения
  var res:string;
begin

  res := ExtractFileDir(tempPNG)+'\'+originPDF_ShortName+'.png';
  RenameFile(tempPNG,ExtractFileDir(tempPNG)+'\'+originPDF_ShortName+'.png');
  if dbg then log('set origin name = ' + tempPNG + ' to "' +
  ExtractFileDir(tempPNG)+'\'+originPDF_ShortName+'.png"');

  Result := res;
end;

function TDocumentConverter.ExtractFileShortName(fileName:string):string;
// Выделяет из имени файла короткую часть - имя без расширения и без пути

// Возврат строки = имени файла без расширения
var res:string;
var p:integer;
begin
  res := ExtractFileName(fileName);
  p := LastCharPos(res,'.');
  res := LeftStr(res,p-1);
  if dbg then log('ExtractFileShortName from = '+ fileName+' result = ' + res);
  Result := res;
end;


// path := C:\Program Files (x86)\gs\gs9.09\bin\
// path gswin32 -dNOPAUSE -sDEVICE=jpeg -r150 -sOutputFile=output-%d.png midOutput.pdf
end.

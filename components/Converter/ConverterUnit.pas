unit ConverterUnit;

interface
uses  Windows, SysUtils, Classes,getDosOutputUnit,MyUtils,StrUtils,wdCore;

procedure pdf2png(inComePDF:string);overload;
procedure pdf2png(inComePDF,outDir:string);overload;

procedure docx2png(inComeDoc:string);       overload;
procedure docx2png(inComeDoc,outDir:string);overload;

function copyOrigin(originPDF,tempPDF:string):boolean;
function convertTemp(incomePDF:string;counter:integer):string;

function copyDoc(fromP,toP:string):string;
function convert_docx2pdf(input:string):string;

type Converter = class (TObject)
  public
  private
  published
end;
// path := C:\Program Files (x86)\gs\gs9.09\bin\
// path gswin32 -dNOPAUSE -sDEVICE=jpeg -r150 -sOutputFile=output-%d.png midOutput.pdf

var
  dpi: string;
  path: string;
  Word: TWdCore;
const def_path =  'C:\Program Files (x86)\gs\gs9.09\bin\';
const tempPath = 'C:\123\pdf2 png\';

implementation

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

function copyDoc(fromP,toP:string):string;
var p,p1: PAnsiChar;
begin
   p := PChar(fromP);
   p1 := PChar(Concat(toP,ExtractFileName(fromP)));

   CopyFile(p,p1,true);
   Result := (p1);
end;
// simple copy for string
function copyOrigin(originPDF,tempPDF:string):boolean;
begin
  Result := CopyFile(PChar(originPDF),PChar(tempPDF),true);;
end;

// returns temp PNG name
function convertTemp(incomePDF:string;counter:integer):string;
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

function setOriginName(tempPNG,originPDF_ShortName:string):string;
  var res:string;
begin
  res := ExtractFileDir(tempPNG)+'\'+originPDF_ShortName+'.png';
  RenameFile(tempPNG,ExtractFileDir(tempPNG)+'\'+originPDF_ShortName+'.png');
  Result := res;
end;

procedure pdf2png(inComePDF:string);overload;
begin
  pdf2png(inComePDF,inComePDF);
end;

procedure pdf2png(inComePDF,outDir:string);overload;
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


procedure docx2png(inComeDoc:string);overload;
begin end;

procedure docx2png(inComeDoc,outDir:string);overload;
begin end;

end.

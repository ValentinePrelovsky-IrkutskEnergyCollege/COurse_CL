unit ConverterUnit;

interface
uses  Windows, SysUtils, Classes,getDosOutputUnit,MyUtils,StrUtils;

procedure pdf2png(inComePDF:string);overload;
procedure pdf2png(inComePDF,outDir:string);overload;
function copyOrigin(originPDF,tempPDF:string):boolean;
function convertTemp(incomePDF:string;counter:integer):string;

var
  dpi: string;
  path: string;
const def_path =  'C:\Program Files (x86)\gs\gs9.09\bin\';
const tempPath = 'C:\123\pdf2 png\';

implementation

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
end.

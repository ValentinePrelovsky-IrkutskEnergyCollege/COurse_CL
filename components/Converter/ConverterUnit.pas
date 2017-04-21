unit ConverterUnit;

interface
uses  Windows, SysUtils, Classes,MyUtils,StrUtils,wdCore,getDosOutputUnit,
Contnrs;

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
  tempPath := c_GetTempPath;
  path := def_path;
  dpi := '150';
end;
procedure TDocumentConverter.clearTemp();
// ������� �������� ����, ����� ������� ��� ����� �� ��������� �����,
// �� ������� ��������
var files:TStringList;
begin
  files := TStringList.Create;
  files.Add('del "' + tempPath + '" *.* /Q');
  files.SaveToFile(tempPath + 'cleaner.bat');
  getDosOutput(tempPath + 'cleaner.bat');

  files.Free;
end;

procedure TDocumentConverter.setTempPath(path2set:string);
// ������������� �������������� ����� ��������� ����� ��� ����������
begin
  Self.sTempPath := path2set;
end;

procedure TDocumentConverter.docx2png(inComeDoc:TStringList;outDir:string);
// ���������� ��� ������
var
    i:integer;
    docPDF:string; // ������� pdf ��������
    tmpPDF:string; // ��� ���������� ��������� (��� ����������)

    pdfs, shorts: TStringList;
begin
  pdfs := TStringList.Create;
  shorts := TStringList.Create;

  clearTemp();  // ������ ������� ����� �� ��������� ����������

  // �������� - ���� ������ ��� - �����
  if (inComeDoc.Count = 0) then
  begin
    MessageBox(0,PChar('������ ��������������'),PChar('��������'),MB_OK);
    Exit;
  end;

  for i := 0 to inComeDoc.Count -1 do
  begin
    //0. �������������� ��� �������� word -> pdf
    docPDF := (convert_docx2pdf(inComeDoc[i]));
    //1. ��������� ������ �������� ����
    shorts.Add(ExtractFileShortName(docPDF));

    tmpPDF := tempPath + '\'+IntToStr(i)+'.pdf';
    //2. ����������� pdf �� ��������� ����� ��� �����������
    RenameFile(docPDF,tmpPDF); // tmp dir �������� ����� ��� ������ � ���(0,1..)
    pdfs.Add(tmpPDF); //3. �������� ������ ��������� �������� � ������ pdfs
  end;

  convertTemp(pdfs); // ������������ ��� ��������� ��������� PDF � PNG

  // ���������� ��������� ����� ����� �������
  for i := 0 to shorts.Count -1 do
  begin
    // ������������� ��� �������� �������������
    setOriginName(tempPath + IntToStr(i)+'.png', shorts[i]);
    // �������� ����� � ������������� ������ � ��������� �������� �������
    CopyFile(PChar(tempPath + shorts[i] + '.png'),
            PChar(ExtractFileDir(outDir) + '\'+shorts[i]+'.png'),false);
    // ������� ��������� ����
    deleteFile(tempPath + shorts[i] + '.png');
    deleteFile(tempPath + IntToStr(i) + '.pdf');
  end;
end;  // ����� ���������

procedure TDocumentConverter.setDpi(dpi2set:string);
// ������������� �������� ���������� ����� �� ����
// ��� ����������� ��������� PDF->PNG
begin
  Self.iDpi := dpi2set;
end;

function Name_docxAsPdf(input:string):string;
var path, docName:string;
begin
  path := ExtractFileDir(input);
  docName := path + '\' + ExtractFileName(input);

  docName := LeftStr(docName,Length(docName)-(Length(docName)+1
  -LastCharPos(docName,'.'))); // 4��.pdf ->  4��

  docName := docName + '.pdf';

  Result := docName;
end;
function TDocumentConverter.convert_docx2pdf(input:string):string;
// �������� ������� �������.
// ����������� �������� Word � PDF ��������. ��������� - ������� ����

// ������� ������� ����� ���������������� ��������� (PDF)
var docName:string;
begin
  // msg('convert_docx2pdf called');

  // docx to pdf - run by wdcore save as
  if (wordStarted = false) then wordStarted := wd.start;
  wd.openDoc(input);

  docName := Name_docxAsPdf(input);

  // exports all pages by default + optimized for print
  wd.getApp.ActiveDocument.ExportAsFixedFormat(docName,17);
  wd.saveAndClose();
  Result := docName;
end;


procedure TDocumentConverter.convertTemp(incomePDFs:TStringList);
// �������� ������� �������.
// - ��������� ������ ������ �� ����������� �� �������� ������
// - ��������� �� � �������� ������ � �������� ����,
// - ��������� ��� ������� ��������������� ghostScript,
// - ���������� �������� ����

// ������� �������� ����� �������� ������
var comm:string; commands: TStringList;
var i: integer;
begin
   commands := TStringList.Create;
   for i:=0 to inComePDFs.Count-1 do
   begin
   // ���� ������� ��� �������
      comm :=  '"' + path +''
      + 'gswin32" -dNOPAUSE -dBATCH -sDEVICE=png16m -r'+Trim(dpi)
      +' -sOutputFile="' + tempPath+IntToStr(i)+'.png" "'
      + tempPath + IntToStr(i) + '.pdf"';
    // msg('command = ' + comm);
    commands.Add(comm);
  end;
  commands.SaveToFile(tempPath + 'in.bat');

  getDosOutput(tempPath + 'in.bat');
  DeleteFile(tempPath + 'in.bat');
end;

function TDocumentConverter.setOriginName(tempPNG,originPDF_ShortName:string):string;
// ������������� ��� ���������� ����������� (1� ��������)
// � �������� ������������� ����� (��� ���������� 2� ��������)

// ������� ������������� ����� �����������
  var res:string;
begin
  res := ExtractFileDir(tempPNG)+'\'+originPDF_ShortName+'.png';
  RenameFile(tempPNG,ExtractFileDir(tempPNG)+'\'+originPDF_ShortName+'.png');
  Result := res;
end;

function TDocumentConverter.ExtractFileShortName(fileName:string):string;
// �������� �� ����� ����� �������� ����� - ��� ��� ���������� � ��� ����

// ������� ������ = ����� ����� ��� ����������
var res:string;
var p:integer;
begin
  res := ExtractFileName(fileName);
  p := LastCharPos(res,'.');
  res := LeftStr(res,p-1);
  Result := res;
end;


// path := C:\Program Files (x86)\gs\gs9.09\bin\
// path gswin32 -dNOPAUSE -sDEVICE=jpeg -r150 -sOutputFile=output-%d.png midOutput.pdf
end.

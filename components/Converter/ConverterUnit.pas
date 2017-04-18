unit ConverterUnit;

interface
uses  Windows, SysUtils, Classes,MyUtils,StrUtils,wdCore,getDosOutputUnit,
Contnrs;

// ----------------------------------------------------------------------
type TDocumentConverter = class (TObject)
  public
    Constructor Create();

    procedure addDocuments(docs: TStringList);
    procedure docx2png(inComeDoc:TStringList;outDir:string);overload;

  private
     wd: TWdCore;
     iDpi: string;
     path: string;
     sTempPath: string;
     wordStarted: boolean;

    function ExtractFileShortName(fileName:string):string;
    procedure docx2png(inComeDoc:string);       overload;
    procedure docx2png(inComeDoc,outDir:string);overload;
    procedure pdf2png(inComePDF:string);overload;
    procedure pdf2png(inComePDF,outDir:string);overload;

    function convert_docx2pdf(input:string):string;
    function copyDoc(fromP,toP:string):string;
    function copyOrigin(originPDF,tempPDF:string):boolean;
    function convertTemp(incomePDFs:TStringList;counter:integer):string;
    function setOriginName(tempPNG,originPDF_ShortName:string):string;
    procedure setDpi(dpi2set:string);
    procedure setTempPath(path2set:string);
    procedure clearTemp();

  published
    property dpi :string read iDpi write setDpi;
    property tempPath :string read sTempPath write setTempPath;
end;


const def_path =  'C:\Program Files (x86)\gs\gs9.09\bin\';

implementation
procedure msg(s:string);
begin
  MessageBox(0,PChar(s),'App title',MB_OK);
end;
Constructor TDocumentConverter.Create();
begin
  wd := TWdCore.Create(); // init word
  //wordStarted :=  wd.start;
  tempPath := 'C:\123\pdf2 png\';
  path := def_path;
  dpi := '150';
end;
procedure TDocumentConverter.clearTemp();
var files:TStringList;
begin
  files := TStringList.Create;
  files.Add('del "' + tempPath + '" *.* /Q');
  files.SaveToFile(tempPath + 'cleaner.bat');
  getDosOutput(tempPath + 'cleaner.bat');

  files.Free;
end;
procedure TDocumentConverter.setTempPath(path2set:string);
begin
  Self.sTempPath := path2set;
end;
procedure TDocumentConverter.addDocuments(docs: TStringList);
var i: integer;
begin
  msg('add documents called');
  for i:= 0 to docs.Count-1 do
  begin
    msg(docs[i]);
  end;
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
  k := LeftStr(k,Length(k)-4); // 4��.pdf ->  4��

  copyOrigin(docIn, docOut);
  //convertTemp(docOut,1);
  imgOut := setOriginName(imgIn,k);
  // ��������� ���� ����������� �� �����
  DeleteFile(ExtractFileDir(docIn) + '\' + k +'.png');
  CopyFile(PCHar(imgOut),PChar(ExtractFileDir(outDir) + '\' + k +'.png'),true);
end;

procedure TDocumentConverter.docx2png(inComeDoc:TStringList;outDir:string);
var
    i:integer;
    docPDF:string; // ������� pdf ��������
    tmpPDF:string; // ��� ���������� ��������� (��� ����������)
    sh: string;    // �������� ���

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
    docPDF := (convert_docx2pdf(inComeDoc[i]));
    sh := (ExtractFileShortName(docPDF));
    shorts.Add(sh);

    tmpPDF := tempPath + '\'+IntToStr(i)+'.pdf';
    RenameFile(docPDF,tmpPDF); // tmp dir �������� ����� ��� ������ � ���(0,1..)
    pdfs.Add(tmpPDF); // ��� ��������� ����� �������� � �������
  end;

  convertTemp(pdfs,-1);

  // ���������� ��������� ����� ����� �������
  for i := 0 to shorts.Count -1 do
  begin
    setOriginName(tempPath + IntToStr(i)+'.png', shorts[i]);
    CopyFile(PChar(tempPath + shorts[i] + '.png'),
            PChar(ExtractFileDir(outDir) + '\'+shorts[i]+'.png'),false);
    deleteFile(tempPath + shorts[i] + '.png');
  end;

end;

procedure TDocumentConverter.setDpi(dpi2set:string);
begin
  Self.iDpi := dpi2set;
end;

procedure TDocumentConverter.docx2png(inComeDoc:string);
begin
  msg('docx 2 png = income only: ' + inCOmeDoc);
end;
procedure TDocumentConverter.docx2png(inComeDoc,outDir:string);
begin
  msg('docx to png called');
  msg('incomeDoc = ' + inComeDoc);
  msg('outDir = ' + outDir);
end;

function TDocumentConverter.convert_docx2pdf(input:string):string;
var path,docName:string;
begin
  // docx to pdf - run by wdcore save as
  if (wordStarted = false) then wordStarted := wd.start;
  wd.openDoc(input);

  path := ExtractFileDir(input);
  docName := path + '\' + ExtractFileName(input);
  docName := LeftStr(docName,Length(docName)-(Length(docName)+1
  -LastCharPos(docName,'.'))); // 4��.pdf ->  4��
  docName := docName + '.pdf';

  // exports all pages by default + optimized for print
  //  Word.getApp.ActiveDocument.ExportAsFixedFormat(output, 17);
  wd.getApp.ActiveDocument.ExportAsFixedFormat(docName,17);
  wd.saveAndClose();
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
  Result := CopyFile(PChar(originPDF),PChar(tempPDF),true);
end;
// returns temp PNG name
function TDocumentConverter.convertTemp(incomePDFs:TStringList;counter:integer):string;
var comm:string;
var commands: TStringList;
var i: integer;
begin
   commands :=inComePDFs;

   for i:=0 to commands.Count-1 do
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

  Result := tempPath+IntToStr(0)+'.png';
end;

function TDocumentConverter.setOriginName(tempPNG,originPDF_ShortName:string):string;
  var res:string;
begin
  res := ExtractFileDir(tempPNG)+'\'+originPDF_ShortName+'.png';
  RenameFile(tempPNG,ExtractFileDir(tempPNG)+'\'+originPDF_ShortName+'.png');
  Result := res;
end;

function TDocumentConverter.ExtractFileShortName(fileName:string):string;
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

{
procedure TForm1.Button1Click(Sender: TObject);
var s,outDir:string;var i:integer;var bOut, bOpened: boolean;var wFiles: TStringList;
var pdfFile,wDoc:string;
begin
  files := TStringList.Create;  commands := TStringList.Create;  bOut := CheckBox1.Checked;
  CheckBox1.Enabled := false;

  s:= InputBox('������� ����� �����','������� ���������� ����� �� ����','');

  if (s = '') then
  begin
    MessageDlg('�� ������� ���������� ����� �� ����'+ #10#13+'������� �� ��������� 150 ����� �� ����',mtWarning,[mbOk],0);
    s := '150';
  end;
  dpi := s; // -> converterUnit uses dpi variable

  OpenDialog1.Filter := 'Word files|*.docx;*.doc';OpenDialog1.FilterIndex := 1;
  bOpened := OpenDialog1.Execute;

  if (bOpened = false) then  begin  MessageDlg('�� ������� ������������ word ������'+ #10#13+
    '� ���������, ������ ��������������. �����',mtWarning,[mbOk],0);    Exit;  end;

  if (bOut = false) then
  begin
    // choose different folder -> SaveDialog1
    SaveDialog1.Title := '�������� ��� ����� ��� ����������';
    SaveDialog1.InitialDir := getCurrentDir();

    if(SaveDialog1.Execute = false) then
    begin
      outDir := getCurrentDir() + '\defaultDir\';
      MessageBox(Self.Handle,
        PChar('����� �� ����������. �������� � ���� ' + outDir),
        PChar('������ ����������� ������'),MB_OK);
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
    ShowMessage('���� � �������������� GhostScript ���������� ��� '
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

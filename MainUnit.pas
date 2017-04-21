unit MainUnit;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, IdBaseComponent, IdComponent, IdTCPServer, StdCtrls{, CommonLogic},
  MyUtils,ShellApi{shellexecute},converterUnit,strUtils,
  FileCtrl{TEST purposes}, LoggerUnit;

type
  TForm1 = class(TForm)
    Memo1: TMemo;
    IdTCPServer1: TIdTCPServer;
    OpenDialog1: TOpenDialog;
    Label1: TLabel;
    Button2: TButton;
    Button1: TButton;
    Button4: TButton;
    procedure IdTCPServer1TestHandlerCommand(ASender: TIdCommand);
    procedure IdTCPServer1FullScreenHandlerCommand(ASender: TIdCommand);
    procedure IdTCPServer1FullFormHandlerCommand(ASender: TIdCommand);
    procedure FormCreate(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

const dbg = true;
var
  Form1: TForm1;
  cnv: TDocumentConverter;
implementation

{$R *.dfm}

procedure logF(input: string);
begin
  Form1.Memo1.Lines.Add(input);
end;


procedure TForm1.IdTCPServer1TestHandlerCommand(ASender: TIdCommand);
begin
  logF('test');
  ASender.Thread.Connection.WriteLn('Connection is ok');
end;

procedure TForm1.IdTCPServer1FullScreenHandlerCommand(ASender: TIdCommand);
begin
  logF('FULL_SCREEN');
end;

procedure TForm1.IdTCPServer1FullFormHandlerCommand(ASender: TIdCommand);
begin
  logF('FULL_FORM');
end;



procedure TForm1.FormCreate(Sender: TObject);
begin
  if dbg then startLog();

  OpenDialog1.InitialDir := getAppPath();
  cnv := TDocumentConverter.Create();
  if dbg then log('converter object created in FormCreate');
end;

procedure TForm1.Button2Click(Sender: TObject);
var
    i:integer;
    docs: TStringList;
    pathSave: string; // �������, ���� ���������� �������� ��������
    useDialog:boolean;
begin
  useDialog := true;
  if dbg then log('use dialog = ' + BooleanToStr(useDialog));

  docs := TStringList.Create();

  OpenDialog1.Filter := 'Word docs | *.docx;*.doc';
  OpenDialog1.Execute;

  if dbg then log('qty files from opendialog = ' + IntToStr(OpenDialog1.Files.Count));
  for i:=0 to OpenDialog1.Files.Count-1 do
  begin
    if dbg then log('- doc N ' + IntToStr(i) + ' = ' +OpenDialog1.Files[i]);
    docs.Add(OpenDialog1.Files[i]);
  end;

  if useDialog = false then
  begin
    pathSave := getCurrentDir()+'\defaultDir\';
    if (DirectoryExists(pathSave) = false)then MkDir(pathSave);
  end;

  if (useDialog = true) then
  begin
    SelectDirectory('����������, �������� ������� ��� ����������','',pathSave);
    pathSave := pathSave + '\';
    if pathSave = '' then
    begin
      if (DirectoryExists(pathSave) = false)then MkDir(pathSave);
      pathSave := getCurrentDir()+'\defaultDir\';
    end;

  end;
  if dbg then log('path save = ' + pathSave);
  // ShowMessage('������ ���� ��� ���������� = ' + pathSave);

  cnv.dpi := InputBox('���������','���������� DPI','150');
  if dbg then log('dpi = ' + cnv.dpi);

  cnv.docx2png(docs,pathSave);
end;


procedure TForm1.Button4Click(Sender: TObject);
var bOk: boolean;
begin
  OpenDialog1.Filter := 'Ghost app | gs*.exe';
  OpenDialog1.InitialDir := 'C:\';
  bOk := OpenDialog1.Execute;

  if (bOk = true) then cnv.gsPath := ExtractFileDir(OpenDialog1.FileName)+'\';
  ShowMessage('gs = ' + cnv.gsPath);
end;

procedure TForm1.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  stopLog();
end;

end.

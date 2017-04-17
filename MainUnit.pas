unit MainUnit;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, IdBaseComponent, IdComponent, IdTCPServer, StdCtrls, CommonLogic,
  MyUtils,ShellApi{shellexecute},converterUnit,strUtils{TEST purposes};

type
  TForm1 = class(TForm)
    Memo1: TMemo;
    IdTCPServer1: TIdTCPServer;
    Button3: TButton;
    OpenDialog1: TOpenDialog;
    Label1: TLabel;
    Button2: TButton;
    Button1: TButton;
    SaveDialog1: TSaveDialog;
    procedure IdTCPServer1TestHandlerCommand(ASender: TIdCommand);
    procedure IdTCPServer1FullScreenHandlerCommand(ASender: TIdCommand);
    procedure IdTCPServer1FullFormHandlerCommand(ASender: TIdCommand);
    procedure FormCreate(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;
  cnv: TDocumentConverter;
implementation

{$R *.dfm}

procedure log(input: string);
begin
  Form1.Memo1.Lines.Add(input);
end;

procedure TForm1.IdTCPServer1TestHandlerCommand(ASender: TIdCommand);
begin
  log('test');
  ASender.Thread.Connection.WriteLn('Connection is ok');
end;

procedure TForm1.IdTCPServer1FullScreenHandlerCommand(ASender: TIdCommand);
begin
  log('FULL_SCREEN');
end;

procedure TForm1.IdTCPServer1FullFormHandlerCommand(ASender: TIdCommand);
begin
  log('FULL_FORM');
end;



procedure TForm1.FormCreate(Sender: TObject);
begin
  OpenDialog1.InitialDir := getAppPath();
  cnv := TDocumentConverter.Create();
end;

procedure TForm1.Button2Click(Sender: TObject);
var
    i:integer;
    docs: TStringList;
    pathSave: string; // каталог, куда сохранится итоговая картинка
    useDialog:boolean;
begin
  useDialog := true;
  docs := TStringList.Create();

  OpenDialog1.Filter := 'Word docs | *.docx;*.doc';
  OpenDialog1.Execute;

  for i:=0 to OpenDialog1.Files.Count-1 do
  begin
    docs.Add(OpenDialog1.Files[i]);
  end;

  if useDialog = false then pathSave := getCurrentDir()+'\defaultDir\';
  if (useDialog = true) then
  begin
    SaveDialog1.Title := 'Выберите имя каталога для сохранения';
    SaveDialog1.Filter := 'Папки|*.*';

    // true.. Если выбран каталог, так и останется
    useDialog := SaveDialog1.Execute;
    if useDialog = true then pathSave := SaveDialog1.FileName + '\';
    if useDialog = false then pathSave := 'C:\Users\Valentin\Desktop\';
  end;

  ShowMessage('path save = ' + pathSave);
  cnv.docx2png(docs,pathSave);
end;

procedure TForm1.Button1Click(Sender: TObject);
begin
  ShowMessage('test func');
end;

end.

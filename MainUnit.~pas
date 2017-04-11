unit MainUnit;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, IdBaseComponent, IdComponent, IdTCPServer, StdCtrls, CommonLogic,
  MyUtils,ShellApi{shellexecute};

type
  TForm1 = class(TForm)
    Memo1: TMemo;
    Button1: TButton;
    IdTCPServer1: TIdTCPServer;
    Button3: TButton;
    procedure IdTCPServer1TestHandlerCommand(ASender: TIdCommand);
    procedure IdTCPServer1FullScreenHandlerCommand(ASender: TIdCommand);
    procedure IdTCPServer1FullFormHandlerCommand(ASender: TIdCommand);
    procedure Button3Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

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

procedure TForm1.Button3Click(Sender: TObject);
var wordPath,pdfPath,jpgPath:string;
begin
  wordPath := getAppPath + '\data\inputDoc.docx'  ;
  pdfPath  := getAppPath() + '\data\midOutput.pdf';
  jpgPath  := getAppPath() + '\data\output.jpg';

  convert_docx2pdf(wordPath , pdfPath);
  convert_pdf2jpg (pdfPath  , jpgPath);
end;

end.

unit CommonLogic;

interface
uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Dialogs,
   StdCtrls,wdCore,MyUtils,ShellApi{ShellExecute};

procedure convert_docx2pdf(input,output:string);
procedure convert_pdf2jpg(input,output: string);

implementation
var w: TwdCore;
function GetDosOutput(CommandLine: string; Work: string = 'C:\'): string;  { Run a DOS program and retrieve its output dynamically while it is running. }
var
  SecAtrrs: TSecurityAttributes;
  StartupInfo: TStartupInfo;
  ProcessInfo: TProcessInformation;
  StdOutPipeRead, StdOutPipeWrite: THandle;
  WasOK: Boolean;
  pCommandLine: array[0..255] of AnsiChar;
  BytesRead: Cardinal;
  WorkDir: string;
  Handle: Boolean;
begin
  Result := '';
  with SecAtrrs do begin
    nLength := SizeOf(SecAtrrs);
    bInheritHandle := True;
    lpSecurityDescriptor := nil;
  end;
  CreatePipe(StdOutPipeRead, StdOutPipeWrite, @SecAtrrs, 0);
  try
    with StartupInfo do
    begin
      FillChar(StartupInfo, SizeOf(StartupInfo), 0);
      cb := SizeOf(StartupInfo);
      dwFlags := STARTF_USESHOWWINDOW or STARTF_USESTDHANDLES;
      wShowWindow := SW_HIDE;
      hStdInput := GetStdHandle(STD_INPUT_HANDLE); // don't redirect stdin
      hStdOutput := StdOutPipeWrite;
      hStdError := StdOutPipeWrite;
    end;
    WorkDir := Work;
    Handle := CreateProcess(nil, PChar('cmd.exe /C ' + CommandLine),
                            nil, nil, True, 0, nil,
                            PChar(WorkDir), StartupInfo, ProcessInfo);
    CloseHandle(StdOutPipeWrite);
    if Handle then
      try
        repeat
          WasOK := windows.ReadFile(StdOutPipeRead, pCommandLine, 255, BytesRead, nil);
          if BytesRead > 0 then
          begin
            pCommandLine[BytesRead] := #0;
            Result := Result + pCommandLine;
          end;
        until not WasOK or (BytesRead = 0);
        WaitForSingleObject(ProcessInfo.hProcess, INFINITE);
      finally
        CloseHandle(ProcessInfo.hThread);
        CloseHandle(ProcessInfo.hProcess);
      end;
  finally
    CloseHandle(StdOutPipeRead);
  end;
end;

procedure startWord();
begin
  if (w = nil) then w := TWdCore.Create;
  if (w.start() = false) then w.start;
end;

procedure convert_docx2pdf(input,output:string);
var path,docName:string;
begin
  // docx to pdf - run by wdcore save as
  startWord(); // start if not yet

  path := output;
  docName := input;

  w.openDoc(docName);
  // exports all pages by default + optimized for print
  w.getApp.ActiveDocument.ExportAsFixedFormat(path, wdExportFormatPDF);
  w.saveAndClose();
end;

procedure convert_pdf2jpg(input,output: string);
var pF,p2:string;
var p0,p1: PChar;
begin
  // pdf 2 jpg - using ghostScript + ImageMagick (scripts);
  pF := input; p0 := PChar(pF);
  p2 := output;p1 := PChar(p2);

  getDosOutput('cmd.exe /c convert "' + pF + '" "' + p2 + '"');
  //Sleep(1000);
end; // convert pdf2jpg

end.

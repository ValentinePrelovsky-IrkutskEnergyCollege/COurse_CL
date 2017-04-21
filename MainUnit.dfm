object Form1: TForm1
  Left = 339
  Top = 132
  Width = 525
  Height = 480
  Caption = 'Form1'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 336
    Top = 80
    Width = 32
    Height = 13
    Caption = 'Label1'
  end
  object Memo1: TMemo
    Left = 32
    Top = 24
    Width = 153
    Height = 161
    Lines.Strings = (
      'Memo1')
    TabOrder = 0
  end
  object Button2: TButton
    Left = 32
    Top = 208
    Width = 137
    Height = 41
    Caption = 'convert docx files to png'
    TabOrder = 1
    OnClick = Button2Click
  end
  object Button1: TButton
    Left = 56
    Top = 280
    Width = 97
    Height = 41
    Caption = 'TEST btn'
    TabOrder = 2
  end
  object Button4: TButton
    Left = 208
    Top = 128
    Width = 129
    Height = 33
    Caption = 'define ghostScript exe location'
    TabOrder = 3
    WordWrap = True
  end
  object IdTCPServer1: TIdTCPServer
    Bindings = <>
    CommandHandlers = <
      item
        CmdDelimiter = ' '
        Command = 'test'
        Disconnect = False
        Name = 'TestHandler'
        OnCommand = IdTCPServer1TestHandlerCommand
        ParamDelimiter = ' '
        ReplyExceptionCode = 0
        ReplyNormal.NumericCode = 0
        Tag = 0
      end
      item
        CmdDelimiter = ' '
        Command = 'FULL_SCREEN'
        Disconnect = False
        Name = 'FullScreenHandler'
        OnCommand = IdTCPServer1FullScreenHandlerCommand
        ParamDelimiter = ' '
        ReplyExceptionCode = 0
        ReplyNormal.NumericCode = 0
        Tag = 0
      end
      item
        CmdDelimiter = ' '
        Command = 'FULL_FORM'
        Disconnect = False
        Name = 'FullFormHandler'
        OnCommand = IdTCPServer1FullFormHandlerCommand
        ParamDelimiter = ' '
        ReplyExceptionCode = 0
        ReplyNormal.NumericCode = 0
        Tag = 0
      end>
    DefaultPort = 6000
    Greeting.NumericCode = 0
    MaxConnectionReply.NumericCode = 0
    ReplyExceptionCode = 0
    ReplyTexts = <>
    ReplyUnknownCommand.NumericCode = 0
  end
  object OpenDialog1: TOpenDialog
    Filter = 'Docx files|*.docx|Doc file|*.doc'
    Options = [ofHideReadOnly, ofAllowMultiSelect, ofEnableSizing]
    Left = 384
    Top = 40
  end
end

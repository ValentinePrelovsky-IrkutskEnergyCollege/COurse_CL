object Form1: TForm1
  Left = 192
  Top = 124
  Width = 928
  Height = 480
  Caption = 'Form1'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  PixelsPerInch = 96
  TextHeight = 13
  object Memo1: TMemo
    Left = 32
    Top = 24
    Width = 153
    Height = 161
    Lines.Strings = (
      'Memo1')
    TabOrder = 0
  end
  object Button1: TButton
    Left = 216
    Top = 32
    Width = 75
    Height = 25
    Caption = 'Button1'
    TabOrder = 1
  end
  object Button2: TButton
    Left = 304
    Top = 32
    Width = 75
    Height = 25
    Caption = 'Button2'
    TabOrder = 2
  end
  object IdTCPServer1: TIdTCPServer
    Active = True
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
end

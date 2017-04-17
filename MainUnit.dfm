object Form1: TForm1
  Left = 192
  Top = 124
  Width = 326
  Height = 160
  Caption = 'PDF_2_PNG'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  OnCreate = FormCreate
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 8
    Top = 0
    Width = 244
    Height = 13
    Caption = #1055#1091#1090#1100' '#1082' '#1080#1085#1090#1077#1088#1087#1088#1077#1090#1072#1090#1086#1088#1091' ghostscript '#1087#1086' '#1091#1084#1086#1083#1095#1072#1085#1080#1102
  end
  object Edit1: TEdit
    Left = 8
    Top = 16
    Width = 257
    Height = 21
    TabOrder = 1
  end
  object Button1: TButton
    Left = 128
    Top = 56
    Width = 169
    Height = 49
    Caption = #1050#1086#1085#1074#1077#1088#1090#1080#1088#1086#1074#1072#1090#1100' '#1092#1072#1081#1083#1099
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -19
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
    TabOrder = 0
    WordWrap = True
    OnClick = Button1Click
  end
  object Button2: TButton
    Left = 272
    Top = 16
    Width = 25
    Height = 17
    Caption = '...'
    TabOrder = 2
    OnClick = Button2Click
  end
  object CheckBox1: TCheckBox
    Left = 8
    Top = 48
    Width = 121
    Height = 57
    Caption = #1056#1077#1079#1091#1083#1100#1090#1072#1090' '#1074' '#1080#1089#1093#1086#1076#1085#1091#1102' '#1087#1072#1087#1082#1091
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -16
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
    TabOrder = 3
    WordWrap = True
  end
  object OpenDialog1: TOpenDialog
    Filter = 'PDF file|*.pdf'
    Options = [ofHideReadOnly, ofAllowMultiSelect, ofEnableSizing]
    Left = 128
    Top = 32
  end
  object SaveDialog1: TSaveDialog
    Left = 160
    Top = 32
  end
end

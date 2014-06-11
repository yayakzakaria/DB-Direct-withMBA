object Form7: TForm7
  Left = 314
  Top = 322
  BorderStyle = bsDialog
  Caption = 'Select Worksheet :'
  ClientHeight = 62
  ClientWidth = 428
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Verdana'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 58
    Top = 11
    Width = 67
    Height = 13
    Caption = 'Workbook :'
  end
  object Label2: TLabel
    Left = 16
    Top = 36
    Width = 109
    Height = 13
    Caption = 'Select Worksheet :'
  end
  object ComboBox1: TComboBox
    Left = 128
    Top = 32
    Width = 169
    Height = 21
    Style = csDropDownList
    ItemHeight = 13
    TabOrder = 0
  end
  object Edit1: TEdit
    Left = 128
    Top = 8
    Width = 169
    Height = 21
    Color = clBtnFace
    ReadOnly = True
    TabOrder = 1
  end
  object BitBtn1: TBitBtn
    Left = 312
    Top = 8
    Width = 105
    Height = 43
    Caption = 'Select'
    ModalResult = 1
    TabOrder = 2
  end
end

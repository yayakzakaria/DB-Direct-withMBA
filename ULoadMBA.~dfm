object Form6: TForm6
  Left = 261
  Top = 214
  Width = 870
  Height = 613
  Caption = 'Load From MBA'
  Color = clBtnFace
  Font.Charset = ANSI_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Verdana'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnShow = FormShow
  DesignSize = (
    854
    575)
  PixelsPerInch = 96
  TextHeight = 13
  object Label3: TLabel
    Left = 391
    Top = 13
    Width = 92
    Height = 13
    Caption = '&Schedule Date :'
    FocusControl = DateTimePicker3
  end
  object SortGrid1: TSortGrid
    Left = 0
    Top = 64
    Width = 854
    Height = 492
    Anchors = [akLeft, akTop, akRight, akBottom]
    ColCount = 14
    FixedColor = 15780784
    FixedCols = 0
    RowCount = 2
    Font.Charset = ANSI_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'Verdana'
    Font.Style = []
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goColSizing]
    ParentFont = False
    TabOrder = 0
    SortOptions.CanSort = False
    SortOptions.SortStyle = ssNormal
    SortOptions.SortCaseSensitive = False
    SortOptions.SortCol = 0
    SortOptions.SortDirection = sdAscending
    SortOptions.UpdateOnSizeChange = False
    PrintOptions.Copies = 0
    PrintOptions.PrintRange = prAll
    PrintOptions.FromRow = 0
    PrintOptions.ToRow = 0
    PrintOptions.PreviewPage = 0
    HideRows = False
    Filtered = False
    MultiSelect = False
    AlignmentHorz = taLeftJustify
    AlignmentVert = taTopJustify
    BevelStyle = cbLowered
    ProportionalScrollBars = False
    ExtendedKeys = False
    OnGetCellFormat = SortGrid1GetCellFormat
    OnSetChecked = SortGrid1SetChecked
    ColWidths = (
      64
      64
      64
      64
      954
      64
      64
      64
      64
      64
      64
      64
      64
      64)
  end
  object BitBtn1: TBitBtn
    Left = 8
    Top = 8
    Width = 113
    Height = 49
    Caption = '&Load from Excel'
    TabOrder = 1
    OnClick = BitBtn1Click
  end
  object DateTimePicker3: TDateTimePicker
    Left = 393
    Top = 32
    Width = 113
    Height = 24
    Date = 41767.519173483790000000
    Time = 41767.519173483790000000
    Font.Charset = ANSI_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'Verdana'
    Font.Style = []
    ParentFont = False
    TabOrder = 2
    OnChange = DateTimePicker3Change
  end
  object BitBtn2: TBitBtn
    Left = 704
    Top = 8
    Width = 137
    Height = 49
    Anchors = [akTop, akRight]
    Caption = 'Add to &Payment List'
    TabOrder = 3
    OnClick = BitBtn2Click
  end
  object LabeledEdit1: TLabeledEdit
    Left = 208
    Top = 8
    Width = 169
    Height = 21
    Color = clBtnFace
    EditLabel.Width = 71
    EditLabel.Height = 13
    EditLabel.BiDiMode = bdLeftToRight
    EditLabel.Caption = 'Workbook : '
    EditLabel.ParentBiDiMode = False
    LabelPosition = lpLeft
    ReadOnly = True
    TabOrder = 4
  end
  object LabeledEdit2: TLabeledEdit
    Left = 208
    Top = 32
    Width = 169
    Height = 21
    Color = clBtnFace
    EditLabel.Width = 70
    EditLabel.Height = 13
    EditLabel.Caption = 'Worksheet :'
    LabelPosition = lpLeft
    ReadOnly = True
    TabOrder = 5
  end
  object BitBtn6: TBitBtn
    Left = 525
    Top = 31
    Width = 116
    Height = 25
    Caption = 'Cl&ear Selection'
    Enabled = False
    TabOrder = 6
    OnClick = BitBtn6Click
  end
  object BitBtn3: TBitBtn
    Left = 525
    Top = 6
    Width = 116
    Height = 25
    Caption = 'Select &All'
    Enabled = False
    TabOrder = 7
    OnClick = BitBtn3Click
  end
  object StatusBar1: TStatusBar
    Left = 0
    Top = 556
    Width = 854
    Height = 19
    Panels = <
      item
        Text = 'Error: 0'
        Width = 72
      end
      item
        Text = 'Count: 0'
        Width = 80
      end
      item
        Text = 'Total Sum: 0'
        Width = 170
      end>
  end
  object OpenDialog1: TOpenDialog
    DefaultExt = 'xls;xlsx'
    Filter = 'Excel Workbook (*.xls;*.xlsx)|*.xls;*.xlsx'
    Options = [ofPathMustExist, ofFileMustExist, ofEnableSizing]
    Title = 'Load MBA file :'
    Left = 16
    Top = 128
  end
  object IBTransaction1: TIBTransaction
    Active = False
    DefaultDatabase = IBDatabase1
    Params.Strings = (
      'concurrency'
      'nowait')
    AutoStopAction = saNone
    Left = 112
    Top = 128
  end
  object IBQuery1: TIBQuery
    Database = IBDatabase1
    Transaction = IBTransaction1
    AutoCalcFields = False
    BufferChunks = 1000
    CachedUpdates = False
    ParamCheck = False
    Left = 80
    Top = 128
  end
  object IBDatabase1: TIBDatabase
    DatabaseName = '10.234.16.3:D:\Powerpro\Data\PowerBO.GDB'
    Params.Strings = (
      'user_name=SYSDBA'
      'password=masterkey')
    LoginPrompt = False
    IdleTimer = 0
    SQLDialect = 1
    TraceFlags = []
    Left = 48
    Top = 128
  end
end

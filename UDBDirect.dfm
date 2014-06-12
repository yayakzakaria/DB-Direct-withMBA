object Form1: TForm1
  Left = 194
  Top = 231
  Width = 1010
  Height = 697
  Anchors = [akLeft, akTop, akRight, akBottom]
  Caption = 'DB-Direct [IDR Only]'
  Color = clBtnFace
  Font.Charset = ANSI_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Verdana'
  Font.Style = []
  Menu = MainMenu1
  OldCreateOrder = False
  Position = poScreenCenter
  OnClose = FormClose
  OnShow = FormShow
  DesignSize = (
    994
    639)
  PixelsPerInch = 96
  TextHeight = 13
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 994
    Height = 620
    Align = alClient
    BevelOuter = bvNone
    Caption = '... [Please Login First] ...'
    Font.Charset = ANSI_CHARSET
    Font.Color = clPurple
    Font.Height = -21
    Font.Name = 'Verdana'
    Font.Style = []
    ParentFont = False
    TabOrder = 6
  end
  object SortGrid1: TSortGrid
    Left = 0
    Top = 120
    Width = 994
    Height = 500
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
    TabOrder = 3
    OnDrawCell = SortGrid1DrawCell
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
  object DateTimePicker3: TDateTimePicker
    Left = 40
    Top = 208
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
    TabOrder = 4
    Visible = False
    OnChange = DateTimePicker3Change
  end
  object StatusBar1: TStatusBar
    Left = 0
    Top = 620
    Width = 994
    Height = 19
    Panels = <
      item
        Text = 'YAYAK'
        Width = 70
      end
      item
        Width = 140
      end
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
      end
      item
        Text = 'GTD Account: 100-024-900-0'
        Width = 200
      end
      item
        Text = 'IDR'
        Width = 55
      end
      item
        Text = 'GTD BANK: DEUTSCHE BANK SURABAYA'
        Width = 50
      end>
  end
  object GroupBox1: TGroupBox
    Left = 8
    Top = 8
    Width = 457
    Height = 49
    Caption = ' Step 1 : (Load Payment List) '
    TabOrder = 0
    object Label1: TLabel
      Left = 8
      Top = 21
      Width = 84
      Height = 13
      Caption = '&Select Period :'
      FocusControl = DateTimePicker1
    end
    object Label2: TLabel
      Left = 210
      Top = 21
      Width = 24
      Height = 13
      Caption = '&until'
      FocusControl = DateTimePicker2
    end
    object DateTimePicker1: TDateTimePicker
      Left = 96
      Top = 16
      Width = 113
      Height = 24
      Date = 41766.677903807870000000
      Time = 41766.677903807870000000
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = 'Verdana'
      Font.Style = []
      ParentFont = False
      TabOrder = 0
    end
    object DateTimePicker2: TDateTimePicker
      Left = 236
      Top = 16
      Width = 113
      Height = 24
      Date = 41766.677981446760000000
      Time = 41766.677981446760000000
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = 'Verdana'
      Font.Style = []
      ParentFont = False
      TabOrder = 1
    end
    object BitBtn1: TBitBtn
      Left = 352
      Top = 14
      Width = 97
      Height = 25
      Caption = '&Load'
      TabOrder = 2
      OnClick = BitBtn1Click
    end
  end
  object GroupBox2: TGroupBox
    Left = 8
    Top = 64
    Width = 457
    Height = 49
    Caption = ' Step 3 : (Generate and Save) '
    TabOrder = 2
    object BitBtn2: TBitBtn
      Left = 328
      Top = 14
      Width = 121
      Height = 25
      Caption = 'E&xport List'
      Enabled = False
      TabOrder = 2
      OnClick = BitBtn2Click
    end
    object BitBtn3: TBitBtn
      Left = 8
      Top = 14
      Width = 113
      Height = 25
      Caption = '&Purge Deleted'
      Enabled = False
      TabOrder = 0
      OnClick = BitBtn3Click
    end
    object BitBtn4: TBitBtn
      Left = 128
      Top = 14
      Width = 145
      Height = 25
      Caption = '&Generate DB Format'
      Enabled = False
      TabOrder = 1
      OnClick = BitBtn4Click
    end
  end
  object GroupBox3: TGroupBox
    Left = 472
    Top = 8
    Width = 513
    Height = 105
    Anchors = [akLeft, akTop, akRight]
    Caption = ' Step 2 : (Balance Calculation) '
    TabOrder = 1
    DesignSize = (
      513
      105)
    object Label3: TLabel
      Left = 7
      Top = 21
      Width = 65
      Height = 13
      Caption = '&Beginning :'
      FocusControl = EditAlign1
    end
    object Label4: TLabel
      Left = 33
      Top = 45
      Width = 39
      Height = 13
      Caption = 'F&ixed :'
      FocusControl = EditAlign2
    end
    object Label5: TLabel
      Left = 267
      Top = 21
      Width = 61
      Height = 13
      Caption = 'Available :'
    end
    object Label6: TLabel
      Left = 281
      Top = 45
      Width = 47
      Height = 13
      Caption = 'Ending :'
    end
    object BitBtn5: TBitBtn
      Left = 384
      Top = 70
      Width = 120
      Height = 25
      Anchors = [akTop, akRight]
      Caption = '&Calculate'
      Enabled = False
      TabOrder = 2
      OnClick = BitBtn5Click
    end
    object BitBtn6: TBitBtn
      Left = 141
      Top = 70
      Width = 115
      Height = 25
      Anchors = [akTop, akRight]
      Caption = 'Cl&ear Selection'
      Enabled = False
      TabOrder = 4
      OnClick = BitBtn6Click
    end
    object CheckBox1: TCheckBox
      Left = 275
      Top = 74
      Width = 105
      Height = 17
      Alignment = taLeftJustify
      Anchors = [akTop, akRight]
      Caption = '&Auto Calculate'
      Enabled = False
      TabOrder = 3
      OnClick = CheckBox1Click
    end
    object EditAlign2: TEditAlign
      Left = 76
      Top = 40
      Width = 172
      Height = 24
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = 'Verdana'
      Font.Style = []
      ParentFont = False
      ReadOnly = True
      TabOrder = 1
      Text = '0'
      OnEnter = Edit1_2Enter
      OnExit = Edit1_2Exit
      OnKeyPress = Edit1_2KeyPress
      Alignment = taRightJustify
    end
    object EditAlign3: TEditAlign
      Left = 332
      Top = 16
      Width = 172
      Height = 24
      Color = clBtnFace
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = 'Verdana'
      Font.Style = []
      ParentFont = False
      ReadOnly = True
      TabOrder = 6
      Text = '0'
      Alignment = taRightJustify
    end
    object EditAlign4: TEditAlign
      Left = 332
      Top = 40
      Width = 172
      Height = 24
      Color = clBtnFace
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = 'Verdana'
      Font.Style = []
      ParentFont = False
      ReadOnly = True
      TabOrder = 7
      Text = '0'
      Alignment = taRightJustify
    end
    object EditAlign1: TEditAlign
      Left = 76
      Top = 16
      Width = 172
      Height = 24
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = 'Verdana'
      Font.Style = []
      ParentFont = False
      ReadOnly = True
      TabOrder = 0
      Text = '0'
      OnEnter = Edit1_2Enter
      OnExit = Edit1_2Exit
      OnKeyPress = Edit1_2KeyPress
      Alignment = taRightJustify
    end
    object BitBtn7: TBitBtn
      Left = 8
      Top = 70
      Width = 129
      Height = 25
      Caption = '&Find Payment Item'
      Enabled = False
      TabOrder = 5
      OnClick = BitBtn7Click
    end
  end
  object IBDatabase1: TIBDatabase
    Connected = True
    DatabaseName = '10.234.16.3:D:\Powerpro\Data\PowerBO.GDB'
    Params.Strings = (
      'user_name=SYSDBA'
      'password=masterkey')
    LoginPrompt = False
    IdleTimer = 0
    SQLDialect = 1
    TraceFlags = []
    Left = 40
    Top = 168
  end
  object IBQuery1: TIBQuery
    Database = IBDatabase1
    Transaction = IBTransaction1
    AutoCalcFields = False
    BufferChunks = 1000
    CachedUpdates = False
    ParamCheck = False
    SQL.Strings = (
      'SELECT * FROM apcash0'
      'WHERE duedate BETWEEN '#39'05/01/2014'#39' AND '#39'05/07/2014'#39
      'AND status = '#39'A'#39' AND bank = '#39'050'#39' AND currency = '#39'IDR'#39
      'ORDER BY duedate')
    Left = 72
    Top = 168
  end
  object IBTransaction1: TIBTransaction
    Active = True
    DefaultDatabase = IBDatabase1
    Params.Strings = (
      'concurrency'
      'nowait')
    AutoStopAction = saNone
    Left = 104
    Top = 168
  end
  object SaveDialog1: TSaveDialog
    DefaultExt = 'csv'
    Filter = 'CSV File|*.csv'
    Options = [ofOverwritePrompt, ofHideReadOnly, ofEnableSizing]
    Left = 288
    Top = 72
  end
  object Timer1: TTimer
    OnTimer = Timer1Timer
    Left = 96
    Top = 600
  end
  object MainMenu1: TMainMenu
    Left = 32
    Top = 8
    object FIle1: TMenuItem
      Caption = 'File'
      object Login1: TMenuItem
        Caption = 'Login'
        OnClick = Login1Click
      end
      object Logout1: TMenuItem
        Caption = 'Logout'
        OnClick = Logout1Click
      end
      object LoadfromMBA1: TMenuItem
        Caption = 'Load from MBA'
        Visible = False
        OnClick = LoadfromMBA1Click
      end
      object N1: TMenuItem
        Caption = '-'
      end
      object Exit1: TMenuItem
        Caption = 'Exit'
        OnClick = Exit1Click
      end
    end
    object Lainnya1: TMenuItem
      Caption = 'Other'
      object Setting1: TMenuItem
        Caption = 'Setting'
        Visible = False
        OnClick = Setting1Click
      end
      object N2: TMenuItem
        Caption = '-'
      end
      object entang1: TMenuItem
        Caption = 'About'
        OnClick = entang1Click
      end
    end
  end
  object PopupMenu1: TPopupMenu
    Left = 216
    Top = 48
    object Generate1: TMenuItem
      Caption = 'Generate DB Format'
      OnClick = Generate1Click
    end
    object N3: TMenuItem
      Caption = '-'
    end
    object SavetoExcel1: TMenuItem
      Caption = 'Save to Excel'
      OnClick = SavetoExcel1Click
    end
    object ExportGeneratedList1: TMenuItem
      Caption = 'Save to CSV'
      OnClick = ExportGeneratedList1Click
    end
  end
  object PopupMenu2: TPopupMenu
    Left = 392
    Top = 48
    object ExporttoExcel1: TMenuItem
      Caption = 'Export List to Excel'
      OnClick = ExporttoExcel1Click
    end
    object ExportListtoCSV1: TMenuItem
      Caption = 'Export List to CSV'
      OnClick = ExportListtoCSV1Click
    end
  end
end

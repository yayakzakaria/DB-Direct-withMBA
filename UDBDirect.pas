unit UDBDirect;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Buttons, ComCtrls, Grids, IBDatabase, DB,
  IBCustomDataSet, IBQuery, DateUtils, MyGrids, SortGrid, ExtCtrls, Mask,
  AlignEdit, Menus;

type
  TForm1 = class(TForm)
    IBDatabase1: TIBDatabase;
    IBQuery1: TIBQuery;
    IBTransaction1: TIBTransaction;
    SortGrid1: TSortGrid;
    DateTimePicker3: TDateTimePicker;
    StatusBar1: TStatusBar;
    GroupBox1: TGroupBox;
    DateTimePicker1: TDateTimePicker;
    DateTimePicker2: TDateTimePicker;
    BitBtn1: TBitBtn;
    Label1: TLabel;
    Label2: TLabel;
    GroupBox2: TGroupBox;
    BitBtn2: TBitBtn;
    BitBtn3: TBitBtn;
    SaveDialog1: TSaveDialog;
    BitBtn4: TBitBtn;
    GroupBox3: TGroupBox;
    BitBtn5: TBitBtn;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    BitBtn6: TBitBtn;
    CheckBox1: TCheckBox;
    Timer1: TTimer;
    EditAlign2: TEditAlign;
    EditAlign3: TEditAlign;
    EditAlign4: TEditAlign;
    EditAlign1: TEditAlign;
    MainMenu1: TMainMenu;
    FIle1: TMenuItem;
    Lainnya1: TMenuItem;
    Login1: TMenuItem;
    Logout1: TMenuItem;
    N1: TMenuItem;
    Exit1: TMenuItem;
    entang1: TMenuItem;
    N2: TMenuItem;
    Setting1: TMenuItem;
    Panel1: TPanel;
    BitBtn7: TBitBtn;
    PopupMenu1: TPopupMenu;
    Generate1: TMenuItem;
    ExportGeneratedList1: TMenuItem;
    N3: TMenuItem;
    SavetoExcel1: TMenuItem;
    PopupMenu2: TPopupMenu;
    ExportListtoCSV1: TMenuItem;
    ExporttoExcel1: TMenuItem;
    LoadfromMBA1: TMenuItem;
    procedure BitBtn1Click(Sender: TObject);
    procedure SortGrid1GetCellFormat(Sender: TObject; Col, Row: Integer;
      State: TGridDrawState; var FormatOptions: TFormatOptions;
      var CheckBox, Combobox, Ellipsis: Boolean);
    procedure SortGrid1DrawCell(Sender: TObject; ACol, ARow: Integer;
      Rect: TRect; State: TGridDrawState);
    procedure SortGrid1SetChecked(Sender: TObject; ACol, ARow: Integer;
      State: Boolean);
    procedure BitBtn2Click(Sender: TObject);
    procedure BitBtn3Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure BitBtn5Click(Sender: TObject);
    procedure BitBtn6Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure Edit1_2KeyPress(Sender: TObject; var Key: Char);
    procedure Edit1_2Exit(Sender: TObject);
    procedure Edit1_2Enter(Sender: TObject);
    procedure Timer1Timer(Sender: TObject);
    procedure DateTimePicker3Change(Sender: TObject);
    procedure BitBtn4Click(Sender: TObject);
    procedure Logout1Click(Sender: TObject);
    procedure Login1Click(Sender: TObject);
    procedure Exit1Click(Sender: TObject);
    procedure Setting1Click(Sender: TObject);
    procedure BitBtn7Click(Sender: TObject);
    procedure entang1Click(Sender: TObject);
    procedure Generate1Click(Sender: TObject);
    procedure CheckBox1Click(Sender: TObject);
    procedure ExportListtoCSV1Click(Sender: TObject);
    procedure ExporttoExcel1Click(Sender: TObject);
    procedure SavetoExcel1Click(Sender: TObject);
    procedure ExportGeneratedList1Click(Sender: TObject);
    procedure LoadfromMBA1Click(Sender: TObject);
  private
    { Private declarations }
    RowHeader, RowChecked, RowError: array of Integer;
    editRowHeader: Integer;
    skippedPayment, GeneratedList: TStringList;
    foundNumberRow: Integer;
    GenerateFileName: String;
    procedure DefaultState;
    procedure ShowList(sSQL: String);
    procedure Calculate(WithCheckRow: Boolean = False);
    procedure SummaryStatusBar;
    function CheckERROR: Integer;
    procedure GenerateDBFile(const FileName: String);
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

uses StrUtils, Math, ULogin, USetting, UProgress, UAbout, Types, ComObj,
  ULoadMBA;

{$R *.dfm}

// ***** http://delphi.about.com/cs/adptips2002/a/bltip0202_2.htm *****
function DeleteLineBreaks(const S: String): String;
var
   Source, SourceEnd: PChar;
begin
   Source := Pointer(S) ;
   SourceEnd := Source + Length(S) ;
   while Source < SourceEnd do
   begin
     case Source^ of
       #10: Source^ := #32;
       #13: Source^ := #32;
     end;
     Inc(Source) ;
   end;
   Result := S;
   Result := StringReplace(Result, '"', '', [rfReplaceAll]);
   Result := StringReplace(Result, '''', '', [rfReplaceAll]);
end;

// ***** http://www.delphipages.com/forum/showthread.php?t=193990 *****
function ExtractNumberInString(sChaine: String): String;
var
  i: Integer;
begin
  Result := '';
  for i := 1 to Length(sChaine) do
  begin
    if sChaine[i] in ['0'..'9'] then
    Result := Result + sChaine[i];
  end;
end;

function TrimSeparator(Src: String): String;
begin
  Result := StringReplace(Src, '.', '', [rfReplaceAll, rfIgnoreCase]);
  Result := StringReplace(Result, '-', '', [rfReplaceAll, rfIgnoreCase]);
  Result := StringReplace(Result, ' ', '', [rfReplaceAll, rfIgnoreCase]);
  if (Pos('.', Result) > 0) or (Pos('-', Result) > 0) or (Pos(' ', Result) > 0) then
    Result := TrimSeparator(Result);
end;

function TrimThousand(Src: String): String;
begin
  Result := StringReplace(Src, ThousandSeparator, '', [rfReplaceAll, rfIgnoreCase]);
  if (Pos(ThousandSeparator, Result) > 0) then
    Result := TrimThousand(Result);
end;

function InArray(const Value: Integer; const ArrayValue: array of Integer): Boolean;
var i: Integer;
begin
  Result := True;
  for i := Low(ArrayValue) to High(ArrayValue) do
    if ArrayValue[i] = Value then Exit;
  Result := False;
end;

function TrimSpace(Src: String): String;
begin
  Result := StringReplace(Src, '  ', ' ', [rfReplaceAll, rfIgnoreCase]);
  if Pos('  ', Result) > 0 then Result := TrimSpace(Result);
end;

procedure LastSplit(const Key: Char; const Str: String; Res: TStringList);
begin
  Res.Add(Trim(Copy(Str, 1, LastDelimiter(Key, Str)-1)));
  Res.Add(Trim(Copy(Str, LastDelimiter(Key, Str)+1, Length(Str) - LastDelimiter(Key, Str))));
end;

procedure ListSupplier(Str: String; Res: TStringList);
var
  Text1, Name2, Value2: String;
  Split: TStringList;
begin
  Split := TStringList.Create;
  LastSplit(':', Str, Split);
  if Split.Count > 1 then Value2 := Split.Strings[1];
  Text1 := Split.Strings[0];

  Split.Clear;
  LastSplit('|', Text1, Split);
  if Split.Count > 1 then Name2 := Split.Strings[1];
  if Length(Name2) > 0 then
    Res.Add(Name2 + '=' + Value2)
  else Res.Add('Notes=' + Value2);
  if Length(Split.Strings[0]) > 0 then
    ListSupplier(Split.Strings[0], Res);
end;

function BreakdownNote(Str: String): TStringList;
var
  KeySplit: TStringList;
  i: Integer;
begin
  KeySplit := TStringList.Create;
  KeySplit.Delimiter := ',';
  KeySplit.DelimitedText := 'Acc,Bank,Name,NPWP,Nama,Alamat,Kliring';

  for i := 0 to KeySplit.Count - 1 do
    Str := StringReplace(Str, KeySplit[i], '|'+KeySplit[i], [rfReplaceAll, rfIgnoreCase]);
  KeySplit.Free;

  Result := TStringList.Create;
  ListSupplier(Str, Result);
end;

procedure TForm1.DefaultState;
begin
  DateTimePicker1.Date := IncMonth(Now, -1);
  DateTimePicker2.Date := IncMonth(Now, 1);

  SortGrid1.Cols[0].Text := 'Delete';
  SortGrid1.Cols[1].Text := 'Schedule Date';
  SortGrid1.Cols[2].Text := 'Number';
  SortGrid1.Cols[3].Text := 'Supplier';
  SortGrid1.Cols[4].Text := 'Note';
  SortGrid1.Cols[5].Text := 'Text';
  SortGrid1.Cols[6].Text := 'Total';
  SortGrid1.Cols[7].Text := 'Open Date';
  SortGrid1.Cols[8].Text := 'Due Date';
  SortGrid1.Cols[9].Text := 'Code';
  SortGrid1.Cols[10].Text := 'Bank Account';
  SortGrid1.Cols[11].Text := 'Bank Name';
  SortGrid1.Cols[12].Text := 'Bank Kliring';
  SortGrid1.Cols[13].Text := 'N.P.W.P.';
  SortGrid1.Cols[14].Text := 'Contact Address';

  EditAlign1.Text := FormatFloat('#,##0.00;-#,##0.00', 0);
  EditAlign2.Text := FormatFloat('#,##0.00;-#,##0.00', 0);
  EditAlign3.Text := FormatFloat('#,##0.00;-#,##0.00', 0);
  EditAlign4.Text := FormatFloat('#,##0.00;-#,##0.00', 0);

  SortGrid1.Clear;
  SortGrid1.ColCount := 15;
  SortGrid1.RowCount := 2;

  SetLength(RowHeader, 0);
  SetLength(RowChecked, 0);
  editRowHeader := -1;
  foundNumberRow := -1;

  BitBtn2.Enabled := False;
  BitBtn3.Enabled := False;
  BitBtn4.Enabled := False;
  BitBtn5.Enabled := False;
  BitBtn6.Enabled := False;
  BitBtn7.Enabled := False;
  CheckBox1.Enabled := False;
  EditAlign1.ReadOnly := True;
  EditAlign2.ReadOnly := True;
end;

procedure TForm1.FormShow(Sender: TObject);
begin
  skippedPayment := TStringList.Create;
  GeneratedList := TStringList.Create;
  GenerateFileName := '';
  Login1.Click;
end;

procedure TForm1.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  skippedPayment.Free;
  GeneratedList.Free;
end;

procedure TForm1.ShowList(sSQL: String);
var
  trnsfr_date, sch_date, temp_date: TDate;
begin
  IBDatabase1.Close;
  IBQuery1.SQL.Text := sSQL;
  IBQuery1.Open;
  IBQuery1.Last;
  Form4.ProgressBar1.Max := IBQuery1.RecordCount;
  Form4.ProgressBar1.Position := 0;
  IBQuery1.First;

  temp_date := 0;
  SortGrid1.Clear;
  SortGrid1.ColCount := 15;
  SortGrid1.RowCount := 2;

  SetLength(RowHeader, 0);
  SetLength(RowChecked, 0);
  SetLength(RowError, 0);
  editRowHeader := -1;
  foundNumberRow := -1;

  Form4.Show;
  while not IBQuery1.Eof do
  begin
    Application.ProcessMessages;

    trnsfr_date := IBQuery1.FieldByName('duedate').AsDateTime;
    sch_date := IncDay(trnsfr_date, IfThen(DayOfTheWeek(trnsfr_date) > 4, 9 - DayOfTheWeek(trnsfr_date), DayOfTheWeek(trnsfr_date) mod 2));
    if sch_date <> temp_date then
    begin
      if SortGrid1.RowCount > 2 then
        SortGrid1.RowCount := SortGrid1.RowCount + 1;
      SetLength(RowHeader, Length(RowHeader)+1);
      RowHeader[Length(RowHeader)-1] := SortGrid1.RowCount-1;
      SortGrid1.Cells[1, SortGrid1.RowCount-1] := DateToStr(sch_date);
      temp_date := sch_date;
    end;
    SortGrid1.RowCount := SortGrid1.RowCount + 1;
    SortGrid1.Cells[1, SortGrid1.RowCount-1] := DateToStr(sch_date);
    SortGrid1.Cells[2, SortGrid1.RowCount-1] := IBQuery1.FieldByName('number').AsString; //number
    SortGrid1.Cells[3, SortGrid1.RowCount-1] := TrimSpace(BreakdownNote(IBQuery1.FieldByName('bankdetail').AsString).Values['Name']); //supplier name
    SortGrid1.Cells[4, SortGrid1.RowCount-1] := IBQuery1.FieldByName('note').AsString; //note
    SortGrid1.Cells[5, SortGrid1.RowCount-1] := IBQuery1.FieldByName('text1').AsString; //text
    SortGrid1.Cells[6, SortGrid1.RowCount-1] := FormatFloat('#,##0;-#,##0', IBQuery1.FieldByName('total').AsFloat); //total
    SortGrid1.Cells[7, SortGrid1.RowCount-1] := IBQuery1.FieldByName('opendate').AsString; //opendate
    SortGrid1.Cells[8, SortGrid1.RowCount-1] := IBQuery1.FieldByName('duedate').AsString; //duedate
    SortGrid1.Cells[9, SortGrid1.RowCount-1] := IBQuery1.FieldByName('supplier').AsString; //supplier code
    SortGrid1.Cells[10, SortGrid1.RowCount-1] := BreakdownNote(IBQuery1.FieldByName('bankdetail').AsString).Values['Acc']; //bank account
    SortGrid1.Cells[11, SortGrid1.RowCount-1] := TrimSpace(BreakdownNote(IBQuery1.FieldByName('bankdetail').AsString).Values['Bank']); //bank name
    SortGrid1.Cells[12, SortGrid1.RowCount-1] := BreakdownNote(IBQuery1.FieldByName('bankdetail').AsString).Values['Kliring']; //bank kliring code
    SortGrid1.Cells[13, SortGrid1.RowCount-1] := BreakdownNote(IBQuery1.FieldByName('bankdetail').AsString).Values['NPWP']; //n.p.w.p.
    SortGrid1.Cells[14, SortGrid1.RowCount-1] := Trim(IBQuery1.FieldByName('address1').AsString + ' ' + //supplier address
      IBQuery1.FieldByName('address2').AsString + ' ' + IBQuery1.FieldByName('address3').AsString);

    Form4.ProgressBar1.Position := Form4.ProgressBar1.Position + 1;
    IBQuery1.Next;
  end;
  Form4.Hide;
  IBDatabase1.Close;

  SortGrid1.ColWidths[0] := 50;
  SortGrid1.ColWidths[1] := 120;
  SortGrid1.ColWidths[2] := 65;
  SortGrid1.ColWidths[3] := 350;
  SortGrid1.ColWidths[4] := 250;
  SortGrid1.ColWidths[5] := 200;
  SortGrid1.ColWidths[6] := 150;
  SortGrid1.ColWidths[7] := 100;
  SortGrid1.ColWidths[8] := 100;
  SortGrid1.ColWidths[9] := 70;
  SortGrid1.ColWidths[10] := 150;
  SortGrid1.ColWidths[11] := 300;
  SortGrid1.ColWidths[12] := 80;
  SortGrid1.ColWidths[13] := 180;
  SortGrid1.AutoSizeCol(14, True);

  SummaryStatusBar;
  if CheckERROR > 0 then begin
    SortGrid1.Repaint;
    MessageDlg('Terdapat ' + IntToStr(Length(RowError)) + ' item yang bermasalah.'#13'Mohon perbaiki atau coret dari daftar Pembayaran terlebih dahulu.', mtError, [mbOK], 0);
  end;
end;

procedure TForm1.Calculate(WithCheckRow: Boolean = False);
var
  SaldoAvail: Double;
  iRow: Integer;
begin
  SaldoAvail := StrToFloat(TrimThousand(EditAlign1.Text)) - StrToFloat(TrimThousand(EditAlign2.Text));
  for iRow := 2 to SortGrid1.RowCount-1 do
    if (not InArray(iRow, RowHeader)) then
    if (not SortGrid1.CellChecked[0, iRow]) then
      if (SaldoAvail > StrToFloat(TrimThousand(SortGrid1.Cells[6, iRow]))) or (not WithCheckRow) then
        SaldoAvail := SaldoAvail - StrToFloat(TrimThousand(SortGrid1.Cells[6, iRow]))
      else
        SortGrid1.CellChecked[0, iRow] := True;
  EditAlign3.Text := FormatFloat('#,##0.00;-#,##0.00', SaldoAvail);
  EditAlign4.Text := FormatFloat('#,##0.00;-#,##0.00', SaldoAvail + StrToFloat(TrimThousand(EditAlign2.Text)));
end;

procedure TForm1.SummaryStatusBar;
var
  SumTotal: Double;
  iRow, SumCount: Integer;
begin
  SumTotal := 0.00; SumCount := 0;
  for iRow := 2 to SortGrid1.RowCount-1 do
    if (not InArray(iRow, RowHeader)) then
      if (not SortGrid1.CellChecked[0, iRow]) then begin
        SumTotal := SumTotal + StrToFloat(TrimThousand(SortGrid1.Cells[6, iRow]));
        SumCount := SumCount + 1;
      end;
  StatusBar1.Panels[3].Text := 'Count: ' + IntToStr(SumCount);
  StatusBar1.Panels[4].Text := 'Total Sum: ' + FormatFloat('#,##0.00;-#,##0.00', SumTotal);
end;

function TForm1.CheckERROR: Integer;
var
  iRow: Integer;
begin
  SetLength(RowError, 0);
  for iRow := 2 to SortGrid1.RowCount-1 do
    if (not InArray(iRow, RowHeader)) then
      if (not SortGrid1.CellChecked[0, iRow]) then
        if (Length(SortGrid1.Cells[10, iRow]) = 0) or (Length(SortGrid1.Cells[12, iRow]) = 0) // if 'Rek.Tujuan' or 'KliringCode' is blank
          or (CompareDate(StrToDate(SortGrid1.Cells[1, iRow]), Now) = LessThanValue) then      // or ScheduleDate is less than Today
            if not InArray(iRow, RowError) then begin
              SetLength(RowError, Length(RowError)+1);
              RowError[Length(RowError)-1] := iRow;
            end;
  Result := Length(RowError);
  StatusBar1.Panels[2].Text := 'Error: ' + IntToStr(Result);
end;

procedure TForm1.SortGrid1GetCellFormat(Sender: TObject; Col, Row: Integer;
  State: TGridDrawState; var FormatOptions: TFormatOptions; var CheckBox,
  Combobox, Ellipsis: Boolean);
begin
  if (Col = 0) and (Row <> 0) and not InArray(Row, RowHeader) then CheckBox := True;
  if (Col in [1,2,7,8,9]) then FormatOptions.AlignmentHorz := taCenter;
  if (Col = 6) then FormatOptions.AlignmentHorz := taRightJustify;
  if (Row mod 2 > 0) then FormatOptions.Brush.Color := $00FFF7EC;
  if InArray(Row, RowHeader) then begin
    FormatOptions.Font.Style := [fsBold];
    FormatOptions.Brush.Color := clMoneyGreen;
    FormatOptions.AlignmentHorz := taLeftJustify;
  end;
  if InArray(Row, RowChecked) then begin
    FormatOptions.Font.Style := [fsStrikeOut];
    FormatOptions.Brush.Color := clSilver;
  end;
  if InArray(Row, RowError) then FormatOptions.Brush.Color := clRed;
  if (Row = foundNumberRow) then begin
    if (Col = 2) then FormatOptions.Font.Style := [];
    FormatOptions.Brush.Color := clLime;
  end;
  if (Row = 0) then FormatOptions.AlignmentHorz := taCenter;
end;

procedure TForm1.SortGrid1DrawCell(Sender: TObject; ACol, ARow: Integer;
  Rect: TRect; State: TGridDrawState);
begin
  if (gdFocused in State) then
  begin
    if (ACol = 1) and InArray(ARow, RowHeader) then
    with DateTimePicker3 do
    begin
      Left := Rect.Left + SortGrid1.Left + 1;
      Top := Rect.Top + SortGrid1.Top + 1;
      Width := Rect.Right - Rect.Left + 2;
      Width := Rect.Right - Rect.Left + 2;
      Height := Rect.Bottom - Rect.Top + 2;
      Date := StrToDate(SortGrid1.Cells[ACol, ARow]);
      editRowHeader := ARow;
      Visible := True;
    end;
  end else begin
    DateTimePicker3.Visible := False;
    editRowHeader := -1;
  end;
end;

procedure TForm1.SortGrid1SetChecked(Sender: TObject; ACol, ARow: Integer;
  State: Boolean);
var
  i: Integer;
begin
  foundNumberRow := -1;
  if State then begin  // --------------------------- if Row is Checked
    if InArray(-1, RowChecked) then begin
      for i := Low(RowChecked) to High(RowChecked) do
        if RowChecked[i] = -1 then begin
          RowChecked[i] := ARow;
          Break;
        end;
    end else begin
      SetLength(RowChecked, Length(RowChecked)+1);
      RowChecked[Length(RowChecked)-1] := ARow;
    end;
  end else begin       // --------------------------- if Row is unChecked
    for i := Low(RowChecked) to High(RowChecked) do
      if RowChecked[i] = ARow then begin
        RowChecked[i] := -1;
        Break;
      end;
  end;
  if CheckBox1.Checked then Calculate;
  SummaryStatusBar;
  CheckERROR;
end;

procedure TForm1.DateTimePicker3Change(Sender: TObject);
var iRow: Integer;
begin
  if InArray(editRowHeader, RowHeader) then begin
    SortGrid1.Cells[1, editRowHeader] := DateToStr(DateTimePicker3.Date);

    iRow := editRowHeader + 1;
    while (not InArray(iRow, RowHeader)) and (iRow < SortGrid1.RowCount) do begin
      SortGrid1.Cells[1, iRow] := DateToStr(DateTimePicker3.Date);
      iRow := iRow + 1;
    end;
  end;
  CheckERROR;
  SortGrid1.Repaint;
end;

procedure TForm1.GenerateDBFile(const FileName: String);
var
   F: TextFile;
   iRow: Integer;
begin
  try
    AssignFile(F, Filename);
    Rewrite(F);

    GeneratedList.Clear;
    for iRow := 2 to SortGrid1.RowCount-1 do
      if (not InArray(iRow, RowHeader)) then
        if (not SortGrid1.CellChecked[0, iRow]) then begin
          Write(F, TrimSeparator(Form3.BankAccount), ';', Form3.BankCurrency, ';'); //our_bank_account_to_be_debited
          Write(F, TrimSeparator(Form3.BankAccount), ';', Form3.BankCurrency, ';CNS;15;;N;IDR;'); //our_bank_account_to_be_debited
          Write(F, FormatDateTime('ddmmyyyy', StrToDate(SortGrid1.Cells[1, iRow])), ';;;;;'); //trx_date - column 10
          Write(F, FormatDateTime('ddmmyyyy', StrToDate(SortGrid1.Cells[1, iRow])), ';N;;'); //schedule_date - column 15
          Write(F, FormatFloat('0;-0', StrToFloat(TrimThousand(SortGrid1.Cells[6, iRow]))), ';;'); //trx_amount - column 18
          Write(F, Trim(UpperCase(LeftStr(TrimSpace(DeleteLineBreaks(SortGrid1.Cells[4, iRow])), 35))), ';;;;;;;'); //trx_detail - column 20 [maxlength 35 chars]
          Write(F, UpperCase(SortGrid1.Cells[3, iRow]), ';'); //supplier_name - column 27
          Write(F, ExtractNumberInString(SortGrid1.Cells[10, iRow]), ';;;'); //bank_account - column 28
          Write(F, '', ';;'); //supplier_address1 - column 31 [not mandatory, can blank]
          Write(F, '', ';;;;'); //supplier_address3 - column 33 [not mandatory, can blank]
          Write(F, UpperCase(SortGrid1.Cells[11, iRow]), ';'); //bank_name - column 37
          Write(F, '', ';'); //bank_swift - column 38 [blank for IDR]
          Write(F, SortGrid1.Cells[12, iRow], ';'); //bank_clearing_code - column 39
          Write(F, '25-ZZZ', ';;;;'); //bank_clearing_code_type - column 40 [always it when IDR]
          Write(F, 'ID', ';;;;;;;;;;;;;;;'); //supplier_bank_country - column 44 [always it when IDR]
          Write(F, Form2.EmailUser, ';;;;;'); //email_address_of_our_cp - column 59 [active if our company has autoemail agreement]
          Write(F, SortGrid1.Cells[2, iRow], ';;;', #13#10); //trx_unique_id - column 64 [use number of cash_payment_doc]
          GeneratedList.Add(SortGrid1.Cells[2, iRow]);
        end;

    CloseFile(F);
    ShowMessage('Generate file done.');
  except
    MessageDlg('Can''t write to file:'#13 + Filename, mtError, [mbOK], 0);
  end;
end;

procedure TForm1.BitBtn1Click(Sender: TObject);
begin
  skippedPayment.Clear;
  ShowList('SELECT c.number, c.note, c.text1, c.total, c.opendate, c.duedate, c.supplier, s.address1, s.address2, s.address3, s.note bankdetail FROM apcash0 c ' +
          'LEFT JOIN syssupplier s ON s.supplier = c.supplier AND s.flag_active = 1 ' +
          'WHERE c.duedate BETWEEN ''' + FormatDateTime('mm/dd/yyyy', DateTimePicker1.Date) + ''' AND ''' + FormatDateTime('mm/dd/yyyy', DateTimePicker2.Date) + ''' ' +
          'AND c.status = ''A'' AND c.bank = ''050'' AND c.currency = ''IDR'' ORDER BY c.duedate, c.number');
  BitBtn2.Enabled := True;
  BitBtn3.Enabled := True;
  BitBtn4.Enabled := True;
  BitBtn5.Enabled := True;
  BitBtn6.Enabled := True;
  BitBtn7.Enabled := True;
  CheckBox1.Enabled := True;
  EditAlign1.ReadOnly := False;
  EditAlign2.ReadOnly := False;
end;

procedure TForm1.BitBtn2Click(Sender: TObject);
var
  button: TBitBtn;
  lowerLeft: TPoint;
begin
  if Sender is TBitBtn then
  begin
    button := TBitBtn(Sender);
    lowerLeft := Point(0, button.Height);
    lowerLeft := button.ClientToScreen(lowerLeft);
    PopupMenu2.Popup(lowerLeft.X, lowerLeft.Y);
  end;
end;

procedure TForm1.BitBtn3Click(Sender: TObject);
var iRow: Integer;
begin
  for iRow := Low(RowChecked) to High(RowChecked) do
    if RowChecked[iRow] > 0 then
      skippedPayment.Add(SortGrid1.Cells[2, RowChecked[iRow]]);
  if skippedPayment.Count = 0 then Exit;
  ShowList('SELECT c.number, c.note, c.text1, c.total, c.opendate, c.duedate, c.supplier, s.address1, s.address2, s.address3, s.note bankdetail FROM apcash0 c ' +
          'LEFT JOIN syssupplier s ON s.supplier = c.supplier AND s.flag_active = 1 ' +
          'WHERE c.duedate BETWEEN ''' + FormatDateTime('mm/dd/yyyy', DateTimePicker1.Date) + ''' AND ''' + FormatDateTime('mm/dd/yyyy', DateTimePicker2.Date) + ''' ' +
          'AND c.status = ''A'' AND c.bank = ''050'' AND c.currency = ''IDR'' AND c.number NOT IN (' + skippedPayment.DelimitedText + ') ORDER BY c.duedate, c.number');
end;

procedure TForm1.BitBtn4Click(Sender: TObject);
var
  button: TBitBtn;
  lowerLeft: TPoint;
begin
  if Sender is TBitBtn then
  begin
    button := TBitBtn(Sender);
    lowerLeft := Point(0, button.Height);
    lowerLeft := button.ClientToScreen(lowerLeft);
    PopupMenu1.Popup(lowerLeft.X, lowerLeft.Y);
  end;
end;

procedure TForm1.BitBtn5Click(Sender: TObject);
begin
  Calculate(True);
  SortGrid1.Repaint;
  GeneratedList.Clear;
end;

procedure TForm1.BitBtn6Click(Sender: TObject);
var
  iRow: Integer;
begin
  for iRow := 2 to SortGrid1.RowCount-1 do
    if (not InArray(iRow, RowHeader)) then SortGrid1.CellChecked[0, iRow] := False;
  SortGrid1.Repaint;
  GeneratedList.Clear;
end;

procedure TForm1.BitBtn7Click(Sender: TObject);
var
  findNumber: String;
  iRow: Integer;
begin
  foundNumberRow := -1;
  if InputQuery('Find Payment Item', 'Please enter the Payment Number', findNumber) then begin
    iRow := SortGrid1.Cols[2].IndexOf(findNumber);
    if iRow > 1 then begin
      foundNumberRow := iRow;
      SortGrid1.MoveTo(2, iRow);
      SortGrid1.Repaint;
    end else ShowMessage('Payment item tidak diketemukan.');
  end;
end;

procedure TForm1.CheckBox1Click(Sender: TObject);
begin
  Calculate;
  SortGrid1.Repaint;
end;

procedure TForm1.Edit1_2KeyPress(Sender: TObject; var Key: Char);
begin
  if not (Key in [#8, '0'..'9', '-', DecimalSeparator]) then
    Key := #0
  else if ((Key = DecimalSeparator) or (Key = '-')) and (Pos(Key, (Sender as TEditAlign).Text) > 0) then
    Key := #0
  else if (Key = '-') and ((Sender as TEditAlign).SelStart <> 0) then
    Key := #0;
end;

procedure TForm1.Edit1_2Exit(Sender: TObject);
begin
  (Sender as TEditAlign).Text := FormatFloat('#,##0.00;-#,##0.00', StrToFloat(TrimThousand((Sender as TEditAlign).Text)));
end;

procedure TForm1.Edit1_2Enter(Sender: TObject);
begin
  (Sender as TEditAlign).Text := TrimThousand((Sender as TEditAlign).Text);
  (Sender as TEditAlign).SelectAll;
end;

procedure TForm1.Timer1Timer(Sender: TObject);
begin
  StatusBar1.Panels[1].Text := FormatDateTime('dd/mm/yyyy [hh:nn:ss]', Now);
end;

procedure TForm1.Logout1Click(Sender: TObject);
begin
  StatusBar1.Panels[0].Text := '';
  StatusBar1.Panels[2].Text := 'Error: 0';
  StatusBar1.Panels[3].Text := 'Count: 0';
  StatusBar1.Panels[4].Text := 'Total Sum: ' + FormatFloat('#,##0.00', 0);
  StatusBar1.Panels[5].Text := 'GTD Account: ';
  StatusBar1.Panels[6].Text := '';
  StatusBar1.Panels[7].Text := 'GTD BANK: ';
  LoadfromMBA1.Visible := False;
  Setting1.Visible := False;

  Panel1.Caption := '... [Please Login First] ...';
  Panel1.BringToFront;
  DefaultState;

  BitBtn1.Enabled := False;
  DateTimePicker1.Enabled := False;
  DateTimePicker2.Enabled := False;
end;

procedure TForm1.Login1Click(Sender: TObject);
begin
  Logout1.Click;
  if Form2.ShowModal = mrOk then begin
    StatusBar1.Panels[0].Text := Form2.Username;
    LoadfromMBA1.Visible := True;
    Setting1.Visible := True;
    Form3.LoadBankCompany;
    StatusBar1.Panels[5].Text := 'GTD Account: ' + Form3.BankAccount;
    StatusBar1.Panels[6].Text := Form3.BankCurrency;
    StatusBar1.Panels[7].Text := 'GTD BANK: ' + Form3.BankName;

    BitBtn1.Enabled := True;
    DateTimePicker1.Enabled := True;
    DateTimePicker2.Enabled := True;

    Panel1.Caption := '';
    Panel1.SendToBack;
  end;
end;

procedure TForm1.Setting1Click(Sender: TObject);
begin
  if Form3.ShowModal = mrOk then begin
    StatusBar1.Panels[5].Text := 'GTD Account: ' + Form3.BankAccount;
    StatusBar1.Panels[6].Text := Form3.BankCurrency;
    StatusBar1.Panels[7].Text := 'GTD BANK: ' + Form3.BankName;
  end;
end;

procedure TForm1.Exit1Click(Sender: TObject);
begin
  Self.Close;
end;

procedure TForm1.entang1Click(Sender: TObject);
begin
  Form5.ShowModal;
end;

procedure TForm1.Generate1Click(Sender: TObject);
begin
  if CheckERROR > 0 then begin
    MessageDlg('Terdapat ' + IntToStr(Length(RowError)) + ' item yang bermasalah.'#13'Mohon perbaiki atau coret dari daftar Pembayaran terlebih dahulu.', mtError, [mbOK], 0);
    Exit;
  end;
  GenerateFileName := FormatDateTime('ddmmyyyyhhnnss', Now);
  SaveDialog1.Title := 'Select destination for DB-Format file';
  SaveDialog1.DefaultExt := 'txt';
  SaveDialog1.Filter := 'DB-Format file|*.txt';
  SaveDialog1.FileName := 'dbi_' + GenerateFileName + '.txt';

  if SaveDialog1.Execute then
    GenerateDBFile(SaveDialog1.FileName);
end;

procedure TForm1.ExporttoExcel1Click(Sender: TObject);
const
  xlHAlignCenter = -4108;
  xlThin = 2;
var
  iExcel: OleVariant;
  iSheet: Variant;
  iRow, iCol: Integer;
  formatXls: String;
begin
  // By using GetActiveOleObject, you use an instance of Word that's already running,
  // if there is one.
  try
    iExcel := GetActiveOleObject('Excel.Application');
  except
    try
      // If no instance of Word is running, try to Create a new Excel Object
      iExcel := CreateOleObject('Excel.Application');
    except
      ShowMessage('Cannot start Excel'#13'Is Excel not installed ?');
      Exit;
    end;
  end;
  formatXls := Format('%s', [iExcel.Version]);
  formatXls := IfThen(StrToInt(LeftStr(formatXls, Pos('.', formatXls)-1)) < 12, 'xls', 'xlsx');

  SaveDialog1.Title := 'Export to Excel file';
  SaveDialog1.DefaultExt := formatXls;
  SaveDialog1.Filter := 'Excel Workbook (*.' + formatXls + ')|*.' + formatXls;
  SaveDialog1.FileName := 'Export List ' + FormatDateTime('ddmmyyyyhhnnss', Now) + '.' + formatXls;
  if not SaveDialog1.Execute then begin
    iExcel.DisplayAlerts := False;  // Discard unsaved files....
    iExcel.Quit;
    Exit;
  end;

  Form4.ProgressBar1.Max := SortGrid1.RowCount-1;
  Form4.ProgressBar1.Position := 0;
  Form4.Show;

  iExcel.Visible := False;
  iSheet := iExcel.Workbooks.Add.ActiveSheet;

  for iCol := 1 to SortGrid1.ColCount-1 do
    iSheet.Cells[1, iCol] := SortGrid1.Cells[iCol, 0];

  iSheet.Range['A1:' + Chr(Ord('A') + SortGrid1.ColCount - 2) + '1'].Font.Bold := True;
  iSheet.Range['A1:' + Chr(Ord('A') + SortGrid1.ColCount - 2) + '1'].Interior.Color := clSkyBlue;
  iSheet.Range['A1:' + Chr(Ord('A') + SortGrid1.ColCount - 2) + '1'].HorizontalAlignment := xlHAlignCenter;

  for iRow := 1 to SortGrid1.RowCount-1 do begin  // ------------------------------Loop iRow
    if InArray(iRow, RowHeader) then begin
      iSheet.Cells[iRow + 1, 1].Font.Bold := True;
      iSheet.Cells[iRow + 1, 1].Interior.Color := clMoneyGreen;
    end else begin
      iSheet.Cells[iRow + 1, 2].HorizontalAlignment := xlHAlignCenter;
      iSheet.Cells[iRow + 1, 2].NumberFormat := '@';  // as Text
      iSheet.Cells[iRow + 1, 10].NumberFormat := '@';  // as Text
      iSheet.Cells[iRow + 1, 12].NumberFormat := '@';  // as Text
    end;

    for iCol := 1 to SortGrid1.ColCount-1 do begin  // ----------------------------Loop iCol
      if iCol = 1 then begin
        iSheet.Cells[iRow + 1, 1] := '''' + SortGrid1.Cells[1, iRow];
        iSheet.Cells[iRow + 1, 1].NumberFormat := LowerCase(ShortDateFormat);  // as Date
        iSheet.Cells[iRow + 1, 1].HorizontalAlignment := xlHAlignCenter;
      end;
      if (not InArray(iRow, RowHeader)) then begin
        if iCol = 6 then begin
          iSheet.Cells[iRow + 1, 6] := FormatFloat('0;-0', StrToFloat(TrimThousand(SortGrid1.Cells[6, iRow])));
          iSheet.Cells[iRow + 1, 6].NumberFormat := '#.##0';  // as Number
        end;

        if SortGrid1.CellChecked[0, iRow] then begin
          iSheet.Cells[iRow + 1, iCol].Font.Strikethrough := True;
          iSheet.Cells[iRow + 1, iCol].Interior.Color := clSilver;
        end else if InArray(iRow, RowError) then
          iSheet.Cells[iRow + 1, iCol].Interior.Color := clRed;
      end;

      if not(iCol in [1, 6]) then
        iSheet.Cells[iRow + 1, iCol] := SortGrid1.Cells[iCol, iRow];
    end;  // ----------------------------------------------------------------------Loop iCol
    Form4.ProgressBar1.Position := iRow;
  end;  // ------------------------------------------------------------------------Loop iRow
  Form4.Hide;

  iSheet.Range['A1:' + Chr(Ord('A') + SortGrid1.ColCount - 2) + '1'].EntireColumn.AutoFit;
  iSheet.Range['A1:'+ Chr(Ord('A') + SortGrid1.ColCount - 2) + IntToStr(SortGrid1.RowCount)].Borders.Weight := xlThin;

  iExcel.ActiveWorkBook.SaveAs(SaveDialog1.FileName);

  iExcel.Visible := True;
  iExcel.UserControl := True;
end;

procedure TForm1.ExportListtoCSV1Click(Sender: TObject);
begin
  SaveDialog1.Title := 'Export CSV file';
  SaveDialog1.DefaultExt := 'csv';
  SaveDialog1.Filter := 'CSV File|*.csv';

  if SaveDialog1.Execute then begin
    SortGrid1.SaveToCSV(SaveDialog1.FileName, ';');
    ShowMessage('Save CSV file done.');
  end;
end;

procedure TForm1.SavetoExcel1Click(Sender: TObject);
const
  xlHAlignCenter = -4108;
  xlRight = -4152;
  xlThin = 2;
var
  iExcel: OleVariant;
  iSheet: Variant;
  iRow, iNum, xlRow, xlRowLast: Integer;
  formatXls: String;
begin
  if GeneratedList.Count = 0 then begin
    MessageDlg('Generate DB Format belum dilakukan.'#13'Lakukan Generate DB Format terlebih dahulu dan kemudian ulangi ini lagi.', mtWarning, [mbOK], 0);
    Exit;
  end;

  // By using GetActiveOleObject, you use an instance of Word that's already running,
  // if there is one.
  try
    iExcel := GetActiveOleObject('Excel.Application');
  except
    try
      // If no instance of Word is running, try to Create a new Excel Object
      iExcel := CreateOleObject('Excel.Application');
    except
      ShowMessage('Cannot start Excel'#13'Is Excel not installed ?');
      Exit;
    end;
  end;
  formatXls := Format('%s', [iExcel.Version]);
  formatXls := IfThen(StrToInt(LeftStr(formatXls, Pos('.', formatXls)-1)) < 12, 'xls', 'xlsx');

  SaveDialog1.Title := 'Save the Last Generated Payment to Excel file';
  SaveDialog1.DefaultExt := formatXls;
  SaveDialog1.Filter := 'Excel File (*.' + formatXls + ')|*.' + formatXls;
  SaveDialog1.FileName := 'Generated PaymentList ' + GenerateFileName + '.' + formatXls;
  if not SaveDialog1.Execute then begin
    iExcel.DisplayAlerts := False;  // Discard unsaved files....
    iExcel.Quit;
    Exit;
  end;

  Form4.ProgressBar1.Max := SortGrid1.RowCount-1;
  Form4.ProgressBar1.Position := 0;
  Form4.Show;

  iExcel.Visible := False;
  iSheet := iExcel.Workbooks.Add.ActiveSheet;

  iSheet.Cells[1, 1] := 'NO';
  iSheet.Cells[1, 2] := 'SCHEDULE DATE';
  iSheet.Cells[1, 3] := 'TRX ID';
  iSheet.Cells[1, 4] := 'BENEFICIARY NAME';
  iSheet.Cells[1, 5] := 'TRX AMOUNT';
  iSheet.Cells[1, 6] := 'BANK ACCOUNT';
  iSheet.Cells[1, 7] := 'BANK NAME';
  iSheet.Cells[1, 8] := 'BANK CODE';
  iSheet.Cells[1, 9] := 'TRX DETAIL';
  iSheet.Cells[1,10] := 'DEBIT ACCOUNT';
  iSheet.Cells[1,11] := 'DEBIT CURRENCY';

  iSheet.Range['A1:K1'].Font.Bold := True;
  iSheet.Range['A1:K1'].Interior.Color := clMoneyGreen;
  iSheet.Range['A1:K1'].HorizontalAlignment := xlHAlignCenter;

  iNum := 1; xlRow := 2; xlRowLast := 0;
  for iRow := 1 to SortGrid1.RowCount-1 do begin
    if InArray(iRow, RowHeader) and (xlRow - xlRowLast > 1) then begin
      iSheet.Range['A' + IntToStr(xlRow) + ':K' + IntToStr(xlRow)].Interior.Color := clSkyBlue;
      xlRowLast := xlRow;
      xlRow := xlRow + 1;
    end;

    if GeneratedList.IndexOf(SortGrid1.Cells[2, iRow]) > -1 then begin
      iSheet.Cells[xlRow, 1].HorizontalAlignment := xlHAlignCenter;
      iSheet.Cells[xlRow, 2].HorizontalAlignment := xlHAlignCenter;
      iSheet.Cells[xlRow, 3].HorizontalAlignment := xlHAlignCenter;
      iSheet.Cells[xlRow, 2].NumberFormat := LowerCase(ShortDateFormat);  // as Date
      iSheet.Cells[xlRow, 5].NumberFormat := '#.##0,00';  // as Currency
      iSheet.Cells[xlRow, 3].NumberFormat := '@';  // as Text
      iSheet.Cells[xlRow, 6].NumberFormat := '@';  // as Text
      iSheet.Cells[xlRow, 8].NumberFormat := '@';  // as Text
      iSheet.Cells[xlRow,10].NumberFormat := '@';  // as Text

      iSheet.Cells[xlRow, 1] := IntToStr(iNum);
      iSheet.Cells[xlRow, 2] := '''' + SortGrid1.Cells[1, iRow];
      iSheet.Cells[xlRow, 3] := SortGrid1.Cells[2, iRow];
      iSheet.Cells[xlRow, 4] := UpperCase(SortGrid1.Cells[3, iRow]);
      iSheet.Cells[xlRow, 5] := FormatFloat('0;-0', StrToFloat(TrimThousand(SortGrid1.Cells[6, iRow])));
      iSheet.Cells[xlRow, 6] := ExtractNumberInString(SortGrid1.Cells[10, iRow]);
      iSheet.Cells[xlRow, 7] := UpperCase(SortGrid1.Cells[11, iRow]);
      iSheet.Cells[xlRow, 8] := SortGrid1.Cells[12, iRow];
      iSheet.Cells[xlRow, 9] := Trim(UpperCase(LeftStr(TrimSpace(DeleteLineBreaks(SortGrid1.Cells[4, iRow])), 35)));
      iSheet.Cells[xlRow,10] := TrimSeparator(Form3.BankAccount);
      iSheet.Cells[xlRow,11] := Form3.BankCurrency;

      iNum := iNum + 1;
      xlRow := xlRow + 1;
    end;
    Form4.ProgressBar1.Position := iRow;
  end;
  Form4.Hide;

  iSheet.Cells[xlRow, 4] := 'TOTAL AMOUNT';
  iSheet.Cells[xlRow, 5].Formula := '=SUM(E3:E' + IntToStr(xlRow-1) + ')';
  iSheet.Cells[xlRow, 4].HorizontalAlignment := xlRight;
  iSheet.Range['D' + IntToStr(xlRow) + ':E'+ IntToStr(xlRow)].Borders.Weight := xlThin;
  iSheet.Range['D' + IntToStr(xlRow) + ':E'+ IntToStr(xlRow)].Font.Bold := True;
  iSheet.Range['D' + IntToStr(xlRow) + ':E'+ IntToStr(xlRow)].Interior.Color := clMoneyGreen;

  iSheet.Range['A1:K1'].EntireColumn.AutoFit;
  iSheet.Range['A1:K'+ IntToStr(xlRow-1)].Borders.Weight := xlThin;

  iExcel.ActiveWorkBook.SaveAs(SaveDialog1.FileName);

  iExcel.Visible := True;
  iExcel.UserControl := True;
end;

procedure TForm1.ExportGeneratedList1Click(Sender: TObject);
var
   F: TextFile;
   iRow, iNum: Integer;
begin
  if GeneratedList.Count = 0 then begin
    MessageDlg('Generate DB Format belum dilakukan.'#13'Lakukan Generate DB Format terlebih dahulu dan kemudian ulangi ini lagi.', mtWarning, [mbOK], 0);
    Exit;
  end;

  SaveDialog1.Title := 'Save the Last Generated Payment to CSV file';
  SaveDialog1.DefaultExt := 'csv';
  SaveDialog1.Filter := 'CSV File|*.csv';
  SaveDialog1.FileName := 'Generated PaymentList ' + GenerateFileName + '.csv';
  if not SaveDialog1.Execute then Exit;

  try
    AssignFile(F, SaveDialog1.Filename);
    Rewrite(F);

    Write(F, 'NO;');
    Write(F, 'SCHEDULE DATE;');
    Write(F, 'TRX ID;');
    Write(F, 'BENEFICIARY NAME;');
    Write(F, 'TRX AMOUNT;');
    Write(F, 'BANK ACCOUNT;');
    Write(F, 'BANK NAME;');
    Write(F, 'BANK CODE;');
    Write(F, 'TRX DETAIL;');
    Write(F, 'DEBIT ACCOUNT;');
    Write(F, 'DEBIT CURRENCY'#13#10);

    iNum := 1;
    for iRow := 1 to SortGrid1.RowCount-1 do begin
      if InArray(iRow, RowHeader) then
        Write(F, ';;;;;;;;;;'#13#10);

      if GeneratedList.IndexOf(SortGrid1.Cells[2, iRow]) > -1 then begin
        Write(F, IntToStr(iNum), ';');
        Write(F, SortGrid1.Cells[1, iRow], ';');
        Write(F, SortGrid1.Cells[2, iRow], ';');
        Write(F, UpperCase(SortGrid1.Cells[3, iRow]), ';');
        Write(F, FormatFloat('0;-0', StrToFloat(TrimThousand(SortGrid1.Cells[6, iRow]))), ';');
        Write(F, ExtractNumberInString(SortGrid1.Cells[10, iRow]), ';');
        Write(F, UpperCase(SortGrid1.Cells[11, iRow]), ';');
        Write(F, SortGrid1.Cells[12, iRow], ';');
        Write(F, Trim(UpperCase(LeftStr(TrimSpace(DeleteLineBreaks(SortGrid1.Cells[4, iRow])), 35))), ';');
        Write(F, TrimSeparator(Form3.BankAccount), ';');
        Write(F, Form3.BankCurrency, #13#10);

        iNum := iNum + 1;
      end;
    end;

    CloseFile(F);
    ShowMessage('Generate file done.');
  except
    MessageDlg('Can''t write to file:'#13 + SaveDialog1.Filename, mtError, [mbOK], 0);
  end;
end;

procedure TForm1.LoadfromMBA1Click(Sender: TObject);
var
  iRow: Integer;
begin
  if Form6.ShowModal = mrOK then begin
    if SortGrid1.RowCount > 2 then
      SortGrid1.RowCount := SortGrid1.RowCount + 1;
    SetLength(RowHeader, Length(RowHeader)+1);
    RowHeader[Length(RowHeader)-1] := SortGrid1.RowCount-1;
    SortGrid1.Cells[1, SortGrid1.RowCount-1] := DateToStr(Form6.DateTimePicker3.Date);

    for iRow := 1 to Form6.SortGrid1.RowCount-1 do
      if (not Form6.SortGrid1.CellChecked[0, iRow]) then begin
        SortGrid1.RowCount := SortGrid1.RowCount + 1;
        SortGrid1.Cells[1, SortGrid1.RowCount-1] := Form6.SortGrid1.Cells[1, iRow];
        SortGrid1.Cells[2, SortGrid1.RowCount-1] := Form6.SortGrid1.Cells[2, iRow]; //number
        SortGrid1.Cells[3, SortGrid1.RowCount-1] := Form6.SortGrid1.Cells[3, iRow]; //supplier name
        SortGrid1.Cells[4, SortGrid1.RowCount-1] := Form6.SortGrid1.Cells[4, iRow]; //note
        SortGrid1.Cells[5, SortGrid1.RowCount-1] := Form6.SortGrid1.Cells[5, iRow]; //text
        SortGrid1.Cells[6, SortGrid1.RowCount-1] := Form6.SortGrid1.Cells[6, iRow]; //total
        SortGrid1.Cells[7, SortGrid1.RowCount-1] := ''; //opendate
        SortGrid1.Cells[8, SortGrid1.RowCount-1] := ''; //duedate
        SortGrid1.Cells[9, SortGrid1.RowCount-1] := Form6.SortGrid1.Cells[7, iRow]; //supplier code
        SortGrid1.Cells[10, SortGrid1.RowCount-1] := Form6.SortGrid1.Cells[8, iRow]; //bank account
        SortGrid1.Cells[11, SortGrid1.RowCount-1] := Form6.SortGrid1.Cells[9, iRow]; //bank name
        SortGrid1.Cells[12, SortGrid1.RowCount-1] := Form6.SortGrid1.Cells[10, iRow]; //bank kliring code
        SortGrid1.Cells[13, SortGrid1.RowCount-1] := ''; //n.p.w.p.
        SortGrid1.Cells[14, SortGrid1.RowCount-1] := ''; //supplier address
      end;

    SortGrid1.ColWidths[0] := 50;
    SortGrid1.ColWidths[1] := 120;
    SortGrid1.ColWidths[2] := 65;
    SortGrid1.ColWidths[3] := 350;
    SortGrid1.ColWidths[4] := 250;
    SortGrid1.ColWidths[5] := 200;
    SortGrid1.ColWidths[6] := 150;
    SortGrid1.ColWidths[7] := 100;
    SortGrid1.ColWidths[8] := 100;
    SortGrid1.ColWidths[9] := 70;
    SortGrid1.ColWidths[10] := 150;
    SortGrid1.ColWidths[11] := 300;
    SortGrid1.ColWidths[12] := 80;
    SortGrid1.ColWidths[13] := 180;
    SortGrid1.AutoSizeCol(14, True);

    SummaryStatusBar;
    BitBtn2.Enabled := True;
    BitBtn3.Enabled := True;
    BitBtn4.Enabled := True;
    BitBtn5.Enabled := True;
    BitBtn6.Enabled := True;
    BitBtn7.Enabled := True;
    CheckBox1.Enabled := True;
    EditAlign1.ReadOnly := False;
    EditAlign2.ReadOnly := False;
  end;
end;

end.

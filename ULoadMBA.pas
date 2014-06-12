unit ULoadMBA;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ComCtrls, Buttons, MyGrids, SortGrid, DateUtils,
  ExtCtrls, DB, IBDatabase, IBCustomDataSet, IBQuery;

type
  TForm6 = class(TForm)
    SortGrid1: TSortGrid;
    BitBtn1: TBitBtn;
    OpenDialog1: TOpenDialog;
    DateTimePicker3: TDateTimePicker;
    Label3: TLabel;
    BitBtn2: TBitBtn;
    LabeledEdit1: TLabeledEdit;
    LabeledEdit2: TLabeledEdit;
    IBTransaction1: TIBTransaction;
    IBQuery1: TIBQuery;
    IBDatabase1: TIBDatabase;
    BitBtn6: TBitBtn;
    BitBtn3: TBitBtn;
    StatusBar1: TStatusBar;
    procedure BitBtn1Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure SortGrid1GetCellFormat(Sender: TObject; Col, Row: Integer;
      State: TGridDrawState; var FormatOptions: TFormatOptions;
      var CheckBox, Combobox, Ellipsis: Boolean);
    procedure SortGrid1SetChecked(Sender: TObject; ACol, ARow: Integer;
      State: Boolean);
    procedure BitBtn6Click(Sender: TObject);
    procedure DateTimePicker3Change(Sender: TObject);
    procedure BitBtn3Click(Sender: TObject);
    procedure BitBtn2Click(Sender: TObject);
  private
    { Private declarations }
    RowChecked, RowError: array of Integer;
    SumCount: Integer;
    procedure SummaryStatusBar;
    function CheckERROR: Integer;
  public
    { Public declarations }
  end;

var
  Form6: TForm6;

implementation

uses Types, ComObj, Math, USelectSheetMBA, UProgress;

{$R *.dfm}

//***** http://www.swissdelphicenter.ch/torry/showcode.php?id=872
function RightPad(S: string; Ch: Char; Len: Integer): string;
var
  RestLen: Integer;
begin
  Result  := S;
  RestLen := Len - Length(s);
  if RestLen < 1 then Exit;
  Result := StringOfChar(Ch, RestLen) + S;
end;

//***** http://www.swissdelphicenter.ch/torry/showcode.php?id=58
function IsStrANumber(const S: String): Boolean;
var
  P: PChar;
begin
  P      := PChar(S);
  Result := False;
  while P^ <> #0 do
  begin
    if not (P^ in ['0'..'9']) then Exit;
    Inc(P);
  end;
  Result := True;
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

procedure TForm6.SummaryStatusBar;
var
  SumTotal: Double;
  iRow: Integer;
begin
  SumTotal := 0.00; SumCount := 0;
  for iRow := 1 to SortGrid1.RowCount-1 do
    if (not SortGrid1.CellChecked[0, iRow]) then begin
      SumTotal := SumTotal + StrToFloat(TrimThousand(SortGrid1.Cells[6, iRow]));
      SumCount := SumCount + 1;
    end;
  StatusBar1.Panels[1].Text := 'Count: ' + IntToStr(SumCount);
  StatusBar1.Panels[2].Text := 'Total Sum: ' + FormatFloat('#,##0.00;-#,##0.00', SumTotal);
end;

function TForm6.CheckERROR: Integer;
var
  iRow: Integer;
begin
  SetLength(RowError, 0);
  for iRow := 1 to SortGrid1.RowCount-1 do
      if (not SortGrid1.CellChecked[0, iRow]) then
        if (Length(SortGrid1.Cells[8, iRow]) = 0) or (Length(SortGrid1.Cells[10, iRow]) = 0) // if 'Rek.Tujuan' or 'KliringCode' is blank
          or (CompareDate(StrToDate(SortGrid1.Cells[1, iRow]), Now) = LessThanValue) then      // or ScheduleDate is less than Today
            if not InArray(iRow, RowError) then begin
              SetLength(RowError, Length(RowError)+1);
              RowError[Length(RowError)-1] := iRow;
            end;
  Result := Length(RowError);
  StatusBar1.Panels[0].Text := 'Error: ' + IntToStr(Result);
end;

procedure TForm6.FormShow(Sender: TObject);
begin
  DateTimePicker3.Date := IncDay(Now, IfThen(DayOfTheWeek(Now) > 4, 9 - DayOfTheWeek(Now), DayOfTheWeek(Now) mod 2));

  SortGrid1.ColCount := 11;
  SortGrid1.RowCount := 2;

  SortGrid1.Cols[0].Text := 'Delete';
  SortGrid1.Cols[1].Text := 'Schedule Date';
  SortGrid1.Cols[2].Text := 'Number';
  SortGrid1.Cols[3].Text := 'Supplier';
  SortGrid1.Cols[4].Text := 'Employee';
  SortGrid1.Cols[5].Text := 'Department';
  SortGrid1.Cols[6].Text := 'Total';
  SortGrid1.Cols[7].Text := 'Code';
  SortGrid1.Cols[8].Text := 'Bank Account';
  SortGrid1.Cols[9].Text := 'Bank Name';
  SortGrid1.Cols[10].Text := 'Bank Kliring';

  LabeledEdit1.Clear;
  LabeledEdit2.Clear;
  DateTimePicker3.Enabled := False;
  BitBtn2.Enabled := False;
  BitBtn3.Enabled := False;
  BitBtn6.Enabled := False;

  StatusBar1.Panels[0].Text := 'Error: 0';
  StatusBar1.Panels[1].Text := 'Count: 0';
  StatusBar1.Panels[2].Text := 'Total Sum: ' + FormatFloat('#,##0.00', 0);
end;

procedure TForm6.SortGrid1GetCellFormat(Sender: TObject; Col, Row: Integer;
  State: TGridDrawState; var FormatOptions: TFormatOptions; var CheckBox,
  Combobox, Ellipsis: Boolean);
begin
  if (Col = 0) and (Row <> 0) then CheckBox := True;
  if (Col in [1,2,7]) then FormatOptions.AlignmentHorz := taCenter;
  if (Col = 6) then FormatOptions.AlignmentHorz := taRightJustify;
  if (Row mod 2 > 0) then FormatOptions.Brush.Color := $00FFF7EC;
  if InArray(Row, RowChecked) then begin
    FormatOptions.Font.Style := [fsStrikeOut];
    FormatOptions.Brush.Color := clSilver;
  end;
  if InArray(Row, RowError) then FormatOptions.Brush.Color := clRed;
  if (Row = 0) then FormatOptions.AlignmentHorz := taCenter;
end;

procedure TForm6.SortGrid1SetChecked(Sender: TObject; ACol, ARow: Integer;
  State: Boolean);
var
  i: Integer;
begin
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
  SummaryStatusBar;
  CheckERROR;
end;

procedure TForm6.BitBtn1Click(Sender: TObject);
const
  xlCellTypeLastCell = 11;
var
  iExcel: OleVariant;
  iSheet: Variant;
  i, maxSheet, maxRow: Integer;
  nameSheet, empID, empName, empTotal: String;
begin
  if not OpenDialog1.Execute then Exit;

  try
    iExcel := CreateOleObject('Excel.Application');
  except
    ShowMessage('Cannot start Excel'#13'Is Excel not installed ?');
    Exit;
  end;

  iExcel.Visible := False;
  iExcel.Workbooks.Open(OpenDialog1.FileName);
  maxSheet := iExcel.ActiveWorkbook.Sheets.Count;
  Form7.Edit1.Text := iExcel.ActiveWorkbook.Name;
  Form7.ComboBox1.Clear;
  for i := 1 to maxSheet do
    Form7.ComboBox1.Items.Add(iExcel.ActiveWorkbook.Sheets[i].Name);
  Form7.ComboBox1.ItemIndex := maxSheet-1;

  if Form7.ShowModal <> mrOk then begin
    iExcel.Quit;
    Exit;
  end;

  nameSheet := Form7.ComboBox1.Text;
  iSheet := iExcel.ActiveWorkbook.WorkSheets[nameSheet];
  iSheet.Activate;
  iSheet.Cells.SpecialCells(xlCellTypeLastCell).Activate;
  maxRow := iExcel.ActiveCell.Row;

  LabeledEdit1.Text := Form7.Edit1.Text;
  LabeledEdit2.Text := nameSheet;
  Form4.ProgressBar1.Max := maxRow;
  Form4.ProgressBar1.Position := 0;

  SortGrid1.Clear;
  SortGrid1.ColCount := 11;
  SortGrid1.RowCount := 2;

  SetLength(RowChecked, 0);
  SetLength(RowError, 0);

  Form4.Show;
  for i := 1 to maxRow do begin
    Application.ProcessMessages;

    empID := iSheet.Cells[i, 3];
    empName := iSheet.Cells[i, 4];
    empTotal := iSheet.Cells[i, 9];
    if (Trim(empID) <> '') and (Trim(empTotal) <> '') and IsStrANumber(empTotal) then begin
      IBDatabase1.Close;
      IBQuery1.SQL.Text := 'SELECT supplier, note FROM syssupplier WHERE name LIKE ''%(' + empID + ')''';// OR name LIKE ''' + StringReplace(empName, '''', '''''', [rfReplaceAll, rfIgnoreCase]) + '%''';
      IBQuery1.Open;

      SortGrid1.Cells[1, SortGrid1.RowCount-1] := DateToStr(DateTimePicker3.Date);
      SortGrid1.Cells[2, SortGrid1.RowCount-1] := FormatDateTime('mmddhhnnss', Now) + RightPad(iSheet.Cells[i, 1], '0', 3);  //number
      SortGrid1.Cells[3, SortGrid1.RowCount-1] := TrimSpace(BreakdownNote(IBQuery1.FieldByName('note').AsString).Values['Name']);  //beneficiary name
      SortGrid1.Cells[4, SortGrid1.RowCount-1] := empID + ' - ' + empName;  //employee name
      SortGrid1.Cells[5, SortGrid1.RowCount-1] := iSheet.Cells[i, 5];  //department
      SortGrid1.Cells[6, SortGrid1.RowCount-1] := FormatFloat('#,##0;-#,##0', StrToFloat(empTotal));  //total
      SortGrid1.Cells[7, SortGrid1.RowCount-1] := IBQuery1.FieldByName('supplier').AsString;  //supplier code
      SortGrid1.Cells[8, SortGrid1.RowCount-1] := BreakdownNote(IBQuery1.FieldByName('note').AsString).Values['Acc'];  //iSheet.Cells[i, 7]; //bank account
      SortGrid1.Cells[9, SortGrid1.RowCount-1] := TrimSpace(BreakdownNote(IBQuery1.FieldByName('note').AsString).Values['Bank']);  //iSheet.Cells[i, 6]; //bank name
      SortGrid1.Cells[10,SortGrid1.RowCount-1] := BreakdownNote(IBQuery1.FieldByName('note').AsString).Values['Kliring'];  //iSheet.Cells[i,10]; //bank kliring code

      SortGrid1.RowCount := SortGrid1.RowCount + 1;
    end;
    Form4.ProgressBar1.Position := i;
  end;
  Form4.Hide;
  SortGrid1.RowCount := SortGrid1.RowCount - 1;
  IBDatabase1.Close;

  if not VarIsEmpty(iExcel) then
  begin
    iExcel.DisplayAlerts := False;  // disable any alert from the Excel
    iExcel.Quit;                    // now quit
    iExcel := Unassigned;           // then unassign all OLE Object has been created
    iSheet := Unassigned;
  end;

  SortGrid1.ColWidths[0] := 50;
  SortGrid1.ColWidths[1] := 120;
  SortGrid1.ColWidths[2] := 130;
  SortGrid1.ColWidths[3] := 350;
  SortGrid1.ColWidths[4] := 250;
  SortGrid1.ColWidths[5] := 200;
  SortGrid1.ColWidths[6] := 150;
  SortGrid1.ColWidths[7] := 70;
  SortGrid1.ColWidths[8] := 150;
  SortGrid1.ColWidths[9] := 300;
  SortGrid1.ColWidths[10] := 80;

  DateTimePicker3.Enabled := True;
  BitBtn2.Enabled := True;
  BitBtn3.Enabled := True;
  BitBtn6.Enabled := True;

  SummaryStatusBar;
  if CheckERROR > 0 then begin
    SortGrid1.Repaint;
    MessageDlg('Terdapat ' + IntToStr(Length(RowError)) + ' item yang bermasalah.'#13'Mohon perbaiki atau coret dari daftar Pembayaran terlebih dahulu.', mtError, [mbOK], 0);
  end;
end;

procedure TForm6.BitBtn2Click(Sender: TObject);
begin
  if CheckERROR > 0 then begin
    MessageDlg('Terdapat ' + IntToStr(Length(RowError)) + ' item yang bermasalah.'#13'Mohon perbaiki atau coret dari daftar Pembayaran terlebih dahulu.', mtError, [mbOK], 0);
    Exit;
  end;
  if SumCount = 0 then begin
    MessageDlg('Tidak ada item terpilih.'#13'Pilih minimal salah satu untuk menambahkan ke daftar Pembayaran.', mtWarning, [mbOK], 0);
    Exit;
  end;
  ModalResult := mrOk;
end;

procedure TForm6.BitBtn3Click(Sender: TObject);
var
  iRow: Integer;
begin
  for iRow := 1 to SortGrid1.RowCount-1 do
    SortGrid1.CellChecked[0, iRow] := True;
  SortGrid1.Repaint;
end;

procedure TForm6.BitBtn6Click(Sender: TObject);
var
  iRow: Integer;
begin
  for iRow := 1 to SortGrid1.RowCount-1 do
    SortGrid1.CellChecked[0, iRow] := False;
  SortGrid1.Repaint;
end;

procedure TForm6.DateTimePicker3Change(Sender: TObject);
var iRow: Integer;
begin
  for iRow := 1 to SortGrid1.RowCount-1 do
    SortGrid1.Cells[1, iRow] := DateToStr(DateTimePicker3.Date);
  CheckERROR;
  SortGrid1.Repaint;
end;

end.

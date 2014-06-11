unit ULogin;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, IBDatabase, DB, IBCustomDataSet, IBQuery, StdCtrls, ExtCtrls,
  jpeg, IniFiles;

type
  TForm2 = class(TForm)
    Image1: TImage;
    Label1: TLabel;
    Image2: TImage;
    GroupBox1: TGroupBox;
    LabeledEdit1: TLabeledEdit;
    LabeledEdit2: TLabeledEdit;
    Button1: TButton;
    Button2: TButton;
    IBQuery1: TIBQuery;
    IBDatabase1: TIBDatabase;
    IBTransaction1: TIBTransaction;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure Button2Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure LabeledEdit1KeyPress(Sender: TObject; var Key: Char);
    procedure LabeledEdit2KeyPress(Sender: TObject; var Key: Char);
    procedure LabeledEdit1_2KeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
  private
    { Private declarations }
    dBTemplate, EquivUser: String;
    procedure LoadDBConfig;
  public
    { Public declarations }
    Database, Username: String;
    EmailUser: String;
  end;

var
  Form2: TForm2;

implementation

{$R *.dfm}

procedure TForm2.LoadDBConfig;
var
  AppINI: TIniFile;
begin
  AppINI := TIniFile.Create('.\PowerAP.ini');
  try
    Database := AppINI.ReadString('Application', 'Database', '');
    dBTemplate := AppINI.ReadString('Application', 'dbRPT', '');
    AppINI.Free;

    AppINI := TIniFile.Create(ChangeFileExt(Application.ExeName, '.ini'));
    EquivUser := AppINI.ReadString('User', 'Equiv', '= BB =');
    EmailUser := AppINI.ReadString('User', 'Email', '');
  finally
    AppINI.Free;
  end;
end;

procedure TForm2.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  IBDatabase1.Close;
end;

procedure TForm2.FormShow(Sender: TObject);
begin
  Button1.Enabled := False;
  LabeledEdit1.Clear;
  LabeledEdit2.Clear;
  LabeledEdit1.SetFocus;
end;

procedure TForm2.LabeledEdit1KeyPress(Sender: TObject; var Key: Char);
begin
  if Key in ['a'..'z'] then Key := UpCase(Key);
  if (Ord(Key) = VK_RETURN) AND Button1.Enabled then
  begin
    Key := #0;
    Button1.Click;
  end;
end;

procedure TForm2.LabeledEdit2KeyPress(Sender: TObject; var Key: Char);
begin
  if (Ord(Key) = VK_RETURN) AND Button1.Enabled then
  begin
    Key := #0;
    Button1.Click;
  end;
end;

procedure TForm2.LabeledEdit1_2KeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  Button1.Enabled := (Length(LabeledEdit1.Text) > 0) AND (Length(LabeledEdit2.Text) > 0);
end;

procedure TForm2.Button2Click(Sender: TObject);
begin
  ModalResult := mrCancel;
end;

procedure TForm2.Button1Click(Sender: TObject);
begin
  LoadDBConfig;
  IBDatabase1.Close;
  IBDatabase1.DatabaseName := dBTemplate;
  //IBQuery1.SQL.Text := 'SELECT userid FROM userid ' +
  //  'WHERE userid = ''' + LabeledEdit1.Text + ''' AND pswd = ''' + LabeledEdit2.Text + ''' AND equivalentuser = ''' + EquivUser + ''' AND useractive = 1';

  // ***** Allowed login is user on AP that has granted rights on "Activities" --> "Cash Payment Status" to "Issued" *****
  IBQuery1.SQL.Text := 'SELECT u.userid FROM userid u LEFT JOIN useridform f ON f.userid = u.equivalentuser ' +
    'AND f.applid = ''AP'' AND f.moduleid = ''Activities'' AND f.formname = ''TfmCashStatus'' ' +
    'WHERE u.userid = ''' + LabeledEdit1.Text + ''' AND u.pswd = ''' + LabeledEdit2.Text + ''' AND f.rights LIKE ''%|S|%'' AND u.useractive = 1';

  IBDatabase1.Open;
  IBQuery1.Open;
  if IBQuery1.RecordCount = 0 then
  begin
    MessageDlg('Access denied or has no granted access!'#13'(Only User or Equivalent that can "Issued" on "AP - Cash Payment Status")'#13#13'Please check your Password or contact IS to check your Granted Rights!!', mtError, [mbOK], 0);
    IBDatabase1.Close;
  end
  else
  begin
    Username := IBQuery1.FieldByName('userid').AsString;
    IBDatabase1.Close;
    ModalResult := mrOk;
  end;
end;

end.

unit USetting;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, IniFiles;

type
  TForm3 = class(TForm)
    GroupBox1: TGroupBox;
    LabeledEdit1: TLabeledEdit;
    LabeledEdit2: TLabeledEdit;
    Button1: TButton;
    Button2: TButton;
    LabeledEdit3: TLabeledEdit;
    LabeledEdit4: TLabeledEdit;
    procedure FormShow(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    BankCustomer, BankAccount, BankName, BankCurrency: String;
    procedure LoadBankCompany;
  end;

var
  Form3: TForm3;

implementation

{$R *.dfm}

procedure TForm3.LoadBankCompany;
var
  IniFile: TIniFile;
begin
  IniFile := TIniFile.Create(ChangeFileExt(Application.ExeName, '.ini'));
  try
    BankCustomer := IniFile.ReadString('Bank', 'Customer', '');
    BankAccount := IniFile.ReadString('Bank', 'Account', '');
    BankName := IniFile.ReadString('Bank', 'Name', '');
    BankCurrency := IniFile.ReadString('Bank', 'Currency', '');
  finally
    IniFile.Free;
  end;
end;

procedure TForm3.FormShow(Sender: TObject);
begin
  LoadBankCompany;
  LabeledEdit1.Text := BankCustomer;
  LabeledEdit2.Text := BankAccount;
  LabeledEdit3.Text := BankName;
  LabeledEdit4.Text := BankCurrency;
end;

procedure TForm3.Button1Click(Sender: TObject);
var
  IniFile: TIniFile;
begin
  IniFile := TIniFile.Create(ChangeFileExt(Application.ExeName, '.ini'));
  try
    IniFile.WriteString('Bank', 'Customer', LabeledEdit1.Text);
    IniFile.WriteString('Bank', 'Account', LabeledEdit2.Text);
    IniFile.WriteString('Bank', 'Name', LabeledEdit3.Text);
    IniFile.WriteString('Bank', 'Currency', LabeledEdit4.Text);
    BankCustomer := LabeledEdit1.Text;
    BankAccount := LabeledEdit2.Text;
    BankName := LabeledEdit3.Text;
    BankCurrency := LabeledEdit4.Text;
  finally
    IniFile.Free;
  end;
  ModalResult := mrOk;
end;

procedure TForm3.Button2Click(Sender: TObject);
begin
  ModalResult := mrCancel;
end;

end.

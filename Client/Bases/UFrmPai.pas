unit UFrmPai;

interface

uses
  Windows,
  Messages,
  SysUtils,
  Variants,
  Graphics,
  Controls,
  Forms,
  Dialogs,
  StdCtrls,
  DBClient;

type
  TFrmPai = class(TForm)
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    fApenasDigitacaoMaiusculo: Boolean;
    fControlaMaximizar: Boolean;
  public
    property ApenasDigitacaoMaiusculo: Boolean read fApenasDigitacaoMaiusculo write fApenasDigitacaoMaiusculo;
    property ControlaMaximizar: Boolean read fControlaMaximizar write fControlaMaximizar;
  end;

var
  FrmPai: TFrmPai;

implementation

uses Constantes, ClassDataSet, UDMConexao;

{$R *.dfm}

procedure TFrmPai.FormCreate(Sender: TObject);
begin
  ApenasDigitacaoMaiusculo  := True;

  //The form remains the size you left it at design time, but
  //is positioned in the center of the form specified by the Owner property.
  //If the Owner property does not specify a form,
  //this position acts like poMainFormCenter.
  Position := poOwnerFormCenter;
end;

procedure TFrmPai.FormClose(Sender: TObject; var Action: TCloseAction);
var
  X: integer;
begin
  with Self do
    for X := 0 to ComponentCount - 1 do
    begin
      if (Components[X] is TClientDataSet) then
        (Components[X] as TClientDataSet).Close;
    end;
end;

end.


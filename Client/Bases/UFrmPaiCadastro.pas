unit UFrmPaiCadastro;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, UFrmPai, Vcl.ExtCtrls, Vcl.ComCtrls;

type
  TFrmPaiCadastro = class(TFrmPai)
    pnlBot: TPanel;
    pnlTop: TPanel;
    pnlButtons: TPanel;
    tbctrlCadastro: TTabControl;
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  FrmPaiCadastro: TFrmPaiCadastro;

implementation

{$R *.dfm}

end.

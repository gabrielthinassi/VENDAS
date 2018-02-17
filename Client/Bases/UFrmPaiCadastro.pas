unit UFrmPaiCadastro;

interface

uses
  Winapi.Windows,
  Winapi.Messages,
  System.SysUtils,
  System.Variants,
  System.Classes,
  Vcl.Graphics,
  Vcl.Controls,
  Vcl.Forms,
  Vcl.Dialogs,
  UFrmPai,
  Vcl.ExtCtrls,
  Vcl.ComCtrls,
  Vcl.StdCtrls,
  Vcl.Mask,
  JvExMask,
  JvToolEdit,
  JvBaseEdits,
  Vcl.Buttons,
  Data.DB,
  Vcl.DBCtrls,
  UDMPaiCadastro;

type
  TFrmPaiCadastro = class(TFrmPai)
    pnlBot: TPanel;
    pnlTop: TPanel;
    pnlButtons: TPanel;
    edtCodigo: TJvCalcEdit;
    btnRelatorio: TSpeedButton;
    btnIncluir: TSpeedButton;
    btnExcluir: TSpeedButton;
    btnGravar: TSpeedButton;
    btnCancelar: TSpeedButton;
    btnPesquisar: TSpeedButton;
    DSCadastro: TDataSource;
    txtAjuda: TDBText;
    pnlNavegar: TPanel;
    btnUltimo: TSpeedButton;
    btnProximo: TSpeedButton;
    btnAnterior: TSpeedButton;
    btnPrimeiro: TSpeedButton;
    pgctrlCadastro: TPageControl;
    pgEdit: TTabSheet;
    procedure FormCreate(Sender: TObject);
    procedure btnIncluirClick(Sender: TObject);
    procedure btnExcluirClick(Sender: TObject);
    procedure btnGravarClick(Sender: TObject);
    procedure btnCancelarClick(Sender: TObject);
    procedure btnPrimeiroClick(Sender: TObject);
    procedure btnAnteriorClick(Sender: TObject);
    procedure btnProximoClick(Sender: TObject);
    procedure btnUltimoClick(Sender: TObject);
    procedure DSCadastroStateChange(Sender: TObject);
  private
    { Private declarations }
    FDMCadastro: TDMPaiCadastro;
  public
    { Public declarations }
  end;

var
  FrmPaiCadastro: TFrmPaiCadastro;

implementation

{$R *.dfm}


procedure TFrmPaiCadastro.btnAnteriorClick(Sender: TObject);
begin
  inherited;
  edtCodigo.AsInteger := FDMCadastro.Anterior(edtCodigo.AsInteger);
end;

procedure TFrmPaiCadastro.btnCancelarClick(Sender: TObject);
begin
  inherited;

  if FDMCadastro.CDSCadastro.State in [dsEdit, dsInsert] then
  begin
    if MessageDlg('Deseja realmente cancelar?',mtWarning,mbYesNo,0) = mrYes then
      FDMCadastro.CDSCadastro.Cancel;
  end;
  FDMCadastro.CDSCadastro.Close;
end;

procedure TFrmPaiCadastro.btnExcluirClick(Sender: TObject);
begin
  inherited;

  if ((not FDMCadastro.CDSCadastro.Active) or (FDMCadastro.CDSCadastro.IsEmpty)) then
  begin
    ShowMessage('Não há o que excluir!');
  end
  else
  begin
    if MessageDlg('Deseja deletar o registro?', mtConfirmation, mbYesNo,0) = mrYes then
      Exit;

    if FDMCadastro.CDSCadastro.State in [dsEdit, dsInsert] then
      Exit;

    try
      FDMCadastro.CDSCadastro.Delete;
      FDMCadastro.CDSCadastro.ApplyUpdates(0);
    except
      FDMCadastro.CDSCadastro.UndoLastChange(True);
      FDMCadastro.CDSCadastro.Close;
      FDMCadastro.CDSCadastro.Open;
      raise Exception.Create('Ocorreram erros ao tentar deletar o Registro!');
    end;
  end;
end;

procedure TFrmPaiCadastro.btnGravarClick(Sender: TObject);
begin
  inherited;

  if FDMCadastro.CDSCadastro.State in [dsEdit, dsInsert] then
    FDMCadastro.CDSCadastro.Post;

  try
    FDMCadastro.CDSCadastro.ApplyUpdates(0);
  except
    raise Exception.Create('Ocorreram erros ao tentar gravar o Registro!');
  end;
end;

procedure TFrmPaiCadastro.btnIncluirClick(Sender: TObject);
begin
  inherited;

  if FDMCadastro.CDSCadastro.State in [dsEdit, dsInsert] then
    Exit;

  try
    edtCodigo.AsInteger            := 0;
    pgctrlCadastro.ActivePageIndex := 0;

    FDMCadastro.CDSCadastro.Close;
    FDMCadastro.CDSCadastro.Open;
    FDMCadastro.CDSCadastro.Insert;
  except
    raise Exception.Create('Ocorreram erros ao tentar incluir o Registro!');
  end;
end;

procedure TFrmPaiCadastro.btnPrimeiroClick(Sender: TObject);
begin
  inherited;
  edtCodigo.AsInteger := FDMCadastro.Primeiro;
end;

procedure TFrmPaiCadastro.btnProximoClick(Sender: TObject);
begin
  inherited;
  edtCodigo.AsInteger := FDMCadastro.Proximo(edtCodigo.AsInteger);
end;

procedure TFrmPaiCadastro.btnUltimoClick(Sender: TObject);
begin
  inherited;
  edtCodigo.AsInteger := FDMCadastro.Ultimo;
end;

procedure TFrmPaiCadastro.DSCadastroStateChange(Sender: TObject);
begin
  inherited;
  if FDMCadastro = nil then
    Exit;

  btnIncluir.Enabled   := (DSCadastro.DataSet.State in [dsBrowse, dsInactive]);
  btnExcluir.Enabled := (DSCadastro.DataSet.State in [dsBrowse]) and DSCadastro.DataSet.IsEmpty;
  btnGravar.Enabled := (DSCadastro.DataSet.State in [dsInsert, dsEdit]);
  btnCancelar.Enabled := (DSCadastro.DataSet.State in [dsInsert, dsEdit]);
  btnPesquisar.Enabled := (DSCadastro.DataSet.State in [dsBrowse, dsInactive]);
  btnRelatorio.Enabled := (DSCadastro.DataSet.State in [dsBrowse, dsInactive]);
end;

procedure TFrmPaiCadastro.FormCreate(Sender: TObject);
begin
  inherited;
  if pgctrlCadastro.PageCount = 1 then
  begin
    pgEdit.TabVisible := False;
    pgEdit.TabStop := False;
  end;
  pgctrlCadastro.ActivePageIndex := 0;

  DSCadastro.DataSet := DMPaiCadastro.CDSCadastro;
end;

end.

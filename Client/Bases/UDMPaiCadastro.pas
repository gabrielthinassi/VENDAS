unit UDMPaiCadastro;

interface

uses
  Windows,
  Messages,
  SysUtils,
  Variants,
  Classes,
  Graphics,
  Controls,
  Forms,
  Dialogs,
  DB,
  DBClient,
  MConnect,
  Typinfo,
  DateUtils,
  DBCtrls,
  Grids,
  DBGrids,
  JvDBGrid,
  Datasnap.DSConnect,
  System.StrUtils,
  System.Math,
  UDMPai,
  ClassPaiCadastro,
  Data.DBXDataSnap,
  Data.DBXCommon,
  IPPeerClient,
  Data.SqlExpr;

type
  TDMPaiCadastro = class(TDMPai)
    CDSCadastro: TClientDataSet;
    DSPCCadastro: TDSProviderConnection;

    procedure DataModuleCreate(Sender: TObject);
    procedure DataModuleDestroy(Sender: TObject);

  private
    FClasseFilha: TClassPaiCadastro;

    function GetIdDetalhe: integer;
    function ProximoCodigo(Tabela: string): int64;
  protected
  public
    property ClasseFilha: TClassPaiCadastro read FClasseFilha write FClasseFilha;

    //property CodigoAtual: integer read FCodigoAtual write FCodigoAtual;
    //property RefreshRecordAfterPost: boolean read FRefreshRecordAfterPost write FRefreshRecordAfterPost;

    property IdDetalhe: integer read GetIdDetalhe;

    function GetClassNameClasseFilha: string;

    function Primeiro: integer; virtual;
    function Anterior(Atual: integer): integer; virtual;
    function Proximo(Atual: integer): integer; virtual;
    function Ultimo: integer; virtual;
    function Novo: integer; virtual;

    function Aplique(Exclusao: boolean = false): integer; virtual;

    procedure AtribuiAutoIncDetalhe(DataSet: TDataSet; Classe: TClassPaiCadastro; CampoChaveEstrangeira: String);
  end;

var
  DMPaiCadastro: TDMPaiCadastro;

implementation

{ %CLASSGROUP 'Vcl.Controls.TControl' }

uses Constantes, ClassDataSet, UDMConexao;

{$R *.dfm}


procedure TDMPaiCadastro.DataModuleCreate(Sender: TObject);
begin
  inherited;
  DSPCCadastro.SQLConnection := DMConexao.ConexaoDS;

  if DSPCCadastro.ServerClassName <> '' then
    CDSCadastro.RemoteServer := DSPCCadastro;
  CDSCadastro.ProviderName := 'DSPCadastro';

  with FClasseFilha do
  begin
    //if Trim(CDSCadastro.CommandText) = '' then
    //  CDSCadastro.CommandText := SQLBaseCadastro;

    CDSCadastro.AdicionarCampos;
  end;
end;

procedure TDMPaiCadastro.DataModuleDestroy(Sender: TObject);
begin
  inherited;
  DSPCCadastro.Close;
  DSPCCadastro.SQLConnection := nil;
end;

function TDMPaiCadastro.Primeiro: integer;
var
  SQL: string;
  Classe: TClassPaiCadastro;
begin
  Classe := FClasseFilha;

  with Classe do
  begin
    SQL := ' SELECT MIN(' + TabelaPrincipal + '.' + CampoChave + ')' +#13+
           ' FROM '  + TabelaPrincipal +#13+
           ' WHERE ' + TabelaPrincipal + '.' + CampoChave + ' <> 0';
  end;
  Result := DMConexao.ExecuteScalar(SQL);
end;

function TDMPaiCadastro.Proximo(Atual: integer): integer;
var
  SQL: string;
  Classe: TClassPaiCadastro;
begin
  Classe := FClasseFilha;

  with Classe do
  begin
    SQL := ' SELECT MIN(' + TabelaPrincipal + '.' + CampoChave + ')' +#13+
           ' FROM '  + TabelaPrincipal +#13+
           ' WHERE ' + TabelaPrincipal + '.' + CampoChave + ' > ' + IntToStr(Atual) +#13+
           ' AND '   + TabelaPrincipal + '.' + CampoChave + ' <> 0)';
  end;
  Result := DMConexao.ExecuteScalar(SQL);
end;


function TDMPaiCadastro.ProximoCodigo(Tabela: string): int64;
begin
  //Busca o Próximo Código no Servidor de Aplicação
  Tabela := AnsiUpperCase(Tabela);
  Result := DMConexao.ExecutaMetodo('TSMConexao.ProximoCodigo', [Tabela]);
end;

function TDMPaiCadastro.Anterior(Atual: integer): integer;
var
  SQL: string;
  Classe: TClassPaiCadastro;
begin
  Classe := FClasseFilha;

  with Classe do
  begin
    SQL := ' SELECT MAX(' + TabelaPrincipal + '.' + CampoChave + ')' +#13+
           ' FROM '  + TabelaPrincipal +#13+
           ' WHERE ' + TabelaPrincipal + '.' + CampoChave + ' < ' + IntToStr(Atual) +#13+
           ' AND '   + TabelaPrincipal + '.' + CampoChave + ' <> 0)';
  end;
  Result := DMConexao.ExecuteScalar(SQL);
end;

function TDMPaiCadastro.Ultimo: integer;
var
  SQL: string;
  Classe: TClassPaiCadastro;
begin
  Classe := FClasseFilha;

  with Classe do
  begin
    SQL := ' SELECT MAX(' + TabelaPrincipal + '.' + CampoChave + ')' +#13+
           ' FROM '  + TabelaPrincipal +#13+
           ' WHERE ' + TabelaPrincipal + '.' + CampoChave + ' <> 0';
  end;
  Result := DMConexao.ExecuteScalar(SQL);
end;



procedure TDMPaiCadastro.VerificarFechamento(DataSet: TDataSet; Classe: TFClassPaiCadastro);
var
  ListaDeCampos: TStrings;
  i: integer;
  DataUltimoFechamento: TDateTime;

  {$IF DEFINED(DEPPESSOAL)}
  CodEmpresa: integer;
  {$IFEND}
begin
  with Classe do
  begin
    if CamposFechamento = '' then
      Exit;

    ListaDeCampos := TStringList.Create;
    try
      ExtractStrings([';'], [' '], PWideChar(CamposFechamento), ListaDeCampos);
      with DataSet do
      begin
        if not IsEmpty then
        begin
{$IF DEFINED(DEPPESSOAL)}
          if ConstanteSistema.Sistema in [ConstanteSistema.cSistemaDepPessoal] then
          begin
            if ListaDeCampos[0] = 'GERAL' then
              CodEmpresa := 0
            else if ListaDeCampos[0] = 'ATUAL' then
              CodEmpresa := DMConexao.SecaoAtual.Empresa.Estabelecimento
            else
              CodEmpresa := FieldByName(ListaDeCampos[0]).AsInteger;

            DataUltimoFechamento := DP_DataFechamento(DMConexao, CodEmpresa, 0, 0, True);
          end;
{$ELSE}
          DataUltimoFechamento := 0;
{$IFEND}
          for i := 0 to ListaDeCampos.Count - 1 do
          begin
            if (FindField(ListaDeCampos[i]) <> nil) and (FieldByName(ListaDeCampos[i]).DataType in [ftTimeStamp, ftDate, ftDateTime]) then
            begin
              if (FieldByName(ListaDeCampos[i]).AsDateTime > 0) and (Trunc(FieldByName(ListaDeCampos[i]).AsDateTime) <= Trunc(DataUltimoFechamento)) then
              begin
                TCaixasDeDialogo.Aviso('Alteração bloqueada para data anterior ao último fechamento - ' + FormatDateTime('dd/mm/yyyy', DataUltimoFechamento));
                Abort;
              end;
            end;
          end;
        end;
      end;
    finally
      FreeAndNil(ListaDeCampos);
    end;
  end;
end;

function TDMPaiCadastro.Novo: integer;
begin
  // Retorna o novo código a ser usado na inserção
  with FClasseFilha do
  begin
    Result := DMConexao.ProximoCodigo(TabelaPrincipal);

    if Result < ValorInicialCampoChave then
      Result := DMConexao.ProximoCodigoAcrescimo(TabelaPrincipal, Quebra, (ValorInicialCampoChave - Result));

    if (ValorFinalCampoChave > 0) and (Result > ValorFinalCampoChave) then
    begin
      ShowMessage('Limite de código superado para o cadastro ' + Descricao + '!');
      Abort;
    end;
  end;
end;

function TDMPaiCadastro.Aplique(Exclusao: boolean = false): integer;
begin
  // O parâmetro "Exclusao" é usado em funções herdadas, não apague.
  Result := CDSCadastro.ApplyUpdates(0);
end;

procedure TDMPaiCadastro.Cancele;
begin
  CDSCadastro.CancelUpdates;
end;

class procedure TDMPaiCadastro.ValidateGeral(Sender: TField; CampoDescricao: string; sClasseBusca: string; Filtro: string = ''; NomeCDSCache: string = ''; ValidaStatus: boolean = True; ValidaStatusBloqueio: boolean = False);
var
  ClasseBusca: TFClassPaiCadastro;
  Persist: TPersistentClass;
begin
  try
    Persist := FindClass(sClasseBusca);
  except
    raise Exception.Create('A classe ' + sClasseBusca + ' não está disponível nesse módulo, entre em contato com a Tek-System');
  end;
  ClasseBusca := TFClassPaiCadastro(Persist);
  ValidateGeral(Sender, CampoDescricao, ClasseBusca, Filtro, NomeCDSCache, ValidaStatus, ValidaStatusBloqueio);
end;

class procedure TDMPaiCadastro.ValidateGeral(Sender: TField; CampoDescricao: string; ClasseBusca: TFClassPaiCadastro; Filtro: string = ''; NomeCDSCache: string = ''; ValidaStatus: boolean = True; ValidaStatusBloqueio: boolean = False);
var
  ValidacaoAnt: TFieldNotifyEvent;
begin
  if not(Sender.DataSet.State in [dsEdit, dsInsert]) then
    Exit;
  with Sender do
  begin
    ValidacaoAnt := OnValidate;
    OnValidate := nil;
    try
      DataSet.FieldByName(CampoDescricao).AsString := DMLookup.BuscaDescricao(ClasseBusca, AsInteger, -1, ValidaStatus, Filtro, NomeCDSCache, ValidaStatusBloqueio);
    finally
      OnValidate := ValidacaoAnt;
    end;
  end;
end;

class procedure TDMPaiCadastro.ValidateGeral(Sender: TField; Campos: array of string;
  ClasseBusca: TFClassPaiCadastro; Empresa: integer = -1; Filtro: string = '');
var
  ValidacaoAnt: TFieldNotifyEvent;
begin
  if not(Sender.DataSet.State in [dsEdit, dsInsert]) then
    Exit;
  with Sender do
  begin
    ValidacaoAnt := OnValidate;
    OnValidate := nil;
    try
      DMLookup.BuscaCampos(ClasseBusca, Campos, DataSet, AsInteger, Empresa, Filtro);
    finally
      OnValidate := ValidacaoAnt;
    end;
  end;
end;

class procedure TDMPaiCadastro.ValidateGeral(CampoBusca: TField; CamposRetorno: array of string;
  NomeCampoWhere: string; ClasseBusca: TFClassPaiCadastro; Filtro: string = '');
var
  ValidacaoAnt: TFieldNotifyEvent;
begin
  if not(CampoBusca.DataSet.State in [dsEdit, dsInsert]) then
    Exit;
  with CampoBusca do
  begin
    ValidacaoAnt := OnValidate;
    OnValidate := nil;
    try
      DMLookup.BuscaCampos(ClasseBusca, NomeCampoWhere, CampoBusca, CamposRetorno, Filtro);
    finally
      OnValidate := ValidacaoAnt;
    end;
  end;
end;

function TDMPaiCadastro.GetClassNameClasseFilha: string;
begin
  if Assigned(FClasseFilha) then
    Result := FClasseFilha.ClassName
  else
    Result := '';
end;

function TDMPaiCadastro.GetIdDetalhe: integer;
begin
  FIdDetalhe := FIdDetalhe - 1;
  Result := FIdDetalhe;
end;

procedure TDMPaiCadastro.AtribuiAutoIncDetalhe(DataSet: TDataSet; Classe: TFClassPaiCadastro; CampoChaveEstrangeira: String);
var
  QtdeAutoIncNegativos: Integer;
  AutoIncDetalhe: Integer;
begin
  with DataSet, Classe do
  begin
    QtdeAutoIncNegativos := 0;
    DisableControls;

    try
      First;
      while not EOF do
      begin
        if FieldByName(CampoChave).AsInteger <= 0 then
          inc(QtdeAutoIncNegativos);
        Next;
      end;

      if QtdeAutoIncNegativos > 0 then
      begin
        AutoIncDetalhe := DMConexao.ProximoCodigoAcrescimo(TabelaPrincipal, 0, QtdeAutoIncNegativos);
        AutoIncDetalhe := AutoIncDetalhe - QtdeAutoIncNegativos;

        First;
        while not Eof do
        begin
          if FieldByName(CampoChave).AsInteger <= 0 then
          begin
            Edit;
            FieldByName(CampoChave).AsInteger := AutoIncDetalhe;
            FieldByName(CampoChaveEstrangeira).AsInteger := CDSCadastro.FieldByName(FClasseFilha.CampoChave).AsInteger;
            Post;
            inc(AutoIncDetalhe);
          end;
          Next;
        end;
      end;
    finally
      EnableControls;
    end;
  end;
end;

end.

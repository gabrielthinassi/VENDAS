unit USMConexao;

interface

uses
  System.SysUtils, System.Classes, System.Variants, Data.SqlExpr, DBXCommon, Data.FMTBcd,
  Winapi.Windows, Data.DB, DBClient, Provider, Vcl.Forms, Data.SqlConst, System.StrUtils, System.Math,
  Datasnap.DSServer, acQBBase, acQBdbExMetaProvider, acAST, acFbSynProvider, Graphics, jpeg, Generics.Collections,
  Data.DBXJSONReflect, Datasnap.DSCommonServer, Datasnap.DSProviderDataModuleAdapter,
  Data.DBXFirebird;

type
  TSMConexao = class(TDSServerModule)
    CDSProximoCodigo: TClientDataSet;
    DSPProximoCodigo: TDataSetProvider;
    SQLDSProximoCodigo: TSQLDataSet;
    ConexaoBD: TSQLConnection;
    procedure DSServerModuleDestroy(Sender: TObject);
  private
  public
    function ExecuteScalar(SQL: string; CriarTransacao: Boolean): OleVariant; virtual;
    function ExecuteReader(SQL: string; TamanhoPacote: Integer; CriarTransacao: Boolean): OleVariant; virtual;
    function ExecuteCommand(SQL: string; CriarTransacao: Boolean): OleVariant; virtual;
    function ExecuteCommand_Update(SQL, Campo: string; Valor: OleVariant; CriarTransacao: Boolean): OleVariant; virtual;
    function ProximoCodigo(Tabela: String): Int64; virtual;
    procedure AcertaProximoCodigo(NomeDaClasse: String; Quebra: Integer; CriarTransacao: Boolean = True); virtual;
  end;

var
  SMConexao: TSMConexao;

implementation

{$R *.dfm}


uses Constantes, ClassPaiCadastro;

procedure TSMConexao.DSServerModuleDestroy(Sender: TObject);
begin
  if Assigned(ConexaoBD) then
  begin
    if (ConexaoBD.Connected) then
      try
        ConexaoBD.Connected := False;
        ConexaoBD := nil;
      except
        // Implementar
      end;
  end;
  inherited;
end;

function TSMConexao.ExecuteScalar(SQL: string): OleVariant;
var
  DataSet   : TSQLDataSet;
  TipoCampo : TFieldType;
begin
  try
    try
      ConexaoBD.Execute(SQL, nil, TDataSet(DataSet));
      DataSet.ParamCheck := False; //Deixa um pouco mais rápido por não checar os Params.

      TipoCampo := DataSet.Fields[0].DataType;

      if (DataSet <> nil) and (not DataSet.eof) then
      case TipoCampo of
        ftSmallint, ftInteger:
          Result := DataSet.Fields[0].AsInteger;
        ftFloat, ftCurrency:
          Result := DataSet.Fields[0].AsCurrency;
        ftString, ftWideString:
          Result := DataSet.Fields[0].AsString;
        ftDate, ftTime, ftDateTime, ftTimeStamp:
          Result := DataSet.Fields[0].AsDateTime;
      else
          Result := DataSet.Fields[0].Value;
      end
      else
        Result := NULL;

    finally
      FreeAndNil(DataSet);
    end;

  except on E: Exception do
    begin
      ConexaoBD.Connected := False;
      raise Exception.Create('Erro na Execução: ' + #13 + E.Message);
    end;
  end;
end;

function TSMConexao.ExecuteReader(SQL: string): OleVariant;
var
  DataSet  : TSQLDataSet;
  Provider : TDataSetProvider;
begin
  try
    DataSet  := TSQLDataSet.Create(nil);
    Provider := TDataSetProvider.Create(nil);
    try
      DataSet.SQLConnection := ConexaoBD;
      DataSet.ParamCheck    := False; //Deixa um pouco mais rápido por não checar os Params.
      DataSet.CommandText   := SQL;
      DataSet.GetMetadata   := False; //Deixa um pouco mais rápido por não carregar o MetaData.

      Provider.Exported    := False; //Deixa um pouco mais rápido, desabilitando a chamada deste Provider em ClientApp.
      Provider.Constraints := False; //Deixa um pouco mais rápido, não enviando as Constraints.
      Provider.DataSet     := DataSet;

      Result      := Provider.Data;

    finally
      FreeAndNil(DataSet);
      FreeAndNil(Provider);
    end;
  except on E: Exception do
    begin
      ConexaoBD.Connected := False;
      raise Exception.Create('Erro na Execução: ' + #13 + E.Message);
    end;
  end;
end;

function TSMConexao.ExecuteCommand(SQL: string): OleVariant;
begin
  try
    try
      ConexaoBD.DBXConnection.CreateCommand.Text := SQL;
      ConexaoBD.DBXConnection.CreateCommand.ExecuteQuery;
      Result := IntToStr(ConexaoBD.DBXConnection.CreateCommand.RowsAffected);
    finally
      ConexaoBD.DBXConnection.CreateCommand.Free;
    end;
  except on E: Exception do
    begin
      ConexaoBD.Connected := False;
      raise Exception.Create('Erro na Execução: ' + #13 + E.Message);
    end;
  end;
end;

function TSMConexao.ExecuteCommand_Update(SQL, Campo: string; Valor: OleVariant; CriarTransacao: Boolean): OleVariant;
var
  DS            : TSQLDataSet;
  DSP           : TDataSetProvider;
  CDS           : TClientDataSet;
  TransacaoLocal: TDBXTransaction;
begin
{  try
    if CriarTransacao then
      TransacaoLocal := Conexao.BeginTransaction(TDBXIsolations.ReadCommitted)
    else if (not Conexao.InTransaction) then
      raise Exception.Create('Toda execução da função ExecuteCommand_Update deve estar dentro de um contexto transacional');

    try
      CDS := TClientDataSet.Create(Self);
      DSP := TDataSetProvider.Create(Self);
      DS  := TSQLDataSet.Create(Self);
      with DS do
      begin
        CommandText   := SQL;
        CommandType   := CtQuery;
        SQLConnection := Conexao;
      end;
      with DSP do
      begin
        DataSet := DS;
        Name    := 'DSProviderECU' + IntToStr(Random(100));
      end;
      with CDS do
      begin
        ProviderName := DSP.Name;
        Open;
        if IsEmpty then
          raise Exception.Create('Não encontrado registro para atualização do campo "' + Campo + '".');
        if RecordCount > 1 then
          raise Exception.Create('Encontrado multiplos registros para atualização do campo "' + Campo + '".');
        Edit;
        FieldByName(Campo).Value := Valor;
        Post;
        Result := ApplyUpdates(0);
      end;
    finally
      FreeAndNil(DS);
      FreeAndNil(DSP);
      FreeAndNil(CDS);

      if CriarTransacao and (Conexao.HasTransaction(TransacaoLocal)) then
        Conexao.CommitFreeAndNil(TransacaoLocal);
    end;
  except
    on E: Exception do
    begin
      if (Pos('Unable to complete network request to host', E.Message) > 0) or
        (Pos('Error writing data to the connection', E.Message) > 0) or
        (Pos('connection shutdown', E.Message) > 0) then
        Conexao.Connected := False
      else if CriarTransacao then
        if (Conexao.HasTransaction(TransacaoLocal)) then
          Conexao.RollbackFreeAndNil(TransacaoLocal)
        else
          Conexao.RollbackIncompleteFreeAndNil(TransacaoLocal);

      raise Exception.Create(FuncoesGeraisServidor.FormataErroNoServidor(Self.ClassName, 'ExecuteCommand_Update',
        'Comando SQL: ' + SQL + #13 + E.Message));
    end;
  end;
}end;


function TSMConexao.ProximoCodigo(Tabela: String): Int64;
const
  Acrescimo = 1;
var
  Transacao: TDBXTransaction;
begin
  if (CDSProximoCodigo.FieldDefs.Count = 0) then
  begin
    CDSProximoCodigo.Close;
    CDSProximoCodigo.FieldDefs.Clear;
    CDSProximoCodigo.FieldDefs.Add('TABELA_AUTOINC', ftString, 31);
    CDSProximoCodigo.FieldDefs.Add('CODIGO_AUTOINC', ftLargeint);
    CDSProximoCodigo.CreateDataSet;
  end;

  Transacao := ConexaoBD.BeginTransaction(TDBXIsolations.ReadCommitted);
  try
    repeat //similar ao DO WHILE
      CDSProximoCodigo.Close;
      CDSProximoCodigo.Params.ParamByName('TABELA').AsString  := Tabela;
      CDSProximoCodigo.Open;
      if CDSProximoCodigo.IsEmpty then
      begin
        CDSProximoCodigo.Insert;
        CDSProximoCodigo.FieldByName('TABELA_AUTOINC').AsString  := Tabela;
        CDSProximoCodigo.FieldByName('CODIGO_AUTOINC').AsInteger := Acrescimo;
      end
      else
      begin
        CDSProximoCodigo.Edit;
        CDSProximoCodigo.FieldByName('CODIGO_AUTOINC').AsInteger := CDSProximoCodigo.FieldByName('CODIGO_AUTOINC').AsInteger + Acrescimo;
      end;
      CDSProximoCodigo.Post;
    until (CDSProximoCodigo.ApplyUpdates(0) = 0);

    Result := CDSProximoCodigo.FieldByName('CODIGO_AUTOINC').AsLargeInt;
    CDSProximoCodigo.Close;

    if ConexaoBD.HasTransaction(Transacao) then
      ConexaoBD.CommitFreeAndNil(Transacao);
  except on E: Exception do
    begin
      if ConexaoBD.HasTransaction(Transacao) then
        ConexaoBD.RollbackFreeAndNil(Transacao);
      raise Exception.Create('Erro na Tabela: ' + Tabela + #13 + E.Message);
    end;
  end;
end;

{$REGION 'AcertaProximoCodigo'}

procedure TSMConexao.AcertaProximoCodigo(NomeDaClasse: String; Quebra: Integer; CriarTransacao: Boolean = True);
var
  SQL   : String;
  Codigo: Int64;
  CDS   : TClientDataSet;
  TD    : TDBXTransaction;
  Classe: TFClassPaiCadastro;
begin
{  if CriarTransacao then
    TD := Conexao.BeginTransaction(TDBXIsolations.ReadCommitted);
  try
    Classe := TFClassPaiCadastro(FindClass(NomeDaClasse));

    if (Classe.TabelaPrincipal = '') or
      (Classe.CampoChave = '') then
      Exit;

    // Montagem da SQL para pegar o último código
    SQL := ' select ';
    if (Classe.CampoEmpresa = '') then
      SQL := SQL + '0'
    else
      SQL := SQL + Classe.TabelaPrincipal + '.' + Classe.CampoEmpresa;
    SQL   := SQL + ' QUEBRA,' +
      ' max(' + Classe.TabelaPrincipal + '.' + Classe.CampoChave + ') ULTIMO' + #13 +
      ' from ' + Classe.TabelaPrincipal;

    if (Classe.CampoEmpresa <> '') then
    begin
      if Quebra > 0 then
        SQL := SQL + #13 +
          'where ' + Classe.TabelaPrincipal + '.' + Classe.CampoEmpresa + ' = ' + IntToStr(Quebra);

      SQL := SQL + #13 +
        'group by ' + Classe.TabelaPrincipal + '.' + Classe.CampoEmpresa;
    end;

    CDS := TClientDataSet.Create(Self);
    try
      with CDS do
      begin
        Data := ExecuteReader(SQL, -1, False);
        First;
        while (not eof) do
        begin
          Codigo := Max(TFuncoesConversao.VariantParaInt64(FieldByName('ULTIMO').Value), 0);
          if Codigo < (Classe.ValorInicialCampoChave - 1) then
            Codigo := (Classe.ValorInicialCampoChave - 1);

          AtualizarProximoCodigo(Classe.TabelaPrincipal, FieldByName('QUEBRA').AsInteger, Codigo, False);

          Next;
        end;
      end;
    finally
      FreeAndNil(CDS);
    end;

    if CriarTransacao and Conexao.HasTransaction(TD) then
      Conexao.CommitFreeAndNil(TD);
  except
    on E: Exception do
    begin
      if CriarTransacao and Conexao.HasTransaction(TD) then
        Conexao.RollbackFreeAndNil(TD);
      raise Exception.Create(FuncoesGeraisServidor.FormataErroNoServidor(Self.ClassName, 'AcertaProximoCodigo', 'Classe: ' + NomeDaClasse + #13 + E.Message));
    end;
  end;
}end;
{$ENDREGION}

end.

unit UDMConexao;

interface

uses
  SysUtils, Classes, FMTBcd, Vcl.Graphics, Controls, Forms, Menus, Dialogs, AppEvnts, Registry, Variants, ExtCtrls, ShellApi, ComCtrls, ActiveX,
  Windows, Messages, IniFiles, StdCtrls, IdStack, DB, DBClient, SqlExpr, Provider, System.Math, System.StrUtils, DateUtils, IndyPeerImpl, Datasnap.DSHTTPCommon,
  Data.DBXCommon, Data.DBXDataSnap, Datasnap.DSConnect, DBXJSONReflect, DSProxy, System.Json, IPPeerClient, System.ImageList, DSHTTPLayer,
  ImgList, DBGrids, JvDBGrid,  Rdprint, ComponentesDPR, NewBtn, KeyNav, JvDesktopAlert, SynEditHighlighter, SynHighlighterPas, SynHighlighterSQL,
  frxClass, frxDsgnIntf, frxNetUtils, frxRes, frxExportCSV, frxExportImage, frxExportText, frxExportXML, frxExportODF, frxExportXLS,
  frxExportHTML, frxExportMail, frxExportPDF, frxDMPExport, frxDCtrl, frxADOComponents, frxDBXComponents, fcxExportCSV, fcxExportDBF, fcxExportHTML, fcxExportBIFF,
  fcxExportODF, fcxCustomExport, fcxExportXML,
  TekPesquisaGrid;

type
  EFuncionalidadeNaoLiberada = class(Exception);

  TDMConexao = class(TDataModule)
    ConexaoDS: TSQLConnection;
    CDSConfigCamposClasses: TClientDataSet;
    DSPCCadAtalho: TDSProviderConnection;
    CDSConfigClasses: TClientDataSet;
    procedure OcultarColuna1Click(Sender: TObject);
    procedure ReexibirColunas1Click(Sender: TObject);
    procedure PesquisaIncremental1Click(Sender: TObject);
    procedure Exportar1Click(Sender: TObject);
    procedure DataModuleCreate(Sender: TObject);
    procedure ApplicationEvents1Message(var Msg: tagMSG; var Handled: Boolean);
    procedure PopupMenuGridPopup(Sender: TObject);
    procedure ApplicationEvents1Exception(Sender: TObject; E: Exception);
    procedure DataModuleDestroy(Sender: TObject);
    procedure ContarRegistros1Click(Sender: TObject);
    procedure SomarOuMediaColuna1Click(Sender: TObject);
    procedure AtribuicaoClick(Sender: TObject);

    procedure MeuBotaoOkB_Click(Sender: TObject);
    procedure frxExportaMailBeginExport(Sender: TObject);

    procedure Todaagrade1Click(Sender: TObject);
    procedure ApenasColunaAtual1Click(Sender: TObject);
    procedure ApenasLinhaAtual1Click(Sender: TObject);
    procedure ConexaoDSAfterConnect(Sender: TObject);
    procedure ConexaoDSAfterDisconnect(Sender: TObject);
    procedure CDSIndicadoresNewRecord(DataSet: TDataSet);
    procedure CDSIndicadoresAfterInsert(DataSet: TDataSet);
    procedure FiltrarRegistros1Click(Sender: TObject);
    function frxExportaMailSendMail(const Server: string; const Port: Integer; const UserField, PasswordField: string; FromField, ToField, SubjectField,
      CompanyField, TextField: WideString; FileNames: TStringList; Timeout: Integer; ConfurmReading: Boolean; MailCc, MailBcc: WideString): string;
    procedure AnalisaremCubodeDeciso1Click(Sender: TObject);
    procedure AutoAjusteColunas1Click(Sender: TObject);
    procedure CDSServidoresRedeGetText(Sender: TField; var Text: string; DisplayText: Boolean);
    procedure CDSServidoresTipoGetText(Sender: TField; var Text: string; DisplayText: Boolean);
    procedure CDSServidoresAfterOpen(DataSet: TDataSet);
    procedure CDSServidoresBeforePost(DataSet: TDataSet);
    procedure ApenasClulaatual1Click(Sender: TObject);
  private
    procedure CarregarListaDeTabelasEProceduresDoBDD;
    procedure CarregarUnitsProtegidas;
    procedure CarregarProcessamentosProtegidos;
  public
    ServidorBDD, VersaoBDD: string;
    ConexaoBDD_Tipo: Integer;

    DestinatarioPadraoEmailRelatorio, AssuntoPadraoEmailRelatorio: string;
    Flag_EnviaEmail_ServidorEmail_EmailSeguro, Flag_EnviaEmail_ServidorEmail_TLS: Boolean;
    Flag_EnviaEmail_Metodo: Integer;
    Flag_EnviandoRelatorioFastReport: Boolean;

    procedure FecharSistema;
    procedure EntrarNoSistema(TrocaInterna: Boolean);
    function ConectaServidorAplicacao(cUsuario, cSenha: string; iQuebra: Integer): Boolean;
    procedure LerServidoresAplicacao;

    //Execução de Métodos
    function ExecutaMetodo(Metodo: string; Parametros_Valor: array of OleVariant): OleVariant;

    function ExecuteReader(sql: string; TamanhoPacote: Integer = 1000; MonitoraSQL: Boolean = True): OleVariant;
    function ExecuteScalar(sql: string; MonitoraSQL: Boolean = True): OleVariant;
    function ExecuteCommand(sql: WideString; MonitoraSQL: Boolean = True): int64;
    function ExecuteCommand_Update(sql: WideString; Campo: string; Valor: OleVariant; MonitoraSQL: Boolean = True): int64;

    function ProximoCodigo(Tabela: string; Quebra: Integer = 0): int64;
    function ProximoCodigoAcrescimo(Tabela: string; Quebra: Integer = 0; Acrescimo: Integer = 1): int64;

    function DataHora: TDateTime;
    function DataHoraServidor(ForcaHoraServidor: Boolean = false): TDateTime;

    procedure RegistraAcao(Descricao: string; Inicio, Fim: TDateTime; Observacao: string);

    procedure CarregaDefaults;
    procedure AtribuirOutrosDefault(DS: TDataSet; Tabela: string);

    function GetCDSConfigCamposClasses: OleVariant;
    function GetCDSConfigClasses: OleVariant;

    function Ler(Campos, Tabela: string; Ordem: Integer; Where: string = ''): OleVariant;
    function Acha(Tabela, Campo: string; Valor: Variant; CampoEmpresa: string = ''; CodigoDaEmpresa: Integer = -1): Boolean;

  end;

var
  DMConexao: TDMConexao;

const
  cServTipo_SoftwareCenter = 0;
  cServTipo_TekServer = 1;

  cServRede_Intranet = 0;
  cServRede_Extranet = 1;

implementation

uses Constantes, ClassDataSet;
{$R *.dfm}

procedure TDMConexao.DataModuleCreate(Sender: TObject);
begin

  ClassArquivoINI.TArquivoINI.NomeArquivoPadrao := Constantes.ArquivoIniClient;
  ClassArquivoINI.TArquivoINI.PathArquivoPadrao := ExtractFilePath(Application.ExeName);

  fConfig := TConfigNovo.Create;
  fSecaoAtual := TClassSecaoNovo.Create;

  CallBack := TClienteCallback.Create;
  CallBack.CallbackMethod := CallbackMethod;

  //TekPesquisaGrid1.ArquivoINI := ExtractFilePath(Application.ExeName) + ArquivoIniClient;

  Constantes.DMConexaoExistente := Self;

  with SQLConexao do
    begin
      Params.Values['Sistema']    := IntToStr(ConstanteSistema.Sistema);
      Params.Values['IP']         := SecaoAtual.IP;
      Params.Values['Host']       := SecaoAtual.Host;
      Params.Values['Plataforma'] := Constantes.sPlataformaAtual;
    end;

  CDSUnitsProtegidas          := TClientDataSet.Create(Self);
  CDSProcessamentosProtegidos := TClientDataSet.Create(Self);

  EntrarNoSistema(false);

  FSplash.Passo(70, 'Configurações Adicionais do DMConexão');

  if FileExists(TFuncoesSistemaOperacional.DiretorioComBarra(Config.DirExe) + ChangeFileExt(Modulos[Sistema, 1], '.CHM')) then
    Application.HelpFile := TFuncoesSistemaOperacional.DiretorioComBarra(Config.DirExe) + ChangeFileExt(Modulos[Sistema, 1], '.CHM')
  else
    if FileExists(TFuncoesSistemaOperacional.DiretorioComBarra(Config.DirExe) + ArquivoHelpGeral) then
    Application.HelpFile := TFuncoesSistemaOperacional.DiretorioComBarra(Config.DirExe) + ArquivoHelpGeral;

  ExtrairModeloRelPai;

  Debug.PopupMenuGradeDebugDataSet := PopupMenuGrid;

  Flag_EnviandoRelatorioFastReport := false;
  DestinatarioPadraoEmailRelatorio := '';
  AssuntoPadraoEmailRelatorio := '';

  try
    PrimeiraVezDoTimer := 0;
    cDataHoraServidorNaEntrada := DataHoraServidor(True);
    cDataHoraServidor := cDataHoraServidorNaEntrada;
  except
    On E: Exception do
      raise Exception.Create('Erro em DataHoraServidor: ' + E.Message);
  end;

  CarregaDefaults;
  CarregaConfigClasses;
  CarregarUnitsProtegidas;
  CarregarProcessamentosProtegidos;

  with CDSAtalhos do
    begin
      RemoteServer := DSPCCadAtalho;
      FetchParams;
    end;

  CarregarListaDeTabelasEProceduresDoBDD;
end;

procedure TDMConexao.DataModuleDestroy(Sender: TObject);
begin
  if Assigned(LogoEmpresa) then
    FreeAndNil(LogoEmpresa);

  if Assigned(CntrlDiasNaoUteis) then
    FreeAndNil(CntrlDiasNaoUteis);

  if Assigned(TekProtClient) then
  begin
    TekProtClient.OnAfterValidate := nil;
    TekProtClient.OnGetCloseApp := nil;
    FreeAndNil(TekProtClient);
  end;

  if Assigned(Config) then
    FreeAndNil(fConfig);

  if SQLConexao.Connected then
    SQLConexao.Close;

  if Assigned(SecaoAtual) then
    FreeAndNil(fSecaoAtual);

  if Assigned(CallBack) then
    FreeAndNil(CallBack);

  if Assigned(CDSUnitsProtegidas) then
    CDSUnitsProtegidas.Free;

  if Assigned(CDSProcessamentosProtegidos) then
    CDSProcessamentosProtegidos.Free;

  Constantes.DMConexaoExistente := nil;
end;


function TDMConexao.ExecutaMetodo(Metodo: string; Parametros_Valor: array of OleVariant): OleVariant;
begin
  Result := FuncoesDataSnap.ExecutaMetodo_Sincrono(SQLConexao, Metodo, Parametros_Valor, false);
end;


function TDMConexao.ExecuteScalar(sql: string; MonitoraSQL: Boolean = True): OleVariant;
var
  Tempo: TTime;
begin
  // Executa a função ExecuteScalar do Servidor de Aplicação, que tem o objetivo de
  // executar um comando no BDD e retornar o primeiro campo do primeiro registro
  // Exemplo: Informacao('Total de Clientes: ' + IntToStr(ExecuteScalar('select count(*) from cliente')));

  Tempo := Time;
  if (FPrincipal <> nil) and (MonitoraSQL) then
    FPrincipal.AdicionaRich('ExecuteScalar', sql);

  Result := ExecutaMetodo('TSMConexao.ExecuteScalar', [Trim(sql), True]);

  if (FPrincipal <> nil) and (MonitoraSQL) then
  begin
    Tempo := Time - Tempo;
    FPrincipal.AdicionaLinhaRich('Tempo Gasto em ExecuteScalar ==> ' + FormatDateTime(' hh:mm:ss.zzz', Tempo), clGreen, [fsBold]);
  end;
end;

function TDMConexao.ExecuteReader(sql: string; TamanhoPacote: Integer = 1000; MonitoraSQL: Boolean = True): OleVariant;
var
  Tempo: TTime;
begin
  // Executa a função ExecuteReader do Servidor de Aplicação, que tem o objetivo de
  // executar um comando no BDD e retornar todos os dados
  // Exemplo: ClientDataSet1.Data := ExecuteReader('select CODIGO_CLI, NOME_CLI from CLIENTE order by NOME_CLI');

  Tempo := Time;
  if (FPrincipal <> nil) and (MonitoraSQL) then
    FPrincipal.AdicionaRich('ExecuteReader', sql);

  Result := ExecutaMetodo('TSMConexao.ExecuteReader', [Trim(sql), TamanhoPacote, True]);

  if (FPrincipal <> nil) and (MonitoraSQL) then
  begin
    Tempo := Time - Tempo;
    FPrincipal.AdicionaLinhaRich('Tempo Gasto em ExecuteReader ==> ' + FormatDateTime(' hh:mm:ss.zzz', Tempo), clGreen, [fsBold]);
  end;
end;

function TDMConexao.ExecuteCommand(sql: WideString; MonitoraSQL: Boolean = True): int64;
var
  Tempo: TTime;
begin
  // Executa a função ExecuteCommand do Servidor de Aplicação, que tem o objetivo de
  // executar comandos que não possuem resultado. Do tipo INSERT, DELETE, UPDATE.
  // Retornando o número de registros afetados
  // Exemplo: Informacao('Registros Atualizados:' + IntToStr(ExecuteCommand('update PRODUTO set SALDOFISICO_PRODUTO = SALDOFISICO_PRODUTO + 1 where EMPRESA_PRODUTO = 1 and CODIGO_PRODUTO = ''XXX''')));

  Tempo := Time;
  if (FPrincipal <> nil) and (MonitoraSQL) then
    FPrincipal.AdicionaRich('ExecuteCommand', sql);

  Result := TFuncoesConversao.VariantParaInt64(ExecutaMetodo('TSMConexao.ExecuteCommand', [Trim(sql), True]));

  if (FPrincipal <> nil) and (MonitoraSQL) then
  begin
    Tempo := Time - Tempo;
    FPrincipal.AdicionaLinhaRich('Tempo Gasto em ExecuteCommand ==> ' + FormatDateTime(' hh:mm:ss.zzz', Tempo), clGreen, [fsBold]);
  end;
end;


procedure TDMConexao.ConexaoDSAfterConnect(Sender: TObject);
var
  Ret: OleVariant;
begin
  if (not ConexaoTratada) then
  begin
    TCaixasDeDialogo.Erro(sErroNaRede + 'Tentativa de Conexão não tratada.');
    FecharSistema;
  end;

  if Assigned(FSplash) then
  begin
    FSplash.Show;
    FSplash.Passo(45, 'Tratando Conexão');
  end;

  Ret := ExecutaMetodo('TSMConexao.InformacoesDaConexaoBDD', []);
  ServidorBDD := Ret[0]; // ExecutaMetodo('TSMConexao.ConexaoBDD', []);
  VersaoBDD := Ret[1]; // ExecutaMetodo('TSMConexao.VersaoBDD', []);
  ConexaoBDD_Tipo := Ret[2]; // ExecutaMetodo('TSMConexao.ConexaoBDD_Tipo', []);
  DriverBDDAtual := Ret[3];
end;

procedure TDMConexao.ConexaoDSAfterDisconnect(Sender: TObject);
begin
  CallBack.DesRegistraCallBack(SecaoAtual.Usuario.Nome);
end;

procedure TDMConexao.CDSIndicadoresAfterInsert(DataSet: TDataSet);
begin
  FPrincipal.PainelSalvo := false;
end;

procedure TDMConexao.CDSIndicadoresNewRecord(DataSet: TDataSet);
begin
  DataSet.AcertarDefaultDinamico(DMConexao.SecaoAtual.Usuario.Nome, DMConexao.cDataHoraServidor);
end;

procedure TDMConexao.CDSServidoresAfterOpen(DataSet: TDataSet);
begin
  CDSServidoresRede.OnSetText := TGetTexts.SetText_Campo1Digito;
  CDSServidoresTipo.OnSetText := TGetTexts.SetText_Campo1Digito;
end;

procedure TDMConexao.CDSServidoresBeforePost(DataSet: TDataSet);
begin
  if (DataSet.FieldByName('Tipo').AsInteger = cServTipo_SoftwareCenter) and
     (DataSet.FieldByName('Rede').AsInteger = cServRede_Extranet) then
  begin
    TCaixasDeDialogo.Aviso('O tipo de servidor informado "SoftwareCenter" atualmente não oferece suporte para o tipo de rede "Extranet".');
    Abort;
  end;
  if (DataSet.FieldByName('Tipo').AsInteger <> cServTipo_SoftwareCenter) and
     (DataSet.FieldByName('Porta').AsInteger = Constantes.TekAgendadorPorta) then
  begin
    TCaixasDeDialogo.Aviso('Verifique a porta informada, pois, pertence à SoftwareCenter.');
    Abort;
  end;
end;

procedure TDMConexao.CDSServidoresRedeGetText(Sender: TField; var Text: string; DisplayText: Boolean);
begin
  if (Sender.DataSet.IsEmpty) then
  begin
    Text := '';
    Exit;
  end;

  case Sender.AsInteger of
    cServRede_Intranet:
      Text := 'Intranet';
    cServRede_Extranet:
      Text := 'Extranet';
    else
      Text := Sender.AsString;
  end;
end;

procedure TDMConexao.CarregarListaDeTabelasEProceduresDoBDD;
begin
  SynSQLSynPadrao.FunctionNames.Text := ExecutaMetodo('TSMConexao.GetProcedureNames', []);
  SynSQLSynPadrao.TableNames.Text    := ExecutaMetodo('TSMConexao.GetTableNames',     [False]);
  SynSQLSynPadrao.TableNames.Add(ExecutaMetodo('TSMConexao.GetTableNames', [True]));
end;

{$REGION 'Carga de Cache'}
procedure TDMConexao.CarregaDefaults;
begin
  CDSDefaults.Data := ExecutaMetodo('TSMConexao.CarregaDefaults', ['']);
  CDSDefaults.IndexFieldNames := 'TABELA';
end;

procedure TDMConexao.AtribuirOutrosDefault(DS: TDataSet; Tabela: string);
begin
  DS.AtribuirOutrosDefault(CDSDefaults, Tabela);
end;

procedure TDMConexao.CarregaPermissoesDoCadastro;
var sql: string;
begin
  sql :=
    'select ' + #13 +
    '  USUARIO_CLASSES.CLASSE_USUCLASS, ' + #13 +
    '  USUARIO_CLASSES.INCLUI_USUCLASS, ' + #13 +
    '  USUARIO_CLASSES.ALTERA_USUCLASS, ' + #13 +
    '  USUARIO_CLASSES.EXCLUI_USUCLASS' + #13 +
    'from USUARIO_CLASSES' + #13 +
    'where USUARIO_CLASSES.USUARIO_USUCLASS = ' + IntToStr(SecaoAtual.Usuario.Codigo);

  CDSPermissoes.Data := DMConexao.ExecuteReader(sql, -1);
  CDSPermissoes.IndexFieldNames := 'CLASSE_USUCLASS';
end;

procedure TDMConexao.CarregaConfigClasses;
var sql: string;
begin
  sql :=
    'select ' + #13 +
    '  CONFIG_CLASSES.CLASSEPAI_CFGC, ' + #13 +
    '  CONFIG_CLASSES.CLASSE_CFGC, ' + #13 +
    '  CONFIG_CLASSES.CONDICAO_CFGC, ' + #13 +
    '  CONFIG_CLASSES.MENSAGEM_CFGC ' + #13 +
    'from CONFIG_CLASSES ';

  CDSConfigClasses.Data := DMConexao.ExecuteReader(sql, -1);
  CDSConfigClasses.IndexFieldNames := 'CLASSEPAI_CFGC;CLASSE_CFGC';
end;

procedure TDMConexao.CarregaConfigCamposClasses;
var sql: string;
begin
  sql :=
    'select ' + #13 +
    '  USUARIO_CLASSES.CLASSE_USUCLASS,   ' + #13 +
    '  USUARIO_CLASSES_CAMPOS.CLASSE_UCC, ' + #13 +
    '  USUARIO_CLASSES_CAMPOS.CAMPO_UCC,  ' + #13 +
    '  ''S'' BLOQUEAR_UCC, ' + #13 +
    '  ''N'' OMITIR_UCC,   ' + #13 +
    '  ''''  CONDICAO_UCC, ' + #13 +
    '  ''''  MENSAGEM_UCC  ' + #13 +
    'from USUARIO_CLASSES_CAMPOS ' + #13 +
    'inner join USUARIO_CLASSES on (USUARIO_CLASSES.AUTOINC_USUCLASS = USUARIO_CLASSES_CAMPOS.CODIGOCLASSE_UCC) ' + #13 +
    'where USUARIO_CLASSES.USUARIO_USUCLASS  = ' + IntToStr(SecaoAtual.Usuario.Codigo) + #13 +
    'union all' + #13 +
    'select ' + #13 +
    '  CONFIG_CLASSES_CAMPOS.CLASSEPAI_CCC, ' + #13 +
    '  CONFIG_CLASSES_CAMPOS.CLASSE_CCC,    ' + #13 +
    '  CONFIG_CLASSES_CAMPOS.CAMPO_CCC,     ' + #13 +
    '  CONFIG_CLASSES_CAMPOS.BLOQUEAR_CCC,  ' + #13 +
    '  CONFIG_CLASSES_CAMPOS.OMITIR_CCC,    ' + #13 +
    '  CONFIG_CLASSES_CAMPOS.CONDICAO_CCC,  ' + #13 +
    '  CONFIG_CLASSES_CAMPOS.MENSAGEM_CCC   ' + #13 +
    'from CONFIG_CLASSES_CAMPOS ';

  CDSConfigCamposClasses.Data := DMConexao.ExecuteReader(sql, -1);
  CDSConfigCamposClasses.IndexFieldNames := 'CLASSE_USUCLASS;CLASSE_UCC;CAMPO_UCC';
end;

procedure TDMConexao.CarregaMensagem;
begin
  StatusDeMensagens := ExecutaMetodo('TSMMensagem.MensagemNaoLida', []);
end;

procedure TDMConexao.CarregaSecaoAtual;
var
  vJson: TJSONValue;
  xUnMarshal: TJSONUnMArshal;
  FCommand: TDBXCommand;
  sSQL: string;
  CDSTemp: TClientDataSet;
begin
  FCommand := SQLConexao.DBXConnection.CreateCommand;
  try
    FCommand.CommandType := TDBXCommandTypes.DSServerMethod;
    FCommand.Text := 'TSMConexao.SecaoAtualSerializadaNovo';
    FCommand.Prepare;
    FCommand.Parameters[0].Value.AsString := '1';
    FCommand.ExecuteUpdate;
    vJson := TJSONValue(FCommand.Parameters[1].Value.GetJSONValue);
  finally
    FreeAndNil(FCommand);
  end;

  if Assigned(vJson) then
  begin
    if Assigned(fSecaoAtual) then
      FreeAndNil(fSecaoAtual);

    xUnMarshal := TJSONUnMArshal.Create;
    try
      fSecaoAtual := (xUnMarshal.UnMarshal(vJson) as TClassSecaoNovo);
    finally
      FreeAndNil(xUnMarshal);
    end;
    FreeAndNil(vJson);
  end;

  CarregaPermissoesDoCadastro;
  CarregaConfigCamposClasses;

  CDSTemp := TClientDataSet.Create(nil);
  try
    sSQL :=
      'select EMPRESA.LOGOMARCA_EMP from EMPRESA' + #13 +
      'where EMPRESA.CODIGO_EMP = ' + IntToStr(SecaoAtual.Empresa.Codigo);
    CDSTemp.Data := ExecuteReader(sSQL, -1, True);

    if not Assigned(LogoEmpresa) then
      LogoEmpresa := TPicture.Create;

    CDSTemp.FieldByName('LOGOMARCA_EMP').Recuperar(LogoEmpresa);
  finally
    FreeAndNil(CDSTemp);
  end;

  if (FPrincipal <> nil) then
    FPrincipal.ConfigurarAmbienteSistema;

{$IF DEFINED(QUALIDADE)}
  DMConexao.ExecutaMetodo('TSMConexao.AtualizarContadorSAC', [DMConexao.SecaoAtual.Usuario.Codigo]);
{$IFEND}

{$Region 'Extração Schemas NFS-e - Provedor pode variar de acordo com município da empresa'}
{$IF DEFINED(FATURAMENTO) OR DEFINED(ESTOQUE) OR DEFINED(LIVROSFISCAIS) OR DEFINED(ANEXO)}
  if not ClassFuncoesSistemaOperacional.IsDebuggerPresent() then
  begin
    TThread.CreateAnonymousThread(
      procedure()
      begin
        TClassConfigNFSe.ExtraiSchemasXMLNFSe(sisERP, NomeSchemaProvedorNFSe(DMConexao.SecaoAtual.Parametro.NFSe_Provedor));
      end
      ).Start;
  end;
{$IFEND}
{$EndRegion}

end;

procedure TDMConexao.CarregaModulosDisp(var CDS: TClientDataSet);
var
  X: Integer;
begin
  with CDS do
  begin
    Close;
    with FieldDefs do
    begin
      Clear;
      Add('Codigo', ftInteger);
      Add('Descricao', ftString, 40);
    end;
    CreateDataSet;
    LogChanges := false;
    IndexFieldNames := 'Descricao';

    // É necessário a inclusão do registro zero para passar na validação do new record
    Insert;
    FieldByName('Codigo').AsInteger := 0;
    FieldByName('Descricao').AsString := 'Indefinido';
    Post;

    for X := Low(Modulos) to High(Modulos) do
    begin
      Insert;
      FieldByName('Codigo').AsInteger := X;
      FieldByName('Descricao').AsString := Copy(Modulos[X, 2], 1, 40);
      Post;
    end;
  end;
end;

function TDMConexao.GetCDSConfigCamposClasses: OleVariant;
begin
  Result := CDSConfigCamposClasses.Data;
end;

function TDMConexao.GetCDSConfigClasses: OleVariant;
begin
  Result := CDSConfigClasses.Data;
end;

{$ENDREGION}

{$REGION 'Atualizar registro pala DLL / Consumir DLL'}


procedure TDMConexao.AtualizarDeptoPessoal(MostrarMensagem: Boolean);
var
  Parametros: OleVariant;
begin
  ExecutaMetodo_ComCallBack('TSMDepPessoal.EP_InicializarTabelas', [Parametros]);

  if MostrarMensagem then
    TCaixasDeDialogo.Informacao(sSucessoEmProcesso);
end;

procedure TDMConexao.AtualizarESocial(MostrarMensagem: Boolean);
var
  Parametros: OleVariant;
begin
  ExecutaMetodo_ComCallBack('TSMDepPessoal.EP_InicializarTabelasESocial', [Parametros]);

  if MostrarMensagem then
    TCaixasDeDialogo.Informacao(sSucessoEmProcesso);
end;

procedure TDMConexao.AtualizarImportaExportaTitulos(MostrarMensagem: Boolean);
begin
  ExecutaMetodo_ComCallBack('TSMCadConfigExportacaoTitulos.AtualizarConfiguracoesImportaExportaTekSystem', []);

  if MostrarMensagem then
    TCaixasDeDialogo.Informacao('Configurações de Importação/Exportação de Títulos atualizados com Sucesso!');
end;

procedure TDMConexao.AtualizarLivroFiscal(MostrarMensagem: Boolean);
var
  Parametros: OleVariant;
begin
  ExecutaMetodo_ComCallBack('TSMLivroFiscal.ExecProcedimento', [LF_Constantes.LF_Proc_InicializarTabelas, Parametros]);

  if MostrarMensagem then
    TCaixasDeDialogo.Informacao(sSucessoEmProcesso);
end;

procedure TDMConexao.AtualizarModelosDataWarehouse(MostrarMensagem: Boolean);
begin
  ExecutaMetodo_ComCallBack('TSMCadDW_Temas.AtualizarModelosDataWarehouse', []);

  if MostrarMensagem then
    TCaixasDeDialogo.Informacao('Temas de Armazéns de Dados (Data Warehouse) Atualizados com Sucesso!');
end;

procedure TDMConexao.AtualizarModelosRelatoriosEspecificos(MostrarMensagem: Boolean);
begin
  ExecutaMetodo_ComCallBack('TSMCadGR_Relatorio.AtualizarModelosRelatorios', []);

  if MostrarMensagem then
    TCaixasDeDialogo.Informacao('Relatórios Específicos (Gerador) atualizados com Sucesso!');
end;

procedure TDMConexao.AtualizarModelosUnidadesCodificacao(MostrarMensagem: Boolean);
begin
  ExecutaMetodo_ComCallBack('TSMCadGR_Unidades_Codificacao.AtualizarModelosUnidadesCodificacao', []);

  if MostrarMensagem then
    TCaixasDeDialogo.Informacao('Modelos de Unidades de Codificação Atualizados com Sucesso!');
end;

procedure TDMConexao.AtualizarModelosProcessamentos(MostrarMensagem: Boolean);
begin
  ExecutaMetodo_ComCallBack('TSMCadTI_Processamentos.AtualizarModelosProcessamentosTekSystem', []);

  if MostrarMensagem then
    TCaixasDeDialogo.Informacao('Processamentos Específicos foram atualizados com sucesso!');
end;

procedure TDMConexao.AtualizarModelosIndicadores(MostrarMensagem: Boolean);
begin
  ExecutaMetodo_ComCallBack('TSMCadGR_Indicadores.AtualizarIndicadoresTekSystem', []);

  if MostrarMensagem then
    TCaixasDeDialogo.Informacao('Modelos de Indicadores Atualizados com Sucesso!');
end;

procedure TDMConexao.AtualizarConfigRemessaRetornoBanc(MostrarMensagem: Boolean);
begin
  ExecutaMetodo_ComCallBack('TSMCadConfigRemessaRetorno.AtualizarConfiguracoesRemessaRetornoBancTekSystem', []);

  if MostrarMensagem then
    TCaixasDeDialogo.Informacao('Configurações de Remessa/Retorno Bancários Atualizados com Sucesso!');
end;

procedure TDMConexao.AtualizarContabilidade(MostrarMensagem: Boolean);
begin
  ExecutaMetodo_ComCallBack('TSMContabilidade.InicializarDadosContabilidade', []);
  if MostrarMensagem then
    TCaixasDeDialogo.Informacao('Tabelas do Módulo Contabilidade Atualizados com Sucesso!');
end;

procedure TDMConexao.AtualizarModelosRelatorios(MostrarMensagem: Boolean);
begin
  ExecutaMetodo_ComCallBack('TSMRelatorio.AtualizaRelatorios', []);

  if MostrarMensagem then
    TCaixasDeDialogo.Informacao('Modelos de relatórios atualizados com sucesso!');
end;

procedure TDMConexao.ExtrairModeloRelPai;
const
  MaxTentativas = 5;
var
  Tentativa: Integer;
  ST: TMemoryStream;
begin
  // No caso da abertura de diversos módulos simultâneos,
  // por exemplo com terminal server
  Tentativa := 1;
  while (Tentativa <= MaxTentativas) do
  begin
    try
      ST := TMemoryStream.Create;
      try
        TFuncoesSistemaOperacional.LerRecursoDLL(sPaiRelatorioGrafico, sNomeDll, ST);
        ST.SaveToFile(Config.DirExe + sPaiRelatorioGrafico + '.fr3');
        Break;
      Finally
        FreeAndNil(ST);
      end;
    except
      on E: Exception do
      begin
        if (Tentativa = MaxTentativas) then
        begin
          if TCaixasDeDialogo.Confirma(
            'Foram feitas ' + IntToStr(MaxTentativas) + ' tentativas de extração do arquivo que contém o modelo pai de relatórios.' +
            'No entanto, não foi possível sua extração, devido o erro:' + #13 + E.Message + #13 +
            'Gostaria de tentar novamente agora?') then
            Tentativa := 1
          else
            Halt;
        end
        else
        begin
          Inc(Tentativa);
          Application.ProcessMessages;
        end;
      end;
    end;
  end;
end;

{$ENDREGION}

{$REGION 'Metodos Comuns com SMConexao'}


function TDMConexao.ProximoCodigo(Tabela: string; Quebra: Integer = 0): int64;
begin
  // Executa a função Proximo do Servidor de Aplicação, que tem o objetivo de
  // retornar o próximo código para a tabela em questão
  Tabela := AnsiUpperCase(Tabela);
  Result := ExecutaMetodo('TSMConexao.ProximoCodigo', [Tabela, Quebra]);
end;

function TDMConexao.ProximoCodigoAcrescimo(Tabela: string; Quebra, Acrescimo: Integer): int64;
begin
  // Executa a função Proximo do Servidor de Aplicação, que tem o objetivo de
  // retornar o próximo código para a tabela em questão, com incremento de acordo com o terceiro parametro
  Tabela := AnsiUpperCase(Tabela);
  Result := ExecutaMetodo('TSMConexao.ProximoCodigoAcrescimo', [Tabela, Quebra, Acrescimo]);
end;

function TDMConexao.DataHora: TDateTime;
begin
  // Posteriormente substituir a função DataHoraServidor por essa;
  Result := DataHoraServidor;
end;

function TDMConexao.DataHoraServidor(ForcaHoraServidor: Boolean = false): TDateTime;
begin
  // Executa a função DataHora do Servidor de Aplicação, que tem o objetivo de
  // Retornar a data e hora do servidor de banco de dados
  // ATENÇÃO: Só é chamada uma vez no sistema, para evitar fluxo ao Servidor de Aplicação.
  // Use a variável cDataHoraServidor que é sempre atualizada
  if (ForcaHoraServidor) or (cDataHoraServidor = 0) then
    Result := ExecutaMetodo('TSMConexao.DataHora', [])
  else
    Result := cDataHoraServidor;
end;

procedure TDMConexao.RegistraAcao(Descricao: string; Inicio, Fim: TDateTime; Observacao: string);
begin
  // Executa a função RegistraAcao do Servidor de Aplicação, que tem o objetivo de
  // Registrar ações monitoradas/perigosas executadas pelo usuário
  ExecutaMetodo('TSMConexao.RegistraAcao', [Descricao, Inicio, Fim, Observacao]);
end;

{$ENDREGION}

{$REGION 'Dias Não Uteis'}


procedure TDMConexao.CarregaDiasNaoUteis(Modulo: Integer);
begin
  if Assigned(CntrlDiasNaoUteis) then
    CntrlDiasNaoUteis.QtdeReferencias := CntrlDiasNaoUteis.QtdeReferencias + 1
  else
    CntrlDiasNaoUteis := TClassCntrlDiasUteis.Create(Self, Modulo, SecaoAtual.Empresa.CidadeCod);
end;

procedure TDMConexao.ReCarregaDiasNaoUteis(Modulo: Integer);
begin
  if Assigned(CntrlDiasNaoUteis) then
    CntrlDiasNaoUteis.CarregaDiasNaoUteis;
end;

procedure TDMConexao.DesCarregaDiasNaoUteis;
begin
  if (CntrlDiasNaoUteis.QtdeReferencias = 1) then
    FreeAndNil(CntrlDiasNaoUteis)
  else
    CntrlDiasNaoUteis.QtdeReferencias := CntrlDiasNaoUteis.QtdeReferencias - 1;
end;

function TDMConexao.DiasEntre(dDataInicial, dDataFinal: TDate; iSistema: Integer; bDiaUtil, bPermiteNegativo: Boolean): Integer;
begin
  CarregaDiasNaoUteis(iSistema);
  try
    Result := CntrlDiasNaoUteis.DiasEntre(dDataInicial, dDataFinal, bDiaUtil, bPermiteNegativo, iSistema);
  finally
    DesCarregaDiasNaoUteis;
  end;
end;

function TDMConexao.DiaUtil(cData: TDate; Modulo: Integer; Cidade: Integer = 0): Boolean;
var
  ClasseCriadaAgora: Boolean;
begin
  ClasseCriadaAgora := not Assigned(CntrlDiasNaoUteis);
  if (ClasseCriadaAgora) then
    CntrlDiasNaoUteis := TClassCntrlDiasUteis.Create(Self, Modulo, Cidade);

  try
    Result := CntrlDiasNaoUteis.DiaUtil(cData);
  finally
    if (ClasseCriadaAgora) then
      FreeAndNil(CntrlDiasNaoUteis);
  end;
end;

function TDMConexao.DiaUtilAnterior(Data: TDate; Modulo, Cidade: Integer): TDate;
var
  ClasseCriadaAgora: Boolean;
begin
  ClasseCriadaAgora := not Assigned(CntrlDiasNaoUteis);
  if (ClasseCriadaAgora) then
    CntrlDiasNaoUteis := TClassCntrlDiasUteis.Create(Self, Modulo, Cidade);

  try
    Result := CntrlDiasNaoUteis.DiaUtilAnterior(Data);
  finally
    if (ClasseCriadaAgora) then
      FreeAndNil(CntrlDiasNaoUteis);
  end;
end;

function TDMConexao.DiasUteisEntre(dDataInicial, dDataFinal: TDateTime; Modulo: Integer; Cidade: Integer = 0): Integer;
var
  ClasseCriadaAgora: Boolean;
begin
  ClasseCriadaAgora := not Assigned(CntrlDiasNaoUteis);
  if (ClasseCriadaAgora) then
    CntrlDiasNaoUteis := TClassCntrlDiasUteis.Create(Self, Modulo, Cidade);

  try
    Result := CntrlDiasNaoUteis.DiasUteisEntre(Trunc(dDataInicial), Trunc(dDataFinal));
  finally
    if (ClasseCriadaAgora) then
      FreeAndNil(CntrlDiasNaoUteis);
  end;
end;

function TDMConexao.ProximoDiaUtil(Data: TDate; Modulo: Integer; Cidade: Integer = 0): TDate;
var
  ClasseCriadaAgora: Boolean;
begin
  ClasseCriadaAgora := not Assigned(CntrlDiasNaoUteis);
  if (ClasseCriadaAgora) then
    CntrlDiasNaoUteis := TClassCntrlDiasUteis.Create(Self, Modulo, Cidade);

  try
    Result := CntrlDiasNaoUteis.ProximoDiaUtil(Data);
  finally
    if (ClasseCriadaAgora) then
      FreeAndNil(CntrlDiasNaoUteis);
  end;
end;

{$ENDREGION}

{$REGION 'Autorização'}


function TDMConexao.VerificaAutorizacao(cOpcao: Integer): Boolean;
begin
  Result := SecaoAtual.Usuario.AutorizacaoEspecial[cOpcao];
end;

function TDMConexao.VerificaAutorizacao(sOpcao: string): Boolean;
begin
  Result := VerificaAutorizacao(StrToInt(sOpcao));
end;

function TDMConexao.VerificaAutorizacao_ComOutroUsuario(cUsuario, cOpcao: Integer): Boolean;
var
  CDSTemp: TClientDataSet;
begin
  CDSTemp := TClientDataSet.Create(nil);
  try
    with ListaDeStrings do
    begin
      Clear;
      Add('select USUARIO_ESPECIAIS.AUTOINC_USUESP');
      Add('from USUARIO_ESPECIAIS');
      Add('where USUARIO_ESPECIAIS.USUARIO_USUESP = ' + IntToStr(cUsuario));
      Add('  and USUARIO_ESPECIAIS.OPCAO_USUESP   = ' + IntToStr(cOpcao));
    end;
    CDSTemp.Data := ExecuteReader(ListaDeStrings.Text, 1);
    Result := not CDSTemp.eof;
  finally
    FreeAndNil(CDSTemp);
  end;
end;

{$ENDREGION}

{$REGION 'Metodos Diversos'}

procedure TDMConexao.EntrarNoSistema(TrocaInterna: Boolean);
var
  sUsuario, sSenha, sRetorno: string;
  iQuebra: Integer;
  tpStrListTekProt: TStringList;
begin
  // if not (TrocaInterna) then
  LerServidoresAplicacao;

  if (TrocaInterna) then
  begin
    if ConstanteSistema.Sistema = cSistemaCaixa then
      iQuebra := SecaoAtual.Empresa.Codigo
    else if (ConstanteSistema.Sistema in [cSistemaDepPessoal, cSistemaContabilidade]) then
      iQuebra := SecaoAtual.Empresa.Estabelecimento
    else
      iQuebra := SecaoAtual.Empresa.Codigo;
  end
  else
  begin
    if ConstanteSistema.Sistema = cSistemaCaixa then
      iQuebra := StrToIntDef(TArquivoINI.Ler('MenuPrincipal', 'CaixaAtual', '1'), 0)
    else if (ConstanteSistema.Sistema in [cSistemaDepPessoal, cSistemaContabilidade]) then
      iQuebra := StrToIntDef(TArquivoINI.Ler('MenuPrincipal', 'EstabelecimentoAtual', '1'), 0)
    else if (ConstanteSistema.Sistema = cSistemaESocial) then
      iQuebra := StrToIntDef(TArquivoINI.Ler('MenuPrincipal', 'EmpregadorAtual', '1'), 0)
    else
      iQuebra := StrToIntDef(TArquivoINI.Ler('MenuPrincipal', 'EmpresaAtual', '1'), 0);
  end;

  if Assigned(FSplash) then
    FSplash.Passo(25, 'Login...');

  if (ParamCount > 0) or (TrocaInterna) then
  begin
    if (TrocaInterna) then
    begin
      sUsuario := SecaoAtual.Usuario.Nome;
      sSenha := SecaoAtual.Usuario.Senha;
    end
    else // (ParamCount > 0)
    begin
      sUsuario := ParamStr(1);
      sSenha := ParamStr(2);
      if FindCmdLineSwitch('CRIPT', ['/'], True) then
        sSenha := TFuncoesCriptografia.Decodifica(sSenha, sChaveCriptografia);

        //Trim(Decode(sSenha));
    end;

    if (not ConectaServidorAplicacao(sUsuario, sSenha, iQuebra)) then
      FecharSistema;
  end
  else
  begin
    Application.CreateForm(TFLogin, FLogin);
    FSplash.Hide;
    try
      FLogin.Quebra := iQuebra;

      if (FLogin.ShowModal <> mrOK) then
        FecharSistema;
    finally
      FreeAndNil(FLogin);
    end;
  end;

  if ServidorBDD = '' then
    FecharSistema;

  TFuncoesBaseDados.AcertarParticularidades(DriverBDDAtual, Constantes.ParticularidadesBDD);

  Application.CreateForm(TDMDownload, DMDownload);
  try
    DMDownload.VerificaArquivosNecessarios;
  finally
    DMDownload.Free;
  end;

  CarregaSecaoAtual;

  CallBack.Configura(
    DMConexao.SQLConexao.Params.Values[TDBXPropertyNames.HostName],
    DMConexao.SQLConexao.Params.Values[TDBXPropertyNames.Port],
    SecaoAtual.Usuario.Nome,
    SecaoAtual.Usuario.Senha,
    Config.ServidorProxy,
    Config.PortaProxy,
    Config.UsuarioProxy,
    Config.SenhaProxy,
    FuncoesCallBack2.cCanal,
    SecaoAtual.Guid);
  CallBack.RegistraCallBack(SecaoAtual.Usuario.Nome);

  VerificaTrocaDeSenha;

  {$region 'Tekprot'}
  if (not Assigned(TekProtClient)) or ((Assigned(TekProtClient)) and ((TekProtClient.Server <> Config.ServidorTekProt) or (TekProtClient.Port <> Config.PortaTekProt))) then
    begin
      if Assigned(FSplash) then
        FSplash.Passo(60, 'Validando a Cópia do Sistema');
      try
        if Assigned(TekProtClient) then
          begin
            TekProtClient.OnAfterValidate := nil;
            TekProtClient.OnGetCloseApp   := nil;
            FreeAndNil(TekProtClient);
          end;

        TekProtClient := TTekProtClient.Create(Self);
        with TekProtClient do
          begin
            ExecName  := Modulos[Sistema, 1];
            ModuleCod := StrToInt(Modulos[Sistema, 3]);

            if (SecaoAtual.Empresa.Codigo = 0) or (SecaoAtual.Sistema in [cSistemaCaixa, cSistemaDepPessoal, cSistemaContabilidade, cSistemaESocial]) then
              EmpID := EmptyStr
            else
              EmpID := TFuncoesString.SoNumero(IfThen(SecaoAtual.Empresa.Natureza = 'J', SecaoAtual.Empresa.CNPJ, SecaoAtual.Empresa.CPF));

            Server := Config.ServidorTekProt;
            Port   := Config.PortaTekProt;

            OnAfterValidate := AposValidar;
            OnGetCloseApp   := QuandoNaoAutorizado;

            validarLicenca;

            // A função abaixo retorna erros no result, assim uma pequena gambia
            // para saber se o retorno que está chegando é um erro ou o valor esperado
            sRetorno := GetLicenseInfo(liClientInfo);
            if Pos('Ocorreram falhas', sRetorno) > 0 then
              raise Exception.Create(sRetorno);

            tpStrListTekProt := TStringList.Create;
            try
              TFuncoesString.DividirParaStringList(tpStrListTekProt, sRetorno, '|');
              with tpStrListTekProt do
                begin
                  FCodigoClienteTek           := Strings[0];
                  FCNPJClienteTek             := Strings[1];
                  FNomeClienteTek             := Strings[2];
                  FEmpresasDisponiveis        := Strings[3];
                  FFuncionalidadesDisponiveis := Strings[4];
                end;
            finally
              if Assigned(tpStrListTekProt) then
                 FreeAndNil(tpStrListTekProt);
            end;
          end;
      except
        on E: Exception do
          begin
            TCaixasDeDialogo.Erro('Ocorreu o seguinte erro ao tentar validar a sua cópia do sistema: ' + #13 +
                                  E.Message + #13 +
                                  'Servidor: ' + TekProtClient.Server + ', porta: ' +
              IntToStr(TekProtClient.Port));

            if TCaixasDeDialogo.Confirma('Deseja configurar servidor de proteção?') then
              ChamarConfig;

            TCaixasDeDialogo.Erro('Sistema será finalizado, acesse novamente se desejar.');

            FecharSistema;
          end;
      end;
    end;
  {$endregion}
end;

function TDMConexao.ConectaServidorAplicacao(cUsuario, cSenha: string; iQuebra: Integer): Boolean;
var
  Primeira: Boolean;
  Servidor, ServidorTP, S: string;
  Porta, PortaTP: Integer;
  SL: TStrings;
begin
  Result := false;

  if Assigned(FSplash) then
  begin
    FSplash.Show;
    FSplash.Passo(30, 'Preparando Conexão...');
  end;

  Porta := 0;
  PortaTP := 0;

  CDSServidores.First;

  Primeira := True;
  while True do
    try
      if Primeira then
      begin
        if Assigned(FSplash) then
          FSplash.Passo(35, 'Conectando ao Servidor de Aplicação Principal');

        SQLConexao.Connected := false;

        with SQLConexao.Params do
        begin
          Values[TDBXPropertyNames.DSAuthenticationUser] := cUsuario;
          Values[TDBXPropertyNames.DSAuthenticationPassword] := TFuncoesCriptografia.Codifica(cSenha, sChaveCriptografia);

          Values['Quebra'] := IntToStr(iQuebra);
        end;
      end
      else
      begin
        if Assigned(FSplash) then
          FSplash.Passo(40, 'Tentando Servidores de Aplicação Secundários');
      end;

      ConexaoTratada := True;
      try
        if CDSServidores.FieldByName('Tipo').AsInteger = cServTipo_SoftwareCenter then
        begin
          SL := TStringList.Create;
          try
            SL.Values[cConfig_Sistema] := IntToStr(86);
            SL.Values[cConfig_Ambiente] := IntToStr(cTipoBDD_Producao);
            S := SQLConexao.Params.Text;
            S := TFuncoesString.Trocar(S, #13#10, '|');
            SL.Values[cConfig_Paramametros] := S;

            try
              SL.Text := TFuncoes_TekConnects.Configuracoes(
                CDSServidores.FieldByName('Servidor').AsString,
                CDSServidores.FieldByName('Porta').AsInteger,
                SL.Text);
            except
              on E: Exception do
              begin
                if not CDSServidores.eof then
                begin
                  CDSServidores.Next;
                  Continue;
                end;
                raise;
              end;
            end;

            Servidor := SL.Values[cConfig_ServAplicHost];
            Porta := StrToIntDef(SL.Values[cConfig_ServAplicPorta], 0);

            ServidorTP := SL.Values[cConfig_ProtecaoHost];
            PortaTP := StrToIntDef(SL.Values[cConfig_ProtecaoPorta], 0);
          finally
            SL.Free;
          end;
        end else begin
          Servidor := CDSServidores.FieldByName('Servidor').AsString;
          Porta := CDSServidores.FieldByName('Porta').AsInteger;
        end;

        SQLConexao.Connected := false;

        with SQLConexao.Params do
        begin
          Values[TDBXPropertyNames.HostName] := Servidor;
          Values[TDBXPropertyNames.Port] := IntToStr(Porta);

          // if Proxy_Utiliza then
          begin
            Values[TDBXPropertyNames.DSProxyHost] := CDSServidores.FieldByName('Proxy_Host').AsString;
            Values[TDBXPropertyNames.DSProxyPort] := CDSServidores.FieldByName('Proxy_Porta').AsString;
            Values[TDBXPropertyNames.DSProxyUsername] := CDSServidores.FieldByName('Proxy_Usuario').AsString;
            Values[TDBXPropertyNames.DSProxyPassword] := CDSServidores.FieldByName('Proxy_Senha').AsString;
          end;
        end;

        SQLConexao.Connected := True;

        if CDSServidores.FieldByName('Tipo').AsInteger = cServTipo_SoftwareCenter then
        begin
          Config.ServidorRelatorio := '';
          Config.PortaRelatorio    := 0;

          Config.ServidorTekProt   := ServidorTP;
          Config.PortaTekProt      := PortaTP;
        end else begin
          Config.ServidorRelatorio := CDSServidores.FieldByName('Secundario_Host').AsString;
          Config.PortaRelatorio    := CDSServidores.FieldByName('Secundario_Porta').AsInteger;

          Config.ServidorTekProt   := CDSServidores.FieldByName('Protecao_Host').AsString;
          Config.PortaTekProt      := CDSServidores.FieldByName('Protecao_Porta').AsInteger;
        end;

        Config.ServidorProxy     := CDSServidores.FieldByName('Proxy_Host').AsString;
        Config.PortaProxy        := CDSServidores.FieldByName('Proxy_Porta').AsInteger;
        Config.UsuarioProxy      := CDSServidores.FieldByName('Proxy_Usuario').AsString;
        Config.SenhaProxy        := CDSServidores.FieldByName('Proxy_Senha').AsString;

        Config.AcessoPelaInternet := CDSServidores.FieldByName('Rede').AsInteger = cServRede_Extranet;

        Result := True;
        Exit;
      finally
        ConexaoTratada := false;
        Primeira := false;
      end;

      Break;
    except
      on E: EIdSocketError do
      begin
        if not CDSServidores.eof then
        begin
          CDSServidores.Next;
          Continue;
        end;

        if TCaixasDeDialogo.Confirma('Servidor de Aplicação não está rodando em ' +
          Servidor + '/' + IntToStr(Porta) + '.' + #13 +
          'Tentar entrar novamente?' + #13#13 +
          'Mensagem Original: ' + E.Message) then
        begin
          LerServidoresAplicacao;
          Primeira := True;
        end
        else
        begin
          if TCaixasDeDialogo.Confirma('Deseja configurar para outros servidores?') then
          begin
            ChamarConfig;
            Primeira := True;
          end
          else
            Exit;
        end;
      end;

      on E: Exception do
      begin
        TCaixasDeDialogo.Erro(E.Message);
        Break;
      end;
    end;
end;

procedure TDMConexao.LerServidoresAplicacao;
const
  MaxTentativas = 5;
var
  Arq: string;
  Count, Tentativa: Integer;
  CDS: TClientDataSet;
begin
  if Assigned(FSplash) then
    FSplash.Passo(20, 'Lendo Servidores de Aplicação');

  // Configurações Padrões
  with Config do
  begin
    RelRemalina := false;
    RelRodapes := True;
    RelZebrado := True;
    RelSalvar := false;
    RelImpressoraEspecial := false;
    RelImpressaoGrafica := false;

    DirExe := ExtractShortPathName(ExtractFilePath(Application.ExeName));
    DirRoot := Copy(DirExe, 1, Length(DirExe) - 1);
    DirRoot := Copy(DirRoot, 1, TFuncoesString.PosDireita('\', DirRoot));

    DirTemp := DirRoot + 'TEMP';
    DirTempCDS := DirRoot + 'TEMP\CDS';
    DirTempDANFe := DirRoot + 'TEMP\NFE\DANF';
    // DirTempRetornoNFe := DirRoot + 'TEMP\NFE\RETORNO';

    AcessoPelaInternet := false;
    MinutosOciosidade := 5;
  end;

  Arq := ExtractFilePath(ParamStr(0)) + '\' + ArquivoServidoresAplicacao;
  CDS := TClientDataSet.Create(nil);
  try
    with CDSServidores do
    begin
      IndexName := '';
      Count := 0;
      while True do
      begin
        Close;
        CreateDataSet;
        LogChanges := false;

        if FileExists(Arq) then
        begin
          // No caso da abertura de diversos módulos simultâneos,
          // por exemplo com terminal server
          Tentativa := 1;
          while (Tentativa <= MaxTentativas) do
            try
              CDS.LoadFromFile(Arq);
              Break;
            except
              on E: Exception do
              begin
                if (Tentativa = MaxTentativas) then
                begin
                  if TCaixasDeDialogo.Confirma(
                    'Foram feitas ' + IntToStr(MaxTentativas) + ' tentativas de leitura do arquivo que contém a lista de servidores de aplicação.' +
                    'No entanto. não foi possível sua abertura, devido o erro:' + #13 + E.Message + #13 +
                    'Gostaria de tentar novamente agora?') then
                    Tentativa := 1
                  else
                    FecharSistema;
                end
                else
                begin
                  Inc(Tentativa);
                  Application.ProcessMessages;
                end;
              end;
            end;

          CDSServidores.DisableConstraints;
          CDSServidores.CopiarRegistros(CDS, True);
          CDSServidores.EnableConstraints;
          Count := RecordCount;

          // Ler Configurações
          with Config do
          begin
            RelRemalina := CDS.GetOptionalParam('Relatorios->Remalina');
            RelRodapes := CDS.GetOptionalParam('Relatorios->Rodapes');
            RelZebrado := CDS.GetOptionalParam('Relatorios->Zebrada');
            RelSalvar := CDS.GetOptionalParam('Relatorios->Salvar');
            RelImpressoraEspecial := CDS.GetOptionalParam('Relatorios->ImpressoraEspecial');
            RelImpressaoGrafica := CDS.GetOptionalParam('Relatorios->ImpressaoGrafica');

            DirTemp := CDS.GetOptionalParam('Diretorios->Temporario');
            DirTempCDS := DirTemp + '\CDS';

            DirTempDANFe := CDS.GetOptionalParam('Diretorios->DANFe');
            if Trim(DirTempDANFe) = '' then
              DirTempDANFe := DirTemp + '\NFE\DANF';

            MinutosOciosidade := CDS.GetOptionalParam('Ociosidade->Minutos');
          end;

        end;
        if (Count = 0) and (TCaixasDeDialogo.Confirma('Servidores de Aplicação não estão configurados, deseja configurá-los agora?')) then
          ChamarConfig(false)
        else
          Break;
      end;
      if (Count = 0) then
        FecharSistema;

      // Adiciona o primeiro servidor como principal
      IndexFieldNames := 'Ordem';
    end;
  finally
    FreeAndNil(CDS);
  end;

  with Config do
  begin
    if not DirectoryExists(DirTemp) then
      ForceDirectories(DirTemp);
    if not DirectoryExists(DirTempCDS) then
      ForceDirectories(DirTempCDS);
    if not DirectoryExists(DirTempDANFe) then
      ForceDirectories(DirTempDANFe);
    // if not DirectoryExists(DirTempRetornoNFe) then
    // ForceDirectories(DirTempRetornoNFe);
  end;
end;

procedure TDMConexao.ChamarConfig(ReLer: Boolean = True);
var
  Xml_Ant, Xml_Dep: WideString;
  ST_Ant, ST_Dep: TStream;
begin
  Application.CreateForm(TFConfigServApl, FConfigServApl); // No Create lê ou cria o CDSServidores
  ST_Ant := TMemoryStream.Create;
  ST_Dep := TMemoryStream.Create;
  try
    CDSServidores.First;
    CDSServidores.SaveToStream(ST_Ant, dfXML); // dfBinary, dfXML, dfXMLUTF8
    CDSServidores.Last;
    if (FConfigServApl.ShowModal = Controls.mrOK) and (ReLer) then
    begin
      CDSServidores.First;
      CDSServidores.SaveToStream(ST_Dep, dfXML);

      ST_Ant.Position := 0;
      ListaDeStrings.LoadFromStream(ST_Ant);
      Xml_Ant := ListaDeStrings.Text;

      ST_Dep.Position := 0;
      ListaDeStrings.LoadFromStream(ST_Dep);
      Xml_Dep := ListaDeStrings.Text;

      if (CompareText(Xml_Ant, Xml_Dep) <> 0){ and Assigned(FPrincipal)} then
        EntrarNoSistema(True);
    end;
  finally
    FConfigServApl.Free;
    FreeAndNil(ST_Ant);
    FreeAndNil(ST_Dep);
  end;
end;

procedure TDMConexao.ExecutaRelatorioGR(CodigoRel: Integer; Filtros: OleVariant; Formato: Integer; ProcessarLocal, EmSegundoPlano: Boolean);
var
  Rel_MS: TMemoryStream;
  Rel_Ole: OleVariant;
  FR: TfrxReport;
  Titulo, ArqTemp: String;
  Parametros: TParametros;
  LocalProcessamento: TLocalProcessamento;
begin
  Rel_MS := TMemoryStream.Create;
  FR := TfrxReport.Create(Self);

  // Até o momento não poderá ser processando no cliente pois há possibilidade de processamento
  // de class de relatório que pode não está disponível no módulo de execução
  ProcessarLocal := false;

  EmSegundoPlano := EmSegundoPlano or (DMConexao.Config.ServidorRelatorio <> '');

  if not EmSegundoPlano then
  begin
    DMConexao.CallBack_AbreTela(Self.ClassName, 'Execução de Relatório');
    EmProcesso := True;
  end;

  try
    Titulo := 'Rel. GR:' + IntToStr(CodigoRel);

    if EmSegundoPlano then
      LocalProcessamento := ClassPaiProcessamento.fLocal_SegundoPlano
    else if ProcessarLocal then
      LocalProcessamento := ClassPaiProcessamento.fLocal_Atual
    else
      LocalProcessamento := ClassPaiProcessamento.fLocal_Servidor;

    if not EmSegundoPlano then
      DMConexao.CallBack_Mensagem(Self.ClassName, 'Gerando Relatório: ' + IntToStr(CodigoRel));

    TGeradorRelatorioEspecifico.GravarParametro(Parametros, 'CodigoRelatorio', CodigoRel);
    TGeradorRelatorioEspecifico.GravarParametro(Parametros, 'Formato', Formato);
    TGeradorRelatorioEspecifico.GravarParametro(Parametros, 'Filtros', Filtros);

    Rel_Ole := TGeradorRelatorioEspecifico.ProcessarClasse(DMConexao, Parametros, Self, False, LocalProcessamento);

    if not EmSegundoPlano then
      DMConexao.CallBack_Mensagem(Self.ClassName, 'Recebendo dados');

    TFuncoesConversao.OleVariantParaStream(Rel_Ole, Rel_MS);

    ArqTemp := TFuncoesSistemaOperacional.DiretorioComBarra(DMConexao.Config.DirTemp) + 'RelatorioGerado-' + IntToStr(CodigoRel) + '.TMP';
    case Formato of
      1: // frArquivoFR:
        begin
          FR.PreviewPages.LoadFromStream(Rel_MS);
          FR.ShowPreparedReport;
        end;
      0: // frArquivoPDF:
        ArqTemp := ChangeFileExt(ArqTemp, '.PDF');
      2: // frArquivoTXT:
        ArqTemp := ChangeFileExt(ArqTemp, '.TXT');
      3: // frArquivoHTML:
        ArqTemp := ChangeFileExt(ArqTemp, '.HTML');
      4: // frArquivoJPEG:
        ArqTemp := ChangeFileExt(ArqTemp, '.JPEG');
      5: // frArquivoCSV:
        ArqTemp := ChangeFileExt(ArqTemp, '.CSV');
    end;

    if (Formato <> 1) then
    begin
      if not EmSegundoPlano then
        DMConexao.CallBack_Mensagem(Self.ClassName, 'Salvando');

      Rel_MS.SaveToFile(ArqTemp);

      // if EmSegundoPlano then
      // begin
      // FPrincipal.Alerta('Disponível: ' + Titulo, ArqTemp, AbrirArquivoDoAlert);
      // end else begin
      if (ShellExecute(0, nil, PWideChar(ArqTemp), nil, nil, SW_SHOWNORMAL) < 32) then
        TCaixasDeDialogo.Informacao('O arquivo está disponível em ' + ArqTemp + ', mas não foi possível abri-lo diretamente');
      // end;
    end;
  finally
    Rel_MS.Free;
    FR.Free;
    if not EmSegundoPlano then
    begin
      DMConexao.CallBack_FechaTela('');
      EmProcesso := false;
    end;
  end;
end;

function TDMConexao.PegaEmpresaDoMovimentoEstoque(iEmp: Integer): Integer;
begin
  Result := ExecuteScalar(
    ' select ' +
    ' CONFIG_SISTEMA_EMPRESA.EMP_MOVESTOQUE_QUAEXTRA_CFSEMP ' +
    ' from CONFIG_SISTEMA_EMPRESA where CONFIG_SISTEMA_EMPRESA.EMPRESA_CFSEMP = ' + IntToStr(iEmp));
  if (Result = 0) then
    Result := iEmp;
end;

function TDMConexao.PegaEmpresasFicticias(iEmp: Integer): string;
var
  CDSTemp: TClientDataSet;
begin
  CDSTemp := TClientDataSet.Create(nil);
  try
    CDSTemp.Data := ExecuteReader(
      'select ' + #13 +
      '  CONFIG_SISTEMA_EMPRESA.EMPRESA_CFSEMP ' + #13 +
      'from CONFIG_SISTEMA_EMPRESA ' + #13 +
      'where CONFIG_SISTEMA_EMPRESA.EMP_MOVESTOQUE_QUAEXTRA_CFSEMP = ' + IntToStr(SecaoAtual.Empresa.Codigo));

    Result := '';
    CDSTemp.First;
    while (not CDSTemp.eof) do
    begin
      Result := Result + CDSTemp.Fields[0].AsString + ',';
      CDSTemp.Next;
    end;

    if (Result <> '') then
      Result := Copy(Result, 1, Length(Result) - 1);
  finally
    FreeAndNil(CDSTemp);
  end;
end;

procedure TDMConexao.TrataOciosidade(var Msg: tagMSG);
var
  X: Integer;
begin
  if ((Msg.Message = WM_MOUSEMOVE) or // qualquer movimento do mouse.
    (Msg.Message = WM_KEYDOWN) or // qualquer tecla pressionada.
    (Msg.Message = WM_LBUTTONDOWN) or // botão esquerdo do mouse
    (Msg.Message = WM_RBUTTONDOWN) or // botão direito do mouse
    (Msg.Message = WM_MOUSEWHEEL) or // Roda do Mouse
    (Msg.Message = WM_SYSKEYDOWN)) and // tecla de sistema
    (not(Assigned(FRegresso))) then
    TempoOcio := GetTickCount
  else if (not EmProcesso) and
    (not Debugando) and
    (not Assigned(FRegresso)) and
    (not(Screen.ActiveForm is TFPainelBordo)) and
    (not(Screen.ActiveForm is TFPainelBordo2)) and
  // (not RelatorioRD.RDAberto) and
  // (not FastAberto) and
  // (not QuickAberto) and
    ((GetTickCount - TempoOcio) > DWORD(Config.MinutosOciosidade * 60 * 1000) - (15 * 1000)) then
  begin
    Application.Restore;
    FRegresso := TFRegresso.Create(Application);
    try
      if (FRegresso.ShowModal = mrOK) then
      begin
        with Screen do
          for X := 0 to ComponentCount - 1 do
            if (Components[X] is TClientDataSet) then
              (Components[X] as TClientDataSet).Close;
        FecharSistema;
      end
      else
        TempoOcio := GetTickCount;
    finally
      FRegresso.Free;
    end;
  end;
end;

procedure TDMConexao.AbrirArquivoDoAlert(Sender: TObject);
var
  ArqTemp: string;
begin
  if not(Sender is TJVDesktopAlert) then
    Exit;

  ArqTemp := TJVDesktopAlert(Sender).MessageText;
  if (ShellExecute(0, nil, PWideChar(ArqTemp), nil, nil, SW_SHOWNORMAL) < 32) then
    TCaixasDeDialogo.Informacao('O arquivo está disponível em ' + ArqTemp + ', mas não foi possível abri-lo diretamente.');
end;

function TDMConexao.GetContadorTransacoesTemporarias: Integer;
begin
  Dec(FContadorTransacoesTemporarias);
  Result := FContadorTransacoesTemporarias;
end;

procedure TDMConexao.ImportarRegistrosParaCDS(BotaoIncluir, BotaoGravar: TNewBtn;
Tabela, NomeDoArquivo: string; CDSDestino: TClientDataSet;
AntesDeGravar: TDataSetNotifyEvent = nil; AntesDeAceitar: TDataSetNotifyEvent = nil);
var
  CDS: TClientDataSet;
begin
  if CDSDestino.State in [dsEdit, dsInsert] then
  begin
    TCaixasDeDialogo.Aviso('Termine a edição do registro atual antes de solicitar importação de registros');
    Exit;
  end;

  with OpenDialogReg do
  begin
    InitialDir := Config.DirTemp;
    Filename := TFuncoesSistemaOperacional.NomeArquivoValido(NomeDoArquivo) + '.CDS';

    if Execute then
    begin
      CDS := TClientDataSet.Create(nil);
      try
        with CDS do
        begin
          LoadFromFile(OpenDialogReg.Filename);

          if (Tabela <> GetOptionalParam('Tabela')) then
          begin
            TCaixasDeDialogo.Aviso('Arquivo incompatível com o cadastro em questão');
            Exit;
          end;

          if Assigned(AntesDeAceitar) then
            AntesDeAceitar(CDS);

          First;
          while not eof do
          begin
            BotaoIncluir.Click;
            CDSDestino.DisableControls;
            try
              CDSDestino.CopiarRegistros(CDS, false, AntesDeGravar);
            finally
              CDSDestino.EnableControls;
            end;
            BotaoGravar.Click;

            Next;
          end;
        end;
      finally
        FreeAndNil(CDS);
      end;
    end;
  end;
end;

procedure TDMConexao.ExportarRegistrosCDS(Tabela, NomeDoArquivo: string; CDSOrigem: TClientDataSet);
begin
  if (not CDSOrigem.Active) or (CDSOrigem.IsEmpty) then
  begin
    TCaixasDeDialogo.Aviso('Tabela deve estar aberta e não deve estar vazia para fazer exportação de registros');
    Exit;
  end;

  if (CDSOrigem.State in [dsEdit, dsInsert]) then
  begin
    TCaixasDeDialogo.Aviso('Salve o registro antes de fazer a exportação do mesmo');
    Exit;
  end;

  with SaveDialogReg do
  begin
    InitialDir := Config.DirTemp;
    Filename := TFuncoesSistemaOperacional.NomeArquivoValido(NomeDoArquivo) + '.CDS';

    if Execute then
    begin
      NomeDoArquivo := Filename;
      case FilterIndex of
        1:
          begin
            NomeDoArquivo := ChangeFileExt(NomeDoArquivo, '.XML');
            CDSOrigem.SaveToFile(NomeDoArquivo, dfXMLUTF8);
          end;
        2:
          begin
            CDSOrigem.SetOptionalParam('Tabela', Tabela);
            NomeDoArquivo := ChangeFileExt(NomeDoArquivo, '.CDS');
            CDSOrigem.SaveToFile(NomeDoArquivo, dfBinary);
          end;
      end;
      TCaixasDeDialogo.Informacao(NomeDoArquivo + ' gerado com sucesso');
    end;
  end;
end;

function TDMConexao.Ler(Campos, Tabela: string; Ordem: Integer; Where: string = ''): OleVariant;
begin
  with ListaDeStrings do
  begin
    Clear;
    Add('select ' + Campos);
    Add(' from ' + Tabela);
    if Length(Trim(Where)) > 0 then
      Add(Where);
    if Ordem >= 0 then
      Add(' order by ' + IntToStr(Ordem));
    Result := DMConexao.ExecuteReader(Text);
  end;
end;

function TDMConexao.Acha(Tabela, Campo: string; Valor: Variant; CampoEmpresa: string = ''; CodigoDaEmpresa: Integer = -1): Boolean;
var
  CDSAcha: TClientDataSet;
  E: Integer;
begin
  if (Tabela = '') or (Campo = '') then
    raise Exception.Create('Função Acha: Nome da tabela e campo deve ser informado.');

  with ListaDeStrings do
  begin
    Clear;
    Add('select ' + Tabela + '.' + Campo);
    Add('from ' + Tabela);
    Add('where');
    if CampoEmpresa <> '' then
    begin
      if CodigoDaEmpresa <> -1 then
        E := CodigoDaEmpresa
      else if (ConstanteSistema.Sistema in [cSistemaDepPessoal, cSistemaContabilidade]) then
        E := SecaoAtual.Empresa.Estabelecimento
      else
        E := SecaoAtual.Empresa.Codigo;
      Add(Tabela + '.' + CampoEmpresa + ' = ' + IntToStr(E));
      Add(' and ');
    end;
    Add(Tabela + '.' + Campo + ' = ' + QuotedStr(Valor));
  end;

  CDSAcha := TClientDataSet.Create(nil);
  try
    CDSAcha.Data := DMConexao.ExecuteReader(ListaDeStrings.Text, 1);
    Result := not CDSAcha.IsEmpty;
  finally
    FreeAndNil(CDSAcha);
  end;
end;

procedure TDMConexao.AnalisaremCubodeDeciso1Click(Sender: TObject);
var
  Grade: TDBGrid;
  CDS: TClientDataSet;
  Descricao: String;
begin
  if not (Screen.ActiveControl is TDBGrid) then
    Exit;

  Grade := (Screen.ActiveControl as TDBGrid);

  if not Assigned(Grade.DataSource) then
    Exit;

  if not (Grade.DataSource.DataSet is TClientDataSet) then
    Exit;

  CDS := (Grade.DataSource.DataSet as TClientDataSet);

  if not CDS.Active then
    Exit;

  Descricao := '';
  if (Grade.Parent is TGroupBox) and
     ((Grade.Parent as TGroupBox).Caption <> '') then
    Descricao := (Grade.Parent as TGroupBox).Caption
  else if (Grade.Parent is TTabSheet) and
          ((Grade.Parent as TTabSheet).TabVisible) and
          ((Grade.Parent as TTabSheet).Caption <> '') then
    Descricao := (Grade.Parent as TTabSheet).Caption
  else if (Grade.Owner is TForm) then
    Descricao := (Grade.Owner as TForm).Caption;

  TFCuboDeDecisao.Abrir(0, Descricao, CDS.Data, False);
end;

procedure TDMConexao.FecharSistema;
begin
  if Assigned(FSplash) then
  begin
    FSplash.Free;
    FSplash := nil;
  end;

  if SQLConexao.Connected then
    SQLConexao.Close;

  Application.Terminate;
  ExitProcess(0);
end;

procedure TDMConexao.SetStatusDeMensagens(const Value: Integer);
begin
  if (FStatusDeMensagens = Value) then
    Exit;

  if (Value < 0) then
  begin
    Windows.Beep(700, 150);
    Windows.Beep(900, 150);
    Windows.Beep(1100, 150);
  end;

  FStatusDeMensagens := Value;
  if (FStatusDeMensagens = 0) then
  begin
    FPrincipal.BotaoMensagem.Images.ActiveIndex := 1;
    FPrincipal.StatusIndicator_Msg.Visible := false;
  end
  else if (FStatusDeMensagens > 0) then
  begin
    FPrincipal.BotaoMensagem.Images.ActiveIndex := 0;
    FPrincipal.StatusIndicator_Msg.Visible := True;
    if Value <= 99 then
      FPrincipal.StatusIndicator_Msg.Caption := IntToStr(Value)
    else
      FPrincipal.StatusIndicator_Msg.Caption := '..';
  end
  else
  begin
    FPrincipal.BotaoMensagem.Images.ActiveIndex := 0;
    FPrincipal.StatusIndicator_Msg.Visible := false;
  end;

  case FStatusDeMensagens of
    cStatusDeMensagens_FechaSistema:
      begin
        FPrincipal.PanelMensagem.Hint := 'Foi solicitado o fechamento do sistema';
        FPrincipal.JvBalloonHint1.ActivateHint(FPrincipal.BotaoMensagem, FPrincipal.PanelMensagem.Hint, 'Mensagem');
        // FPrincipal.StatusIndicator_Msg.Cor := clRed;
        TrataMensagemDeFechamento(True);
      end;
    cStatusDeMensagens_PedidoAutorizacao:
      begin
        FPrincipal.PanelMensagem.Hint := 'Há processos aguardando a sua autorização';
        FPrincipal.JvBalloonHint1.ActivateHint(FPrincipal.BotaoMensagem, 'Novo processo aguardando a sua autorização', 'Mensagem');
        // FPrincipal.StatusIndicator_Msg.Cor := $000080FF;
      end;
    cStatusDeMensagens_Autorizacao:
      begin
        FPrincipal.PanelMensagem.Hint := 'Foi concedida a autorização para execução de algum processo solicitado';
        FPrincipal.JvBalloonHint1.ActivateHint(FPrincipal.BotaoMensagem, FPrincipal.PanelMensagem.Hint, 'Mensagem');
        // FPrincipal.StatusIndicator_Msg.Cor := clLime;
      end;
    cStatusDeMensagens_NegacaoAutorizacao:
      begin
        FPrincipal.PanelMensagem.Hint := 'Foi negada a autorização para execução de algum processo solicitado';
        // FPrincipal.JvBalloonHint1.ActivateHint(FPrincipal.BotaoMensagem, FPrincipal.PanelMensagem.Hint, 'Mensagem');
        // FPrincipal.StatusIndicator_Msg.Cor := clBlack;
      end;
    cStatusDeMensagens_SemMensagem:
      begin
        FPrincipal.PanelMensagem.Hint := 'Não há novas mensagens';
        // FPrincipal.StatusIndicator_Msg.Cor := clBtnFace;
      end;
    cStatusDeMensagens_ExisteMensagem:
      begin
        FPrincipal.PanelMensagem.Hint := Format('Há %d mensagem(ns) não lida(s)', [Value]);
        // FPrincipal.JvBalloonHint1.ActivateHint(FPrincipal.BotaoMensagem, 'Há Nova(s) mensagem(ns)', 'Mensagem');
        // FPrincipal.StatusIndicator_Msg.Cor := clBlue;
      end;
  else
    FPrincipal.PanelMensagem.Hint := '';
  end
end;

procedure TDMConexao.SetStatusSAC(const Value: Integer);
begin
  if FStatusSAC = Value then
    Exit;

  if (Value < 0) then
  begin
    Windows.Beep(700, 150);
    Windows.Beep(900, 150);
    Windows.Beep(1100, 150);
  end;

  FStatusSAC := Value;
  if (FStatusSAC < 0) then
    Exit;

{$IF Defined(QUALIDADE)}
  with FPrincipal do
  begin
    if FStatusSAC > 0 then
    begin
      BotaoSAC.Images.ActiveIndex := 4;
      BotaoSAC.Hint := Format('Há %d atendimento(s) em andamento', [FStatusSAC]);
      StatusIndicator_SAC.Visible := True;
      if FStatusSAC <= 99 then
        StatusIndicator_SAC.Caption := IntToStr(Value)
      else
        StatusIndicator_SAC.Caption := '..';

      // JvBalloonHint1.ActivateHint(BotaoSAC, 'Existem Atendimentos em Andamento', 'SAC');
    end
    else
    begin
      BotaoSAC.Images.ActiveIndex := 3;
      StatusIndicator_SAC.Visible := false;
      BotaoSAC.Hint := 'SAC';
    end;
  end;
{$IFEND}
end;

procedure TDMConexao.TrataHelp(var Msg: tagMSG);
begin
  if ((Screen.ActiveForm <> nil) and (Screen.ActiveForm.ClassName <> 'TMessageForm') and
    (Msg.Message = WM_KEYDOWN) and (Msg.wParam = VK_F1) and
    (GetKeyState(VK_SHIFT) < 0)) then
  begin
    AbrirHelp
  end;
end;

procedure TDMConexao.AbrirHelp;
var
  CDSTemp: TClientDataSet;
  ArquivoHelp, NomeForm, HelpContextoForm, HelpContextoComponente, ComandoHelp: string;
  AbriuHelp: Boolean;
  ST: TMemoryStream;
begin
  if (LendoHelp) then
    Exit;

  LendoHelp := True;
  try
    AbriuHelp := false;
    ArquivoHelp := TFuncoesSistemaOperacional.DiretorioComBarra(Config.DirExe) + ChangeFileExt(Modulos[Sistema, 1], '.CHM');
    NomeForm := Screen.ActiveForm.Name;
    HelpContextoForm := IfThen(Screen.ActiveForm.HelpKeyword = '', '-', Screen.ActiveForm.HelpKeyword);
    HelpContextoComponente := IfThen(Screen.ActiveControl.HelpKeyword = '', '-', Screen.ActiveControl.HelpKeyword);

    ST := TMemoryStream.Create;
    CDSTemp := TClientDataSet.Create(Self);
    try
      TFuncoesSistemaOperacional.LerRecursoDLL('MAPEAMENTOHELP', sNomeDll, ST);
      CDSTemp.LoadFromStream(ST);
      CDSTemp.IndexFieldNames := 'Formulario;HelpKeyword_Form;HelpKeyword_Componente';
      CDSTemp.First;
      if CDSTemp.FindKey([NomeForm, HelpContextoForm, HelpContextoComponente]) then
      begin
        ComandoHelp := ArquivoHelp + '::/' + StringReplace(CDSTemp.FieldByName('Caminho').AsString, '\', '/', [rfReplaceAll]);
        AbriuHelp := HtmlHelp(0, PAnsiChar(AnsiString(ComandoHelp)), HH_DISPLAY_TOPIC, 0) <> 0;
      end;
      CDSTemp.Close;
    finally
      FreeAndNil(ST);
      FreeAndNil(CDSTemp);
    end;

    if (not AbriuHelp) then
      AbriuHelp := HtmlHelp(0, PAnsiChar(AnsiString(ArquivoHelp)), HH_DISPLAY_TOPIC, 0) <> 0;

    if (not AbriuHelp) then
      HtmlHelp(Application.Handle, PAnsiChar(AnsiString(ArquivoHelpGeral)), HH_DISPLAY_TOPIC, 0);
  finally
    LendoHelp := false;
  end;
end;

procedure TDMConexao.TrataMensagemDeFechamento(TemMensagem: Boolean);
var
  CDSTemp: TClientDataSet;
  AutoInc, Mens: string;
  DataHoraEnvio: TDateTime;
begin
  if TemMensagem then
  begin
    CDSTemp := TClientDataSet.Create(Self);
    try
      // 1 - Ler a mensagem de fechamento
      CDSTemp.Data := ExecutaMetodo('TSMMensagem.BuscaCabecalhosMensagens', [1, cStatusDeMensagens_FechaSistema, 'S']);
      CDSTemp.Data := ExecutaMetodo('TSMMensagem.BuscaDetalhesMensagem', [CDSTemp.FieldByName('AUTOINC_MENSAGEM').AsString]);
      with CDSTemp do
      begin
        AutoInc := FieldByName('AUTOINC_MENSAGEM').AsString;
        DataHoraEnvio := FieldByName('DATAHORAENVIO_MENSAGEM').AsDateTime;
        Mens :=
          ' O usuário ' + FieldByName('REMETENTE').AsString +
          ' solicitou o fechamento do sistema em ' + FieldByName('DATAHORAENVIO_MENSAGEM').AsString + #13 +
          ' Assunto: ' + TFuncoesCriptografia.DeCodifica(FieldByName('ASSUNTO_MENSAGEM').AsString, sChaveCriptografia);
        if (Trim(FieldByName('TEXTO_MENSAGEM').AsString) <> '') then
          Mens := Mens + #13#13 + Trim(TFuncoesCriptografia.DeCodifica(FieldByName('TEXTO_MENSAGEM').AsString, sChaveCriptografia));
        Close;
      end;
    finally
      FreeAndNil(CDSTemp);
    end;

    // 2 - Marcar como lida
    ExecutaMetodo('TSMMensagem.MarcaMensagemComoLida', [AutoInc]);
  end
  else
  begin
    DataHoraEnvio := DataHoraServidor;
    Mens := 'Foi solicitado o fechamento do sistema (ShutDown)';
  end;

  // Se o usuário receber a mensagem com mais de 3 minutos de atraso
  // é porque ele estava desconectado. Então não tem valia para ele
  if (DataHoraServidor - DataHoraEnvio) < (1 / 24 / 60) * 3 then
  begin
    Application.ProcessMessages;

    try
      FAguarde.Desativar('');
    except
    end;

    try
      FAguarde2.Desativar;
    except
    end;

    // 3 - Desconectar do servidor
    SQLConexao.Connected := false;

    // 4 - Exibir a mensagem
    TCaixasDeDialogo.Aviso(Mens);

    // 5 - Finalizar o sistema
    FecharSistema;
  end;
end;

procedure TDMConexao.HabilitarOpcaoDeFiltrarGrade(Grade: TDBGrid);
var
  MenuItem: TMenuItem;
begin
  MenuItem := Grade.PopupMenu.Items.Find(FiltrarRegistros1.Caption);

  if Assigned(MenuItem) then
    MenuItem.Visible := True;
end;

procedure TDMConexao.DBGridToClipBoard(DBGrid: TDBGrid; ComCabecalho, ApenasLinhaAtual, ApenasColunaAtual: Boolean);
begin
  if ApenasColunaAtual then
    DBGrid.GridParaClipBoard(ComCabecalho, ApenasLinhaAtual, DBGrid.SelectedField.FieldName)
  else
    DBGrid.GridParaClipBoard(ComCabecalho, ApenasLinhaAtual);

  TCaixasDeDialogo.Informacao('Informações transferidas para a área de transferência. Agora você pode colá-las em outros programas.');
end;

procedure TDMConexao.VerificaTrocaDeSenha;
var
  DtUltModificacao: TDateTime;
begin
  if SecaoAtual.Parametro.Seg_DiasTrocaSenhas <= 0 then
    Exit;

  DtUltModificacao := ExecuteScalar('select ULTIMATROCASENHA_USUARIO from USUARIO where CODIGO_USUARIO = ' + IntToStr(SecaoAtual.Usuario.Codigo));

  if (DtUltModificacao = 0) or
     (DtUltModificacao + SecaoAtual.Parametro.Seg_DiasTrocaSenhas < DataHora) then
  begin
    TCaixasDeDialogo.Informacao('Necessário efetuar a troca de senha.');
    Application.CreateForm(TFTrocaSenha, FTrocaSenha);
    try
      if (FTrocaSenha.ShowModal <> mrOK) then
        FecharSistema;
    finally
      FreeAndNil(FTrocaSenha);
    end;
  end;
end;

{$ENDREGION}

{$REGION 'Metodos compatibilidade externa'}


function TDMConexao.Funcao_AcbrExecuteCommand(s: string): int64;
begin
  Result := ExecuteCommand(s);
end;

function TDMConexao.Funcao_AcbrExecuteReader(s: string): OleVariant;
begin
  Result := ExecuteReader(s);
end;

function TDMConexao.Funcao_AcbrExecuteScalar(s: string): OleVariant;
begin
  Result := ExecuteScalar(s);
end;

function TDMConexao.Funcao_AcbrProximoCodigo(s: string): OleVariant;
begin
  Result := ProximoCodigo(s);
end;

{$ENDREGION}

{$REGION 'Deprecated - retirar futuramente'}

procedure TDMConexao.MostrarLog(TextoDoLog, Titulo: string; MostrarNoRichEdit: Boolean; ExibirLandscape: Boolean);
begin
  ULogSistema.Mostrar_Texto(TextoDoLog, Titulo, MostrarNoRichEdit, ExibirLandscape);
end;

procedure TDMConexao.MostrarLog(DataSet: TClientDataSet; NomeDoCampo: string; Titulo: string = ''; ExibirLandscape: Boolean = false);
begin
  ULogSistema.Mostrar_DataSet(DataSet, NomeDoCampo, Titulo, ExibirLandscape);
end;

procedure TDMConexao.MostrarLog(TextoDoLog: TStrings; MostrarNoRichEdit: Boolean = True; ExibirLandscape: Boolean = false);
begin
  ULogSistema.Mostrar_Texto(TextoDoLog.Text, '', MostrarNoRichEdit, ExibirLandscape);
end;

procedure TDMConexao.MostrarLog(MostrarNoRichEdit: Boolean; NomeDoArquivoDeLog: WideString; ExibirLandscape: Boolean);
begin
  ULogSistema.Mostrar_Arquivo(NomeDoArquivoDeLog, '', MostrarNoRichEdit, ExibirLandscape);
end;

procedure TDMConexao.MostrarLog(TextoDoLog: string; MostrarNoRichEdit: Boolean = True; ExibirLandscape: Boolean = false);
begin
  ULogSistema.Mostrar_Texto(TextoDoLog, '', MostrarNoRichEdit, ExibirLandscape);
end;

{$ENDREGION}

end.

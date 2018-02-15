unit UDMConexao;

interface

uses
  SysUtils, Classes, FMTBcd, Vcl.Graphics, Controls, Forms, Menus, Dialogs, AppEvnts, Registry, Variants, ExtCtrls, ShellApi, ComCtrls, ActiveX,
  Windows, Messages, IniFiles, StdCtrls, IdStack, DB, DBClient, SqlExpr, Provider, System.Math, System.StrUtils, DateUtils, IndyPeerImpl, Datasnap.DSHTTPCommon,
  Data.DBXCommon, Data.DBXDataSnap, Datasnap.DSConnect, DBXJSONReflect, DSProxy, System.Json, IPPeerClient, System.ImageList, DSHTTPLayer,
  ImgList, DBGrids, JvDBGrid;

type
  EFuncionalidadeNaoLiberada = class(Exception);

  TDMConexao = class(TDataModule)
    ConexaoDS: TSQLConnection;
    CDSConfigCamposClasses: TClientDataSet;
    CDSConfigClasses: TClientDataSet;
    procedure DataModuleCreate(Sender: TObject);
    procedure DataModuleDestroy(Sender: TObject);
  private
    procedure CarregarListaDeTabelasEProceduresDoBDD;
    procedure CarregarUnitsProtegidas;
    procedure CarregarProcessamentosProtegidos;
  public
    function ExecutaMetodo(Metodo: string; Parametros_Valor: array of OleVariant): OleVariant;
    function ExecuteReader(sql: string; TamanhoPacote: Integer = 1000; MonitoraSQL: Boolean = True): OleVariant;
    function ExecuteScalar(sql: string; MonitoraSQL: Boolean = True): OleVariant;
    function ExecuteCommand(sql: WideString; MonitoraSQL: Boolean = True): int64;
    function ExecuteCommand_Update(sql: WideString; Campo: string; Valor: OleVariant; MonitoraSQL: Boolean = True): int64;
    function ProximoCodigo(Tabela: string; Quebra: Integer = 0): int64;

    procedure CarregaDefaults;
    function GetCDSConfigCamposClasses: OleVariant;
    function GetCDSConfigClasses: OleVariant;
  end;

var
  DMConexao: TDMConexao;

implementation

uses Constantes, ClassDataSet;
{$R *.dfm}

procedure TDMConexao.DataModuleCreate(Sender: TObject);
begin

  CarregaDefaults;
  //CarregaConfigClasses;
  CarregarListaDeTabelasEProceduresDoBDD;
end;

procedure TDMConexao.DataModuleDestroy(Sender: TObject);
begin
  if ConexaoDS.Connected then
    ConexaoDS.Close;
end;


function TDMConexao.ExecutaMetodo(Metodo: string; Parametros_Valor: array of OleVariant): OleVariant;
begin
  //Result := FuncoesDataSnap.ExecutaMetodo_Sincrono(ConexaoDS, Metodo, Parametros_Valor, false);
end;

function TDMConexao.ProximoCodigo(Tabela: string): int64;
begin
  Tabela := AnsiUpperCase(Tabela);
  Result := ExecutaMetodo('TSMConexao.ProximoCodigo', [Trim(Tabela)]);
end;

function TDMConexao.ExecuteScalar(sql: string): OleVariant;
begin
  Result := ExecutaMetodo('TSMConexao.ExecuteScalar', [Trim(sql)]);
end;

function TDMConexao.ExecuteReader(sql: string): OleVariant;
begin
  Result := ExecutaMetodo('TSMConexao.ExecuteReader', [Trim(sql)]);
end;


function TDMConexao.ExecuteCommand(sql: WideString; MonitoraSQL: Boolean = True): int64;
begin
  //Result := TFuncoesConversao.VariantParaInt64(ExecutaMetodo('TSMConexao.ExecuteCommand', [Trim(sql), True]));
end;

procedure TDMConexao.CarregarListaDeTabelasEProceduresDoBDD;
begin
  //SynSQLSynPadrao.FunctionNames.Text := ExecutaMetodo('TSMConexao.GetProcedureNames', []);
  //SynSQLSynPadrao.TableNames.Text    := ExecutaMetodo('TSMConexao.GetTableNames',     [False]);
  //SynSQLSynPadrao.TableNames.Add(ExecutaMetodo('TSMConexao.GetTableNames', [True]));
end;

end.

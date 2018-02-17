program VendasClient;

uses
  Vcl.Forms,
  UFrmPrincipal in 'UFrmPrincipal.pas' {FrmPrincipal},
  UDMPai in 'Bases\UDMPai.pas' {DMPai: TDataModule},
  UDMPaiCadastro in 'Bases\UDMPaiCadastro.pas' {DMPaiCadastro: TDataModule},
  UFrmPai in 'Bases\UFrmPai.pas' {FrmPai},
  UPaiCadastro in 'Bases\UPaiCadastro.pas' {FPaiCadastro},
  ClassDataSet in '..\Class\ClassDataSet.pas',
  ClassExpositorDeClasses in '..\Class\ClassExpositorDeClasses.pas',
  ClassPai in '..\Class\ClassPai.pas',
  ClassPaiCadastro in '..\Class\ClassPaiCadastro.pas',
  ClassStatus in '..\Class\ClassStatus.pas',
  Constantes in '..\Class\Constantes.pas',
  UDMConexao in 'Bases\UDMConexao.pas' {DMConexao: TDataModule},
  UFrmPaiCadastro in 'Bases\UFrmPaiCadastro.pas' {FrmPaiCadastro};

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TFrmPai, FrmPai);
  Application.CreateForm(TFrmPaiCadastro, FrmPaiCadastro);
  Application.Run;
end.

object DMConexao: TDMConexao
  OldCreateOrder = False
  OnCreate = DataModuleCreate
  OnDestroy = DataModuleDestroy
  Height = 276
  Width = 507
  object ConexaoDS: TSQLConnection
    DriverName = 'Datasnap'
    LoginPrompt = False
    Params.Strings = (
      'DriverUnit=Data.DBXDataSnap'
      'HostName=localhost'
      'Port=211'
      'CommunicationProtocol=tcp/ip'
      'DatasnapContext=datasnap/'
      
        'DriverAssemblyLoader=Borland.Data.TDBXClientDriverLoader,Borland' +
        '.Data.DbxClientDriver,Version=16.0.0.0,Culture=neutral,PublicKey' +
        'Token=91d62ebb5b0d1b1b'
      'Filters={}')
    Left = 50
    Top = 65
  end
  object CDSConfigCamposClasses: TClientDataSet
    Aggregates = <>
    Params = <>
    Left = 348
    Top = 62
  end
  object CDSConfigClasses: TClientDataSet
    Aggregates = <>
    Params = <>
    Left = 213
    Top = 62
  end
end

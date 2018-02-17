inherited DMPaiCadastro: TDMPaiCadastro
  OldCreateOrder = True
  OnCreate = DataModuleCreate
  Width = 225
  object CDSCadastro: TClientDataSet
    Aggregates = <>
    Params = <>
    RemoteServer = DSPCCadastro
    Left = 48
    Top = 53
  end
  object DSPCCadastro: TDSProviderConnection
    Left = 49
    Top = 5
  end
end

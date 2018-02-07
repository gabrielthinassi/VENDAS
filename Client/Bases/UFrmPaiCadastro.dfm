inherited FrmPaiCadastro: TFrmPaiCadastro
  Caption = 'Formul'#225'rio de Cadastro'
  ClientHeight = 350
  ClientWidth = 500
  ExplicitWidth = 506
  ExplicitHeight = 379
  PixelsPerInch = 96
  TextHeight = 13
  object pnlBot: TPanel
    Left = 0
    Top = 309
    Width = 500
    Height = 41
    Align = alBottom
    BevelEdges = []
    BevelInner = bvLowered
    BevelKind = bkFlat
    BevelOuter = bvNone
    DoubleBuffered = False
    ParentDoubleBuffered = False
    TabOrder = 0
  end
  object pnlTop: TPanel
    Left = 0
    Top = 0
    Width = 500
    Height = 41
    Align = alTop
    BevelEdges = []
    BevelInner = bvLowered
    BevelKind = bkFlat
    BevelOuter = bvNone
    DoubleBuffered = False
    ParentDoubleBuffered = False
    TabOrder = 1
    object edtCodigo: TJvCalcEdit
      Left = 12
      Top = 12
      Width = 99
      Height = 21
      BevelInner = bvLowered
      BevelKind = bkFlat
      BevelOuter = bvNone
      Flat = False
      ParentFlat = False
      ImageKind = ikEllipsis
      ButtonWidth = 34
      TabOrder = 0
      DecimalPlacesAlwaysShown = False
    end
  end
  object pnlButtons: TPanel
    Left = 0
    Top = 41
    Width = 126
    Height = 268
    Align = alLeft
    BevelEdges = []
    BevelInner = bvLowered
    BevelKind = bkFlat
    BevelOuter = bvNone
    DoubleBuffered = False
    ParentDoubleBuffered = False
    TabOrder = 2
  end
  object tbctrlCadastro: TTabControl
    Left = 126
    Top = 41
    Width = 374
    Height = 268
    Align = alClient
    TabOrder = 3
    Tabs.Strings = (
      'Edit')
    TabIndex = 0
  end
end

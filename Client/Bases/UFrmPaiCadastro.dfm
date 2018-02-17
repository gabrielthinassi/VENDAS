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
    object txtAjuda: TDBText
      Left = 1
      Top = 1
      Width = 498
      Height = 39
      Align = alClient
      ExplicitLeft = 61
      ExplicitTop = 6
      ExplicitWidth = 65
      ExplicitHeight = 17
    end
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
    object pnlNavegar: TPanel
      Left = 315
      Top = 1
      Width = 184
      Height = 39
      Align = alRight
      BevelEdges = []
      BevelInner = bvLowered
      BevelKind = bkFlat
      BevelOuter = bvNone
      DoubleBuffered = False
      ParentDoubleBuffered = False
      TabOrder = 1
      object btnUltimo: TSpeedButton
        Left = 138
        Top = 1
        Width = 45
        Height = 37
        Align = alRight
        Caption = '>|'
        OnClick = btnUltimoClick
        ExplicitHeight = 32
      end
      object btnProximo: TSpeedButton
        Left = 93
        Top = 1
        Width = 45
        Height = 37
        Align = alRight
        Caption = '>'
        OnClick = btnProximoClick
        ExplicitHeight = 32
      end
      object btnAnterior: TSpeedButton
        Left = 48
        Top = 1
        Width = 45
        Height = 37
        Align = alRight
        Caption = '<'
        OnClick = btnAnteriorClick
        ExplicitTop = -3
      end
      object btnPrimeiro: TSpeedButton
        Left = 3
        Top = 1
        Width = 45
        Height = 37
        Align = alRight
        Caption = '|<'
        OnClick = btnPrimeiroClick
        ExplicitHeight = 32
      end
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
    object btnRelatorio: TSpeedButton
      Left = 1
      Top = 221
      Width = 124
      Height = 44
      Align = alTop
      Caption = '&Relat'#243'rio'
    end
    object btnIncluir: TSpeedButton
      Left = 1
      Top = 1
      Width = 124
      Height = 44
      Align = alTop
      Caption = '&Incluir'
      OnClick = btnIncluirClick
      ExplicitTop = 2
    end
    object btnExcluir: TSpeedButton
      Left = 1
      Top = 45
      Width = 124
      Height = 44
      Align = alTop
      Caption = '&Excluir'
      OnClick = btnExcluirClick
    end
    object btnGravar: TSpeedButton
      Left = 1
      Top = 89
      Width = 124
      Height = 44
      Align = alTop
      Caption = '&Gravar'
      OnClick = btnGravarClick
    end
    object btnCancelar: TSpeedButton
      Left = 1
      Top = 133
      Width = 124
      Height = 44
      Align = alTop
      Caption = '&Cancelar'
      OnClick = btnCancelarClick
    end
    object btnPesquisar: TSpeedButton
      Left = 1
      Top = 177
      Width = 124
      Height = 44
      Align = alTop
      Caption = '&Pesquisar'
    end
  end
  object pgctrlCadastro: TPageControl
    Left = 126
    Top = 41
    Width = 374
    Height = 268
    ActivePage = pgEdit
    Align = alClient
    TabOrder = 3
    object pgEdit: TTabSheet
      Caption = 'Edit'
    end
  end
  object DSCadastro: TDataSource
    OnStateChange = DSCadastroStateChange
    Left = 160
    Top = 75
  end
end

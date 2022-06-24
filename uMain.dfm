object frmMain: TfrmMain
  Left = 0
  Top = 0
  Caption = 'Gerenciador Microsoft Access'
  ClientHeight = 447
  ClientWidth = 724
  Color = clBtnFace
  Constraints.MinHeight = 486
  Constraints.MinWidth = 740
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -12
  Font.Name = 'Segoe UI'
  Font.Style = []
  Position = poDesktopCenter
  OnShow = FormShow
  PixelsPerInch = 96
  DesignSize = (
    724
    447)
  TextHeight = 15
  object Label1: TLabel
    Left = 6
    Top = 6
    Width = 38
    Height = 15
    Caption = 'Tabelas'
  end
  object Label2: TLabel
    Left = 203
    Top = 6
    Width = 44
    Height = 15
    Caption = 'Campos'
  end
  object Label3: TLabel
    Left = 434
    Top = 6
    Width = 58
    Height = 15
    Caption = 'Comandos'
  end
  object cbTabelas: TComboBox
    Left = 50
    Top = 3
    Width = 145
    Height = 23
    Style = csDropDownList
    TabOrder = 0
    OnChange = cbTabelasChange
  end
  object cbCampos: TComboBox
    Left = 253
    Top = 3
    Width = 145
    Height = 23
    Style = csDropDownList
    TabOrder = 1
    OnEnter = cbCamposEnter
  end
  object PageControl1: TPageControl
    Left = 8
    Top = 37
    Width = 714
    Height = 402
    ActivePage = TabSheet1
    Anchors = [akLeft, akTop, akRight, akBottom]
    TabOrder = 2
    OnChange = PageControl1Change
    ExplicitWidth = 852
    object TabSheet1: TTabSheet
      Caption = 'Manuten'#231#227'o SQL'
      DesignSize = (
        706
        372)
      object Memo1: TMemo
        Left = 3
        Top = 3
        Width = 629
        Height = 74
        Anchors = [akLeft, akTop, akRight]
        TabOrder = 0
        ExplicitWidth = 484
      end
      object btnExecuta: TBitBtn
        Left = 638
        Top = 1
        Width = 65
        Height = 77
        Anchors = [akTop, akRight]
        Caption = 'Executa'
        TabOrder = 1
        OnClick = btnExecutaClick
        ExplicitLeft = 493
      end
      object DBGrid1: TDBGrid
        Left = 3
        Top = 85
        Width = 695
        Height = 252
        Anchors = [akLeft, akTop, akRight, akBottom]
        DataSource = DataSource
        ParentShowHint = False
        ShowHint = False
        TabOrder = 2
        TitleFont.Charset = DEFAULT_CHARSET
        TitleFont.Color = clWindowText
        TitleFont.Height = -12
        TitleFont.Name = 'Segoe UI'
        TitleFont.Style = []
      end
      object BitBtn1: TBitBtn
        Left = 3
        Top = 343
        Width = 75
        Height = 25
        Caption = 'Post'
        TabOrder = 3
        OnClick = BitBtn1Click
      end
      object edtBancoPath: TEdit
        Left = 96
        Top = 344
        Width = 521
        Height = 23
        Anchors = [akLeft, akTop, akRight]
        ReadOnly = True
        TabOrder = 4
      end
      object BitBtn2: TBitBtn
        Left = 623
        Top = 343
        Width = 78
        Height = 25
        Caption = 'Base de dados'
        TabOrder = 5
        OnClick = BitBtn2Click
      end
    end
    object tabEstrutura: TTabSheet
      Caption = 'Estrutura'
      ImageIndex = 1
      DesignSize = (
        706
        372)
      object DBGrid2: TDBGrid
        Left = 4
        Top = 3
        Width = 696
        Height = 366
        Anchors = [akLeft, akTop, akRight, akBottom]
        DataSource = dsEstrutura
        ReadOnly = True
        TabOrder = 0
        TitleFont.Charset = DEFAULT_CHARSET
        TitleFont.Color = clWindowText
        TitleFont.Height = -12
        TitleFont.Name = 'Segoe UI'
        TitleFont.Style = []
        Columns = <
          item
            Expanded = False
            FieldName = 'NomeCampo'
            Title.Caption = 'Campo'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'TipoCampo'
            Title.Caption = 'Tipo'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'Tamanho'
            Visible = True
          end>
      end
    end
  end
  object cbComando: TComboBox
    Left = 498
    Top = 3
    Width = 119
    Height = 23
    Style = csDropDownList
    TabOrder = 3
    Items.Strings = (
      'Selecionar Dados'
      'Excluir Dados'
      'Criar Tabela'
      'Excluir Tabela'
      'Adicionar Campo'
      'Alterar Campo'
      'Apagar Campo'
      '')
  end
  object btnGerar: TBitBtn
    Left = 623
    Top = 2
    Width = 101
    Height = 25
    Caption = 'Script comando'
    TabOrder = 4
    OnClick = btnGerarClick
  end
  object CheckBox1: TCheckBox
    Left = 404
    Top = 6
    Width = 24
    Height = 17
    Hint = 'Mostrar inverso'
    ParentShowHint = False
    ShowHint = True
    TabOrder = 5
    OnClick = CheckBox1Click
  end
  object ADOConnection: TADOConnection
    LoginPrompt = False
    Mode = cmShareDenyNone
    Provider = 'Microsoft.Jet.OLEDB.4.0'
    Left = 304
    Top = 224
  end
  object ADOQuery: TADOQuery
    Connection = ADOConnection
    CursorType = ctStatic
    Parameters = <>
    SQL.Strings = (
      'SELECT * FROM Config')
    Left = 280
    Top = 144
  end
  object DataSource: TDataSource
    DataSet = ADOQuery
    Left = 416
    Top = 168
  end
  object QryEstrutura: TADOQuery
    Connection = ADOConnection
    CursorType = ctStatic
    Parameters = <>
    SQL.Strings = (
      'SELECT * FROM ESTRUTURA')
    Left = 416
    Top = 320
    object QryEstruturaCODIGO: TIntegerField
      FieldName = 'CODIGO'
    end
    object QryEstruturaTabela: TWideStringField
      FieldName = 'Tabela'
      Size = 30
    end
    object QryEstruturaNomeCampo: TWideStringField
      FieldName = 'NomeCampo'
      Size = 30
    end
    object QryEstruturaTipoCampo: TWideStringField
      FieldName = 'TipoCampo'
      Size = 25
    end
    object QryEstruturaTamanho: TIntegerField
      FieldName = 'Tamanho'
    end
  end
  object dsEstrutura: TDataSource
    DataSet = QryEstrutura
    Left = 512
    Top = 320
  end
  object qryUtil: TADOQuery
    Connection = ADOConnection
    CursorType = ctStatic
    Parameters = <>
    SQL.Strings = (
      'SELECT * FROM Config')
    Left = 200
    Top = 328
  end
  object OpenDialog1: TOpenDialog
    Filter = 'Arquivos Ms Access|*.mdb'
    Left = 552
    Top = 168
  end
  object SaveDialog1: TSaveDialog
    Left = 544
    Top = 232
  end
end

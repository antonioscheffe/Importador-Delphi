object FrmPrincipal: TFrmPrincipal
  Left = 409
  Top = 246
  Width = 620
  Height = 369
  Caption = 'Importar Rotas'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poDesktopCenter
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object lblSituacao: TLabel
    Left = 25
    Top = 128
    Width = 8
    Height = 29
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -24
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
  end
  object Label2: TLabel
    Left = 90
    Top = 53
    Width = 42
    Height = 13
    Caption = 'Arquivo :'
  end
  object lblNomeTransportadora: TLabel
    Left = 272
    Top = 47
    Width = 3
    Height = 13
  end
  object lbl1: TLabel
    Left = 56
    Top = 24
    Width = 83
    Height = 13
    Caption = 'Tipo Importa'#231#227'o :'
  end
  object btnImportar: TButton
    Left = 112
    Top = 86
    Width = 75
    Height = 25
    Caption = 'Importar'
    TabOrder = 0
    OnClick = btnImportarClick
  end
  object edtcaminho: TEdit
    Left = 141
    Top = 47
    Width = 441
    Height = 21
    Enabled = False
    TabOrder = 1
    Text = 
      'C:\Users\antonio.carlos\Desktop\Projeto Importador\BLUKIT-SC.xls' +
      'x'
  end
  object XStringGrid: TStringGrid
    Left = 448
    Top = 8
    Width = 320
    Height = 120
    TabOrder = 2
    Visible = False
  end
  object ProgressBar1: TProgressBar
    Left = 0
    Top = 310
    Width = 612
    Height = 28
    Align = alBottom
    TabOrder = 3
  end
  object btncarregar: TBitBtn
    Left = 24
    Top = 86
    Width = 75
    Height = 25
    Caption = 'Carregar'
    TabOrder = 4
    OnClick = btncarregarClick
  end
  object btnCancelar: TButton
    Left = 200
    Top = 86
    Width = 75
    Height = 25
    Caption = 'Cancelar'
    TabOrder = 5
    OnClick = btnCancelarClick
  end
  object GroupBox1: TGroupBox
    Left = 6
    Top = 176
    Width = 576
    Height = 129
    Caption = 'Log de Erros'
    TabOrder = 6
    object mmLog: TMemo
      Left = 7
      Top = 20
      Width = 560
      Height = 93
      Enabled = False
      TabOrder = 0
    end
  end
  object combotipoImportacao: TComboBox
    Left = 143
    Top = 17
    Width = 145
    Height = 21
    ItemHeight = 13
    ItemIndex = 0
    TabOrder = 7
    Text = 'Atualiza'#231#227'o'
    Items.Strings = (
      'Atualiza'#231#227'o'
      'Nova Importacao')
  end
  object connection1: TSQLConnection
    ConnectionName = 'OracleConnection'
    DriverName = 'Oracle'
    GetDriverFunc = 'getSQLDriverORACLE'
    LibraryName = 'dbexpora.dll'
    LoginPrompt = False
    Params.Strings = (
      'DriverName=Oracle'
      'DataBase=dbprod'
      'User_Name=wis50'
      'Password=wis50'
      'RowsetSize=20'
      'BlobSize=-1'
      'ErrorResourceFile='
      'LocaleCode=0000'
      'Oracle TransIsolation=ReadCommited'
      'OS Authentication=False'
      'Multiple Transaction=False'
      'Trim Char=False')
    VendorLib = 'oci.dll'
    Connected = True
    Left = 368
    Top = 88
  end
  object qryConsulta: TSQLQuery
    MaxBlobSize = -1
    Params = <>
    SQLConnection = connection1
    Left = 424
    Top = 88
  end
  object OpenDialog1: TOpenDialog
    Left = 352
    Top = 8
  end
end

VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmNFeG 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Notas Fiscais Eletr�nicas - NFe"
   ClientHeight    =   8595
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9615
   Icon            =   "frmNFeG.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   9615
   Begin MSComDlg.CommonDialog cdlgImprimirNota 
      Left            =   9720
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9720
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNFeG.frx":030A
            Key             =   "Alterar"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNFeG.frx":0466
            Key             =   "Anterior"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNFeG.frx":05C2
            Key             =   "Cancelar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNFeG.frx":071E
            Key             =   "Gravar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNFeG.frx":0882
            Key             =   "Imprimir"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNFeG.frx":0E1E
            Key             =   "Excluir"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNFeG.frx":13BA
            Key             =   "Localizar"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNFeG.frx":1516
            Key             =   "Novo"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNFeG.frx":1672
            Key             =   "Inicio"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNFeG.frx":17CE
            Key             =   "Proximo"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNFeG.frx":192A
            Key             =   "Fim"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNFeG.frx":1A86
            Key             =   "Outros"
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7890
      Left            =   60
      TabIndex        =   0
      Top             =   420
      Width           =   9510
      _ExtentX        =   16775
      _ExtentY        =   13917
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Nota Fiscal"
      TabPicture(0)   =   "frmNFeG.frx":1DA0
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblChaveAcesso"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblProtocolo"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblCancelada"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fraProdutos"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "fraCliente"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "fraCabecalho"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "fraDatas"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "fraEndereco"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtChaveAcesso"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtProtocolo"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "Informa��es Adicionais"
      TabPicture(1)   =   "frmNFeG.frx":1DBC
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtChaveAcessoDevolucao"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "fraNFeComplementoICMS"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "fraInformacoesCorpoNota"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "fraFatura"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "fraTransportadorVolumes"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "fraDadosAdicionais"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "lblChaveAcessoDevolucao"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).ControlCount=   7
      Begin VB.TextBox txtChaveAcessoDevolucao 
         Height          =   315
         Left            =   -72600
         TabIndex        =   114
         Top             =   6600
         Width           =   6780
      End
      Begin VB.Frame fraNFeComplementoICMS 
         Caption         =   "Complemento de ICMS - NFe"
         Height          =   1035
         Left            =   -74940
         TabIndex        =   109
         Top             =   5520
         Width           =   9375
         Begin VB.TextBox txtChaveAcessoNFeComplementar 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1500
            TabIndex        =   112
            Top             =   600
            Width           =   6780
         End
         Begin MSMask.MaskEdBox mskNumeroNFeComplementar 
            Height          =   285
            Left            =   1500
            TabIndex        =   110
            Top             =   300
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Format          =   "0"
            PromptChar      =   " "
         End
         Begin VB.Label lblChaveAcessoComplementar 
            AutoSize        =   -1  'True
            Caption         =   "Chave de Acesso"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   135
            TabIndex        =   113
            Top             =   660
            Width           =   1260
         End
         Begin VB.Label lblNumeroNFeComplementar 
            AutoSize        =   -1  'True
            Caption         =   "N�mero"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   840
            TabIndex        =   111
            Top             =   345
            Width           =   555
         End
      End
      Begin VB.TextBox txtProtocolo 
         Enabled         =   0   'False
         Height          =   315
         Left            =   6900
         TabIndex        =   104
         Top             =   660
         Width           =   2460
      End
      Begin VB.TextBox txtChaveAcesso 
         Enabled         =   0   'False
         Height          =   315
         Left            =   75
         TabIndex        =   103
         Top             =   660
         Width           =   6780
      End
      Begin VB.Frame fraEndereco 
         Caption         =   "Nome Fantasia/Endere�o"
         Height          =   1760
         Left            =   5580
         TabIndex        =   101
         Top             =   1050
         Width           =   3855
         Begin VB.Label lblDescricaoEndereco 
            ForeColor       =   &H80000008&
            Height          =   1440
            Left            =   120
            TabIndex        =   102
            Top             =   240
            Width           =   3615
         End
      End
      Begin VB.Frame fraInformacoesCorpoNota 
         Caption         =   "Informa��es no Corpo da Nota"
         Height          =   1275
         Left            =   -74940
         TabIndex        =   88
         Top             =   4200
         Width           =   9375
         Begin VB.TextBox txtInformacoesCorpo 
            Height          =   915
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   89
            Top             =   240
            Width           =   9135
         End
      End
      Begin VB.Frame fraFatura 
         Caption         =   "Faturamento"
         Height          =   2385
         Left            =   -70320
         TabIndex        =   90
         Top             =   1800
         Width           =   4755
         Begin VB.CommandButton cmdIncluirFatura 
            BackColor       =   &H80000004&
            Height          =   315
            Left            =   3960
            Picture         =   "frmNFeG.frx":1DD8
            Style           =   1  'Graphical
            TabIndex        =   97
            Top             =   405
            Width           =   360
         End
         Begin VB.CommandButton cmdExcluirFatura 
            BackColor       =   &H80000004&
            Height          =   315
            Left            =   4320
            Picture         =   "frmNFeG.frx":228E
            Style           =   1  'Graphical
            TabIndex        =   98
            Top             =   405
            Width           =   360
         End
         Begin MSMask.MaskEdBox mskNumeroBoleto 
            Height          =   285
            Left            =   60
            TabIndex        =   92
            Top             =   420
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   503
            _Version        =   393216
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskVencimentoBoleto 
            Height          =   285
            Left            =   1710
            TabIndex        =   94
            Top             =   420
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "99/99/9999"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskValorBoleto 
            Height          =   285
            Left            =   2700
            TabIndex        =   96
            Top             =   420
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   15
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "R$ #,##0.00"
            PromptChar      =   " "
         End
         Begin MSComctlLib.ListView lvwBoletos 
            Height          =   1590
            Left            =   60
            TabIndex        =   99
            Top             =   720
            Width           =   4620
            _ExtentX        =   8149
            _ExtentY        =   2805
            View            =   3
            LabelEdit       =   1
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483624
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
         Begin VB.Label lblValorBoleto 
            AutoSize        =   -1  'True
            Caption         =   "Valor"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   2700
            TabIndex        =   95
            Top             =   210
            Width           =   360
         End
         Begin VB.Label lblVencimentoBoleto 
            AutoSize        =   -1  'True
            Caption         =   "Vencimento"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   1710
            TabIndex        =   93
            Top             =   210
            Width           =   840
         End
         Begin VB.Label lblBoleto 
            AutoSize        =   -1  'True
            Caption         =   "Boleto"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   60
            TabIndex        =   91
            Top             =   210
            Width           =   450
         End
      End
      Begin VB.Frame fraTransportadorVolumes 
         Caption         =   "Transportador/Volumes"
         Height          =   1335
         Left            =   -74940
         TabIndex        =   65
         Top             =   420
         Width           =   9375
         Begin VB.ComboBox cmbUFPlaca 
            Height          =   315
            ItemData        =   "frmNFeG.frx":2744
            Left            =   7920
            List            =   "frmNFeG.frx":2746
            Style           =   2  'Dropdown List
            TabIndex        =   79
            Top             =   570
            Width           =   735
         End
         Begin VB.TextBox txtVolumeEspecie 
            Height          =   315
            Left            =   5400
            MaxLength       =   50
            TabIndex        =   85
            Top             =   900
            Width           =   1920
         End
         Begin VB.TextBox txtVolumeQuantidade 
            Height          =   315
            Left            =   1020
            MaxLength       =   50
            TabIndex        =   73
            Top             =   570
            Width           =   1215
         End
         Begin VB.TextBox txtVolumeNumero 
            Height          =   315
            Left            =   5400
            MaxLength       =   50
            TabIndex        =   77
            Top             =   570
            Width           =   1920
         End
         Begin VB.TextBox txtVolumeMarca 
            Height          =   315
            Left            =   2880
            MaxLength       =   50
            TabIndex        =   75
            Top             =   570
            Width           =   1860
         End
         Begin VB.ComboBox cmbFreteConta 
            Height          =   315
            ItemData        =   "frmNFeG.frx":2748
            Left            =   5400
            List            =   "frmNFeG.frx":2758
            Style           =   2  'Dropdown List
            TabIndex        =   69
            Top             =   240
            Width           =   1935
         End
         Begin VB.ComboBox cmbTransportador 
            Height          =   315
            ItemData        =   "frmNFeG.frx":2796
            Left            =   1020
            List            =   "frmNFeG.frx":2798
            Style           =   2  'Dropdown List
            TabIndex        =   67
            Top             =   240
            Width           =   3735
         End
         Begin MSMask.MaskEdBox mskPlaca 
            Height          =   285
            Left            =   7920
            TabIndex        =   71
            Top             =   255
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   8
            Mask            =   "CCC-9999"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskVolumePesoBruto 
            Height          =   315
            Left            =   1020
            TabIndex        =   81
            Top             =   900
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   11
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "##,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskVolumePesoLiquido 
            Height          =   315
            Left            =   3345
            TabIndex        =   83
            Top             =   900
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   11
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "##,##0.00"
            PromptChar      =   " "
         End
         Begin VB.Label lblUFCaminhao 
            AutoSize        =   -1  'True
            Caption         =   "UF"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   7620
            TabIndex        =   78
            Top             =   600
            Width           =   210
         End
         Begin VB.Label lblVolumeEspecie 
            AutoSize        =   -1  'True
            Caption         =   "Esp�cie"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   4785
            TabIndex        =   84
            Top             =   960
            Width           =   570
         End
         Begin VB.Label lblVolumePesoLiquido 
            AutoSize        =   -1  'True
            Caption         =   "Peso L�quido"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   2340
            TabIndex        =   82
            Top             =   960
            Width           =   945
         End
         Begin VB.Label lblVolumePesoBruto 
            AutoSize        =   -1  'True
            Caption         =   "Peso Bruto"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   165
            TabIndex        =   80
            Top             =   960
            Width           =   780
         End
         Begin VB.Label lblVolumeNumero 
            AutoSize        =   -1  'True
            Caption         =   "N�mero"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   4800
            TabIndex        =   76
            Top             =   630
            Width           =   555
         End
         Begin VB.Label lblVolumeMarca 
            AutoSize        =   -1  'True
            Caption         =   "Marca"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   2340
            TabIndex        =   74
            Top             =   630
            Width           =   450
         End
         Begin VB.Label lblVolumeQuantidade 
            AutoSize        =   -1  'True
            Caption         =   "Quantidade"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   72
            Top             =   630
            Width           =   825
         End
         Begin VB.Label lblTransportador 
            AutoSize        =   -1  'True
            Caption         =   "Nome"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   525
            TabIndex        =   66
            Top             =   300
            Width           =   420
         End
         Begin VB.Label lblPlaca 
            AutoSize        =   -1  'True
            Caption         =   "Placa"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   7440
            TabIndex        =   70
            Top             =   300
            Width           =   405
         End
         Begin VB.Label lblFreteConta 
            AutoSize        =   -1  'True
            Caption         =   "Frete"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   4995
            TabIndex        =   68
            Top             =   300
            Width           =   360
         End
      End
      Begin VB.Frame fraDadosAdicionais 
         Caption         =   "Dados Adicionais/Informa��es Complementares"
         Height          =   2385
         Left            =   -74940
         TabIndex        =   86
         Top             =   1800
         Width           =   4575
         Begin VB.TextBox txtDadosAdicionais 
            Height          =   2055
            Left            =   75
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   87
            Top             =   240
            Width           =   4410
         End
      End
      Begin VB.Frame fraDatas 
         Height          =   1560
         Left            =   7080
         TabIndex        =   25
         Top             =   2775
         Width           =   2355
         Begin MSMask.MaskEdBox mskDataEmissao 
            Height          =   285
            Left            =   1275
            TabIndex        =   29
            Top             =   195
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "99/99/9999"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskDataVencimento 
            Height          =   285
            Left            =   1275
            TabIndex        =   30
            Top             =   495
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "99/99/9999"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskHora 
            Height          =   285
            Left            =   1275
            TabIndex        =   31
            Top             =   795
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   8
            Mask            =   "99:99:99"
            PromptChar      =   " "
         End
         Begin VB.Label lblDataEmissao 
            AutoSize        =   -1  'True
            Caption         =   "Emiss�o"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   615
            TabIndex        =   26
            Top             =   240
            Width           =   585
         End
         Begin VB.Label lblDataVencimento 
            AutoSize        =   -1  'True
            Caption         =   "Sa�da"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   765
            TabIndex        =   27
            Top             =   540
            Width           =   435
         End
         Begin VB.Label lblHora 
            AutoSize        =   -1  'True
            Caption         =   "Hora"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   855
            TabIndex        =   28
            Top             =   840
            Width           =   345
         End
      End
      Begin VB.Frame fraCabecalho 
         Height          =   1560
         Left            =   60
         TabIndex        =   1
         Top             =   2775
         Width           =   6975
         Begin VB.TextBox txtObservacao 
            Height          =   315
            Left            =   810
            MaxLength       =   20
            TabIndex        =   24
            Top             =   1125
            Width           =   6060
         End
         Begin VB.TextBox txtDocumento 
            Height          =   315
            Left            =   4380
            MaxLength       =   20
            TabIndex        =   22
            Top             =   795
            Width           =   2490
         End
         Begin VB.ComboBox cmbFormaPagamento 
            Height          =   315
            ItemData        =   "frmNFeG.frx":279A
            Left            =   4380
            List            =   "frmNFeG.frx":279C
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   480
            Width           =   2505
         End
         Begin VB.ComboBox cmbNaturezaOperacao 
            Height          =   315
            ItemData        =   "frmNFeG.frx":279E
            Left            =   4380
            List            =   "frmNFeG.frx":27A0
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   165
            Width           =   2505
         End
         Begin VB.CommandButton cmdPesquisaCFOP 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1320
            Picture         =   "frmNFeG.frx":27A2
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   465
            Width           =   360
         End
         Begin MSMask.MaskEdBox mskNumero 
            Height          =   285
            Left            =   810
            TabIndex        =   7
            Top             =   180
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Format          =   "0"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskCFOP 
            Height          =   285
            Left            =   810
            TabIndex        =   13
            Top             =   495
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   5
            Mask            =   "9.999"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskCupom 
            Height          =   285
            Left            =   2550
            TabIndex        =   9
            Top             =   180
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Format          =   "0"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskDescontoGeral 
            Height          =   315
            Left            =   810
            TabIndex        =   18
            Top             =   795
            Width           =   645
            _ExtentX        =   1138
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskBonificacao 
            Height          =   315
            Left            =   2550
            TabIndex        =   20
            Top             =   795
            Visible         =   0   'False
            Width           =   645
            _ExtentX        =   1138
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#0.00"
            PromptChar      =   " "
         End
         Begin VB.Label lblBonificacao 
            AutoSize        =   -1  'True
            Caption         =   "Bonifica��o"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   1635
            TabIndex        =   19
            Top             =   855
            Visible         =   0   'False
            Width           =   840
         End
         Begin VB.Label lblDescontoGeral 
            AutoSize        =   -1  'True
            Caption         =   "Desconto"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   45
            TabIndex        =   17
            Top             =   840
            Width           =   690
         End
         Begin VB.Label lblObservacao 
            AutoSize        =   -1  'True
            Caption         =   "Obs."
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   405
            TabIndex        =   23
            Top             =   1185
            Width           =   330
         End
         Begin VB.Label lblDocumento 
            AutoSize        =   -1  'True
            Caption         =   "Documento"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   3465
            TabIndex        =   21
            Top             =   855
            Width           =   825
         End
         Begin VB.Label lblTipoPagamento 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Pg."
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   3690
            TabIndex        =   15
            Top             =   540
            Width           =   600
         End
         Begin VB.Label lblNumero 
            AutoSize        =   -1  'True
            Caption         =   "N�mero"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   75
            TabIndex        =   2
            Top             =   225
            Width           =   660
         End
         Begin VB.Label lblNaturezaOperacao 
            AutoSize        =   -1  'True
            Caption         =   "Natureza"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   3645
            TabIndex        =   10
            Top             =   225
            Width           =   645
         End
         Begin VB.Label lblCFOP 
            AutoSize        =   -1  'True
            Caption         =   "CFOP"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   315
            TabIndex        =   12
            Top             =   585
            Width           =   420
         End
         Begin VB.Label lblCupom 
            AutoSize        =   -1  'True
            Caption         =   "Cupom"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   1980
            TabIndex        =   8
            Top             =   225
            Width           =   495
         End
      End
      Begin VB.Frame fraCliente 
         Caption         =   "Cliente"
         Height          =   1725
         Left            =   60
         TabIndex        =   32
         Top             =   1050
         Width           =   5475
         Begin VB.TextBox txtCodigoCliente 
            Height          =   315
            Left            =   105
            MaxLength       =   20
            TabIndex        =   3
            Top             =   210
            Width           =   855
         End
         Begin VB.ComboBox cmbCliente 
            Height          =   960
            Left            =   105
            Style           =   1  'Simple Combo
            TabIndex        =   6
            Text            =   "cmbCliente"
            Top             =   540
            Width           =   5265
         End
         Begin VB.CommandButton cmdPesquisaCliente 
            Height          =   315
            Left            =   2700
            Picture         =   "frmNFeG.frx":28EC
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   210
            Width           =   360
         End
         Begin MSMask.MaskEdBox mskCNPJ_CPF 
            Height          =   315
            Left            =   990
            TabIndex        =   4
            Top             =   210
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   " "
         End
      End
      Begin VB.Frame fraProdutos 
         Height          =   3465
         Left            =   60
         TabIndex        =   33
         Top             =   4320
         Width           =   9375
         Begin VB.CommandButton cmdAnotacoes 
            BackColor       =   &H00E0E0E0&
            Height          =   315
            Left            =   8160
            Picture         =   "frmNFeG.frx":2A36
            Style           =   1  'Graphical
            TabIndex        =   45
            ToolTipText     =   "Notas"
            Top             =   405
            Width           =   360
         End
         Begin VB.CommandButton cmdExcluir 
            BackColor       =   &H80000004&
            Height          =   315
            Left            =   8925
            Picture         =   "frmNFeG.frx":2E08
            Style           =   1  'Graphical
            TabIndex        =   47
            Top             =   405
            Width           =   360
         End
         Begin VB.TextBox txtProduto 
            Height          =   315
            Left            =   1590
            MaxLength       =   50
            TabIndex        =   38
            Top             =   405
            Width           =   4005
         End
         Begin VB.CommandButton cmdProduto 
            BackColor       =   &H00E0E0E0&
            Height          =   315
            Left            =   1200
            Picture         =   "frmNFeG.frx":32BE
            Style           =   1  'Graphical
            TabIndex        =   36
            Top             =   420
            Width           =   360
         End
         Begin VB.CommandButton cmdIncluir 
            BackColor       =   &H00E0E0E0&
            Height          =   315
            Left            =   8535
            Picture         =   "frmNFeG.frx":3408
            Style           =   1  'Graphical
            TabIndex        =   46
            Top             =   405
            Width           =   360
         End
         Begin VB.TextBox txtCodigo 
            Height          =   315
            Left            =   75
            MaxLength       =   20
            TabIndex        =   35
            Top             =   405
            Width           =   1095
         End
         Begin MSMask.MaskEdBox mskQuantidade 
            Height          =   315
            Left            =   5610
            TabIndex        =   40
            Top             =   405
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   11
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "##,##0.000"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskDesconto 
            Height          =   315
            Left            =   6480
            TabIndex        =   42
            Top             =   405
            Width           =   645
            _ExtentX        =   1138
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskUnitario 
            Height          =   315
            Left            =   7140
            TabIndex        =   44
            Top             =   405
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            MaxLength       =   15
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSComctlLib.ListView lvwProdutos 
            Height          =   1560
            Left            =   60
            TabIndex        =   48
            Top             =   735
            Width           =   9225
            _ExtentX        =   16272
            _ExtentY        =   2752
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483624
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
         Begin MSMask.MaskEdBox mskBaseCalculoICMS 
            Height          =   285
            Left            =   1140
            TabIndex        =   50
            Top             =   2325
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   13
            Format          =   "###,##0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskBaseICMSSubstituicao 
            Height          =   285
            Left            =   1140
            TabIndex        =   58
            Top             =   2625
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   13
            Format          =   "###,##0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskValorTotalProdutos 
            Height          =   285
            Left            =   8055
            TabIndex        =   56
            Top             =   2325
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   13
            Format          =   "###,##0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskValorTotalNota 
            Height          =   285
            Left            =   8055
            TabIndex        =   64
            Top             =   2625
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   13
            Format          =   "###,##0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskOutrasDespesas 
            Height          =   285
            Left            =   5850
            TabIndex        =   62
            Top             =   2625
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   13
            Format          =   "###,##0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskValorICMS 
            Height          =   285
            Left            =   3510
            TabIndex        =   52
            Top             =   2325
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   13
            Format          =   "###,##0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskValorFrete 
            Height          =   285
            Left            =   5850
            TabIndex        =   54
            Top             =   2325
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   13
            Format          =   "###,##0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskValorICMSSubstituicao 
            Height          =   285
            Left            =   3510
            TabIndex        =   60
            Top             =   2625
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   13
            Format          =   "###,##0.00"
            PromptChar      =   "_"
         End
         Begin VB.Label lblTributos 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   120
            TabIndex        =   116
            Top             =   3000
            Width           =   9135
         End
         Begin VB.Label lblDespesas 
            AutoSize        =   -1  'True
            Caption         =   "Despesas"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   5040
            TabIndex        =   61
            Top             =   2700
            Width           =   705
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Vl. ICMS Sub."
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   2460
            TabIndex        =   59
            Top             =   2670
            Width           =   990
         End
         Begin VB.Label lblTotalNota 
            AutoSize        =   -1  'True
            Caption         =   "Total Nota"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   7245
            TabIndex        =   63
            Top             =   2670
            Width           =   750
         End
         Begin VB.Label lblValorTotalProdutos 
            AutoSize        =   -1  'True
            Caption         =   "Total Produtos"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   6960
            TabIndex        =   55
            Top             =   2370
            Width           =   1035
         End
         Begin VB.Label lblBaseCalculoICMS 
            AutoSize        =   -1  'True
            Caption         =   "Base do ICMS"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   60
            TabIndex        =   49
            Top             =   2370
            Width           =   1020
         End
         Begin VB.Label lblValorICMS 
            AutoSize        =   -1  'True
            Caption         =   "Valor do ICMS"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   2430
            TabIndex        =   51
            Top             =   2370
            Width           =   1020
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Base Subst."
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   225
            TabIndex        =   57
            Top             =   2670
            Width           =   855
         End
         Begin VB.Label lblValorFrete 
            AutoSize        =   -1  'True
            Caption         =   "Valor do Frete"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   4785
            TabIndex        =   53
            Top             =   2370
            Width           =   990
         End
         Begin VB.Label lblUnitario 
            AutoSize        =   -1  'True
            Caption         =   "Vl. Unit�rio"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   7320
            TabIndex        =   43
            Top             =   180
            Width           =   765
         End
         Begin VB.Label lblDesconto 
            AutoSize        =   -1  'True
            Caption         =   "Desc."
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   6705
            TabIndex        =   41
            Top             =   180
            Width           =   420
         End
         Begin VB.Label lblQuantidade 
            AutoSize        =   -1  'True
            Caption         =   "Quantidade"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   5610
            TabIndex        =   39
            Top             =   180
            Width           =   825
         End
         Begin VB.Label lblProduto 
            AutoSize        =   -1  'True
            Caption         =   "Produto"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   1620
            TabIndex        =   37
            Top             =   180
            Width           =   555
         End
         Begin VB.Label lblCodigo 
            AutoSize        =   -1  'True
            Caption         =   "C�digo"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   90
            TabIndex        =   34
            Top             =   180
            Width           =   495
         End
      End
      Begin VB.Label lblChaveAcessoDevolucao 
         AutoSize        =   -1  'True
         Caption         =   "Chave de Acesso - Devolu��o"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   -74940
         TabIndex        =   115
         Top             =   6660
         Width           =   2175
      End
      Begin VB.Label lblCancelada 
         AutoSize        =   -1  'True
         Caption         =   "CANCELADA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   8400
         TabIndex        =   107
         Top             =   60
         Width           =   1110
      End
      Begin VB.Label lblProtocolo 
         AutoSize        =   -1  'True
         Caption         =   "Protocolo de Autoriza��o"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   6900
         TabIndex        =   106
         Top             =   420
         Width           =   1785
      End
      Begin VB.Label lblChaveAcesso 
         AutoSize        =   -1  'True
         Caption         =   "Chave de Acesso"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   75
         TabIndex        =   105
         Top             =   405
         Width           =   1260
      End
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   100
      Top             =   0
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   15
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Novo"
            Object.ToolTipText     =   "Novo Cadastro"
            ImageKey        =   "Novo"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Alterar"
            Object.ToolTipText     =   "Altera Cadastro"
            ImageKey        =   "Alterar"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Excluir"
            Object.ToolTipText     =   "Exclui Cadastro"
            ImageKey        =   "Excluir"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Gravar"
            Object.ToolTipText     =   "Gravar"
            ImageKey        =   "Gravar"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cancelar"
            Object.ToolTipText     =   "Cancelar"
            ImageKey        =   "Cancelar"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Inicio"
            Object.ToolTipText     =   "Primeiro Registro"
            ImageKey        =   "Inicio"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Anterior"
            Object.ToolTipText     =   "Registro Anterior"
            ImageKey        =   "Anterior"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Proximo"
            Object.ToolTipText     =   "Pr�ximo Registro"
            ImageKey        =   "Proximo"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Fim"
            Object.ToolTipText     =   "�ltimo Registro"
            ImageKey        =   "Fim"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Localizar"
            Object.ToolTipText     =   "Localiza Registros"
            ImageKey        =   "Localizar"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Localiza"
                  Text            =   "Localiza Registro"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Visualiza"
                  Text            =   "Visualiza Campos"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Imprimir"
            Object.ToolTipText     =   "Imprimir"
            ImageKey        =   "Imprimir"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Nota"
                  Text            =   "Imprime Nota"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Boleto"
                  Text            =   "Imprime Boleto"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Outros"
            Object.ToolTipText     =   "Outras Op��es"
            ImageKey        =   "Outros"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "AtualizarNFe"
                  Text            =   "AtualizarNFe"
               EndProperty
            EndProperty
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Label lblAtualizarNFe 
      Caption         =   "AtualizarNFe"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   180
      TabIndex        =   108
      Top             =   8340
      Visible         =   0   'False
      Width           =   1395
   End
End
Attribute VB_Name = "frmNFeG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ItemList As ListItem
Dim PagamentoAtual As String
Dim NotaAtual As String
Dim RegistroAtual As Double 'Para posicionar o ponteiro depois de pesquisas
Dim rsNFe As New ADODB.Recordset
Dim rsNFeItens As New ADODB.Recordset
Dim rsNFeBoletos As New ADODB.Recordset
Dim rsProdutos As New ADODB.Recordset
Dim rsEmpresa As New ADODB.Recordset
Dim rsTransportador As New ADODB.Recordset
Dim rsCFOPs As New ADODB.Recordset
Dim rsClientes As New ADODB.Recordset
Dim rsTemp As New ADODB.Recordset
Dim rsTemp2 As New ADODB.Recordset
Dim rsNaturezasOperacao As New ADODB.Recordset
Dim rsContasBancarias As New ADODB.Recordset
Dim rsSaldoProdutos As New ADODB.Recordset
Dim rsUnidadesMedida As New ADODB.Recordset
Dim rsSituacoesTributarias As New ADODB.Recordset
Dim rsCFOPReferencias As New ADODB.Recordset
Dim strDescricaoTemp As String
Dim intItensNota As Integer
Dim rsLogradouros As New ADODB.Recordset
Dim rsMunicipios As New ADODB.Recordset
Dim rsUnidades As New ADODB.Recordset
Dim rsFormasPagamento As New ADODB.Recordset
Dim rsFretes As New ADODB.Recordset

Private Sub Form_Load()
   I_TituloForm = Me.Caption
   On Error GoTo Erro
   SSTab1.Tab = 0 ' Posiciona no primeiro tab
   Status = 0
   RegistroAtual = 0
   Centraliza frmNFeG
   
   Set rsNFe = cnSistema.Execute("Select * from NFe Order By Numero")
   Set rsContasBancarias = cnSistema.Execute("Select * from ContasBancarias")
   Set rsEmpresa = cnSistema.Execute("Select * from Empresa")
          
'''   rsNFe.Open "Select * from NFe Order By Numero", cnSistema, adOpenForwardOnly, adLockOptimistic, 1
'''   rsContasBancarias.Open "Select * from ContasBancarias", cnSistema, adOpenForwardOnly, adLockOptimistic, 1
'''   rsEmpresa.Open "Select * from Empresa", cnSistema, adOpenForwardOnly, adLockOptimistic, 1

   lvwProdutos.ColumnHeaders.Add , , "C�digo", 850
   lvwProdutos.ColumnHeaders.Add , , "Produto", 3200
   lvwProdutos.ColumnHeaders.Add , , "Quantidade", 1000, lvwColumnRight
   lvwProdutos.ColumnHeaders.Add , , "Vl. Unit�rio", 1050, lvwColumnRight
   lvwProdutos.ColumnHeaders.Add , , "Desc.", 700, lvwColumnRight
   lvwProdutos.ColumnHeaders.Add , , "Valor Total", 1050, lvwColumnRight
   lvwProdutos.ColumnHeaders.Add , , "Vl. L�quido", 1050, lvwColumnRight

   lvwBoletos.ColumnHeaders.Add , , "Boleto", 1500
   lvwBoletos.ColumnHeaders.Add , , "Vencimento", 1300
   lvwBoletos.ColumnHeaders.Add , , "Valor", 1500, lvwColumnRight

   If LerArquivoINI("SEPD", "TipoImpressao", CaminhoINI & "\System.ini") = 0 Then
      intItensNota = Val(LerArquivoINI("Notas Fiscais", "ItensNota", CaminhoINI & "\Notas.ini"))
   Else
      intItensNota = Val(LerArquivoINI("Notas Fiscais", "ItensNota", CaminhoINI & "\NotasManuais.ini"))
   End If

   Carrega_Combos
   If Registros2("NFe") = 0 Then
      Botoes 3, frmNFeG
   Else
      Botoes 1, frmNFeG
      rsNFe.MoveLast
      Prencher_Campos
   End If

   If I_Acesso = 3 Then ' Controle N�veis de Acesso
      Toolbar.Buttons(2).Visible = False
      Toolbar.Buttons(3).Visible = False
   End If

   Campos False
'''   If Registros(cnSistema, "NFe") = 0 Then
   If Registros2("NFe") = 0 Then
      Toolbar.Buttons(15).Enabled = False
   End If
'   MDISistema.StatusBar.Panels(1).text = "Cadastro Notas de Sa�da Eletr�nicas"

Exit Sub
Erro:
   If Err.Number = -2147467259 Then
      rsErro = True
      Beep
      MsgBox "Erro na Abertura do Arquivo de Dados" & Chr(13) & "Algum usu�rio est� com o Arquivo em modo Exclusivo", vbExclamation, "Erro"
      Exit Sub
   Else
      rsErro = True
      Beep
      MsgBox "Verificar: " & Err.Number & Chr(13) & Err.Description, vbExclamation, "Sistema"
      Exit Sub
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'   If Not rsErro Then rsNFe.Close
'   If Not rsErro Then rsContasBancarias.Close
'   If Not rsErro Then rsEmpresa.Close
   
   Set rsNFe = Nothing
   Set rsContasBancarias = Nothing
   Set rsEmpresa = Nothing
End Sub

Private Sub mskCNPJ_CPF_LostFocus()
Dim Verifica_CPF As String
Dim intCliente, Contador As Integer

   If mskCNPJ_CPF.text <> Empty Then
      Verifica_CPF = CNPJ_CPF(mskCNPJ_CPF.text)
      If Verifica_CPF <> "ERRO" Then
          mskCNPJ_CPF.text = CNPJ_CPF(mskCNPJ_CPF.text)
          Set rsTemp = cnSistema.Execute("Select idCliente From Clientes Where CNPJ_CPF = '" & mskCNPJ_CPF.text & "'")
          If Not rsTemp.BOF And Not rsTemp.EOF Then
             intCliente = rsTemp!idCliente

             For Contador = 0 To (cmbCliente.ListCount - 1)
                 If cmbCliente.ItemData(Contador) = intCliente Then
                    cmbCliente.ListIndex = Contador
                    Exit For
                 End If
             Next
          End If
          If cmbCliente.Enabled Then cmbCliente.SetFocus
      Else
          mskCNPJ_CPF.SelStart = 0
          mskCNPJ_CPF.SelLength = Len(mskCNPJ_CPF.text)
          mskCNPJ_CPF.SetFocus
      End If
   End If
End Sub

Private Sub mskQuantidade_LostFocus()
Dim Contador As Integer

'''   Set rsTemp = cnSistema.Execute("Select * From TabelaProdutos Where Codigo = '" & txtCodigo.Text & "'")
'''   If Not rsTemp.EOF Then
'''      If rsTemp!Saldo < Val(mskQuantidade.Text) Then
'''         Beep
'''         Set rsTemp2 = cnSistema.Execute("Select * From SaldosConferencias Where idProduto = " & rsTemp!idProduto)
'''         Dim sTruncados As String
'''         If Not rsTemp2.EOF Then
'''            sTruncados = Chr(13) & "Mais: " & rsTemp2!Quantidade & " Truncados"
'''         Else
'''            sTruncados = ""
'''         End If
'''
'''         MsgBox "Saldo Menor que Quantidade." & Chr(13) & "Atual: " & rsTemp!Saldo & sTruncados, vbExclamation, "Erro"
'''         mskDesconto.SetFocus
'''         Exit Sub
'''      End If
'''   End If

   If frmNFeG.txtCodigo.text <> "" Then
    ' Carrega Combos
      Set rsTemp = cnSistema.Execute("Select * from UnidadesMedida Order By Descricao")
      frmNFeGComplemento.cmbUnidade.Clear
      Do While Not rsTemp.EOF
         frmNFeGComplemento.cmbUnidade.AddItem rsTemp!Descricao
         frmNFeGComplemento.cmbUnidade.ItemData(frmNFeGComplemento.cmbUnidade.NewIndex) = rsTemp!idUnidadeMedida
         rsTemp.MoveNext
      Loop

    ' Preencher Campos
      Set rsTemp = cnSistema.Execute("Select * From Produtos Where Codigo = '" & SQLCheck(frmNFeG.txtCodigo.text) & "'")
      If Not rsTemp.EOF Then
         ' ICMS
         frmNFeGComplemento.mskICMSProduto.text = rsTemp!ICMS
         ' Unidade
         For Contador = 0 To (frmNFeGComplemento.cmbUnidade.ListCount - 1)
            If frmNFeGComplemento.cmbUnidade.ItemData(Contador) = rsTemp!idUnidade Then
               frmNFeGComplemento.cmbUnidade.ListIndex = Contador
               Exit For
            End If
         Next
      End If
   End If

End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then Sendkeys "{TAB}"
End Sub

Private Sub txtCodigoCliente_LostFocus()
Dim intCliente, Contador As Integer

   If txtCodigoCliente.text <> Empty Then
      Set rsTemp = cnSistema.Execute("Select * From Clientes Where Codigo = " & txtCodigoCliente.text)
      If Not rsTemp.BOF And Not rsTemp.EOF Then
         mskCNPJ_CPF.text = rsTemp!CNPJ_CPF
         intCliente = rsTemp!idCliente

         For Contador = 0 To (cmbCliente.ListCount - 1)
             If cmbCliente.ItemData(Contador) = intCliente Then
                cmbCliente.ListIndex = Contador
                Exit For
             End If
         Next
         If cmbCliente.Enabled Then cmbCliente.SetFocus
      Else
          MsgBox "C�digo n�o encontrado", vbOKOnly, "Visualiza"
          txtCodigoCliente.text = Empty
          txtCodigoCliente.SetFocus
      End If
   End If
End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
   Select Case Button.Key
      Case "Novo"
         Status = 1
         Limpa_Campos
         Botoes 2, frmNFeG
         RegistroAtual = IIf(rsNFe.EOF, 0, rsNFe!idNFe)
         Campos True
         txtCodigoCliente.SetFocus

      Case "Cancelar"
         Status = 0
         If rsNFe.EOF Then
            Botoes 3, frmNFeG
         Else
            Botoes 1, frmNFeG
         End If
         Campos False
'''         If Registros(cnSistema, "NFe") <> 0 Then
         If Registros2("NFe") <> 0 Then
            If RegistroAtual <> 0 Then
               rsNFe.MoveFirst
               rsNFe.Find "idNFe = " & RegistroAtual
            End If
            Prencher_Campos
         Else
            Limpa_Campos
         End If

      Case "Gravar"
         Call Gravar

      Case "Alterar"
         RegistroAtual = rsNFe!idNFe
         NotaAtual = mskNumero.text
         Status = 2
         Call Alterar

      Case "Excluir"
         Status = 3
         Call Excluir
         If rsNFe.EOF Then
            Botoes 3, frmNFeG
         Else
            Botoes 1, frmNFeG
         End If

      Case "Localizar"
         RegistroAtual = rsNFe!idNFe
         Status = 4
         Call Localizar

      Case "Inicio"
         rsNFe.MoveFirst
         Prencher_Campos

      Case "Fim"
         rsNFe.MoveLast
         Prencher_Campos

      Case "Proximo"
         rsNFe.MoveNext
         If rsNFe.EOF Then rsNFe.MoveLast
         Prencher_Campos

      Case "Anterior"
         rsNFe.MovePrevious
         If rsNFe.BOF Then rsNFe.MoveFirst
         Prencher_Campos

      Case "Imprimir"
         Call ImprimirNota
   End Select
End Sub

Private Sub Toolbar_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Dim Contador As Integer
Dim Total_Registros As Integer
Dim PaginaInicial, Paginafinal, NumeroCopias, i

   RegistroAtual = rsNFe!idNFe
   Select Case ButtonMenu.Key
      Case "Localiza"
'         Status = 4
'         Call Localizar

         Status = 4
         Registro_Selecionado = False
         Screen.MousePointer = vbDefault
         frmNFeGPesquisa.Show vbModal
         rsNFe.MoveFirst
         If Registro_Selecionado Then
            rsNFe.Find "idNFe = " & Val(Mid(frmNFeGPesquisa.lvwDados.SelectedItem.Key, 2, Len(frmNFeGPesquisa.lvwDados.SelectedItem.Key)))
         Else
            rsNFe.Find "idNFe = " & RegistroAtual
         End If
         Prencher_Campos

      Case "Visualiza"
'''         Total_Registros = Registros(cnSistema, "NFe")
         Total_Registros = Registros2("NFe")
         If Total_Registros = 0 Then
            MsgBox "N�o existe nenhum Registro na Tabela", vbOKOnly, "Visualiza"
            Exit Sub
         End If
         Contador = 1
         Status = 4
         Screen.MousePointer = vbHourglass
         frmVisualiza.lvwDados.ColumnHeaders.Clear
         frmVisualiza.lvwDados.ColumnHeaders.Add , , "Vencimento", 1100
         frmVisualiza.lvwDados.ColumnHeaders.Add , , "N�mero", 1100
         frmVisualiza.lvwDados.ColumnHeaders.Add , , "Cliente", 3800
         rsNFe.MoveFirst
         frmVisualiza.lvwDados.ListItems.Clear
         Do While Not rsNFe.EOF
            frmNFeG.Caption = "Processando " & StrZero(Contador, 8) & " de " & StrZero(Total_Registros, 8)
            Set rsTemp = cnSistema.Execute("Select * From Clientes Where idCliente = " & rsNFe!idCliente)
            Set ItemList = frmVisualiza.lvwDados.ListItems.Add(, "R" & CStr(rsNFe!idNFe), rsNFe!DataEmissao)
                ItemList.SubItems(1) = rsNFe!Numero
                ItemList.SubItems(2) = rsTemp!Nome
            rsNFe.MoveNext
            Contador = Contador + 1
         Loop
         rsNFe.MoveFirst
         Registro_Selecionado = False
         Screen.MousePointer = vbDefault
         Me.Caption = I_TituloForm
         frmVisualiza.Show vbModal
         rsNFe.MoveFirst
         If Registro_Selecionado Then
            rsNFe.Find "idNFe = " & Val(Mid(frmVisualiza.lvwDados.SelectedItem.Key, 2, Len(frmVisualiza.lvwDados.SelectedItem.Key)))
         Else
            rsNFe.Find "idNFe = " & RegistroAtual
         End If
         Prencher_Campos

      Case "Nota"
         Call ImprimirNota

      Case "Boleto"
         Call ImprimirBoleto

      Case "Copiar"
         Call CopiarNota

      Case "AtualizarNFe"
         lblAtualizarNFe.Tag = rsNFe!idNFe
         frmNFeGAtualizar.Show
         lblAtualizarNFe.Tag = ""

   End Select
End Sub

Private Sub mskNumero_LostFocus()
   If Status = 1 Then
      Set rsTemp = cnSistema.Execute("SELECT * FROM NFe WHERE Numero = " & mskNumero.text)
      If Not rsTemp.EOF Then
         RegistroAtual = rsNFe!idNFe
         Status = 2
         rsNFe.MoveFirst
         rsNFe.Find "idNFe = " & rsTemp!idNFe
         Prencher_Campos
      End If
   End If

   If Status = 4 Then
      RegistroAtual = IIf(rsNFe.EOF, 0, rsNFe!idNFe)
      If mskNumero.text = Empty Then
         MsgBox "Digite um N�mero para a Consulta", vbOKOnly, "Localizar"
         If mskNumero.Enabled Then mskNumero.SetFocus
         Exit Sub
      End If
      rsNFe.MoveFirst
      rsNFe.Find "Numero Like " & Trim(mskNumero.text)
      If Not rsNFe.EOF Then
         Botoes 1, frmNFeG
         Prencher_Campos
         Campos False
      Else
         MsgBox "N�mero n�o Encontrado", vbOKOnly + vbExclamation, "Localizar"
         mskNumero.SetFocus
         mskNumero.SelStart = 0
         mskNumero.SelLength = Len(mskNumero.text)
         rsNFe.MoveFirst
         If RegistroAtual <> 0 Then rsNFe.Find "idNFe = " & RegistroAtual
      End If
   End If
End Sub

Sub Excluir()
On Error GoTo ErroIntegridade
   If MsgBox("Confirma Excluir o registro atual? ", vbYesNo + vbInformation, "Excluir") = vbYes Then
      Atividade "Exclus�o: " & mskNumero.text, Me.Caption
      cnSistema.Execute "Delete * from NFeItens Where idNFe = " & rsNFe!idNFe  ' Itens da Nota de Entrada
      cnSistema.Execute "Delete from NFe Where idNFe=" & rsNFe!idNFe           ' Nota de Entrada
      rsNFe.Requery
'''      If Registros(cnSistema, "NFe") = 0 Then
      If Registros2("NFe") = 0 Then
         Limpa_Campos
      Else
         Prencher_Campos
      End If
   End If

On Error GoTo 0
Exit Sub
ErroIntegridade:
   If Err.Number = 0 Then
      ' Opera��o Ok
   ElseIf Err.Number = -2147467259 Then
      Beep
      MsgBox "N�o � poss�vel Excluir este Registro" & Chr(13) & "Existe lan�amentos relacionados com este Registro", vbInformation + vbOKOnly, "Excluir"
      Exit Sub
   Else
      Beep
      MsgBox "Verificar: " & Err.Number & Chr(13) & Err.Description, vbExclamation, "Excluir"
      Exit Sub
   End If
End Sub

Sub Alterar()
   Status = 2
   Botoes 2, frmNFeG
   Campos True
   txtCodigoCliente.SetFocus
End Sub

Sub Localizar()
   Campos True
   Botoes 4, frmNFeG
   Limpa_Campos
   mskNumero.SetFocus
End Sub

Private Function Verifica_Campos()
Dim strMensagem As String
Verifica_Campos = True

   If mskNumero.text = Empty Then strMensagem = strMensagem & "N�mero" & Chr(13)
   If Not IsDate(mskDataEmissao.text) Or Val(Mid(mskDataEmissao.text, 7, 4)) < 1900 Then strMensagem = strMensagem & "Emiss�o" & Chr(13)
   If Not IsDate(mskDataVencimento.text) Or Val(Mid(mskDataVencimento.text, 7, 4)) < 1900 Then strMensagem = strMensagem & "Vencimento" & Chr(13)
   If cmbNaturezaOperacao.ListIndex = -1 Then strMensagem = strMensagem & "Natureza da Opera��o" & Chr(13)
   If cmbFormaPagamento.ListIndex = -1 Then strMensagem = strMensagem & "Forma de Pagamento" & Chr(13)
   If cmbCliente.ListIndex = -1 Then strMensagem = strMensagem & "Cliente" & Chr(13)
   If Len(Trim(mskPlaca.text)) > 1 Then
      If Len(cmbUFPlaca.text) <> 2 Then strMensagem = strMensagem & "UF da Placa" & Chr(13)
   End If

   mskValorTotalNota.text = 0
   mskValorTotalNota.text = Format(Val(Substitui(mskValorTotalNota.text, ",", ".")) + Val(Substitui(mskValorTotalProdutos.text, ",", ".")), "###,##0.00")
   mskValorTotalNota.text = Format(Val(Substitui(mskValorTotalNota.text, ",", ".")) + Val(Substitui(mskValorFrete.text, ",", ".")), "###,##0.00")
   mskValorTotalNota.text = Format(Val(Substitui(mskValorTotalNota.text, ",", ".")) + Val(Substitui(mskOutrasDespesas.text, ",", ".")), "###,##0.00")

   Set rsCFOPs = cnSistema.Execute("Select * From CFOPs Where CFOP = '" & mskCFOP.text & "'")
   If rsCFOPs.BOF Or rsCFOPs.EOF Then strMensagem = strMensagem & "CFOP n�o Encontrado" & Chr(13)

   Set rsTemp = cnSistema.Execute("Select * From NFeInutilizadas Where Numero = " & mskNumero.text)
   If Not rsTemp.EOF Then strMensagem = strMensagem & "N�mero de Nota Fiscal Inutilizada" & Chr(13)

   ' Verifica Numeros
   If Status = 1 Then
      Dim intUltimoNumero As Double
      If Not rsNFe.BOF Or Not rsNFe.EOF Then
         rsNFe.MoveLast
         intUltimoNumero = rsNFe!Numero
      End If

      If ((Val(mskNumero.text) - intUltimoNumero) > 1) Or ((Val(mskNumero.text) - intUltimoNumero) < 1) Then
         strMensagem = strMensagem & "N�mero da nota fora da sequ�ncia" & Chr(13)
      End If
   ElseIf Status = 2 Then
      If mskNumero.text <> NotaAtual Then
         strMensagem = strMensagem & "N�mero da nota n�o pode ser alterado" & Chr(13)
      End If
   End If


'''   If cmbNaturezaOperacao.ListIndex >= 0 And cmbCliente.ListIndex >= 0 Then
'''      Set rsClientes = cnSistema.Execute("Select * From Clientes Where idCliente = " & cmbCliente.ItemData(cmbCliente.ListIndex))
'''      Set rsTemp = cnSistema.Execute("Select * From NaturezasOperacao Where idNaturezaOperacao = " & cmbNaturezaOperacao.ItemData(cmbNaturezaOperacao.ListIndex))
'''      If rsClientes!UF = rsEmpresa!UF Then
'''         mskCFOP.Text = rsTemp!CFOPDentroUF
'''      Else
'''         mskCFOP.Text = rsTemp!CFOPForaUF
'''      End If
'''   End If

   If Not strMensagem = Empty Then
      MsgBox "Verifique os Seguintes Campos:" & Chr(13) & strMensagem, vbExclamation + vbOKOnly, "Campos Obrigat�rios"
      Verifica_Campos = False
      Exit Function
   End If

End Function

Sub Campos(Parametro As Boolean)
   mskNumero.Enabled = Parametro
   mskCupom.Enabled = Parametro
   cmbNaturezaOperacao.Enabled = Parametro
   mskCFOP.Enabled = Parametro
   mskDataEmissao.Enabled = Parametro
   mskDataVencimento.Enabled = Parametro
   mskHora.Enabled = Parametro
   txtCodigoCliente.Enabled = Parametro
   mskCNPJ_CPF.Enabled = Parametro
   cmdPesquisaCliente.Enabled = Parametro
   cmbCliente.Enabled = Parametro
   mskBaseCalculoICMS.Enabled = Parametro
   mskValorICMS.Enabled = Parametro
   mskValorFrete.Enabled = Parametro
   mskValorTotalProdutos.Enabled = Parametro
   mskBaseICMSSubstituicao.Enabled = Parametro
   mskValorICMSSubstituicao.Enabled = Parametro
   mskOutrasDespesas.Enabled = Parametro
   mskValorTotalNota.Enabled = Parametro
   txtDadosAdicionais.Enabled = Parametro
   cmbFormaPagamento.Enabled = Parametro
   mskDescontoGeral.Enabled = Parametro
   mskBonificacao.Enabled = Parametro
   txtDocumento.Enabled = Parametro
   txtObservacao.Enabled = Parametro

   cmbTransportador.Enabled = Parametro
   cmbFreteConta.Enabled = Parametro
   mskPlaca.Enabled = Parametro
   cmbUFPlaca.Enabled = Parametro
   txtVolumeQuantidade.Enabled = Parametro
   txtVolumeMarca.Enabled = Parametro
   txtVolumeEspecie.Enabled = Parametro
   txtVolumeNumero.Enabled = Parametro
   mskVolumePesoBruto.Enabled = Parametro
   mskVolumePesoLiquido.Enabled = Parametro
   txtInformacoesCorpo.Enabled = Parametro

   txtCodigo.Enabled = Not Parametro
   cmdProduto.Enabled = Not Parametro
   txtProduto.Enabled = Not Parametro
   mskQuantidade.Enabled = Not Parametro
   mskDesconto.Enabled = Not Parametro
   mskUnitario.Enabled = Not Parametro
   cmdIncluir.Enabled = Not Parametro
   cmdExcluir.Enabled = Not Parametro
   lvwProdutos.Enabled = Not Parametro

   mskNumeroBoleto.Enabled = Not Parametro
   mskVencimentoBoleto.Enabled = Not Parametro
   mskValorBoleto.Enabled = Not Parametro
   cmdIncluirFatura.Enabled = Not Parametro
   cmdExcluirFatura.Enabled = Not Parametro
   lvwBoletos.Enabled = Not Parametro

   mskNumeroNFeComplementar.Enabled = Parametro
   txtChaveAcessoNFeComplementar.Enabled = Parametro
   txtChaveAcessoDevolucao.Enabled = Parametro

   Toolbar.Buttons(15).Enabled = Not Parametro
End Sub

Sub Limpa_Campos()

   Set rsTemp = cnSistema.Execute("SELECT NFeInutilizadas.Numero FROM NFeInutilizadas ORDER BY NFeInutilizadas.Numero DESC")

   If Status <> 4 Then
      If Not rsNFe.BOF Or Not rsNFe.EOF Then
         rsNFe.MoveLast
         mskNumero.text = rsNFe!Numero + 1

         If Not rsTemp.EOF Then
            If mskNumero.text = rsTemp!Numero Then
               mskNumero.text = Val(mskNumero.text) + 1
            End If
         End If
      Else
         mskNumero.text = 1
      End If
   Else
      mskNumero.text = Empty
   End If

   txtChaveAcesso.text = Empty
   txtProtocolo.text = Empty
   mskCupom.text = Empty
   cmbNaturezaOperacao.ListIndex = -1
   mskCFOP.text = " .   "
   mskDataEmissao.text = Date
   mskDataVencimento.text = Date
   mskHora.text = Time
   txtCodigoCliente.text = Empty
   mskCNPJ_CPF.text = Empty
   cmbCliente.ListIndex = -1
   mskBaseCalculoICMS.text = Empty
   mskValorICMS.text = Empty
   mskValorFrete.text = Empty
   mskValorTotalProdutos.text = Empty
   mskBaseICMSSubstituicao.text = Empty
   mskValorICMSSubstituicao.text = Empty
   mskOutrasDespesas.text = Empty
   mskValorTotalNota.text = Empty
   txtDadosAdicionais.text = LerArquivoINI("Notas Fiscais", "DadosAdicionais", CaminhoINI & "\System.ini")
   txtInformacoesCorpo.text = LerArquivoINI("Notas Fiscais", "InformacoesCorpo", CaminhoINI & "\System.ini")
   cmbFormaPagamento.ListIndex = -1
   mskDescontoGeral.text = Empty
   mskBonificacao.text = Empty
   txtDocumento.text = Empty
   txtObservacao.text = Empty

   lblCancelada.Visible = False

   cmbTransportador.ListIndex = -1
   cmbFreteConta.ListIndex = -1
   mskPlaca.text = "   -    "
   cmbUFPlaca.ListIndex = -1
   txtVolumeQuantidade.text = Empty
   txtVolumeMarca.text = Empty
   txtVolumeEspecie.text = Empty
   txtVolumeNumero.text = Empty
   mskVolumePesoBruto.text = Empty
   mskVolumePesoLiquido.text = Empty

   lvwProdutos.ListItems.Clear
   lvwBoletos.ListItems.Clear

   lblDescricaoEndereco.Caption = Empty

   mskNumeroNFeComplementar.text = Empty
   txtChaveAcessoNFeComplementar.text = Empty
   txtChaveAcessoDevolucao.text = Empty

End Sub

Sub Prencher_Campos()
Dim Contador As Integer
Dim intCliente As Integer
Dim curValorPagar As Currency

   Set rsClientes = cnSistema.Execute("Select * From Clientes Where idCliente = " & rsNFe!idCliente)

   txtChaveAcesso.text = IIf(Trim(rsNFe!ChaveNFe) = "" Or IsNull(rsNFe!ChaveNFe), Empty, rsNFe!ChaveNFe)
   txtProtocolo.text = IIf(Trim(rsNFe!Protocolo) = "" Or IsNull(rsNFe!Protocolo), Empty, rsNFe!Protocolo)

   mskNumero.text = IIf(Trim(rsNFe!Numero) = "" Or IsNull(rsNFe!Numero), Empty, rsNFe!Numero)
   mskCupom.text = IIf(Trim(rsNFe!Cupom) = "" Or IsNull(rsNFe!Cupom), Empty, rsNFe!Cupom)
   Set rsCFOPs = cnSistema.Execute("Select * From CFOPs Where idCFOP = " & rsNFe!idCFOP)
   If Not rsCFOPs.EOF Then
      mskCFOP.text = rsCFOPs!CFOP
   End If
   mskDataEmissao.text = IIf(IsNull(rsNFe!DataEmissao), "  /  /    ", rsNFe!DataEmissao)
   mskDataVencimento.text = IIf(IsNull(rsNFe!DataVencimento), "  /  /    ", rsNFe!DataVencimento)
   mskHora.text = IIf(Trim(rsNFe!Hora) = "" Or IsNull(rsNFe!Hora), "  :  :  ", rsNFe!Hora)
   txtCodigoCliente.text = rsClientes!Codigo
   mskCNPJ_CPF.text = IIf(Not IsNull(rsClientes!CNPJ_CPF), rsClientes!CNPJ_CPF, "")
   mskBaseCalculoICMS.text = IIf(Trim(rsNFe!BaseCalculoICMS) = "" Or IsNull(rsNFe!BaseCalculoICMS), Empty, rsNFe!BaseCalculoICMS)
   mskValorICMS.text = IIf(Trim(rsNFe!ValorICMS) = "" Or IsNull(rsNFe!ValorICMS), Empty, rsNFe!ValorICMS)
   mskValorFrete.text = IIf(Trim(rsNFe!ValorFrete) = "" Or IsNull(rsNFe!ValorFrete), Empty, rsNFe!ValorFrete)
   mskBaseICMSSubstituicao.text = IIf(Trim(rsNFe!BaseICMSSubstituicao) = "" Or IsNull(rsNFe!BaseICMSSubstituicao), Empty, rsNFe!BaseICMSSubstituicao)
   mskValorICMSSubstituicao.text = IIf(Trim(rsNFe!ValorICMSSubstituicao) = "" Or IsNull(rsNFe!ValorICMSSubstituicao), Empty, rsNFe!ValorICMSSubstituicao)
   mskOutrasDespesas.text = IIf(Trim(rsNFe!OutrasDespesas) = "" Or IsNull(rsNFe!OutrasDespesas), Empty, rsNFe!OutrasDespesas)
   txtDadosAdicionais.text = IIf(Trim(rsNFe!DadosAdicionais) = "" Or IsNull(rsNFe!DadosAdicionais), Empty, rsNFe!DadosAdicionais)
   mskDescontoGeral.text = IIf(Trim(rsNFe!DescontoGeral) = "" Or IsNull(rsNFe!DescontoGeral), Empty, rsNFe!DescontoGeral)
   mskBonificacao.text = IIf(Trim(rsNFe!Bonificacao) = "" Or IsNull(rsNFe!Bonificacao), Empty, rsNFe!Bonificacao)
   txtDocumento.text = IIf(Trim(rsNFe!Documento) = "" Or IsNull(rsNFe!Documento), Empty, rsNFe!Documento)
   txtObservacao.text = IIf(Trim(rsNFe!Observacao) = "" Or IsNull(rsNFe!Observacao), Empty, rsNFe!Observacao)

   mskPlaca.text = IIf(Trim(rsNFe!PlacaVeiculo) = "" Or IsNull(rsNFe!PlacaVeiculo), "   -    ", rsNFe!PlacaVeiculo)
'   txtUFCaminhao.Text = IIf(Trim(rsNFe!UFCaminhao) = "" Or IsNull(rsNFe!UFCaminhao), "", rsNFe!UFCaminhao)
   txtVolumeQuantidade.text = IIf(Trim(rsNFe!VolumeQuantidade) = "" Or IsNull(rsNFe!VolumeQuantidade), Empty, rsNFe!VolumeQuantidade)
   txtVolumeMarca.text = IIf(Trim(rsNFe!VolumeMarca) = "" Or IsNull(rsNFe!VolumeMarca), Empty, rsNFe!VolumeMarca)
   txtVolumeEspecie.text = IIf(Trim(rsNFe!VolumeEspecie) = "" Or IsNull(rsNFe!VolumeEspecie), Empty, rsNFe!VolumeEspecie)
   txtVolumeNumero.text = IIf(Trim(rsNFe!VolumeNumero) = "" Or IsNull(rsNFe!VolumeNumero), Empty, rsNFe!VolumeNumero)
   mskVolumePesoBruto.text = IIf(Trim(rsNFe!VolumePesoBruto) = "" Or IsNull(rsNFe!VolumePesoBruto), Empty, rsNFe!VolumePesoBruto)
   mskVolumePesoLiquido.text = IIf(Trim(rsNFe!VolumePesoLiquido) = "" Or IsNull(rsNFe!VolumePesoLiquido), Empty, rsNFe!VolumePesoLiquido)
   txtInformacoesCorpo.text = IIf(Trim(rsNFe!InformacoesCorpo) = "" Or IsNull(rsNFe!InformacoesCorpo), Empty, rsNFe!InformacoesCorpo)
   If rsNFe!UFCaminhao <> "" Then cmbUFPlaca.text = rsNFe!UFCaminhao
   If rsNFe!Situacao = 3 Then
      lblCancelada.Visible = True
   Else
      lblCancelada.Visible = False
   End If

   mskNumeroNFeComplementar.text = IIf(Trim(rsNFe!NumeroNFeComplementar) = "" Or IsNull(rsNFe!NumeroNFeComplementar), Empty, rsNFe!NumeroNFeComplementar)
   txtChaveAcessoNFeComplementar.text = IIf(Trim(rsNFe!ChaveAcessoNFeComplementar) = "" Or IsNull(rsNFe!ChaveAcessoNFeComplementar), Empty, rsNFe!ChaveAcessoNFeComplementar)
   txtChaveAcessoDevolucao.text = IIf(Trim(rsNFe!ChaveAcessoDevolucao) = "" Or IsNull(rsNFe!ChaveAcessoDevolucao), Empty, rsNFe!ChaveAcessoDevolucao)

   For Contador = 0 To (cmbNaturezaOperacao.ListCount - 1)
      If cmbNaturezaOperacao.ItemData(Contador) = rsNFe!idNaturezaOperacao Then
         cmbNaturezaOperacao.ListIndex = Contador
         Exit For
      End If
   Next

   For Contador = 0 To (cmbTransportador.ListCount - 1)
      If cmbTransportador.ItemData(Contador) = rsNFe!idTransportador Then
         cmbTransportador.ListIndex = Contador
         Exit For
      End If
   Next

   For Contador = 0 To (cmbFormaPagamento.ListCount - 1)
      If cmbFormaPagamento.ItemData(Contador) = rsNFe!idFormaPagamento Then
         cmbFormaPagamento.ListIndex = Contador
         Exit For
      End If
   Next

   For Contador = 0 To (cmbFreteConta.ListCount - 1)
      If cmbFreteConta.ItemData(Contador) = rsNFe!FreteConta Then
         cmbFreteConta.ListIndex = Contador
         Exit For
      End If
   Next

'   cmbFormaPagamento.ListIndex = rsNFe!idFormaPagamento

   For Contador = 0 To (cmbCliente.ListCount - 1)
       If cmbCliente.ItemData(Contador) = rsNFe!idCliente Then
          cmbCliente.ListIndex = Contador
          Exit For
       End If
   Next

   If cmbCliente.ListIndex <> -1 Then
      Set rsTemp = cnSistema.Execute("Select * From Clientes Where idCliente = " & cmbCliente.ItemData(cmbCliente.ListIndex))
      lblDescricaoEndereco.Caption = Trim(rsTemp!NomeFantasia) & Chr(13) & _
                                     Trim(rsTemp!Endereco) & ", " & Trim(rsTemp!Bairro) & Chr(13) & _
                                     Trim(rsTemp!Cidade) & " - " & rsTemp!UF & Chr(13) & _
                                     "CEP: " & rsTemp!CEP & _
                                     " Fone: " & rsTemp!Telefone1 & Chr(13) & _
                                     "CNPJ/CPF: " & rsTemp!CNPJ_CPF & Chr(13) & _
                                     "IE: " & rsTemp!IE_CI
   End If

'  Produtos

   Dim cValorBruto As Currency
   Dim cValorDesconto As Currency
   Dim cValorBonificacao As Currency
   Dim cValorLiquido As Currency

   Set rsTemp = cnSistema.Execute("SELECT * FROM NFeItens WHERE NFeItens.idNFe = " & rsNFe!idNFe & " ORDER BY idNFeItem")

   lvwProdutos.ListItems.Clear
   Do While Not rsTemp.EOF
      Set rsProdutos = cnSistema.Execute("SELECT * FROM Produtos WHERE Produtos.idProduto = " & rsTemp!idProduto)
      If Not rsProdutos.EOF Then
         cValorBruto = (rsTemp!Quantidade * rsTemp!ValorUnitario)
         cValorDesconto = (((rsTemp!Quantidade * rsTemp!ValorUnitario) * rsTemp!Desconto) / 100)
         cValorBonificacao = (((cValorBruto - cValorDesconto) * rsNFe!Bonificacao) / 100)
         cValorLiquido = (cValorBruto - cValorDesconto - cValorBonificacao)

'         Set ItemList = lvwProdutos.ListItems.Add(, "R" & CStr(rsTemp!idProduto), rsProdutos!Codigo)
         Set ItemList = lvwProdutos.ListItems.Add(, "R" & CStr(rsTemp!idNFeItem), rsProdutos!Codigo)
         ItemList.SubItems(1) = Trim(rsProdutos!Descricao) & " " & Trim(rsTemp!DescricaoComplementar)
         ItemList.SubItems(2) = Format(rsTemp!Quantidade, mskQuantidade.Format)
         ItemList.SubItems(3) = Format(rsTemp!ValorUnitario, mskUnitario.Format)
         ItemList.SubItems(4) = Format(rsTemp!Desconto, "##,##0.00")
         ItemList.SubItems(5) = Format((rsTemp!Quantidade * rsTemp!ValorUnitario), "##,##0.00")
'         ItemList.SubItems(6) = Format((rsTemp!Quantidade * rsTemp!ValorUnitario) * (1 - ((rsTemp!Desconto) / 100)), "##,##0.00")
         ItemList.SubItems(6) = Format(cValorLiquido, "##,##0.00")
      End If

      rsTemp.MoveNext
   Loop

'  Boletos
   Set rsTemp = cnSistema.Execute("SELECT * FROM NFeBoletos WHERE NFeBoletos.idNFe = " & rsNFe!idNFe)

   lvwBoletos.ListItems.Clear
   Do While Not rsTemp.EOF
      Set ItemList = lvwBoletos.ListItems.Add(, "R" & rsTemp!Numero, rsTemp!Numero)
      ItemList.SubItems(1) = rsTemp!Vencimento
      ItemList.SubItems(2) = Format(rsTemp!Valor, "##,##0.00")

      rsTemp.MoveNext
   Loop

'  Total da Nota
   Set rsTemp = cnSistema.Execute("Select * From TotalNFe Where Numero = " & mskNumero.text)
   If Not rsTemp.EOF Then
      mskValorTotalProdutos.text = Format(rsTemp!Total, "###,##0.00")
      mskValorFrete.text = Format(rsTemp!TotalFrete, "###,##0.00")
      mskValorTotalNota.text = Format(rsTemp!Total + IIf(Not IsNull(rsTemp!TotalFrete), rsTemp!TotalFrete, 0) + rsNFe!OutrasDespesas, "###,##0.00")
      mskBaseCalculoICMS.text = Format(IIf(rsTemp!ValorICMS > 0, rsTemp!BaseCalculo, 0), "###,##0.00")
      mskValorICMS.text = Format(rsTemp!ValorICMS, "###,##0.00")
'      mskVolumePesoBruto.Text = Format(rsNFe!VolumePesoBruto, "###,##0.00")
'      mskVolumePesoLiquido.Text = Format(rsNFe!VolumePesoLiquido, "###,##0.00")

''''      mskVolumePesoBruto.Text = Format(rsTemp!PesoBruto, "###,##0.00")
''''      mskVolumePesoLiquido.Text = Format(rsTemp!PesoLiquido, "###,##0.00")
   Else
      mskValorTotalProdutos.text = Format(0, "###,##0.00")
      mskValorFrete.text = Format(0, "###,##0.00")
      mskValorTotalNota.text = Format(0, "###,##0.00")
      mskBaseCalculoICMS.text = Format(0, "###,##0.00")
      mskValorICMS.text = Format(0, "###,##0.00")
'      mskVolumePesoBruto.Text = Format(0, "###,##0.00")
'      mskVolumePesoLiquido.Text = Format(0, "###,##0.00")
   End If

End Sub

Sub Gravar()
Dim CNPJ_CPF As String
Dim dHora As String
Dim strCFOP As Integer
Dim iTransportador As Integer, iFreteConta As Integer
Dim bGeradaNFe As Boolean

   If Not Verifica_Campos() Then Exit Sub

   If cmbTransportador.ListIndex = -1 Then
      iTransportador = 1
   Else
      iTransportador = cmbTransportador.ItemData(cmbTransportador.ListIndex)
   End If

   If cmbFreteConta.ListIndex = -1 Then
      iFreteConta = 0
   Else
'      iFreteConta = IIf(cmbFreteConta.ListIndex = 0, 1, 2)
      iFreteConta = cmbFreteConta.ItemData(cmbFreteConta.ListIndex)
   End If

   If mskHora.text = "  :  :  " Then
      dHora = "00:00:00"
   Else
      dHora = mskHora.text
   End If

   Set rsCFOPs = cnSistema.Execute("Select * From CFOPs Where CFOP = '" & mskCFOP.text & "'")
   If Not rsCFOPs.EOF Then strCFOP = rsCFOPs!idCFOP
   Select Case Status
      Case 1 'Inclus�o
         If MsgBox("Confirma Incluir o registro atual", vbYesNo + vbQuestion, "Inclus�o") = vbYes Then
            cnSistema.Execute "Insert Into NFe (Numero,Cupom,idCliente,idNaturezaOperacao,idCFOP,DadosAdicionais,DataEmissao,DataCaixa,DataVencimento,Hora,BaseCalculoICMS,ValorICMS,ValorFrete,ValorTotalProdutos,BaseICMSSubstituicao,ValorICMSSubstituicao,OutrasDespesas,ValorTotalNota,idTransportador,FreteConta,PlacaVeiculo,UFCaminhao,VolumeQuantidade,VolumeMarca,VolumeEspecie,VolumeNumero,VolumePesoBruto,VolumePesoLiquido,InformacoesCorpo,idFormaPagamento,DescontoGeral,Bonificacao,Documento,Observacao,GeradaNFe,Situacao,NumeroNFeComplementar,ChaveAcessoNFeComplementar,ChaveAcessoDevolucao) " & _
                              "Values (" & Val(mskNumero.text) & "," & Val(mskCupom.text) & "," & cmbCliente.ItemData(cmbCliente.ListIndex) & "," & cmbNaturezaOperacao.ItemData(cmbNaturezaOperacao.ListIndex) & "," & strCFOP & ",'" & txtDadosAdicionais.text & "','" & mskDataEmissao.text & "','" & mskDataEmissao.text & "','" & mskDataVencimento.text & "','" & dHora & "','" & Val(Substitui(mskBaseCalculoICMS.ClipText, ",", ".")) & "','" & Val(Substitui(mskValorICMS.ClipText, ",", ".")) & "'," & _
                                      "'" & Val(Substitui(mskValorFrete.ClipText, ",", ".")) & "','" & Val(Substitui(mskValorTotalProdutos.ClipText, ",", ".")) & "','" & Val(Substitui(mskBaseICMSSubstituicao.ClipText, ",", ".")) & "','" & Val(Substitui(mskValorICMSSubstituicao.ClipText, ",", ".")) & "','" & Val(Substitui(mskOutrasDespesas.ClipText, ",", ".")) & "'," & _
                                      "'" & Val(Substitui(mskValorTotalNota.ClipText, ",", ".")) & "'," & iTransportador & "," & iFreteConta & ",'" & UCase(mskPlaca.text) & "','" & cmbUFPlaca.text & "','" & txtVolumeQuantidade.text & "','" & txtVolumeMarca.text & "','" & txtVolumeEspecie.text & "','" & txtVolumeNumero.text & "','" & Val(Substitui(mskVolumePesoBruto.ClipText, ",", ".")) & "','" & Val(Substitui(mskVolumePesoLiquido.ClipText, ",", ".")) & "','" & txtInformacoesCorpo.text & "'" & _
                                      "," & cmbFormaPagamento.ItemData(cmbFormaPagamento.ListIndex) & ",'" & Val(Substitui(mskDescontoGeral.ClipText, ",", ".")) & "','" & Val(Substitui(mskBonificacao.ClipText, ",", ".")) & "','" & txtDocumento.text & "','" & SQLCheck(txtObservacao.text) & "'," & bGeradaNFe & ",0," & CStrValor(mskNumeroNFeComplementar.text) & ",'" & txtChaveAcessoNFeComplementar.text & "','" & txtChaveAcessoDevolucao.text & "')"
            Atividade "Inclus�o: " & mskNumero.text, Me.Caption
            rsNFe.Requery
            rsNFe.Find "Numero = '" & mskNumero.text & "'"
         End If

      Case 2 'Alterac�o
         If MsgBox("Confirma Alterar o registro atual", vbYesNo + vbQuestion, "Altera��o") = vbYes Then

          ' Altera Desconto nos Produtos
            AlterarDescontos

            cnSistema.Execute "Update NFe set " & _
                  "Numero = " & Val(mskNumero.text) & ", " & "Cupom = " & Val(mskCupom.text) & ", " & _
                  "idCliente = " & cmbCliente.ItemData(cmbCliente.ListIndex) & ", " & "idNaturezaOperacao = " & cmbNaturezaOperacao.ItemData(cmbNaturezaOperacao.ListIndex) & ", " & _
                  "idCFOP = " & strCFOP & ", " & _
                  "DataEmissao = '" & mskDataEmissao.text & "', " & "DataCaixa = '" & mskDataEmissao.text & "', " & "DataVencimento = '" & mskDataVencimento.text & "', " & "Hora = '" & dHora & "', " & "DadosAdicionais = '" & SQLCheck(txtDadosAdicionais.text) & "', " & _
                  "BaseCalculoICMS = '" & Val(Substitui(mskBaseCalculoICMS.ClipText, ",", ".")) & "', " & "ValorICMS = '" & Val(Substitui(mskValorICMS.ClipText, ",", ".")) & "', " & _
                  "ValorFrete = '" & Val(Substitui(mskValorFrete.ClipText, ",", ".")) & "', " & "ValorTotalProdutos = '" & Val(Substitui(mskValorTotalProdutos.ClipText, ",", ".")) & "', " & _
                  "BaseICMSSubstituicao = '" & Val(Substitui(mskBaseICMSSubstituicao.ClipText, ",", ".")) & "', " & _
                  "ValorICMSSubstituicao = '" & Val(Substitui(mskValorICMSSubstituicao.ClipText, ",", ".")) & "', " & _
                  "OutrasDespesas = '" & Val(Substitui(mskOutrasDespesas.ClipText, ",", ".")) & "', " & _
                  "ValorTotalNota = '" & Val(Substitui(mskValorTotalNota.ClipText, ",", ".")) & "', " & _
                  "idTransportador = " & iTransportador & ", " & "FreteConta = " & iFreteConta & ", " & _
                  "PlacaVeiculo = '" & UCase(mskPlaca.text) & "', " & "UFCaminhao = '" & cmbUFPlaca.text & "', " & _
                  "VolumeQuantidade = '" & txtVolumeQuantidade.text & "', " & "VolumeMarca = '" & txtVolumeMarca.text & "', " & "VolumeNumero = '" & txtVolumeNumero.text & "', " & "VolumeEspecie = '" & txtVolumeEspecie.text & "', " & _
                  "VolumePesoBruto = '" & Val(Substitui(mskVolumePesoBruto.ClipText, ",", ".")) & "', " & _
                  "VolumePesoLiquido = '" & Val(Substitui(mskVolumePesoLiquido.ClipText, ",", ".")) & "', " & _
                  "InformacoesCorpo = '" & txtInformacoesCorpo.text & "', " & _
                  "idFormaPagamento = " & cmbFormaPagamento.ItemData(cmbFormaPagamento.ListIndex) & ", " & _
                  "DescontoGeral = '" & Val(Substitui(mskDescontoGeral.ClipText, ",", ".")) & "', " & _
                  "Bonificacao = '" & Val(Substitui(mskBonificacao.ClipText, ",", ".")) & "', " & _
                  "Documento = '" & txtDocumento.text & "', " & _
                  "Observacao = '" & SQLCheck(txtObservacao.text) & "', GeradaNFe = " & bGeradaNFe & ", " & _
                  "NumeroNFeComplementar = " & CStrValor(mskNumeroNFeComplementar.text) & ", ChaveAcessoNFeComplementar = '" & txtChaveAcessoNFeComplementar.text & "', ChaveAcessoDevolucao = '" & txtChaveAcessoDevolucao.text & "' " & _
                  "Where idNFe = " & rsNFe!idNFe
            Atividade "Alterar: " & mskNumero.text, Me.Caption
            rsNFe.Requery
            rsNFe.Find "Numero = '" & mskNumero.text & "'"
         End If
   End Select
   Prencher_Campos
   Botoes 1, frmNFeG
   Campos False
   Status = 0
   SSTab1.Tab = 0 ' Posiciona no primeiro tab
   txtCodigo.SetFocus
End Sub

Private Sub AlterarDescontos()

   If rsNFe!DescontoGeral <> Val(Substitui(mskDescontoGeral.text, ",", ".")) Then
      If MsgBox("Alterar Desconto dos Produtos", vbYesNo + vbQuestion, "Altera��o") = vbYes Then
         Set rsTemp = cnSistema.Execute("SELECT * FROM NFeItens WHERE NFeItens.idNFe = " & rsNFe!idNFe)

         lvwProdutos.ListItems.Clear
         Dim nDesconto As Double, nValorTotalNota As Double, nValorTotalProdutos As Double
         nValorTotalNota = 0
         nValorTotalProdutos = 0
         Do While Not rsTemp.EOF
            Set rsProdutos = cnSistema.Execute("SELECT * FROM Produtos WHERE Produtos.idProduto = " & rsTemp!idProduto)
            If Not rsProdutos!Marca Then
               If rsTemp!Desconto <> rsNFe!DescontoGeral Then
                  nDesconto = rsTemp!Desconto
               Else
                  nDesconto = Val(Substitui(mskDescontoGeral.text, ",", "."))
               End If
            Else
               nDesconto = 0
            End If

          ' Altera Desconto de Cada Produto
            cnSistema.Execute "Update NFeItens set " & _
                  "Desconto = '" & nDesconto & "' " & _
                  "Where idNFe = " & rsNFe!idNFe & " And idProduto = " & rsTemp!idProduto

            Set ItemList = lvwProdutos.ListItems.Add(, "R" & CStr(rsTemp!idProduto), rsProdutos!Codigo)
            ItemList.SubItems(1) = rsProdutos!Descricao
            ItemList.SubItems(2) = Format(rsTemp!Quantidade, mskQuantidade.Format)
            ItemList.SubItems(3) = Format(rsTemp!ValorUnitario, mskUnitario.Format)
            ItemList.SubItems(4) = Format(rsTemp!Desconto, "##,##0.00")
            ItemList.SubItems(5) = Format((rsTemp!Quantidade * rsTemp!ValorUnitario), "##,##0.00")
            ItemList.SubItems(6) = Format((rsTemp!Quantidade * rsTemp!ValorUnitario) * (1 - (rsTemp!Desconto / 100)), "##,##0.00")

            nValorTotalProdutos = nValorTotalProdutos + ((rsTemp!Quantidade * rsTemp!ValorUnitario) * (1 - (nDesconto / 100)))
            nValorTotalNota = nValorTotalNota + ((rsTemp!Quantidade * rsTemp!ValorUnitario) * (1 - (nDesconto / 100)))

            mskValorTotalProdutos.text = nValorTotalProdutos
            mskValorTotalNota = nValorTotalNota
            rsTemp.MoveNext
         Loop

       ' Altera Valor Total da Nota
         cnSistema.Execute "Update NFe set " & _
               "ValorTotalProdutos = '" & nValorTotalProdutos & "', " & _
               "ValorTotalNota = '" & nValorTotalNota & "' " & _
               "Where idNFe = " & rsNFe!idNFe

'         rsNFe.Requery
      End If
   End If
End Sub

Private Sub Carrega_Combos()

'  Clientes

'   Set rsTemp = cnSistema.Execute("Select * from Clientes Where Situacao = 0 Order By Nome")
   Set rsTemp = cnSistema.Execute("Select * from Clientes Order By Nome")
   cmbCliente.Clear
   Do While Not rsTemp.EOF
      cmbCliente.AddItem rsTemp!Nome
      cmbCliente.ItemData(cmbCliente.NewIndex) = rsTemp!idCliente
      rsTemp.MoveNext
   Loop

'  Naturezas de Opera��o

   Set rsTemp = cnSistema.Execute("Select * from NaturezasOperacao ORDER BY Descricao")
   cmbNaturezaOperacao.Clear
   Do While Not rsTemp.EOF
      cmbNaturezaOperacao.AddItem rsTemp!Descricao
      cmbNaturezaOperacao.ItemData(cmbNaturezaOperacao.NewIndex) = rsTemp!idNaturezaOperacao
      rsTemp.MoveNext
   Loop

'  Transportadores

   Set rsTemp = cnSistema.Execute("Select * from Transportadores ORDER BY Nome")
   cmbTransportador.Clear
   Do While Not rsTemp.EOF
      cmbTransportador.AddItem rsTemp!Nome
      cmbTransportador.ItemData(cmbTransportador.NewIndex) = rsTemp!idTransportador
      rsTemp.MoveNext
   Loop

'  Formas de Pagamento

   Set rsTemp = cnSistema.Execute("Select * from FormasPagamento ORDER BY Descricao")
   cmbFormaPagamento.Clear
   Do While Not rsTemp.EOF
      cmbFormaPagamento.AddItem rsTemp!Descricao
      cmbFormaPagamento.ItemData(cmbFormaPagamento.NewIndex) = rsTemp!idFormaPagamento
      rsTemp.MoveNext
   Loop

'  UFs
   Set rsTemp = cnSistema.Execute("Select * from UFs ORDER BY Sigla")
   cmbUFPlaca.Clear
   Do While Not rsTemp.EOF
      cmbUFPlaca.AddItem rsTemp!Sigla
      cmbUFPlaca.ItemData(cmbUFPlaca.NewIndex) = rsTemp!idUF
      rsTemp.MoveNext
   Loop

'  Frete Conta
   Set rsTemp = cnSistema.Execute("Select * from FreteConta ORDER BY Descricao")
   cmbFreteConta.Clear
   Do While Not rsTemp.EOF
      cmbFreteConta.AddItem rsTemp!Descricao
      cmbFreteConta.ItemData(cmbFreteConta.NewIndex) = rsTemp!idFreteConta
      rsTemp.MoveNext
   Loop

   rsTemp.Close
End Sub

Private Sub cmdExcluir_Click()
   Beep
   If MsgBox("Deseja excluir este item", vbYesNo + vbQuestion, "Excluir") = vbYes Then
      cnSistema.Execute "Update NFe set " & _
            "ValorTotalProdutos = '" & mskValorTotalProdutos.text & "', " & _
            "ValorTotalNota = '" & mskValorTotalNota.text & "' " & _
            "Where idNFe = " & rsNFe!idNFe

''      cnSistema.Execute "Delete from NFeItens Where idNFe = " & rsNFe!idNFe & " And idProduto = " & Val(Mid(lvwProdutos.SelectedItem.Key, 2, Len(lvwProdutos.SelectedItem.Key)))
      cnSistema.Execute "Delete from NFeItens Where idNFe = " & rsNFe!idNFe & " And idNFeItem = " & Val(Mid(lvwProdutos.SelectedItem.Key, 2, Len(lvwProdutos.SelectedItem.Key)))
      lvwProdutos.ListItems.Remove (lvwProdutos.SelectedItem.Index)

      Set rsTemp = cnSistema.Execute("Select * From TotalNFe Where Numero = " & mskNumero.text)
      If Not rsTemp.EOF Then
         mskValorTotalProdutos.text = Format(rsTemp!Total, "###,##0.00")
         mskValorTotalNota.text = Format(rsTemp!Total, "###,##0.00")
         mskBaseCalculoICMS.text = Format(rsTemp!BaseCalculo, "###,##0.00")
         mskValorICMS.text = Format(rsTemp!ValorICMS, "###,##0.00")
         mskVolumePesoBruto.text = Format(rsTemp!PesoBruto, "###,##0.00")
         mskVolumePesoLiquido.text = Format(rsTemp!PesoLiquido, "###,##0.00")
      Else
         mskValorTotalProdutos.text = Format(0, "###,##0.00")
         mskValorTotalNota.text = Format(0, "###,##0.00")
         mskBaseCalculoICMS.text = Format(0, "###,##0.00")
         mskValorICMS.text = Format(0, "###,##0.00")
         mskVolumePesoBruto.text = Format(0, "###,##0.00")
         mskVolumePesoLiquido.text = Format(0, "###,##0.00")
      End If
   End If
End Sub

Private Sub cmdIncluir_Click()
   If lvwProdutos.ListItems.Count = intItensNota Then
      Beep
      MsgBox "Total de Itens da Nota Excedido", vbExclamation, "Erro"

      txtCodigo.text = Empty
      txtProduto.text = Empty
      mskQuantidade.text = Empty
      mskDesconto.text = Empty
      mskUnitario.text = Empty
      txtCodigo.SetFocus

      Exit Sub
   End If

   If Verifica_Campos_Produtos() Then
      Set rsTemp = cnSistema.Execute("Select * From Produtos Where Codigo = '" & txtCodigo.text & "'")
      If Not rsTemp!Marca Then
         If mskDesconto.text <> "" Then
            mskDesconto.text = Val(Substitui(mskDesconto.text, ",", "."))
         Else
            mskDesconto.text = Val(Substitui(mskDescontoGeral.text, ",", "."))
         End If
      Else
         mskDesconto.text = 0
      End If

'      Set ItemList = lvwProdutos.ListItems.Add(, "R" & CStr(rsTemp!idProduto), txtCodigo.text)
'      ItemList.SubItems(1) = Trim(txtProduto.text) & " " & Trim(frmNFeGComplemento.txtDescricaoComplementar.text)
'      ItemList.SubItems(2) = Format(mskQuantidade.text, mskQuantidade.Format)
'      ItemList.SubItems(3) = Format(mskUnitario.text, mskUnitario.Format)
'      ItemList.SubItems(4) = Format(mskDesconto.text, "##0.00")
'      ItemList.SubItems(5) = Format((mskQuantidade.text * mskUnitario.text), "##,##0.00")
'      ItemList.SubItems(6) = Format((mskQuantidade.text * mskUnitario.text) * (1 - ((Val(Substitui(mskDesconto.text, ",", ".")) + Val(Substitui(mskBonificacao.text, ",", "."))) / 100)), "##0.00")

      mskQuantidade.text = Val(Substitui(mskQuantidade.text, ",", "."))
      mskDesconto.text = Val(Substitui(mskDesconto.text, ",", "."))
      mskUnitario.text = Val(Substitui(mskUnitario.text, ",", "."))
      frmNFeGComplemento.mskBaseReduzidaICMS.text = Val(Substitui(frmNFeGComplemento.mskBaseReduzidaICMS.text, ",", "."))

      Dim iUnidade As Integer, iSituacaoTributaria As Integer

      If frmNFeGComplemento.cmbUnidade.ListIndex = -1 Then
         iUnidade = 1
      Else
         iUnidade = frmNFeGComplemento.cmbUnidade.ItemData(frmNFeGComplemento.cmbUnidade.ListIndex)
      End If

      If frmNFeGComplemento.cmbSituacaoTributaria.ListIndex = -1 Then
         iSituacaoTributaria = rsTemp!idSituacaoTributaria
      Else
         iSituacaoTributaria = frmNFeGComplemento.cmbSituacaoTributaria.ItemData(frmNFeGComplemento.cmbSituacaoTributaria.ListIndex)
      End If

      If frmNFeGComplemento.mskCFOP.text = " .   " Then frmNFeGComplemento.mskCFOP.text = mskCFOP.text

      If Val(Substitui(frmNFeGComplemento.mskBaseReduzidaICMS.text, ",", ".")) = 0 Then
         If Not rsEmpresa.EOF Then
            If rsClientes!UF = rsEmpresa!UF Then
               frmNFeGComplemento.mskBaseReduzidaICMS.text = rsTemp!BaseReduzidaICMSdUF
            Else
               frmNFeGComplemento.mskBaseReduzidaICMS.text = rsTemp!BaseReduzidaICMSfUF
            End If
         End If
      End If

      cnSistema.Execute "Insert Into NFeItens (idNFe,idProduto,Data,Quantidade,Desconto,ValorUnitario,ICMS,BaseReduzida,DescricaoComplementar,idUnidade,idSituacaoTributaria,DiscriminacaoProduto,IPI,BaseReduzidaIPI,ClassificacaoFiscal,ValorFrete,CFOP) " & _
                        "Values (" & rsNFe!idNFe & "," & rsTemp!idProduto & ",'" & mskDataEmissao.text & _
                        "','" & CStrValor(mskQuantidade.text) & "','" & CStrValor(mskDesconto.text) & "','" & CStrValor(mskUnitario.text) & "','" & _
                        CStrValor(frmNFeGComplemento.mskICMSProduto.text) & "','" & CStrValor(frmNFeGComplemento.mskBaseReduzidaICMS.text) & "','" & SQLCheck(frmNFeGComplemento.txtDescricaoComplementar.text) & "'," & iUnidade & "," & iSituacaoTributaria & ",'" & _
                        SQLCheck(frmNFeGComplemento.txtDiscriminacaoProduto.text) & "','" & CStrValor(frmNFeGComplemento.mskIPIProduto.text) & "','" & CStrValor(frmNFeGComplemento.mskBaseReduzidaIPI.text) & "','" & SQLCheck(frmNFeGComplemento.txtClassificacaoFiscal.text) & "','" & CStrValor(frmNFeGComplemento.mskValorFrete.text) & "','" & frmNFeGComplemento.mskCFOP.text & "')"

      cnSistema.Execute "Update NFe set " & _
            "ValorTotalProdutos = '" & mskValorTotalProdutos.text & "', " & _
            "ValorTotalNota = '" & mskValorTotalNota.text & "' " & _
            "Where idNFe = " & rsNFe!idNFe

      Set rsTemp = cnSistema.Execute("Select * From TotalNFe Where Numero = " & mskNumero.text)
      If Not rsTemp.EOF Then
         mskValorTotalProdutos.text = Format(rsTemp!Total, "###,##0.00")
         mskValorFrete.text = Format(rsTemp!TotalFrete, "###,##0.00")
         mskValorTotalNota.text = Format(rsTemp!Total + rsTemp!TotalFrete, "###,##0.00")
         mskBaseCalculoICMS.text = Format(rsTemp!BaseCalculo, "###,##0.00")
         mskValorICMS.text = Format(rsTemp!ValorICMS, "###,##0.00")
         mskVolumePesoBruto.text = Format(rsTemp!PesoTotal, "###,##0.00")
         mskVolumePesoLiquido.text = Format(rsTemp!PesoTotal, "###,##0.00")
      Else
         mskValorTotalProdutos.text = Format(0, "###,##0.00")
         mskValorFrete.text = Format(0, "###,##0.00")
         mskValorTotalNota.text = Format(0, "###,##0.00")
         mskBaseCalculoICMS.text = Format(0, "###,##0.00")
         mskValorICMS.text = Format(0, "###,##0.00")
         mskVolumePesoBruto.text = Format(0, "###,##0.00")
         mskVolumePesoLiquido.text = Format(0, "###,##0.00")
      End If

      ' Atualiza Corpo da Nota
      Set rsCFOPs = cnSistema.Execute("Select * From CFOPs Where CFOP = '" & mskCFOP.text & "'")
      Set rsProdutos = cnSistema.Execute("Select * From Produtos Where Codigo = '" & txtCodigo.text & "'")

      If Not rsCFOPs.EOF Then
         Set rsCFOPReferencias = cnSistema.Execute("Select * From CFOPReferencias Where idCFOP = " & rsCFOPs!idCFOP & " AND idProduto = " & rsProdutos!idProduto)
         If Not rsCFOPReferencias.EOF Then
            cnSistema.Execute "Update NFe set " & _
                  "InformacoesCorpo = '" & rsCFOPReferencias!InformacoesCorpo & "' " & _
                  "Where idNFe = " & rsNFe!idNFe
         End If

'         Prencher_Campos
      End If

      Prencher_Campos

      txtCodigo.text = Empty
      txtProduto.text = Empty
      mskQuantidade.text = Empty
      mskDesconto.text = Empty
      mskUnitario.text = Empty

      frmNFeGComplemento.mskICMSProduto.text = Empty
      frmNFeGComplemento.mskBaseReduzidaICMS.text = Empty
      frmNFeGComplemento.txtDescricaoComplementar.text = Empty
      frmNFeGComplemento.cmbUnidade.ListIndex = -1
      frmNFeGComplemento.cmbSituacaoTributaria.ListIndex = -1
      frmNFeGComplemento.txtDiscriminacaoProduto.text = Empty
      frmNFeGComplemento.mskIPIProduto.text = Empty
      frmNFeGComplemento.mskBaseReduzidaIPI.text = Empty
      frmNFeGComplemento.mskValorFrete.text = Empty
      frmNFeGComplemento.txtClassificacaoFiscal.text = Empty

      txtCodigo.SetFocus
   End If
End Sub

Private Function Verifica_Campos_Produtos()
Dim strMensagem As String
Dim ProcuraItem As ListItem
Verifica_Campos_Produtos = True

   If txtCodigo.text = Empty Then strMensagem = strMensagem & "C�digo" & Chr(13)
   If txtProduto.text = Empty Then strMensagem = strMensagem & "Produto" & Chr(13)
   If Val(Substitui(mskQuantidade.text, ",", ".")) = 0 Then strMensagem = strMensagem & "Quantidade" & Chr(13)

   Set rsProdutos = cnSistema.Execute("SELECT * FROM Produtos WHERE Codigo = '" & txtCodigo.text & "'")
   If Not rsProdutos.EOF Then
'      Set rsTemp = cnSistema.Execute("SELECT * FROM SaldoProdutos WHERE idProduto = " & rsProdutos!idProduto)
'      If rsTemp!Saldo < Val(mskQuantidade.Text) Then strMensagem = strMensagem & "Saldo de Produtos inferior a quantidade. Saldo Atual " & Round(rsTemp!Saldo, 2) & Chr(13)
   End If

'   Set rsTemp = cnSistema.Execute("Select * From Clientes Where idCliente = " & cboCliente.ItemData(cmbCliente.ListIndex))
'   If Not rsTemp.EOF Then
'      If Len(Trim(rsTemp!Endereco)) = 0 Then strMensagem = strMensagem & "Endere�o do cliente n�o cadastrado" & Chr(13)
'      If Len(Trim(rsTemp!Bairro)) = 0 Then strMensagem = strMensagem & "Bairro do cliente n�o cadastrado" & Chr(13)
'      If Len(Trim(RemoveCaracteres(rsTemp!CNPJ_CPF))) = 0 Then strMensagem = strMensagem & "CPF/CNPJ do cliente n�o cadastrado" & Chr(13)
'   End If

   If Not strMensagem = Empty Then
      Beep
      MsgBox "Verifique os Seguintes Campos:" & Chr(13) & strMensagem, vbExclamation + vbOKOnly, "Campos Obrigat�rios"
      Verifica_Campos_Produtos = False
      Exit Function
   End If

   Set ProcuraItem = lvwProdutos.FindItem(txtCodigo.text)
   If Not ProcuraItem Is Nothing Then
      Beep

      If MsgBox("Produto j� existe na Rela��o" & Chr(13) & "Deseja substitu�-lo", vbYesNo + vbQuestion, "Produtos") = vbNo Then
         Verifica_Campos_Produtos = False
         txtCodigo.text = Empty
         txtProduto.text = Empty
         mskQuantidade.text = Empty
         mskDesconto.text = Empty
         mskUnitario.text = Empty
         txtCodigo.SetFocus

         Exit Function
      Else
         mskValorTotalProdutos.text = Format(rsNFe!ValorTotalProdutos - Val(Substitui(lvwProdutos.ListItems(lvwProdutos.SelectedItem.Index).ListSubItems(6), ",", ".")), "###,##0.00")
         mskValorTotalNota.text = Format(rsNFe!ValorTotalNota - Val(Substitui(lvwProdutos.ListItems(lvwProdutos.SelectedItem.Index).ListSubItems(6), ",", ".")), "###,##0.00")

         cnSistema.Execute "Update NFe set " & _
               "ValorTotalProdutos = '" & mskValorTotalProdutos.text & "', " & _
               "ValorTotalNota = '" & mskValorTotalNota.text & "' " & _
               "Where idNFe = " & rsNFe!idNFe

         Set rsTemp = cnSistema.Execute("Select * From Produtos Where Codigo = '" & txtCodigo.text & "'")
         cnSistema.Execute "Delete from NFeItens Where idNFe = " & rsNFe!idNFe & " and idProduto = " & rsTemp!idProduto
         lvwProdutos.ListItems.Remove ProcuraItem.Index
      End If
   End If
End Function

Private Sub cmdProduto_Click()
'   frmPesquisaProduto.Show vbModal

   Registro_Selecionado = False
   Screen.MousePointer = vbDefault
   frmPesquisaProduto.Show vbModal
   If Registro_Selecionado Then
      Set rsTemp = cnSistema.Execute("Select * From Produtos Where idProduto = " & Val(Mid(frmPesquisaProduto.lvwDados.SelectedItem.Key, 2, Len(frmPesquisaProduto.lvwDados.SelectedItem.Key))))
   End If

'   rsProdutos.MoveFirst
'   If Registro_Selecionado Then
'      rsProdutos.Find "idProduto = " & Val(Mid(frmPesquisaProduto.lvwDados.SelectedItem.Key, 2, Len(frmPesquisaProduto.lvwDados.SelectedItem.Key)))
'   End If
'
   If frmPesquisaProduto.lvwDados.ListItems.Count <> 0 Then
'      txtCodigo.Text = Mid(frmPesquisaProduto.lvwDados.SelectedItem.Key, 2, Len(frmPesquisaProduto.lvwDados.SelectedItem.Key))
      txtCodigo.text = rsTemp!Codigo
      txtCodigo.SetFocus
      Sendkeys "{TAB}"
   End If
'
End Sub

Private Sub txtCodigo_LostFocus()
   If txtCodigo.text <> Empty Then
      Set rsNaturezasOperacao = cnSistema.Execute("Select * From NaturezasOperacao Where Descricao = '" & cmbNaturezaOperacao.text & "'")
      Set rsTemp = cnSistema.Execute("Select * From Produtos Where Codigo = '" & SQLCheck(txtCodigo.text) & "'")
      If Not rsTemp.EOF Then
         If rsTemp!Situacao = 0 Then
            Set rsTemp2 = cnSistema.Execute("Select * From NFeItens Where idNFe = " & rsNFe!idNFe & " And idProduto = " & rsTemp!idProduto)
            If Not rsTemp2.EOF Then
               mskUnitario.text = rsTemp2!ValorUnitario
               mskQuantidade.text = rsTemp2!Quantidade
               mskDesconto.text = rsTemp2!Desconto
            End If

            txtCodigo.text = txtCodigo.text
            txtProduto.text = rsTemp!Descricao
            If rsNaturezasOperacao!Tipo <> 1 Then
               mskUnitario.text = rsTemp!Preco
            Else
               mskUnitario.text = rsTemp!ValorCusto
            End If
            mskQuantidade.SetFocus
         Else
            Beep
            MsgBox "Este Produto est� Inativo", vbOKOnly + vbInformation, "Produtos"
            txtCodigo.SelStart = 0
            txtCodigo.SelLength = Len(txtCodigo.text)
            txtCodigo.SetFocus
         End If
      Else
         Beep
         MsgBox "N�o existe Produto com este C�digo", vbOKOnly + vbInformation, "Produtos"
         txtCodigo.SelStart = 0
         txtCodigo.SelLength = Len(txtCodigo.text)
         txtCodigo.SetFocus
      End If
   End If
End Sub

Private Sub cmbCliente_Click()
   If cmbCliente.ListIndex <> -1 Then
      Set rsTemp = cnSistema.Execute("Select * From Clientes Where idCliente = " & cmbCliente.ItemData(cmbCliente.ListIndex))
      lblDescricaoEndereco.Caption = Trim(rsTemp!NomeFantasia) & Chr(13) & Trim(rsTemp!Endereco) & ", " & Trim(rsTemp!Bairro) & Chr(13) & Trim(rsTemp!Cidade) & " - " & rsTemp!UF & Chr(13) & "CEP: " & rsTemp!CEP & " Fone: " & rsTemp!Telefone1
   End If
End Sub

Private Sub cmbCliente_LostFocus()
   If Status = 1 Then
      If mskNumero.text <> "" And cmbCliente.ListIndex >= 0 Then
         Set rsTemp = cnSistema.Execute("Select * From NFe Where Numero = " & mskNumero.text & " and idCliente = " & cmbCliente.ItemData(cmbCliente.ListIndex))
         If Not rsTemp.EOF Then
            Beep
            MsgBox "Nota fiscal j� lan�ada", vbExclamation + vbOKOnly, "Localiza��o"
            mskNumero.text = "      "
            mskNumero.SetFocus
            Exit Sub
         End If
      End If

      Set rsClientes = cnSistema.Execute("Select * From Clientes Where idCliente = " & cmbCliente.ItemData(cmbCliente.ListIndex))

    ' Preencher Campos
      If rsClientes!DescontoMaximo <> 0 Then
         mskDescontoGeral.text = rsClientes!DescontoMaximo
      End If

    ' Pesquisa Clientes
      Dim Contador As Integer
      For Contador = 0 To (cmbFormaPagamento.ListCount - 1)
         If cmbFormaPagamento.ItemData(Contador) = rsClientes!idFormaPagamento Then
            cmbFormaPagamento.ListIndex = Contador
            Exit For
         End If
      Next

      Dim strMensagem As String
      strMensagem = "Forma de Pagamento: " & cmbFormaPagamento.text & Chr(13)
      strMensagem = strMensagem & "Desconto M�ximo: " & rsClientes!DescontoMaximo

'      MsgBox "Verifique os Seguintes Campos:" & Chr(13) & strMensagem, vbExclamation + vbOKOnly, "Campos Obrigat�rios"

    ' Situa��o Pendente
      If rsClientes!Situacao <> 0 Then
         strMensagem = rsClientes!MotivoInativacao
         MsgBox "Cliente Inativo e com situa��o Pendente: " & Chr(13) & strMensagem, vbExclamation + vbOKOnly, "Campos Obrigat�rios"
      End If
   End If
End Sub

Private Sub mskValorFrete_LostFocus()
   mskValorTotalNota.text = 0
   mskValorTotalNota.text = Format(CStrValor(mskValorTotalNota.text) + CStrValor(mskValorTotalProdutos.text), "###,##0.00")
   mskValorTotalNota.text = Format(CStrValor(mskValorTotalNota.text) + CStrValor(mskValorFrete.text), "###,##0.00")
   mskValorTotalNota.text = Format(CStrValor(mskValorTotalNota.text) + CStrValor(mskOutrasDespesas.text), "###,##0.00")
End Sub

Private Sub mskOutrasDespesas_LostFocus()
   mskValorTotalNota.text = 0
   mskValorTotalNota.text = Format(CStrValor(mskValorTotalNota.text) + CStrValor(mskValorTotalProdutos.text), "###,##0.00")
   mskValorTotalNota.text = Format(CStrValor(mskValorTotalNota.text) + CStrValor(mskValorFrete.text), "###,##0.00")
   mskValorTotalNota.text = Format(CStrValor(mskValorTotalNota.text) + CStrValor(mskOutrasDespesas.text), "###,##0.00")
End Sub

Private Sub mskValorTotalProdutos_LostFocus()
   mskValorTotalNota.text = 0
   mskValorTotalNota.text = Format(CStrValor(mskValorTotalNota.text) + CStrValor(mskValorTotalProdutos.text), "###,##0.00")
   mskValorTotalNota.text = Format(CStrValor(mskValorTotalNota.text) + CStrValor(mskValorFrete.text), "###,##0.00")
   mskValorTotalNota.text = Format(CStrValor(mskValorTotalNota.text) + CStrValor(mskOutrasDespesas.text), "###,##0.00")
End Sub

Private Sub cmbNaturezaOperacao_LostFocus()
   If cmbNaturezaOperacao.ListIndex >= 0 And cmbCliente.ListIndex >= 0 Then
      Set rsClientes = cnSistema.Execute("Select * From Clientes Where idCliente = " & cmbCliente.ItemData(cmbCliente.ListIndex))
      Set rsTemp = cnSistema.Execute("Select * From NaturezasOperacao Where idNaturezaOperacao = " & cmbNaturezaOperacao.ItemData(cmbNaturezaOperacao.ListIndex))
      If Not rsEmpresa.EOF Then
         If rsClientes!UF = rsEmpresa!UF Then
            mskCFOP.text = rsTemp!CFOPDentroUF
         Else
            mskCFOP.text = rsTemp!CFOPForaUF
         End If
      End If
   End If
End Sub

Private Sub txtProduto_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then Sendkeys "{TAB}"
End Sub

Private Sub txtProduto_LostFocus()
   If mskNumero.Enabled Then Exit Sub

   If txtCodigo.text <> Empty And strDescricaoTemp = txtProduto.text Then Exit Sub
   If txtCodigo.text = Empty And txtProduto.text = Empty Then
'      Beep
'      MsgBox "Digite o C�digo ou a Descri��o do Produto", vbOKOnly + vbInformation, "Produtos"
'      txtCodigo.SetFocus
'      Exit Sub
'      Set rsTemp = cnSistema.Execute("Select * From Produtos Order By Descricao")
      Set rsTemp = cnSistema.Execute("Select * From TabelaProdutos Where Situacao = 0 Order By Descricao")
   Else
      Set rsTemp = cnSistema.Execute("Select * From TabelaProdutos Where Situacao = 0 And Descricao Like '%" & txtProduto.text & "%' Order By Descricao")
   End If

   If rsTemp.EOF Then
      Beep
      MsgBox "N�o existem Produtos com esta Descri��o", vbOKOnly + vbInformation, "Produtos"
      txtProduto.SelStart = 0
      txtProduto.SelLength = Len(txtProduto.text)
      txtProduto.SetFocus
      Exit Sub
   End If

   Screen.MousePointer = vbHourglass
   frmVisualiza.lvwDados.ColumnHeaders.Clear
   frmVisualiza.lvwDados.ColumnHeaders.Add , , "Produto", 3680
   frmVisualiza.lvwDados.ColumnHeaders.Add , , "C�digo", 970
   frmVisualiza.lvwDados.ColumnHeaders.Add , , "Valor", 700, lvwColumnRight
   frmVisualiza.lvwDados.ColumnHeaders.Add , , "Saldo", 700, lvwColumnRight
   frmVisualiza.lvwDados.ListItems.Clear
   Do While Not rsTemp.EOF
      Set ItemList = frmVisualiza.lvwDados.ListItems.Add(, "R" & CStr(rsTemp!idProduto), rsTemp!Descricao)
      ItemList.SubItems(1) = rsTemp!Codigo
      ItemList.SubItems(2) = Format(rsTemp!Preco, "###,##0.00")
      ItemList.SubItems(3) = Format(rsTemp!Saldo, "###,##0.00")
      rsTemp.MoveNext
   Loop
   Registro_Selecionado = False
   Screen.MousePointer = vbDefault
   frmVisualiza.Show vbModal
   If Registro_Selecionado Then
      Set rsNaturezasOperacao = cnSistema.Execute("Select * From NaturezasOperacao Where Descricao = '" & cmbNaturezaOperacao.text & "'")
      Set rsTemp = cnSistema.Execute("Select * From Produtos Where idProduto = " & Val(Mid(frmVisualiza.lvwDados.SelectedItem.Key, 2, Len(frmVisualiza.lvwDados.SelectedItem.Key))))
      txtCodigo.text = rsTemp!Codigo
      txtProduto.text = rsTemp!Descricao
      If rsNaturezasOperacao!Tipo <> 1 Then
         mskUnitario.text = rsTemp!Preco
      Else
         mskUnitario.text = rsTemp!ValorCusto
      End If

    ' Carrega Combos
      Set rsTemp = cnSistema.Execute("Select * from UnidadesMedida Order By Descricao")
      frmNFeGComplemento.cmbUnidade.Clear
      Do While Not rsTemp.EOF
         frmNFeGComplemento.cmbUnidade.AddItem rsTemp!Descricao
         frmNFeGComplemento.cmbUnidade.ItemData(frmNFeGComplemento.cmbUnidade.NewIndex) = rsTemp!idUnidadeMedida
         rsTemp.MoveNext
      Loop

''    ' Preencher Campos
''      Set rsTemp = cnSistema.Execute("Select * From Produtos Where Codigo = '" & SQLCheck(frmNFeG.txtCodigo.Text) & "'")
''      If Not rsTemp.EOF Then
''         ' ICMS
''         frmNFeGComplemento.mskICMSProduto.Text = rsTemp!ICMS
''         ' Unidade
''         For Contador = 0 To (frmNFeGComplemento.cmbUnidade.ListCount - 1)
''            If frmNFeGComplemento.cmbUnidade.ItemData(Contador) = rsTemp!idUnidade Then
''               frmNFeGComplemento.cmbUnidade.ListIndex = Contador
''               Exit For
''            End If
''         Next
''      End If
   End If
End Sub

Private Sub mskCFOP_LostFocus()
   Set rsTemp = cnSistema.Execute("Select * From CFOPs Where CFOP = '" & mskCFOP.text & "'")
End Sub

Private Sub cmdIncluirFatura_Click()
   If Verifica_Campos_Boletos() Then
      Set ItemList = lvwBoletos.ListItems.Add(, "R" & mskNumeroBoleto.text, mskNumeroBoleto.text)
      ItemList.SubItems(1) = mskVencimentoBoleto.text
      ItemList.SubItems(2) = Format(mskValorBoleto.text, "##,##0.00")

      cnSistema.Execute "Insert Into NFeBoletos (idNFe,Numero,Vencimento,Valor) " & _
                        "Values (" & rsNFe!idNFe & ",'" & mskNumeroBoleto.text & "','" & mskVencimentoBoleto.text & _
                        "'," & Substitui(mskValorBoleto.text, ",", ".") & ")"

      mskNumeroBoleto.text = Empty
      mskVencimentoBoleto.text = "  /  /    "
      mskValorBoleto.text = Empty
      mskNumeroBoleto.SetFocus
   End If
End Sub

Private Sub cmdExcluirFatura_Click()
   Beep
   If MsgBox("Deseja excluir este item", vbYesNo + vbQuestion, "Excluir") = vbYes Then
      cnSistema.Execute "Delete from NFeBoletos Where Numero = '" & Mid(lvwBoletos.SelectedItem.Key, 2, Len(lvwBoletos.SelectedItem.Key)) & "'"
      lvwBoletos.ListItems.Remove (lvwBoletos.SelectedItem.Index)
   End If
End Sub

Private Function Verifica_Campos_Boletos()
Dim strMensagem As String
Dim ProcuraItem As ListItem
Verifica_Campos_Boletos = True

   If mskNumeroBoleto.text = Empty Then strMensagem = strMensagem & "Boleto" & Chr(13)
   If mskVencimentoBoleto.text = Empty Then strMensagem = strMensagem & "Vencimento" & Chr(13)
   If Val(Substitui(mskValorBoleto.text, ",", ".")) = 0 Then strMensagem = strMensagem & "Valor" & Chr(13)

   If Not strMensagem = Empty Then
      Beep
      MsgBox "Verifique os Seguintes Campos:" & Chr(13) & strMensagem, vbExclamation + vbOKOnly, "Campos Obrigat�rios"
      Verifica_Campos_Boletos = False
      Exit Function
   End If

   Set ProcuraItem = lvwBoletos.FindItem(mskNumeroBoleto.text)
End Function

Private Sub cmdPesquisaCliente_Click()
   Registro_Selecionado = False
   Screen.MousePointer = vbDefault
   frmPesquisaClientes.Show vbModal
   If Registro_Selecionado Then
      Set rsTemp = cnSistema.Execute("Select * From Clientes Where idCliente = " & Val(Mid(frmPesquisaClientes.lvwDados.SelectedItem.Key, 2, Len(frmPesquisaClientes.lvwDados.SelectedItem.Key))))
   End If

   If frmPesquisaClientes.lvwDados.ListItems.Count <> 0 Then
      mskCNPJ_CPF.text = rsTemp!CNPJ_CPF
      mskCNPJ_CPF.SetFocus
      Sendkeys "{TAB}"
   End If
   cmbCliente.SetFocus
End Sub

Private Function SaltarLinha(iParametro As Integer) As Integer
Dim Contador As Integer

   If iParametro > 0 Then
      For Contador = 1 To (iParametro - 1)
          Print #1, ""
      Next
   End If
End Function

Private Sub mskQuantidade_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then Sendkeys "{TAB}"
   If KeyAscii = 46 Then KeyAscii = 44
   If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 44 And KeyAscii <> 8 Then
      KeyAscii = 0
   End If
   If KeyAscii = 44 Then
      If InStr(mskQuantidade.ClipText, ",") <> 0 Then
         KeyAscii = 0
      End If
   End If
   If Len(mskQuantidade.text) > 6 And KeyAscii <> 8 And KeyAscii <> 44 Then
      If InStr(mskQuantidade.ClipText, ",") = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub mskDesconto_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then Sendkeys "{TAB}"
   If KeyAscii = 46 Then KeyAscii = 44
   If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 44 And KeyAscii <> 8 Then
      KeyAscii = 0
   End If
   If KeyAscii = 44 Then
      If InStr(mskDesconto.ClipText, ",") <> 0 Then
         KeyAscii = 0
      End If
   End If
   If Len(mskDesconto.text) > 1 And KeyAscii <> 8 And KeyAscii <> 44 Then
      If InStr(mskDesconto.ClipText, ",") = 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub mskUnitario_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then Sendkeys "{TAB}"
   If KeyAscii = 46 Then KeyAscii = 44
   If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 44 And KeyAscii <> 8 Then
      KeyAscii = 0
   End If
   If KeyAscii = 44 Then
      If InStr(mskUnitario.ClipText, ",") <> 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Function FData(dData As String) As Boolean
   If Not IsDate(dData) Or Val(Mid(dData, 7, 4)) < 1900 Then
      MsgBox "Data Inv�lida", vbOKOnly + vbInformation, "Valida��o"
      FData = False
      Exit Function
   Else
      FData = True
   End If
End Function

Private Sub cmdAnotacoes_Click()
Dim Contador As Integer

   If frmNFeG.txtCodigo.text <> "" Then
    ' Carrega Combos
      ' Unidades de Medida
      Set rsTemp = cnSistema.Execute("Select * from UnidadesMedida Order By Descricao")
      frmNFeGComplemento.cmbUnidade.Clear
      Do While Not rsTemp.EOF
         frmNFeGComplemento.cmbUnidade.AddItem rsTemp!Descricao
         frmNFeGComplemento.cmbUnidade.ItemData(frmNFeGComplemento.cmbUnidade.NewIndex) = rsTemp!idUnidadeMedida
         rsTemp.MoveNext
      Loop

      ' Situacoes Tributarias
      Set rsTemp = cnSistema.Execute("Select * from SituacoesTributarias Order By Descricao")
      frmNFeGComplemento.cmbSituacaoTributaria.Clear
      Do While Not rsTemp.EOF
         frmNFeGComplemento.cmbSituacaoTributaria.AddItem rsTemp!Descricao
         frmNFeGComplemento.cmbSituacaoTributaria.ItemData(frmNFeGComplemento.cmbSituacaoTributaria.NewIndex) = rsTemp!idSituacaoTributaria
         rsTemp.MoveNext
      Loop

    ' Preencher Campos
      Set rsTemp = cnSistema.Execute("Select * From Produtos Where Codigo = '" & SQLCheck(frmNFeG.txtCodigo.text) & "'")
      If Not rsTemp.EOF Then
         ' ICMS
         frmNFeGComplemento.mskICMSProduto.text = rsTemp!ICMS
         If Not rsEmpresa.EOF Then
            If rsClientes!UF = rsEmpresa!UF Then
               frmNFeGComplemento.mskBaseReduzidaICMS.text = rsTemp!BaseReduzidaICMSdUF
            Else
               frmNFeGComplemento.mskBaseReduzidaICMS.text = rsTemp!BaseReduzidaICMSfUF
            End If
         End If
         frmNFeGComplemento.mskCFOP.text = mskCFOP.text

         ' Unidade
         For Contador = 0 To (frmNFeGComplemento.cmbUnidade.ListCount - 1)
            If frmNFeGComplemento.cmbUnidade.ItemData(Contador) = rsTemp!idUnidade Then
               frmNFeGComplemento.cmbUnidade.ListIndex = Contador
               Exit For
            End If
         Next

         ' Situacao Tributaria
         For Contador = 0 To (frmNFeGComplemento.cmbSituacaoTributaria.ListCount - 1)
            If frmNFeGComplemento.cmbSituacaoTributaria.ItemData(Contador) = rsTemp!idSituacaoTributaria Then
               frmNFeGComplemento.cmbSituacaoTributaria.ListIndex = Contador
               Exit For
            End If
         Next
      End If

      frmNFeGComplemento.Show vbModal
      cmdIncluir.SetFocus
   End If
End Sub

Sub ImprimirNota()

   If LerArquivoINI("SEPD", "TipoImpressao", CaminhoINI & "\System.ini") = 1 Then
      Call ImprimirDanfe
   ElseIf LerArquivoINI("SEPD", "TipoImpressao", CaminhoINI & "\System.ini") = 2 Then
      Call FolhaSolta
   ElseIf LerArquivoINI("SEPD", "TipoImpressao", CaminhoINI & "\System.ini") = 3 Then
      Call NotaServico
   End If
End Sub

Sub FolhaSolta()
Dim iItensNota As Integer
Dim iPosLinNumero As Integer
Dim iPosColMarcaEntrada As Integer
Dim iPosColMarcaSaida As Integer
Dim iPosColNumero As Integer
Dim iPosLinNatureza As Integer
Dim iPosColNatureza As Integer
Dim iPosColCBO As Integer
Dim iPosLinCliente As Integer
Dim iPosColCliente As Integer
Dim iPosColEmissao As Integer
Dim iPosLinEndereco As Integer
Dim iPosColEndereco As Integer
Dim iPosColBairro As Integer
Dim iPosColSaida As Integer
Dim iPosLinCidade As Integer
Dim iPosColCidade As Integer
Dim iPosColTelefone As Integer
Dim iPosColUF As Integer
Dim iPosColIE_CI As Integer
Dim iPosLinBairro As Integer
Dim iPosColCEP As Integer
Dim iPosColCNPJ As Integer
Dim iPosLinCobranca As Integer
Dim iPosLinProdutos As Integer
Dim iPosColCodigo As Integer
Dim iPosColDescricao As Integer
Dim iPosColUnidade As Integer
Dim iPosColQuantidade As Integer
Dim iPosColVlUnitario As Integer
Dim iPosColVlLiquido As Integer
Dim iPosColICMS As Integer
Dim iPosLinInfoCN As Integer
Dim iPosColInfoCN As Integer
Dim iPosLinBase As Integer
Dim iPosColBase As Integer
Dim iPosColValorICMS As Integer
Dim iPosColTotalProdutos As Integer
Dim iPosLinTotalNota As Integer
Dim iPosColTotalNota As Integer
Dim iPosLinTransportador As Integer
Dim iPosColTransportador As Integer
Dim iPosColFreteConta As Integer
Dim iPosColPlacaVeiculo As Integer
Dim iPosColUFPlaca As Integer
Dim iPosColCNPJ_CPFTrans As Integer
Dim iPosLinEndTrans As Integer
Dim iPosColEndTrans As Integer
Dim iPosColCidTrans As Integer
Dim iPosColUFTrans As Integer
Dim iPosColIE_CITrans As Integer
Dim iPosLinVolQuant As Integer
Dim iPosColVolQuant As Integer
Dim iPosColVolMarca As Integer
Dim iPosColVolNumero As Integer
Dim iPosColPesoBruto As Integer
Dim iPosColPesoLiquido As Integer
Dim iPosLinDadosAdic As Integer
Dim iPosColDadosAdic As Integer
Dim iPosLinNumeroFim As Integer
Dim iPosColNumeroFim As Integer
Dim iPosLinProximaNota As Integer

   Set rsNFeItens = cnSistema.Execute("SELECT * FROM NFeItens WHERE idNFe = " & rsNFe!idNFe)
   If rsNFeItens.EOF Then
      MsgBox "Nota n�o pode ser impressa sem Produtos ", vbExclamation + vbOKOnly, "Campos Obrigat�rios"
      Exit Sub
   End If

 ' Marca e Numero
   iPosLinNumero = LerArquivoINI("Notas Fiscais", "PosLinNumero", CaminhoINI & "\NotasManuais.ini")
   iPosColMarcaEntrada = LerArquivoINI("Notas Fiscais", "PosColMarcaEntrada", CaminhoINI & "\NotasManuais.ini")
   iPosColMarcaSaida = LerArquivoINI("Notas Fiscais", "PosColMarcaSaida", CaminhoINI & "\NotasManuais.ini")
   iPosColNumero = LerArquivoINI("Notas Fiscais", "PosColNumero", CaminhoINI & "\NotasManuais.ini")

 ' Natureza e CBO
   iPosLinNatureza = LerArquivoINI("Notas Fiscais", "PosLinNatureza", CaminhoINI & "\NotasManuais.ini")
   iPosColNatureza = LerArquivoINI("Notas Fiscais", "PosColNatureza", CaminhoINI & "\NotasManuais.ini")
   iPosColCBO = LerArquivoINI("Notas Fiscais", "PosColCBO", CaminhoINI & "\NotasManuais.ini")

 ' Cliente, CNPJ e Emiss�o
   iPosLinCliente = LerArquivoINI("Notas Fiscais", "PosLinCliente", CaminhoINI & "\NotasManuais.ini")
   iPosColCliente = LerArquivoINI("Notas Fiscais", "PosColCliente", CaminhoINI & "\NotasManuais.ini")
   iPosColEmissao = LerArquivoINI("Notas Fiscais", "PosColEmissao", CaminhoINI & "\NotasManuais.ini")

 ' Endere�o e Data de Saida
   iPosLinEndereco = LerArquivoINI("Notas Fiscais", "PosLinEndereco", CaminhoINI & "\NotasManuais.ini")
   iPosColEndereco = LerArquivoINI("Notas Fiscais", "PosColEndereco", CaminhoINI & "\NotasManuais.ini")
   iPosColSaida = LerArquivoINI("Notas Fiscais", "PosColSaida", CaminhoINI & "\NotasManuais.ini")

 ' Cidade, Telefone, UF e Inscri��o Estadual
   iPosLinCidade = LerArquivoINI("Notas Fiscais", "PosLinCidade", CaminhoINI & "\NotasManuais.ini")
   iPosColCidade = LerArquivoINI("Notas Fiscais", "PosColCidade", CaminhoINI & "\NotasManuais.ini")
   iPosColTelefone = LerArquivoINI("Notas Fiscais", "PosColTelefone", CaminhoINI & "\NotasManuais.ini")
   iPosColUF = LerArquivoINI("Notas Fiscais", "PosColUF", CaminhoINI & "\NotasManuais.ini")
   iPosColIE_CI = LerArquivoINI("Notas Fiscais", "PosColIE_CI", CaminhoINI & "\NotasManuais.ini")

 ' Bairro, CEP e CNPJ
   iPosLinBairro = LerArquivoINI("Notas Fiscais", "PosLinBairro", CaminhoINI & "\NotasManuais.ini")
   iPosColBairro = LerArquivoINI("Notas Fiscais", "PosColBairro", CaminhoINI & "\NotasManuais.ini")
   iPosColCEP = LerArquivoINI("Notas Fiscais", "PosColCEP", CaminhoINI & "\NotasManuais.ini")
   iPosColCNPJ = LerArquivoINI("Notas Fiscais", "PosColCNPJ", CaminhoINI & "\NotasManuais.ini")

 ' Cobranca
   iPosLinCobranca = LerArquivoINI("Notas Fiscais", "PosLinCobranca", CaminhoINI & "\NotasManuais.ini")

 ' Produtos
   iPosLinProdutos = LerArquivoINI("Notas Fiscais", "PosLinProdutos", CaminhoINI & "\NotasManuais.ini")
   iPosColCodigo = LerArquivoINI("Notas Fiscais", "PosColCodigo", CaminhoINI & "\NotasManuais.ini")
   iPosColDescricao = LerArquivoINI("Notas Fiscais", "PosColDescricao", CaminhoINI & "\NotasManuais.ini")
   iPosColUnidade = LerArquivoINI("Notas Fiscais", "PosColUnidade", CaminhoINI & "\NotasManuais.ini")
   iPosColQuantidade = LerArquivoINI("Notas Fiscais", "PosColQuantidade", CaminhoINI & "\NotasManuais.ini")
   iPosColVlUnitario = LerArquivoINI("Notas Fiscais", "PosColVlUnitario", CaminhoINI & "\NotasManuais.ini")
   iPosColVlLiquido = LerArquivoINI("Notas Fiscais", "PosColVlLiquido", CaminhoINI & "\NotasManuais.ini")
   iPosColICMS = LerArquivoINI("Notas Fiscais", "PosColICMS", CaminhoINI & "\NotasManuais.ini")

 ' Informa��es do Corpo da Nota
   iPosLinInfoCN = LerArquivoINI("Notas Fiscais", "PosLinInfoCN", CaminhoINI & "\NotasManuais.ini")
   iPosColInfoCN = LerArquivoINI("Notas Fiscais", "PosColInfoCN", CaminhoINI & "\NotasManuais.ini")

 ' Base de Calculo, Valor do ICMS e Valor Total dos Produtos
   iPosLinBase = LerArquivoINI("Notas Fiscais", "PosLinBase", CaminhoINI & "\NotasManuais.ini")
   iPosColBase = LerArquivoINI("Notas Fiscais", "PosColBase", CaminhoINI & "\NotasManuais.ini")
   iPosColValorICMS = LerArquivoINI("Notas Fiscais", "PosColValorICMS", CaminhoINI & "\NotasManuais.ini")
   iPosColTotalProdutos = LerArquivoINI("Notas Fiscais", "PosColTotalProdutos", CaminhoINI & "\NotasManuais.ini")

 ' Valor Total da Nota
   iPosLinTotalNota = LerArquivoINI("Notas Fiscais", "PosLinTotalNota", CaminhoINI & "\NotasManuais.ini")
   iPosColTotalNota = LerArquivoINI("Notas Fiscais", "PosColTotalNota", CaminhoINI & "\NotasManuais.ini")

 ' Transportador
   iPosLinTransportador = LerArquivoINI("Notas Fiscais", "PosLinTransportador", CaminhoINI & "\NotasManuais.ini")
   iPosColTransportador = LerArquivoINI("Notas Fiscais", "PosColTransportador", CaminhoINI & "\NotasManuais.ini")
   iPosColFreteConta = LerArquivoINI("Notas Fiscais", "PosColFreteConta", CaminhoINI & "\NotasManuais.ini")
   iPosColPlacaVeiculo = LerArquivoINI("Notas Fiscais", "PosColPlacaVeiculo", CaminhoINI & "\NotasManuais.ini")
   iPosColUFPlaca = LerArquivoINI("Notas Fiscais", "PosColUFPlaca", CaminhoINI & "\NotasManuais.ini")
   iPosColCNPJ_CPFTrans = LerArquivoINI("Notas Fiscais", "PosColCNPJ_CPFTrans", CaminhoINI & "\NotasManuais.ini")

   iPosLinEndTrans = LerArquivoINI("Notas Fiscais", "PosLinEndTrans", CaminhoINI & "\NotasManuais.ini")
   iPosColEndTrans = LerArquivoINI("Notas Fiscais", "PosColEndTrans", CaminhoINI & "\NotasManuais.ini")
   iPosColCidTrans = LerArquivoINI("Notas Fiscais", "PosColCidTrans", CaminhoINI & "\NotasManuais.ini")
   iPosColUFTrans = LerArquivoINI("Notas Fiscais", "PosColUFTrans", CaminhoINI & "\NotasManuais.ini")
   iPosColIE_CITrans = LerArquivoINI("Notas Fiscais", "PosColIE_CITrans", CaminhoINI & "\NotasManuais.ini")

 ' Volume
   iPosLinVolQuant = LerArquivoINI("Notas Fiscais", "PosLinVolQuant", CaminhoINI & "\NotasManuais.ini")
   iPosColVolQuant = LerArquivoINI("Notas Fiscais", "PosColVolQuant", CaminhoINI & "\NotasManuais.ini")
   iPosColVolMarca = LerArquivoINI("Notas Fiscais", "PosColVolMarca", CaminhoINI & "\NotasManuais.ini")
   iPosColVolNumero = LerArquivoINI("Notas Fiscais", "PosColVolNumero", CaminhoINI & "\NotasManuais.ini")
   iPosColPesoBruto = LerArquivoINI("Notas Fiscais", "PosColPesoBruto", CaminhoINI & "\NotasManuais.ini")
   iPosColPesoLiquido = LerArquivoINI("Notas Fiscais", "PosColPesoLiquido", CaminhoINI & "\NotasManuais.ini")

 ' Dados Adicionais
   iPosLinDadosAdic = LerArquivoINI("Notas Fiscais", "PosLinDadosAdic", CaminhoINI & "\Notas.ini")
   iPosColDadosAdic = LerArquivoINI("Notas Fiscais", "PosColDadosAdic", CaminhoINI & "\Notas.ini")

 ' Numero Final
   iPosLinNumeroFim = LerArquivoINI("Notas Fiscais", "PosLinNumeroFim", CaminhoINI & "\NotasManuais.ini")
   iPosColNumeroFim = LerArquivoINI("Notas Fiscais", "PosColNumeroFim", CaminhoINI & "\NotasManuais.ini")
   iPosLinProximaNota = LerArquivoINI("Notas Fiscais", "PosLinProximaNota", CaminhoINI & "\NotasManuais.ini")

   If rsNFe!Impressa Then
      Beep
      MsgBox "Nota Fiscal j� Impressa", vbExclamation, "Aviso"
   End If

   cdlgImprimirNota.CancelError = True

   Registro_Selecionado = False
   VisualizarImpressao
   If Registro_Selecionado Then
      Open LerArquivoINI("Impressoras", "Notas", CaminhoINI & "\System.ini") For Output As #1
'      Open caminhoini & "\teste.txt" For Output As #1

      Set rsCFOPs = cnSistema.Execute("Select * From CFOPs Where idCFOP = " & rsNFe!idCFOP)
      Set rsClientes = cnSistema.Execute("Select * From Clientes Where idCliente = " & rsNFe!idCliente)
      Set rsNaturezasOperacao = cnSistema.Execute("Select * From NaturezasOperacao Where idNaturezaOperacao = " & rsNFe!idNaturezaOperacao)

    ' Numero da Nota
      Print #1, Chr(27) & "x0"; Chr$(27) & Chr(69); Chr(15)
      SaltarLinha (iPosLinNumero)
      Print #1, Tab(iPosColMarcaSaida); "X"
      SaltarLinha (iPosLinNatureza)

    ' Codigo Fiscal de Opera��o
      Print #1, Tab(iPosColNatureza); rsNaturezasOperacao!Descricao; Tab(iPosColCBO); rsCFOPs!CFOP
      SaltarLinha (iPosLinCliente)

    ' Cliente e Emiss�o
'''      Print #1, Tab(iPosColCliente); RemoveAcentos(rsClientes!Nome); Tab(iPosColCNPJ); rsClientes!CNPJ_CPF; Tab(iPosColEmissao); rsNFe!DataEmissao
      Print #1, Tab(iPosColCliente); RemoveAcentos(rsClientes!Nome); Tab(iPosColEmissao); rsNFe!DataEmissao
      SaltarLinha (iPosLinEndereco)

    ' Endere�o e Saida
'''      Print #1, Tab(iPosColEndereco); RemoveAcentos(rsClientes!Endereco); Tab(iPosColBairro); RemoveAcentos(rsClientes!Bairro); Tab(iPosColCEP); rsClientes!CEP; Tab(iPosColSaida); rsNFe!DataEmissao
      Print #1, Tab(iPosColEndereco); RemoveAcentos(IIf(Not IsNull(rsClientes!Endereco), rsClientes!Endereco, "")); Tab(iPosColSaida); rsNFe!DataEmissao
      SaltarLinha (iPosLinCidade)

    ' Cidade, Telefone, UF e Inscri��o Estadual
'''      Print #1, Tab(iPosColCidade); RemoveAcentos(rsClientes!Cidade); Tab(iPosColTelefone); rsClientes!Telefone1; Tab(iPosColUF); rsClientes!UF; Tab(iPosColIE_CI); rsClientes!IE_CI
      Print #1, Tab(iPosColCidade); RemoveAcentos(IIf(Not IsNull(rsClientes!Cidade), rsClientes!Cidade, "")); Tab(iPosColTelefone); rsClientes!Telefone1; Tab(iPosColUF); IIf(Not IsNull(rsClientes!UF), rsClientes!UF, ""); Tab(iPosColIE_CI); IIf(Not IsNull(rsClientes!IE_CI), rsClientes!IE_CI, "")
      SaltarLinha (iPosLinBairro)

    ' Bairro, CEP e CNPJ/CPF
'''      Print #1, Tab(iPosColCidade); RemoveAcentos(rsClientes!Cidade); Tab(iPosColTelefone); rsClientes!Telefone1; Tab(iPosColUF); rsClientes!UF; Tab(iPosColIE_CI); rsClientes!IE_CI
      Print #1, Tab(iPosColBairro); RemoveAcentos(IIf(Not IsNull(rsClientes!Bairro), rsClientes!Bairro, "")); Tab(iPosColCEP); IIf(Not IsNull(rsClientes!CEP), rsClientes!CEP, ""); Tab(iPosColCNPJ); IIf(Not IsNull(rsClientes!CNPJ_CPF), rsClientes!CNPJ_CPF, "")
      SaltarLinha (iPosLinCobranca)

    ' Cobranca
'            Print #1, Tab(5); rsNFe!Documento; Spc(5); cmbFormaPagamento.Text; Spc(5); rsNFe!DataVencimento
      SaltarLinha (iPosLinProdutos)

    ' Produtos

      Dim cValorBruto As Currency
      Dim cValorDesconto As Currency
      Dim cValorBonificacao As Currency
      Dim cValorLiquido As Currency

      Dim iLinhasProdutos As Integer
      iLinhasProdutos = LerArquivoINI("Notas Fiscais", "ItensNota", CaminhoINI & "\NotasManuais.ini")
      Set rsNFeItens = cnSistema.Execute("SELECT NFeItens.idNFe, NFeItens.idProduto, NFeItens.idUnidade, Produtos.Descricao, NFeItens.idClassificacaoFiscal, NFeItens.idSituacaoTributaria, NFeItens.Data, NFeItens.Quantidade, NFeItens.ValorUnitario, NFeItens.Desconto, NFeItens.IPI, NFeItens.ICMS, NFeItens.DescricaoComplementar " & _
                                                  "FROM NFeItens INNER JOIN Produtos ON NFeItens.idProduto = Produtos.idProduto " & _
                                                  "Where NFeItens.idNFe = " & rsNFe!idNFe & " ORDER BY Produtos.Descricao")

      Do While Not rsNFeItens.EOF
         Set rsProdutos = cnSistema.Execute("SELECT * FROM Produtos WHERE idProduto = " & rsNFeItens!idProduto)
         Set rsUnidadesMedida = cnSistema.Execute("SELECT * FROM UnidadesMedida WHERE idUnidadeMedida = " & rsNFeItens!idUnidade)
         Dim sDescricao As String, sUnidade As String
         sDescricao = RemoveAcentos(Mid(Trim(rsProdutos!Descricao), 1, 50) & " " & Trim(rsNFeItens!DescricaoComplementar))
         If Not rsUnidadesMedida.EOF Then
            sUnidade = rsUnidadesMedida!Sigla
         Else
            sUnidade = " "
         End If

         cValorBruto = (rsNFeItens!Quantidade * rsNFeItens!ValorUnitario)
         cValorDesconto = (((rsNFeItens!Quantidade * rsNFeItens!ValorUnitario) * rsNFeItens!Desconto) / 100)
         cValorBonificacao = (((cValorBruto - cValorDesconto) * rsNFe!Bonificacao) / 100)
         cValorLiquido = (cValorBruto - cValorDesconto - cValorBonificacao)

         Print #1, Tab(iPosColCodigo); rsProdutos!Codigo; _
                   Tab(iPosColDescricao); sDescricao; _
                   Tab(iPosColUnidade); sUnidade; _
                   Tab(iPosColQuantidade); FormataTXT(Format(rsNFeItens!Quantidade, "##,##0.00"), 2.1, 10); _
                   Tab(iPosColVlUnitario); FormataTXT(Format(rsNFeItens!ValorUnitario, "##,##0.00"), 2.1, 10); _
                   Tab(iPosColVlLiquido); FormataTXT(Format(cValorLiquido, "##,##0.00"), 2.1, 12); _
                   Tab(iPosColICMS); Format(rsNFeItens!ICMS, "##0")

         iLinhasProdutos = iLinhasProdutos - 1
         rsNFeItens.MoveNext
      Loop

    ' Informa��es do Corpo da Nota
      SaltarLinha (iLinhasProdutos + iPosLinInfoCN)
      Dim sInformacoesCorpo1 As String, sInformacoesCorpo2 As String, sInformacoesCorpo3 As String
      If Len(rsNFe!InformacoesCorpo) >= 240 Then
         sInformacoesCorpo1 = Mid(rsNFe!InformacoesCorpo, 1, 120)
         sInformacoesCorpo2 = Mid(rsNFe!InformacoesCorpo, 121, 120)
         sInformacoesCorpo3 = Mid(rsNFe!InformacoesCorpo, 241, 120)

         Print #1, Tab(iPosColInfoCN); sInformacoesCorpo1
         Print #1, Tab(iPosColInfoCN); sInformacoesCorpo2
         Print #1, Tab(iPosColInfoCN); sInformacoesCorpo3
         SaltarLinha (iPosLinBase - 3)
      Else
         If Len(rsNFe!InformacoesCorpo) >= 120 Then
            sInformacoesCorpo1 = Mid(rsNFe!InformacoesCorpo, 1, 120)
            sInformacoesCorpo2 = Mid(rsNFe!InformacoesCorpo, 121, 120)

            Print #1, Tab(iPosColInfoCN); sInformacoesCorpo1
            Print #1, Tab(iPosColInfoCN); sInformacoesCorpo2
            SaltarLinha (iPosLinBase - 2)
         Else
            If Len(rsNFe!InformacoesCorpo) >= 1 Then
               sInformacoesCorpo1 = rsNFe!InformacoesCorpo

               Print #1, Tab(iPosColInfoCN); sInformacoesCorpo1
               SaltarLinha (iPosLinBase - 1)
            Else
               SaltarLinha (iPosLinBase)
            End If
         End If
      End If

'      Print #1, Tab(iPosColInfoCN); rsNFe!InformacoesCorpo

    ' Total
      Set rsTemp = cnSistema.Execute("Select * From TotalNFe Where Numero = " & mskNumero.text)
      Dim dBaseCalculo As Double, dValorICMS As Double
      If rsTemp!ValorICMS > 0 Then
         dBaseCalculo = rsTemp!BaseCalculo
         dValorICMS = rsTemp!ValorICMS
      Else
         dBaseCalculo = 0
         dValorICMS = 0
      End If

      If Not rsTemp.EOF Then
'         Print #1, Tab(10); Format(rsNFe!BaseCalculoICMS, "#0.00"); Tab(40); Format(rsNFe!ValorICMS, "#0.00"); Tab(130); Format(rsTemp!Total, "#0.00")
         Print #1, Tab(iPosColBase); Format(dBaseCalculo, "##,##0.00"); Tab(iPosColValorICMS); Format(dValorICMS, "##,##0.00"); Tab(iPosColTotalProdutos); Format(rsTemp!Total, "##,##0.00")
         SaltarLinha (iPosLinTotalNota)
         Print #1, Tab(iPosColTotalNota); Format(rsTemp!Total, "##,##0.00")
         SaltarLinha (iPosLinTransportador)
      Else
'         Print #1, Tab(10); Format(rsNFe!BaseCalculoICMS, "#0.00"); Tab(40); Format(rsNFe!ValorICMS, "#0.00"); Tab(130); Format(0, "#0.00")
         Print #1, Tab(iPosColBase); Format(dBaseCalculo, "##,##0.00"); Tab(iPosColValorICMS); Format(dValorICMS, "##,##0.00"); Tab(iPosColTotalProdutos); Format(0, "##,##0.00")
         SaltarLinha (iPosLinTotalNota)
         Print #1, Tab(iPosColTotalNota); Format(0, "##,##0.00")
         SaltarLinha (iPosLinTransportador)
      End If

    ' Transportador
      Set rsTransportador = cnSistema.Execute("SELECT * FROM Transportadores WHERE idTransportador = " & rsNFe!idTransportador)
      If Not rsTransportador.EOF Then
         Dim sFreteConta As String
         sFreteConta = IIf(rsNFe!FreteConta > 0, rsNFe!FreteConta, "")

         Print #1, Tab(iPosColTransportador); rsTransportador!Nome; Tab(iPosColFreteConta); sFreteConta; Tab(iPosColPlacaVeiculo); rsNFe!PlacaVeiculo; Tab(iPosColUFPlaca); rsTransportador!UFPlaca; Tab(iPosColCNPJ_CPFTrans); ; IIf(Len(rsTransportador!CNPJ_CPF) > 4, rsTransportador!CNPJ_CPF, "")
         SaltarLinha (iPosLinEndTrans)
         Print #1, Tab(iPosColEndTrans); rsTransportador!Endereco; Tab(iPosColCidTrans); rsTransportador!Cidade; Tab(iPosColUFTrans); rsTransportador!UF; Tab(iPosColIE_CITrans); rsTransportador!IE_CI
         SaltarLinha (iPosLinVolQuant)
      Else
         SaltarLinha (6)
      End If
      Print #1, Tab(iPosColVolQuant); rsNFe!VolumeQuantidade; Tab(iPosColVolMarca); rsNFe!VolumeMarca; Tab(iPosColVolNumero); rsNFe!VolumeNumero; Tab(iPosColPesoBruto); Format(rsTemp!PesoTotal, "##,##0.00"); Tab(iPosColPesoLiquido); Format(rsTemp!PesoTotal, "##,##0.00")

    ' Dados Adicionais
      SaltarLinha (iPosLinDadosAdic)
      Dim sDadosAdicionais1 As String, sDadosAdicionais2 As String, sDadosAdicionais3 As String
      If Len(rsNFe!DadosAdicionais) >= 120 Then
         sDadosAdicionais1 = RemoveAcentos(Mid(rsNFe!DadosAdicionais, 1, 60))
         sDadosAdicionais2 = RemoveAcentos(Mid(rsNFe!DadosAdicionais, 61, 60))
         sDadosAdicionais3 = RemoveAcentos(Mid(rsNFe!DadosAdicionais, 121, 60))

         Print #1, Tab(iPosColDadosAdic); sDadosAdicionais1
         Print #1, Tab(iPosColDadosAdic); sDadosAdicionais2
         Print #1, Tab(iPosColDadosAdic); sDadosAdicionais3
         SaltarLinha (iPosLinNumeroFim - 3)
      Else
         If Len(rsNFe!DadosAdicionais) >= 60 Then
            sDadosAdicionais1 = RemoveAcentos(Mid(rsNFe!DadosAdicionais, 1, 60))
            sDadosAdicionais2 = RemoveAcentos(Mid(rsNFe!DadosAdicionais, 61, 60))

            Print #1, Tab(iPosColDadosAdic); sDadosAdicionais1
            Print #1, Tab(iPosColDadosAdic); sDadosAdicionais2
            SaltarLinha (iPosLinNumeroFim - 2)
         Else
            If Len(rsNFe!DadosAdicionais) >= 1 Then
               sDadosAdicionais1 = RemoveAcentos(rsNFe!DadosAdicionais)

               Print #1, Tab(iPosColDadosAdic); sDadosAdicionais1
               SaltarLinha (iPosLinNumeroFim - 1)
            Else
               SaltarLinha (iPosLinNumeroFim)
            End If
         End If
      End If

    ' Numero da Nota
'      SaltarLinha (iPosLinNumeroFim)
      Print #1, Tab(iPosColNumeroFim); Chr$(27) & Chr(70); " " & Chr(18)
      SaltarLinha (iPosLinProximaNota)

      Close #1
      cnSistema.Execute "Update NFe set " & _
            "Impressa = " & True & " " & _
            "Where idNFe = " & rsNFe!idNFe
   End If

End Sub

Sub NotaServico()
Dim iItensNota As Integer

Dim iPosLinEmissao As Integer
Dim iPosColEmissao As Integer
Dim iPosColNumero  As Integer

Dim iPosLinCliente As Integer
Dim iPosColCliente As Integer

Dim iPosLinEndereco As Integer
Dim iPosColEndereco As Integer

Dim iPosLinCidade As Integer
Dim iPosColCidade As Integer
Dim iPosColUF As Integer
Dim iPosColTelefone As Integer

Dim iPosLinCNPJ As Integer
Dim iPosColCNPJ As Integer
Dim iPosColIE As Integer

Dim iPosLinCobranca As Integer
Dim iPosColDocBol1 As Integer
Dim iPosColVenBol1 As Integer
Dim iPosColValBol1 As Integer
Dim iPosColDocBol2 As Integer
Dim iPosColVenBol2 As Integer
Dim iPosColValBol2 As Integer
Dim iPosColDocBol3 As Integer
Dim iPosColVenBol3 As Integer
Dim iPosColValBol3 As Integer
Dim iPosColDocBol4 As Integer
Dim iPosColVenBol4 As Integer
Dim iPosColValBol4 As Integer

Dim iPosLinProdutos As Integer
Dim iPosColQuantidade As Integer
Dim iPosColUnidade As Integer
Dim iPosColDescricao As Integer
Dim iPosColVlUnitario As Integer
Dim iPosColVlLiquido As Integer

Dim iPosLinTotalNota As Integer
Dim iPosColTotalNota As Integer

Dim iPosLinNumeroFim As Integer
Dim iPosColNumeroFim As Integer
Dim iPosLinProximaNota As Integer

   Set rsNFeItens = cnSistema.Execute("SELECT * FROM NFeItens WHERE idNFe = " & rsNFe!idNFe)
   If rsNFeItens.EOF Then
      MsgBox "Nota n�o pode ser impressa sem Produtos ", vbExclamation + vbOKOnly, "Campos Obrigat�rios"
      Exit Sub
   End If

 ' Emissao e Numero
   iPosLinEmissao = LerArquivoINI("Notas Fiscais", "PosLinEmissao", CaminhoINI & "\NotasServico.ini")
   iPosColEmissao = LerArquivoINI("Notas Fiscais", "PosColEmissao", CaminhoINI & "\NotasServico.ini")
   iPosColNumero = LerArquivoINI("Notas Fiscais", "PosColNumero", CaminhoINI & "\NotasServico.ini")

 ' Cliente
   iPosLinCliente = LerArquivoINI("Notas Fiscais", "PosLinCliente", CaminhoINI & "\NotasServico.ini")
   iPosColCliente = LerArquivoINI("Notas Fiscais", "PosColCliente", CaminhoINI & "\NotasServico.ini")

 ' Endereco
   iPosLinEndereco = LerArquivoINI("Notas Fiscais", "PosLinEndereco", CaminhoINI & "\NotasServico.ini")
   iPosColEndereco = LerArquivoINI("Notas Fiscais", "PosColEndereco", CaminhoINI & "\NotasServico.ini")

 ' Cidade, UF e Telefone
   iPosLinCidade = LerArquivoINI("Notas Fiscais", "PosLinCidade", CaminhoINI & "\NotasServico.ini")
   iPosColCidade = LerArquivoINI("Notas Fiscais", "PosColCidade", CaminhoINI & "\NotasServico.ini")
   iPosColUF = LerArquivoINI("Notas Fiscais", "PosColUF", CaminhoINI & "\NotasServico.ini")
   iPosColTelefone = LerArquivoINI("Notas Fiscais", "PosColTelefone", CaminhoINI & "\NotasServico.ini")

 ' CNPJ e IE
   iPosLinCNPJ = LerArquivoINI("Notas Fiscais", "PosLinCNPJ", CaminhoINI & "\NotasServico.ini")
   iPosColCNPJ = LerArquivoINI("Notas Fiscais", "PosColCNPJ", CaminhoINI & "\NotasServico.ini")
   iPosColIE = LerArquivoINI("Notas Fiscais", "PosColIE", CaminhoINI & "\NotasServico.ini")

 ' Cobranca
   iPosLinCobranca = LerArquivoINI("Notas Fiscais", "PosLinCobranca", CaminhoINI & "\NotasServico.ini")
   iPosColDocBol1 = LerArquivoINI("Notas Fiscais", "PosColDocBol1", CaminhoINI & "\NotasServico.ini")
   iPosColVenBol1 = LerArquivoINI("Notas Fiscais", "PosColVenBol1", CaminhoINI & "\NotasServico.ini")
   iPosColValBol1 = LerArquivoINI("Notas Fiscais", "PosColValBol1", CaminhoINI & "\NotasServico.ini")
   iPosColDocBol2 = LerArquivoINI("Notas Fiscais", "PosColDocBol2", CaminhoINI & "\NotasServico.ini")
   iPosColVenBol2 = LerArquivoINI("Notas Fiscais", "PosColVenBol2", CaminhoINI & "\NotasServico.ini")
   iPosColValBol2 = LerArquivoINI("Notas Fiscais", "PosColValBol2", CaminhoINI & "\NotasServico.ini")
   iPosColDocBol3 = LerArquivoINI("Notas Fiscais", "PosColDocBol3", CaminhoINI & "\NotasServico.ini")
   iPosColVenBol3 = LerArquivoINI("Notas Fiscais", "PosColVenBol3", CaminhoINI & "\NotasServico.ini")
   iPosColValBol3 = LerArquivoINI("Notas Fiscais", "PosColValBol3", CaminhoINI & "\NotasServico.ini")
   iPosColDocBol4 = LerArquivoINI("Notas Fiscais", "PosColDocBol4", CaminhoINI & "\NotasServico.ini")
   iPosColVenBol4 = LerArquivoINI("Notas Fiscais", "PosColVenBol4", CaminhoINI & "\NotasServico.ini")
   iPosColValBol4 = LerArquivoINI("Notas Fiscais", "PosColValBol4", CaminhoINI & "\NotasServico.ini")

 ' Produtos
   iPosLinProdutos = LerArquivoINI("Notas Fiscais", "PosLinProdutos", CaminhoINI & "\NotasServico.ini")
   iPosColQuantidade = LerArquivoINI("Notas Fiscais", "PosColQuantidade", CaminhoINI & "\NotasServico.ini")
   iPosColUnidade = LerArquivoINI("Notas Fiscais", "PosColUnidade", CaminhoINI & "\NotasServico.ini")
   iPosColDescricao = LerArquivoINI("Notas Fiscais", "PosColDescricao", CaminhoINI & "\NotasServico.ini")
   iPosColVlUnitario = LerArquivoINI("Notas Fiscais", "PosColVlUnitario", CaminhoINI & "\NotasServico.ini")
   iPosColVlLiquido = LerArquivoINI("Notas Fiscais", "PosColVlLiquido", CaminhoINI & "\NotasServico.ini")

 ' Valor Total da Nota
   iPosLinTotalNota = LerArquivoINI("Notas Fiscais", "PosLinTotalNota", CaminhoINI & "\NotasServico.ini")
   iPosColTotalNota = LerArquivoINI("Notas Fiscais", "PosColTotalNota", CaminhoINI & "\NotasServico.ini")

 ' Numero Final
   iPosLinNumeroFim = LerArquivoINI("Notas Fiscais", "PosLinNumeroFim", CaminhoINI & "\NotasServico.ini")
   iPosColNumeroFim = LerArquivoINI("Notas Fiscais", "PosColNumeroFim", CaminhoINI & "\NotasServico.ini")
   iPosLinProximaNota = LerArquivoINI("Notas Fiscais", "PosLinProximaNota", CaminhoINI & "\NotasServico.ini")

   If rsNFe!Impressa Then
      Beep
      MsgBox "Nota Fiscal j� Impressa", vbExclamation, "Aviso"
   End If

   cdlgImprimirNota.CancelError = True

   Registro_Selecionado = False
''*   VisualizarImpressao
''*   If Registro_Selecionado Then

      Open LerArquivoINI("Impressoras", "Notas", CaminhoINI & "\System.ini") For Output As #1
'      Open CaminhoINI & "\teste.txt" For Output As #1

      Set rsClientes = cnSistema.Execute("Select * From Clientes Where idCliente = " & rsNFe!idCliente)

    ' Emissao e Numero
      Print #1, Chr(27) & "x0"; Chr(15)
      SaltarLinha (iPosLinEmissao)
      Print #1, Tab(iPosColEmissao); rsNFe!DataEmissao; Tab(iPosColNumero); StrZero(rsNFe!Numero, 6)
      SaltarLinha (iPosLinCliente)

    ' Cliente
      Print #1, Tab(iPosColCliente); RemoveAcentos(IIf(Not IsNull(rsClientes!Nome), rsClientes!Nome, ""))
      SaltarLinha (iPosLinEndereco)

    ' Endere�o
      Print #1, Tab(iPosColEndereco); RemoveAcentos(IIf(Not IsNull(rsClientes!Endereco), rsClientes!Endereco, ""))
      SaltarLinha (iPosLinCidade)

    ' Cidade, UF e Telefone
      Print #1, Tab(iPosColCidade); RemoveAcentos(IIf(Not IsNull(rsClientes!Cidade), rsClientes!Cidade, "")); Tab(iPosColUF); IIf(Not IsNull(rsClientes!UF), rsClientes!UF, ""); Tab(iPosColTelefone); IIf(Not IsNull(rsClientes!Telefone1), rsClientes!Telefone1, "")
      SaltarLinha (iPosLinCNPJ)

    ' CNPJ e IE
      Print #1, Tab(iPosColCNPJ); IIf(Not IsNull(rsClientes!CNPJ_CPF), rsClientes!CNPJ_CPF, ""); Tab(iPosColIE); IIf(Not IsNull(rsClientes!IE_CI), rsClientes!IE_CI, "")
      SaltarLinha (iPosLinCobranca)

    ' Cobranca
      Dim ContCob As Integer
      ContCob = 1

      Dim DocBol1 As String
      Dim VenBol1 As String
      Dim ValBol1 As String
      Dim DocBol2 As String
      Dim VenBol2 As String
      Dim ValBol2 As String
      Dim DocBol3 As String
      Dim VenBol3 As String
      Dim ValBol3 As String
      Dim DocBol4 As String
      Dim VenBol4 As String
      Dim ValBol4 As String

      Set rsNFeBoletos = cnSistema.Execute("SELECT * FROM NFeBoletos WHERE idNFe = " & rsNFe!idNFe)
      Do While Not rsNFeBoletos.EOF
         If ContCob = 1 Then
            DocBol1 = rsNFeBoletos!Numero
            VenBol1 = FormataTXT(Format(rsNFeBoletos!Valor, "##,##0.00"), 2.1, 10)
            ValBol1 = rsNFeBoletos!Vencimento

         ElseIf ContCob = 2 Then
            DocBol2 = rsNFeBoletos!Numero
            VenBol2 = FormataTXT(Format(rsNFeBoletos!Valor, "##,##0.00"), 2.1, 10)
            ValBol2 = rsNFeBoletos!Vencimento

         ElseIf ContCob = 3 Then
            DocBol3 = rsNFeBoletos!Numero
            VenBol3 = FormataTXT(Format(rsNFeBoletos!Valor, "##,##0.00"), 2.1, 10)
            ValBol3 = rsNFeBoletos!Vencimento

         ElseIf ContCob = 4 Then
            DocBol4 = rsNFeBoletos!Numero
            VenBol4 = FormataTXT(Format(rsNFeBoletos!Valor, "##,##0.00"), 2.1, 10)
            ValBol4 = rsNFeBoletos!Vencimento
         End If

         rsNFeBoletos.MoveNext
         ContCob = ContCob + 1
      Loop

      Print #1, Tab(iPosColDocBol1); DocBol1; _
                Tab(iPosColVenBol1); VenBol1; _
                Tab(iPosColValBol1); ValBol1; _
                Tab(iPosColDocBol2); DocBol2; _
                Tab(iPosColVenBol2); VenBol2; _
                Tab(iPosColValBol2); ValBol2; _
                Tab(iPosColDocBol3); DocBol3; _
                Tab(iPosColVenBol3); VenBol3; _
                Tab(iPosColValBol3); ValBol3; _
                Tab(iPosColDocBol4); DocBol4; _
                Tab(iPosColVenBol4); VenBol4; _
                Tab(iPosColValBol4); ValBol4

      SaltarLinha (iPosLinProdutos)

    ' Produtos
      Dim cValorBruto As Currency
      Dim cValorDesconto As Currency
      Dim cValorBonificacao As Currency
      Dim cValorLiquido As Currency

      Dim iLinhasProdutos As Integer
      iLinhasProdutos = LerArquivoINI("Notas Fiscais", "ItensNota", CaminhoINI & "\NotasServico.ini")
      Set rsNFeItens = cnSistema.Execute("SELECT NFeItens.idNFe, NFeItens.idProduto, NFeItens.idUnidade, Produtos.Descricao, NFeItens.idClassificacaoFiscal, NFeItens.idSituacaoTributaria, NFeItens.Data, NFeItens.Quantidade, NFeItens.ValorUnitario, NFeItens.Desconto, NFeItens.IPI, NFeItens.ICMS, NFeItens.DescricaoComplementar " & _
                                                  "FROM NFeItens INNER JOIN Produtos ON NFeItens.idProduto = Produtos.idProduto " & _
                                                  "Where NFeItens.idNFe = " & rsNFe!idNFe & " ORDER BY Produtos.Descricao")

      Do While Not rsNFeItens.EOF
         Set rsProdutos = cnSistema.Execute("SELECT * FROM Produtos WHERE idProduto = " & rsNFeItens!idProduto)
         Set rsUnidadesMedida = cnSistema.Execute("SELECT * FROM UnidadesMedida WHERE idUnidadeMedida = " & rsNFeItens!idUnidade)
         Dim sDescricao As String, sUnidade As String
         sDescricao = RemoveAcentos(Mid(Trim(rsProdutos!Descricao), 1, 50) & " " & Trim(rsNFeItens!DescricaoComplementar))
         If Not rsUnidadesMedida.EOF Then
            sUnidade = rsUnidadesMedida!Sigla
         Else
            sUnidade = " "
         End If

         cValorBruto = (rsNFeItens!Quantidade * rsNFeItens!ValorUnitario)
         cValorDesconto = (((rsNFeItens!Quantidade * rsNFeItens!ValorUnitario) * rsNFeItens!Desconto) / 100)
         cValorBonificacao = (((cValorBruto - cValorDesconto) * rsNFe!Bonificacao) / 100)
         cValorLiquido = (cValorBruto - cValorDesconto - cValorBonificacao)

         Print #1, Tab(iPosColQuantidade); FormataTXT(Format(rsNFeItens!Quantidade, "##,##0.00"), 2.1, 10); _
                   Tab(iPosColUnidade); sUnidade; _
                   Tab(iPosColDescricao); sDescricao; _
                   Tab(iPosColVlUnitario); FormataTXT(Format(rsNFeItens!ValorUnitario, "##,##0.00"), 2.1, 10); _
                   Tab(iPosColVlLiquido); FormataTXT(Format(cValorLiquido, "##,##0.00"), 2.1, 12)

         iLinhasProdutos = iLinhasProdutos - 1
         rsNFeItens.MoveNext
      Loop

      ' Informa��es do Corpo da Nota
      SaltarLinha (iLinhasProdutos + iPosLinTotalNota)

'      Print #1, Tab(iPosColInfoCN); rsNFe!InformacoesCorpo
'      SaltarLinha (iPosLinBase)

    ' Total
      Set rsTemp = cnSistema.Execute("Select * From TotalNFe Where Numero = " & mskNumero.text)
      If Not rsTemp.EOF Then
         Print #1, Tab(iPosColTotalNota); Format(rsTemp!Total, "##,##0.00")
         SaltarLinha (iPosLinNumeroFim)
      Else
         Print #1, Tab(iPosColTotalNota); Format(0, "##,##0.00")
         SaltarLinha (iPosLinNumeroFim)
      End If

    ' Numero da Nota
      Print #1, Tab(iPosColNumeroFim); StrZero(rsNFe!Numero, 6) & Chr(18)
      SaltarLinha (iPosLinProximaNota)

      Close #1
      cnSistema.Execute "Update NFe set " & _
            "Impressa = " & True & " " & _
            "Where idNFe = " & rsNFe!idNFe
''*   End If

End Sub

Private Sub VisualizarImpressao()
Dim Contador As Integer
   Screen.MousePointer = vbHourglass
   frmVisualizaImpressao.lvwDados.ColumnHeaders.Clear
   frmVisualizaImpressao.lvwDados.ColumnHeaders.Add , , "Nota Eletr�nica", 11450
   frmVisualizaImpressao.lvwDados.ListItems.Clear

   Set rsCFOPs = cnSistema.Execute("Select * From CFOPs Where idCFOP = " & rsNFe!idCFOP)
   Set rsClientes = cnSistema.Execute("Select * From Clientes Where idCliente = " & rsNFe!idCliente)
   Set rsNaturezasOperacao = cnSistema.Execute("Select * From NaturezasOperacao Where idNaturezaOperacao = " & rsNFe!idNaturezaOperacao)

'  Numero da Nota
   Contador = 1

   Set ItemList = frmVisualizaImpressao.lvwDados.ListItems.Add(, "R" & CStr(Contador), Space(99) & StrZero(rsNFe!Numero, 6))

'  Codigo Fiscal de Opera��o
   Contador = Contador + 1
   Set ItemList = frmVisualizaImpressao.lvwDados.ListItems.Add(, "R" & CStr(Contador), rsNaturezasOperacao!Descricao & Space(20) & rsCFOPs!CFOP)

'  Cliente, CNPJ e Emiss�o
   Contador = Contador + 1
   Set ItemList = frmVisualizaImpressao.lvwDados.ListItems.Add(, "R" & CStr(Contador), rsClientes!Codigo & "-" & rsClientes!Nome & Space(10) & rsClientes!CNPJ_CPF & Space(26) & rsNFe!DataEmissao)

'  Endere�o, Bairro e CEP
   Contador = Contador + 1
   Set ItemList = frmVisualizaImpressao.lvwDados.ListItems.Add(, "R" & CStr(Contador), rsClientes!Endereco & Space(10) & rsClientes!Bairro & Space(10) & rsClientes!CEP)

'  Cidade, Telefone, UF e Inscri��o Estadual
   Contador = Contador + 1
   Set ItemList = frmVisualizaImpressao.lvwDados.ListItems.Add(, "R" & CStr(Contador), rsClientes!Cidade & Space(10) & rsClientes!Telefone1 & Space(10) & rsClientes!UF & Space(10) & rsClientes!IE_CI)

'  Documento
   Contador = Contador + 1
   Set ItemList = frmVisualizaImpressao.lvwDados.ListItems.Add(, "R" & CStr(Contador), rsNFe!Documento & Space(10) & cmbFormaPagamento.text & Space(56) & rsNFe!DataVencimento)

   Contador = Contador + 1
   Set ItemList = frmVisualizaImpressao.lvwDados.ListItems.Add(, "R" & CStr(Contador), "")

'  Produtos

   Dim cValorBruto As Currency
   Dim cValorDesconto As Currency
   Dim cValorBonificacao As Currency
   Dim cValorLiquido As Currency

   Dim ContadorProdutos As Integer
   ContadorProdutos = 0
   Set rsNFeItens = cnSistema.Execute("SELECT NFeItens.idNFe, NFeItens.idProduto,NFeItens.idUnidade, Produtos.Descricao, NFeItens.idClassificacaoFiscal, NFeItens.idSituacaoTributaria, NFeItens.Data, NFeItens.Quantidade, NFeItens.ValorUnitario, NFeItens.Desconto, NFeItens.IPI, NFeItens.ICMS, NFeItens.DescricaoComplementar " & _
                                               "FROM NFeItens INNER JOIN Produtos ON NFeItens.idProduto = Produtos.idProduto " & _
                                               "Where NFeItens.idNFe = " & rsNFe!idNFe & " ORDER BY Produtos.Descricao")

   Do While Not rsNFeItens.EOF
      Contador = Contador + 1
      ContadorProdutos = ContadorProdutos + 1

      Set rsProdutos = cnSistema.Execute("SELECT * FROM Produtos WHERE idProduto = " & rsNFeItens!idProduto)
      Set rsUnidadesMedida = cnSistema.Execute("SELECT * FROM UnidadesMedida WHERE idUnidadeMedida = " & rsNFeItens!idUnidade)
      Set rsSituacoesTributarias = cnSistema.Execute("SELECT * FROM SituacoesTributarias WHERE idSituacaoTributaria = " & rsNFeItens!idSituacaoTributaria)
      Dim sDescricao As String, sUnidade As String, sCST As String
      sDescricao = Mid(Trim(rsProdutos!Descricao), 1, 50) & " " & Trim(rsNFeItens!DescricaoComplementar)
      If Not rsUnidadesMedida.EOF Then
         sUnidade = rsUnidadesMedida!Sigla
      Else
         sUnidade = " "
      End If

      If Not rsSituacoesTributarias.EOF Then
         sCST = rsSituacoesTributarias!Codigo
      Else
         sCST = " "
      End If

      cValorBruto = (rsNFeItens!Quantidade * rsNFeItens!ValorUnitario)
      cValorDesconto = (((rsNFeItens!Quantidade * rsNFeItens!ValorUnitario) * rsNFeItens!Desconto) / 100)
      cValorBonificacao = (((cValorBruto - cValorDesconto) * rsNFe!Bonificacao) / 100)
      cValorLiquido = (cValorBruto - cValorDesconto - cValorBonificacao)

      Set ItemList = frmVisualizaImpressao.lvwDados.ListItems.Add(, "R" & CStr(Contador), rsProdutos!Codigo & _
                Space(2) & FormataTXT(sDescricao, 1, 44) & _
                Space(3) & sCST & _
                Space(3) & sUnidade & _
                Space(3) & FormataTXT(Format(rsNFeItens!Quantidade, "##,##0.00"), 2.1, 10) & _
                Space(3) & FormataTXT(Format(rsNFeItens!ValorUnitario, "##,##0.00"), 2.1, 10) & _
                Space(3) & FormataTXT(Format(rsNFeItens!Desconto, "#0.00"), 2.1, 4) & IIf(rsNFe!Bonificacao = 0, "", "/" & FormataTXT(Format(rsNFe!Bonificacao, "#0.00"), 2.1, 4)) & _
                Space(3) & FormataTXT(Format(cValorLiquido, "##,##0.00"), 2.1, 10) & _
                Space(3) & Format(rsNFeItens!ICMS, "##0"))

      rsNFeItens.MoveNext
   Loop

   Contador = Contador + 1
   Set ItemList = frmVisualizaImpressao.lvwDados.ListItems.Add(, "R" & CStr(Contador), "")

   Contador = Contador + 1
'   Set ItemList = frmVisualizaImpressao.lvwDados.ListItems.Add(, "R" & CStr(Contador), rsNFe!InformacoesCorpo)

   ' Informa��es do Corpo da Nota
   Dim sInformacoesCorpo1 As String, sInformacoesCorpo2 As String, sInformacoesCorpo3 As String
   If Len(rsNFe!InformacoesCorpo) >= 240 Then
      sInformacoesCorpo1 = RemoveAcentos(Mid(rsNFe!InformacoesCorpo, 1, 120))
      sInformacoesCorpo2 = RemoveAcentos(Mid(rsNFe!InformacoesCorpo, 121, 120))
      sInformacoesCorpo3 = RemoveAcentos(Mid(rsNFe!InformacoesCorpo, 241, 120))

      Contador = Contador + 1
      Set ItemList = frmVisualizaImpressao.lvwDados.ListItems.Add(, "R" & CStr(Contador), sInformacoesCorpo1)
      Contador = Contador + 1
      Set ItemList = frmVisualizaImpressao.lvwDados.ListItems.Add(, "R" & CStr(Contador), sInformacoesCorpo2)
      Contador = Contador + 1
      Set ItemList = frmVisualizaImpressao.lvwDados.ListItems.Add(, "R" & CStr(Contador), sInformacoesCorpo3)
   Else
      If Len(rsNFe!InformacoesCorpo) >= 120 Then
         sInformacoesCorpo1 = RemoveAcentos(Mid(rsNFe!InformacoesCorpo, 1, 120))
         sInformacoesCorpo2 = RemoveAcentos(Mid(rsNFe!InformacoesCorpo, 121, 120))

         Contador = Contador + 1
         Set ItemList = frmVisualizaImpressao.lvwDados.ListItems.Add(, "R" & CStr(Contador), sInformacoesCorpo1)
         Contador = Contador + 1
         Set ItemList = frmVisualizaImpressao.lvwDados.ListItems.Add(, "R" & CStr(Contador), sInformacoesCorpo2)
      Else
         If Len(rsNFe!InformacoesCorpo) >= 1 Then
            sInformacoesCorpo1 = RemoveAcentos(rsNFe!InformacoesCorpo)

            Contador = Contador + 1
            Set ItemList = frmVisualizaImpressao.lvwDados.ListItems.Add(, "R" & CStr(Contador), sInformacoesCorpo1)
         End If
      End If
   End If

   Contador = Contador + 1
   Set ItemList = frmVisualizaImpressao.lvwDados.ListItems.Add(, "R" & CStr(Contador), "")

'  Total

   Set rsTemp = cnSistema.Execute("Select * From TotalNFe Where Numero = " & mskNumero.text)
   Dim dBaseCalculo As Double, dValorICMS As Double
   If rsTemp!ValorICMS > 0 Then
      dBaseCalculo = rsTemp!BaseCalculo
      dValorICMS = rsTemp!ValorICMS
   Else
      dBaseCalculo = 0
      dValorICMS = 0
   End If

   If Not rsTemp.EOF Then
      Contador = Contador + 1
      Set ItemList = frmVisualizaImpressao.lvwDados.ListItems.Add(, "R" & CStr(Contador), "Total de Produtos: " & StrZero(ContadorProdutos, 3) & Space(31) & Format(dBaseCalculo, "##,##0.00") & Space(10) & Format(dValorICMS, "##,##0.00") & Space(10) & Format(rsTemp!Total, "##,##0.00"))

      Contador = Contador + 1
      Set ItemList = frmVisualizaImpressao.lvwDados.ListItems.Add(, "R" & CStr(Contador), Space(90) & Format(rsTemp!Total, "##,##0.00"))
   Else
      Contador = Contador + 1
'      Set ItemList = frmVisualizaImpressao.lvwDados.ListItems.Add(, "R" & CStr(Contador), Space(71) & Format(rsNFe!BaseCalculoICMS, "#0.00") & Space(10) & Format(rsNFe!ValorICMS, "#0.00") & Space(10) & Format(0, "#0.00"))
      Set ItemList = frmVisualizaImpressao.lvwDados.ListItems.Add(, "R" & CStr(Contador), Space(71) & Format(dBaseCalculo, "##,##0.00") & Space(10) & Format(dValorICMS, "##,##0.00") & Space(10) & Format(0, "##,##0.00"))

      Contador = Contador + 1
      Set ItemList = frmVisualizaImpressao.lvwDados.ListItems.Add(, "R" & CStr(Contador), Space(99) & Format(0, "##,##0.00"))
   End If

'  Transportador
   Set rsTransportador = cnSistema.Execute("SELECT * FROM Transportadores WHERE idTransportador = " & rsNFe!idTransportador)
   If Not rsTransportador.EOF Then
      Contador = Contador + 1
      Set ItemList = frmVisualizaImpressao.lvwDados.ListItems.Add(, "R" & CStr(Contador), rsTransportador!Nome & Space(10) & rsNFe!PlacaVeiculo & Space(10) & rsTransportador!UFPlaca & Space(10) & rsTransportador!CNPJ_CPF)

      Contador = Contador + 1
      Set ItemList = frmVisualizaImpressao.lvwDados.ListItems.Add(, "R" & CStr(Contador), rsTransportador!Endereco & Space(10) & rsTransportador!Cidade & Space(10) & rsTransportador!UF & Space(10) & rsTransportador!IE_CI)
   Else
      Contador = Contador + 2
   End If
   Contador = Contador + 1
   Set ItemList = frmVisualizaImpressao.lvwDados.ListItems.Add(, "R" & CStr(Contador), rsNFe!VolumeQuantidade & Space(10) & rsNFe!VolumeMarca & Space(10) & rsNFe!VolumeNumero & Space(10) & Format(rsNFe!VolumePesoBruto, "##,##0") & Space(10) & Format(rsNFe!VolumePesoLiquido, "##,##0"))

   ' Volume
   Contador = Contador + 1
   Set ItemList = frmVisualizaImpressao.lvwDados.ListItems.Add(, "R" & CStr(Contador), Space(5) & rsNFe!VolumeQuantidade & Space(5) & rsNFe!VolumeEspecie & Space(5) & rsNFe!VolumeMarca & Space(5) & rsNFe!VolumeNumero & Space(5) & Format(rsTemp!PesoBruto, "##,##0.00") & Space(5) & Format(rsTemp!PesoLiquido, "##,##0.00"))
   Contador = Contador + 1
   Set ItemList = frmVisualizaImpressao.lvwDados.ListItems.Add(, "R" & CStr(Contador), "")

   ' Dados Adicionais
   Dim sDadosAdicionais1 As String, sDadosAdicionais2 As String, sDadosAdicionais3 As String
   If Len(rsNFe!DadosAdicionais) >= 120 Then
      sDadosAdicionais1 = RemoveAcentos(Mid(rsNFe!DadosAdicionais, 1, 60))
      sDadosAdicionais2 = RemoveAcentos(Mid(rsNFe!DadosAdicionais, 61, 60))
      sDadosAdicionais3 = RemoveAcentos(Mid(rsNFe!DadosAdicionais, 121, 60))

      Contador = Contador + 1
      Set ItemList = frmVisualizaImpressao.lvwDados.ListItems.Add(, "R" & CStr(Contador), sDadosAdicionais1)
      Contador = Contador + 1
      Set ItemList = frmVisualizaImpressao.lvwDados.ListItems.Add(, "R" & CStr(Contador), sDadosAdicionais2)
      Contador = Contador + 1
      Set ItemList = frmVisualizaImpressao.lvwDados.ListItems.Add(, "R" & CStr(Contador), sDadosAdicionais3)
      Contador = Contador + 1
      Set ItemList = frmVisualizaImpressao.lvwDados.ListItems.Add(, "R" & CStr(Contador), "")
   Else
      If Len(rsNFe!DadosAdicionais) >= 60 Then
         sDadosAdicionais1 = RemoveAcentos(Mid(rsNFe!DadosAdicionais, 1, 60))
         sDadosAdicionais2 = RemoveAcentos(Mid(rsNFe!DadosAdicionais, 61, 60))

         Contador = Contador + 1
         Set ItemList = frmVisualizaImpressao.lvwDados.ListItems.Add(, "R" & CStr(Contador), sDadosAdicionais1)
         Contador = Contador + 1
         Set ItemList = frmVisualizaImpressao.lvwDados.ListItems.Add(, "R" & CStr(Contador), sDadosAdicionais2)
         Contador = Contador + 1
         Set ItemList = frmVisualizaImpressao.lvwDados.ListItems.Add(, "R" & CStr(Contador), "")
      Else
         If Len(rsNFe!DadosAdicionais) >= 1 Then
            sDadosAdicionais1 = RemoveAcentos(rsNFe!DadosAdicionais)

            Contador = Contador + 1
            Set ItemList = frmVisualizaImpressao.lvwDados.ListItems.Add(, "R" & CStr(Contador), sDadosAdicionais1)
            Contador = Contador + 1
            Set ItemList = frmVisualizaImpressao.lvwDados.ListItems.Add(, "R" & CStr(Contador), "")
         End If
      End If
   End If

'  Numero da Nota
   Contador = Contador + 1
   Set ItemList = frmVisualizaImpressao.lvwDados.ListItems.Add(, "R" & CStr(Contador), Space(99) & StrZero(rsNFe!Numero, 6))

   Screen.MousePointer = vbDefault
   frmNFeG.Caption = "Visualiza Impress�o de Notas Eletr�nicas"
   frmVisualizaImpressao.Show vbModal
End Sub

Sub ImprimirBoleto()
Dim Contador As Integer
Dim Total_Registros As Integer
Dim PaginaInicial, Paginafinal, NumeroCopias, i
Dim B_Local As String
Dim B_Vencimento As String
Dim B_Emissao As Date
Dim B_Documento As String
Dim B_Valor As Double
Dim B_Taxa As Double
Dim B_Inst1 As String
Dim B_Inst2 As String
Dim B_Inst3 As String
Dim B_Inst4 As String
Dim B_Inst5 As String
Dim B_Inst6 As String
Dim B_CliCNPJ As String
Dim B_Endereco As String

Dim iPosLinLocal As Integer
Dim iPosColLocal As Integer
Dim iPosColVencimento As Integer
Dim iPosLinEmissao As Integer
Dim iPosColEmissao As Integer
Dim iPosColDocumento As Integer
Dim iPosLinValor As Integer
Dim iPosColValor As Integer
Dim iPosLinInstrucoes As Integer
Dim iPosColInstrucoes As Integer
Dim iPosLinCliente As Integer
Dim iPosColCliente As Integer
Dim iPosLinProximo As Integer

'''     Total_Registros = Registros(cnSistema, "ContasBancarias")
     Total_Registros = Registros2("ContasBancarias")
     If Total_Registros = 0 Then
        MsgBox "N�o existe nenhum Registro na Tabela", vbOKOnly, "Visualiza"
        Exit Sub
     End If
     Contador = 1
     Status = 4
     Screen.MousePointer = vbHourglass
     frmVisualiza.lvwDados.ColumnHeaders.Clear
     frmVisualiza.lvwDados.ColumnHeaders.Add , , "Banco", 6000
     rsContasBancarias.MoveFirst
     frmVisualiza.lvwDados.ListItems.Clear
     Do While Not rsContasBancarias.EOF
        frmNFeG.Caption = "Processando " & StrZero(Contador, 8) & " de " & StrZero(Total_Registros, 8)
        Set ItemList = frmVisualiza.lvwDados.ListItems.Add(, "R" & CStr(rsContasBancarias!idContaBancaria), rsContasBancarias!Descricao)
        rsContasBancarias.MoveNext
        Contador = Contador + 1
     Loop
     rsContasBancarias.MoveFirst
     Registro_Selecionado = False
     Screen.MousePointer = vbDefault
     frmNFeG.Caption = "Notas de Saida Eletr�nicas"
     frmVisualiza.Show vbModal
     If Registro_Selecionado Then
        rsContasBancarias.Find "idContaBancaria = " & Val(Mid(frmVisualiza.lvwDados.SelectedItem.Key, 2, Len(frmVisualiza.lvwDados.SelectedItem.Key)))
     End If

   ' Mostra a janela para impressora

   ' Captura os valores definidos pelo usu�rio na janela

     iPosLinLocal = rsContasBancarias!PosicaoLocalLinha
     iPosColLocal = rsContasBancarias!PosicaoLocalColuna
     iPosColVencimento = rsContasBancarias!PosicaoVencimentoColuna
     iPosLinEmissao = rsContasBancarias!PosicaoEmissaoLinha
     iPosColEmissao = rsContasBancarias!PosicaoEmissaoColuna
     iPosColDocumento = rsContasBancarias!PosicaoDocumentoColuna
     iPosLinValor = rsContasBancarias!PosicaoValorLinha
     iPosColValor = rsContasBancarias!PosicaoValorColuna
     iPosLinInstrucoes = rsContasBancarias!PosicaoInstrucoesLinha
     iPosColInstrucoes = rsContasBancarias!PosicaoInstrucoesColuna
     iPosLinCliente = rsContasBancarias!PosicaoClienteLinha
     iPosColCliente = rsContasBancarias!PosicaoClienteColuna
     iPosLinProximo = rsContasBancarias!PosicaoProximoLinha

     Set rsTemp = cnSistema.Execute("SELECT * From NFe WHERE NFe.Documento = '" & rsNFe!Documento & "'")

     Do While Not rsTemp.EOF
        B_Documento = B_Documento & StrZero(rsTemp!Numero, 6) & "/"

        Set rsTemp2 = cnSistema.Execute("Select * From TotalNFe Where Numero = " & rsTemp!Numero)
        If Not rsTemp2.EOF Then
           B_Valor = B_Valor + rsTemp2!Total
        Else
           B_Valor = B_Valor + 0
        End If
        rsTemp.MoveNext
     Loop

     NumeroCopias = InputBox("Digite a Quantidade de Parcelas", "Quantidade", 1)
     If Val(NumeroCopias) = 0 Then Exit Sub

     B_Local = rsContasBancarias!LocalPagamento
     B_Emissao = rsNFe!DataEmissao
     B_Valor = B_Valor / NumeroCopias
     B_Taxa = rsContasBancarias!TaxaBancaria
     B_Inst1 = rsContasBancarias!Instrucoes1
     B_Inst2 = rsContasBancarias!Instrucoes2
     B_Inst3 = rsContasBancarias!Instrucoes3
     B_Inst4 = rsContasBancarias!Instrucoes4
     B_Inst5 = rsContasBancarias!Instrucoes5
     B_CliCNPJ = Trim(rsClientes!Nome) & " CPF/CNPJ: " & rsClientes!CNPJ_CPF
     B_Endereco = Trim(rsClientes!Endereco) & " " & Trim(rsClientes!Bairro) & " " & Trim(rsClientes!Cidade) & " " & Trim(rsClientes!CEP)

     Dim sMsg As String, sVencimento As String
     sMsg = "Digite o Vencimento"

     Open LerArquivoINI("Impressoras", "Boletos1", CaminhoINI & "\System.ini") For Output As #1

     For i = 1 To NumeroCopias
       ' Digitar o Vencimento
         Dim bValidaData As Boolean
         bValidaData = False
         Do While Not bValidaData
            B_Vencimento = InputBox("Digite o Vencimento", "Vencimento", "")
            If Val(B_Vencimento) = 0 Then
               Close #1
               Exit Sub
            End If

            If FData(B_Vencimento) Then
               bValidaData = True
            End If
         Loop

       ' Local de Pagamento e Vencimento
         SaltarLinha (iPosLinLocal)
         Print #1, Chr(27) & Chr(15); Tab(iPosColLocal); B_Local; Tab(iPosColVencimento); B_Vencimento

       ' Emiss�o e Documento
         SaltarLinha (iPosLinEmissao)
         Print #1, Tab(iPosColEmissao); B_Emissao; Tab(iPosColDocumento); B_Documento; Chr(27) & Chr(8)

       ' Valor
         SaltarLinha (iPosLinValor)
         Print #1, Tab(iPosColValor); Format(B_Valor, "Standard")

       ' Instru��es
         SaltarLinha (iPosLinInstrucoes)
         Print #1, Tab(iPosColInstrucoes); B_Inst1
         Print #1, Tab(iPosColInstrucoes); B_Inst2
         Print #1, Tab(iPosColInstrucoes); B_Inst3
         Print #1, Tab(iPosColInstrucoes); B_Inst4
         Print #1, Tab(iPosColInstrucoes); B_Inst5; Chr(27) & Chr(15)

       ' Instru��es
         SaltarLinha (iPosLinCliente)
         Print #1, Tab(iPosColCliente); B_CliCNPJ
         Print #1, Tab(iPosColCliente); B_Endereco; Chr(27) & Chr(8)

       ' Pr�ximo
         SaltarLinha (iPosLinProximo + 1)
     Next
     Close #1

End Sub

Sub CopiarNota()

   If MsgBox("Confirma C�pia", vbYesNo + vbQuestion, "Inclus�o") = vbYes Then
      RegistroAtual = IIf(rsNFe.EOF, 0, rsNFe!idNFe)
      Dim iNumero As Long

   '  Notas Eletronicas
      rsNFe.MoveLast
      iNumero = rsNFe!Numero + 1
      If RegistroAtual <> 0 Then
         rsNFe.MoveFirst
         rsNFe.Find "idNFe = " & RegistroAtual
      End If

      cnSistema.Execute "Insert Into NFe (Numero,Cupom,idCliente,idNaturezaOperacao,idCFOP,DadosAdicionais,DataEmissao,DataVencimento,Hora,BaseCalculoICMS,ValorICMS,ValorFrete,ValorTotalProdutos,BaseICMSSubstituicao,ValorICMSSubstituicao,OutrasDespesas,ValorTotalNota,idTransportador,FreteConta,PlacaVeiculo,VolumeQuantidade,VolumeMarca,VolumeNumero,VolumePesoBruto,VolumePesoLiquido,InformacoesCorpo,idFormaPagamento,DescontoGeral,Bonificacao,Documento,Observacao) " & _
               "Values (" & iNumero & "," & rsNFe!Cupom & "," & rsNFe!idCliente & "," & rsNFe!idNaturezaOperacao & "," & rsNFe!idCFOP & ",'" & rsNFe!DadosAdicionais & "','" & rsNFe!DataEmissao & "','" & rsNFe!DataVencimento & "','" & rsNFe!Hora & "','" & rsNFe!BaseCalculoICMS & "','" & rsNFe!ValorICMS & "'," & _
                       "'" & rsNFe!ValorFrete & "','" & rsNFe!ValorTotalProdutos & "','" & rsNFe!BaseICMSSubstituicao & "','" & rsNFe!ValorICMSSubstituicao & "','" & rsNFe!OutrasDespesas & "'," & _
                       "'" & rsNFe!ValorTotalNota & "'," & rsNFe!idTransportador & "," & rsNFe!FreteConta & ",'" & rsNFe!PlacaVeiculo & "','" & rsNFe!VolumeQuantidade & "','" & rsNFe!VolumeMarca & "','" & rsNFe!VolumeNumero & "','" & rsNFe!VolumePesoBruto & "','" & rsNFe!VolumePesoLiquido & "','" & rsNFe!InformacoesCorpo & "'," & 0 & _
                       "," & rsNFe!idFormaPagamento & ",'" & rsNFe!DescontoGeral & "','" & rsNFe!Bonificacao & "','" & rsNFe!Documento & "','" & rsNFe!Observacao & "')"

      rsNFe.Requery

   '  Notas Eletronicas Itens
      rsNFe.MoveLast
      iNumero = rsNFe!idNFe
      If RegistroAtual <> 0 Then
         rsNFe.MoveFirst
         rsNFe.Find "idNFe = " & RegistroAtual
      End If

      Set rsNFeItens = cnSistema.Execute("Select * From NFeItens Where idNFe = " & rsNFe!idNFe)
      Do While Not rsNFeItens.EOF
         cnSistema.Execute "Insert Into NFeItens (idNFe,idProduto,Data,Quantidade,Desconto,ValorUnitario,ICMS,BaseReduzida,DescricaoComplementar) " & _
                  "Values (" & iNumero & "," & rsNFeItens!idProduto & ",'" & rsNFeItens!Data & _
                  "','" & rsNFeItens!Quantidade & "','" & rsNFeItens!Desconto & "','" & rsNFeItens!ValorUnitario & "','" & rsNFeItens!ICMS & _
                  "','" & rsNFeItens!BaseReduzida & "','" & rsNFeItens!DescricaoComplementar & "')"

         rsNFeItens.MoveNext
      Loop

      rsNFe.MoveLast
      Prencher_Campos
   End If
End Sub

Private Function ImprimirDanfe()
Dim sPercMargAdICMSST As String
Dim sUF As String
Dim sChaveAcesso As String
Dim sNaturezaOperacao As String
Dim sFormaPagamento As String
Dim sModelo As String
Dim sSerie As String
Dim sNumero As String
Dim sDataEmissao As String
Dim sDataSaida As String
Dim sTipoNF As String
Dim sCodigoMunicipio As String
Dim sFormatoDANFE As String
Dim sTipoEmissao As String
Dim sDVChaveAcesso As String
Dim sidAmbiente As String
Dim sFinalidade As String
Dim sProcessoEmissao As String
Dim sVersaoAplicativo As String
Dim sNomeArquivo As String

   If MsgBox("Confirma Emiss�o da Nota", vbYesNo + vbInformation, "Confirma��o") = vbNo Then
      Exit Function
   End If

'  Identifica��o do Arquivo
   Set rsEmpresa = cnSistema.Execute("Select * From Empresa")
   sNomeArquivo = StrZero(Val(mskNumero.text), 6) & "_" & RemoveCaracteres(rsEmpresa!CNPJ_CPF) & "_55_" & Mid(mskDataEmissao.text, 1, 2) & "_" & Mid(mskDataEmissao.text, 4, 2) & "_" & Mid(mskDataEmissao.text, 7, 4) & "-nfe.txt"

   Open "C:\NFe\XML\" & sNomeArquivo For Output As #1

'   Set rsNFe = cnSistema.Execute("Select * From NFe WHERE GeradaNFe=False")
   Print #1, "NOTAFISCAL|1"

   ' Cabecalho
   Print #1, "A|1.10|NFe"

   ' Identificadores
   Set rsNaturezasOperacao = cnSistema.Execute("Select * From NaturezasOperacao WHERE idNaturezaOperacao=" & rsNFe!idNaturezaOperacao)
   Set rsFormasPagamento = cnSistema.Execute("Select * From FormasPagamento WHERE idFormaPagamento=" & rsNFe!idFormaPagamento)
   Set rsCFOPs = cnSistema.Execute("Select * From CFOPs WHERE idCFOP=" & rsNFe!idCFOP)

   Set rsTemp = cnSistema.Execute("Select * From UFs WHERE Sigla=" & rsEmpresa!UF)
   If Not rsTemp.EOF Then
      sUF = rsTemp!Codigo ' Minas Gerais
   Else
      sUF = "  " ' Minas Gerais
   End If

   sChaveAcesso = ""
   sNaturezaOperacao = IIf(Not rsNaturezasOperacao.EOF, rsNaturezasOperacao!Descricao, "VENDA")
   If Not rsFormasPagamento.EOF Then
      If rsFormasPagamento!TipoPagamento <= 1 Then
         sFormaPagamento = "0"
      Else
         sFormaPagamento = "1"
      End If
   Else
      sFormaPagamento = "0"
   End If
   sModelo = "55"
   sSerie = "1"
   sNumero = rsNFe!Numero
   sDataEmissao = Format(rsNFe!DataEmissao, "yyyy-mm-dd")
   sDataSaida = Format(rsNFe!DataVencimento, "yyyy-mm-dd")
   sTipoNF = IIf(rsCFOPs!Tipo = 0, 0, 1)

   Set rsTemp = cnSistema.Execute("Select * From Municipios WHERE Nome=" & rsEmpresa!Nome)
   If Not rsTemp.EOF Then
      sCodigoMunicipio = RemoveCaracteres(rsTemp!Codigo)
   Else
      sCodigoMunicipio = "5212501"
   End If

   sFormatoDANFE = "1"
   sTipoEmissao = LerArquivoINI("Notas Fiscais", "TipoEmissaoNFe", CaminhoINI & "\System.ini")
   sDVChaveAcesso = ""
   sidAmbiente = LerArquivoINI("Notas Fiscais", "idAmbienteNFe", CaminhoINI & "\System.ini") ' Trocar pra 1 no Oficial
   sFinalidade = "1" ' NFe Normal
'   sProcessoEmissao = "3" ' Utilizando Software do Fisco
   sProcessoEmissao = "0" ' Utilizando Aplicativo do contribuinte
   sVersaoAplicativo = "1.4.1"
   '      sVersaoAplicativo = "TESTE 1.4.0"

   Print #1, "B|" & _
             sUF & "|" & _
             sChaveAcesso & "|" & _
             sNaturezaOperacao & "|" & _
             sFormaPagamento & "|" & _
             sModelo & "|" & _
             sSerie & "|" & _
             sNumero & "|" & _
             sDataEmissao & "|" & _
             sDataSaida & "|" & _
             sTipoNF & "|" & _
             sCodigoMunicipio & "|" & _
             sFormatoDANFE & "|" & _
             sTipoEmissao & "|" & _
             sDVChaveAcesso & "|" & _
             sidAmbiente & "|" & _
             sFinalidade & "|" & _
             sProcessoEmissao & "|" & _
             sVersaoAplicativo

   ' Emitente
   Dim sERazaoSocial As String
   Dim sEFantasia As String
   Dim sEIE As String
   Dim sEIEST As String
   Dim sEIM As String
   Dim sECNAE As String

   sERazaoSocial = rsEmpresa!Nome
   sEFantasia = ""
   sEIE = IIf(rsEmpresa!IE_CI <> "ISENTO", Trim(RemoveCaracteres(rsEmpresa!IE_CI)), "ISENTO")
   sEIEST = ""
   sEIM = ""
   sECNAE = ""

   Print #1, "C|" & _
             sERazaoSocial & "|" & _
             sEFantasia & "|" & _
             sEIE & "|" & _
             sEIEST & "|" & _
             sEIM & "|" & _
             sECNAE

   Dim sECNPJ As String
   sECNPJ = RemoveCaracteres(rsEmpresa!CNPJ_CPF)

   Print #1, "C02|" & _
             sECNPJ

   Dim sELogradouro As String
   Dim sENumero As String
   Dim sEComplemento As String
   Dim sEBairro As String
   Dim sECodigoMunicipio As String
   Dim sEMunicipio As String
   Dim sEUF As String
   Dim sECEP As String
   Dim sECodigoPais As String
   Dim sEPais As String
   Dim sETelefone As String

   sELogradouro = "RUA"
   sENumero = "25"
   sEComplemento = rsEmpresa!Endereco
   sEBairro = rsEmpresa!Bairro
   sECodigoMunicipio = "5212501"
   sEMunicipio = rsEmpresa!Cidade
   sEUF = rsEmpresa!UF
   sECEP = RemoveCaracteres(rsEmpresa!CEP)
   sECodigoPais = "1058"
   sEPais = "BRASIL"
   sETelefone = RemoveCaracteres(rsEmpresa!Telefone1)

   Print #1, "C05|" & _
             sELogradouro & "|" & _
             sENumero & "|" & _
             sEComplemento & "|" & _
             sEBairro & "|" & _
             sECodigoMunicipio & "|" & _
             sEMunicipio & "|" & _
             sEUF & "|" & _
             sECEP & "|" & _
             sECodigoPais & "|" & _
             sEPais & "|" & _
             sETelefone

   ' Destinatario
   Set rsClientes = cnSistema.Execute("Select * From Clientes WHERE idCliente=" & rsNFe!idCliente)

   Dim sDRazaoSocial As String
   Dim sDIE As String
   Dim sDISUF As String

   sDRazaoSocial = rsClientes!Nome
   sDIE = IIf(rsClientes!IE_CI <> "ISENTO", Trim(RemoveCaracteres(rsClientes!IE_CI)), "ISENTO")
   sDISUF = ""

   Print #1, "E|" & _
             sDRazaoSocial & "|" & _
             sDIE & "|" & _
             sDISUF

   Dim sDCNPJ As String
   sDCNPJ = RemoveCaracteres(rsClientes!CNPJ_CPF)

   If Len(Trim(sDCNPJ)) > 11 Then
      Print #1, "E02|" & _
                sDCNPJ
   Else
      Print #1, "E03|" & _
                sDCNPJ
   End If

   ' Endereco
   Set rsLogradouros = cnSistema.Execute("Select * From Logradouros WHERE idLogradouro=" & rsClientes!idLogradouro)
   Set rsMunicipios = cnSistema.Execute("Select * From Municipios WHERE idMunicipio=" & rsClientes!idMunicipio)

   Dim sDLogradouro As String
   Dim sDNumero As String
   Dim sDComplemento As String
   Dim sDBairro As String
   Dim sDCodigoMunicipio As String
   Dim sDMunicipio As String
   Dim sDUF As String
   Dim sDCEP As String
   Dim sDCodigoPais As String
   Dim sDPais As String
   Dim sDTelefone As String

   sDLogradouro = IIf(Not rsLogradouros.EOF, rsLogradouros!Abreviacao, ".")
   sDNumero = Trim(rsClientes!Numero)
   sDComplemento = Trim(rsClientes!Endereco)
   sDBairro = rsClientes!Bairro
   sDCodigoMunicipio = Trim(rsClientes!CodigoMunicipio)
   sDMunicipio = rsMunicipios!Nome
   sDUF = rsClientes!UF
   sDCEP = RemoveCaracteres(rsClientes!CEP)
   sDCodigoPais = "1058"
   sDPais = "BRASIL"
   sDTelefone = StrZero(Val(rsClientes!PrefixoFone1), 2) & Trim(FormataTXT(RemoveCaracteres(rsClientes!Telefone1), 1, 10))

   Print #1, "E05|" & _
             sDLogradouro & "|" & _
             sDNumero & "|" & _
             sDComplemento & "|" & _
             sDBairro & "|" & _
             sDCodigoMunicipio & "|" & _
             sDMunicipio & "|" & _
             sDUF & "|" & _
             sDCEP & "|" & _
             sDCodigoPais & "|" & _
             sDPais & "|" & _
             sDTelefone

   ' Itens
   Dim Contador As Integer
   Contador = 1

   Dim dValorTotalBC As Double
   Dim dValorTotalICMS As Double
   Dim dValorTotalBCST As Double
   Dim dValorTotalICMSST As Double
   Dim dValorTotalProdutos As Double
   Dim dValorTotalFrete As Double
   Dim dValorTotalSeguro As Double
   Dim dValorTotalDesconto As Double
   Dim dValorTotalII As Double
   Dim dValorTotalIPI As Double
   Dim dValorTotalPIS As Double
   Dim dValorTotalCofins As Double
   Dim dValorTotalOutro As Double
   Dim dValorTotalNFe As Double

   dValorTotalBC = 0
   dValorTotalICMS = 0
   dValorTotalBCST = 0
   dValorTotalICMSST = 0
   dValorTotalProdutos = 0
   dValorTotalFrete = 0
   dValorTotalSeguro = 0
   dValorTotalDesconto = 0
   dValorTotalII = 0
   dValorTotalIPI = 0
   dValorTotalPIS = 0
   dValorTotalCofins = 0
   dValorTotalOutro = 0
   dValorTotalNFe = 0

   Set rsNFeItens = cnSistema.Execute("SELECT * FROM NFeItens WHERE NFeItens.idNFe = " & rsNFe!idNFe)
   Do While Not rsNFeItens.EOF
      Set rsProdutos = cnSistema.Execute("SELECT * FROM Produtos WHERE idProduto = " & rsNFeItens!idProduto)
      Set rsUnidades = cnSistema.Execute("Select * From UnidadesMedida WHERE idUnidadeMedida=" & rsProdutos!idUnidade)
      Set rsSituacoesTributarias = cnSistema.Execute("Select * from SituacoesTributarias WHERE idSituacaoTributaria=" & rsNFeItens!idSituacaoTributaria)

      Print #1, "H|" & _
                Contador & "|" & _
                ""

      Dim sCodigoProduto As String
      Dim sCodigoBarras As String
      Dim sDescricaoProduto As String
      Dim sCodigoNCM As String
      Dim sEXTIPI As String
      Dim sGenero As String
      Dim sCFOP As String
      Dim sUnidComercial As String
      Dim sQuantidadeComercial As String
      Dim sVlUnitarioComercial As String
      Dim sVlTotalBruto As String
      Dim sCodigoBarrasTrib As String
      Dim sUnidTrib As String
      Dim sQuantidadeTrib As String
      Dim sVlUnitarioTrib As String
      Dim sVlFrete As String
      Dim sVlSeguro As String
      Dim sVlDesconto As String

      sCodigoProduto = StrZero(Val(rsProdutos!Codigo), 5)
      sCodigoBarras = ""
      sDescricaoProduto = IIf(Trim(rsProdutos!DiscriminacaoProduto) = "", RemoveAcentos(rsProdutos!Descricao), RemoveAcentos(rsProdutos!DiscriminacaoProduto))
      sCodigoNCM = ""
      sEXTIPI = ""
      sGenero = ""
      sCFOP = RemoveCaracteres(rsNFeItens!CFOP)
      sUnidComercial = IIf(Not rsUnidades.EOF, rsUnidades!Sigla, "UN")
      sQuantidadeComercial = Substitui(Format(rsNFeItens!Quantidade, "#######0.0000"), ",", ".")
      sVlUnitarioComercial = Substitui(Format(rsNFeItens!ValorUnitario, "#######0.0000"), ",", ".")
      sVlTotalBruto = Substitui(Format(rsNFeItens!Quantidade * rsNFeItens!ValorUnitario, "#######0.00"), ",", ".")
      sCodigoBarrasTrib = ""
      sUnidTrib = IIf(Not rsUnidades.EOF, rsUnidades!Sigla, "UN")
      sQuantidadeTrib = Substitui(Format(rsNFeItens!Quantidade, "#######0.0000"), ",", ".")
      sVlUnitarioTrib = Substitui(Format(rsNFeItens!ValorUnitario, "#######0.0000"), ",", ".")
      sVlFrete = ""
      sVlSeguro = ""
      sVlDesconto = IIf(rsNFeItens!Desconto > 0, Substitui(Format((((rsNFeItens!Quantidade * rsNFeItens!ValorUnitario) * rsNFeItens!Desconto) / 100), "#######0.00"), ",", "."), "")

      Print #1, "I|" & _
                sCodigoProduto & "|" & _
                sCodigoBarras & "|" & _
                sDescricaoProduto & "|" & _
                sCodigoNCM & "|" & _
                sEXTIPI & "|" & _
                sGenero & "|" & _
                sCFOP & "|" & _
                sUnidComercial & "|" & _
                sQuantidadeComercial & "|" & _
                sVlUnitarioComercial & "|" & _
                sVlTotalBruto & "|" & _
                sCodigoBarrasTrib & "|" & _
                sUnidTrib & "|" & _
                sQuantidadeTrib & "|" & _
                sVlUnitarioTrib & "|" & _
                sVlFrete & "|" & _
                sVlSeguro & "|" & _
                sVlDesconto

      ' Tributos Incidentes
      Print #1, "M"
      Print #1, "N"

      Dim sCST As String
      sCST = Mid(rsSituacoesTributarias!Codigo, 2, 2)

      Dim sOrigem As String
      Dim sModalidadeBC As String
      Dim sPercRedBC As String
      Dim sValorBC As String
      Dim sICMS As String
      Dim sValorICMS As String
      Dim dCalculoBC As Double

      Dim Nx As String
      Select Case sCST
             Case "00" ' Tributada Integralmente
                  dCalculoBC = (rsNFeItens!Quantidade * rsNFeItens!ValorUnitario) - (((rsNFeItens!Quantidade * rsNFeItens!ValorUnitario) * rsNFeItens!Desconto) / 100)

                  sOrigem = "0"
                  sModalidadeBC = "3"
                  sValorBC = Substitui(Format(dCalculoBC, "#######0.00"), ",", ".")
                  sICMS = Substitui(Format(rsNFeItens!ICMS, "###0.00"), ",", ".")
                  sValorICMS = Substitui(Format(((dCalculoBC * rsNFeItens!ICMS) / 100), "#######0.00"), ",", ".")

                  Print #1, "N02|" & _
                            sOrigem & "|" & _
                            sCST & "|" & _
                            sModalidadeBC & "|" & _
                            sValorBC & "|" & _
                            sICMS & "|" & _
                            sValorICMS & "|"

            Case "10"
                 Nx = "03"
            Case "20"
                 dCalculoBC = Round(((((rsNFeItens!Quantidade * rsNFeItens!ValorUnitario) - (((rsNFeItens!Quantidade * rsNFeItens!ValorUnitario) * rsNFeItens!Desconto) / 100)) * rsNFeItens!BaseReduzida) / 100), 2)

                  sOrigem = "0"
                  sModalidadeBC = "3"
                  sPercRedBC = Substitui(Format(rsNFeItens!BaseReduzida, "#########0.00"), ",", ".")
                  sValorBC = Substitui(Format(dCalculoBC, "#######0.00"), ",", ".")
                  sICMS = Substitui(Format(rsNFeItens!ICMS, "###0.00"), ",", ".")
                  sValorICMS = Substitui(Format(((dCalculoBC * rsNFeItens!ICMS) / 100), "#######0.00"), ",", ".")

                  Print #1, "N04|" & _
                            sOrigem & "|" & _
                            sCST & "|" & _
                            sModalidadeBC & "|" & _
                            sPercRedBC & "|" & _
                            sValorBC & "|" & _
                            sICMS & "|" & _
                            sValorICMS

             Case "30"
                  Nx = "05"
             Case "40"
                  Nx = "06"
             Case "51"
                  Nx = "07"
             Case "60"
                  Nx = "08"
             Case "70"
                  Nx = "09"
             Case "90"
                  Nx = "10"
      End Select

      ' Totalizar
      dValorTotalBC = dValorTotalBC + dCalculoBC
      dValorTotalICMS = dValorTotalICMS + ((dCalculoBC * rsNFeItens!ICMS) / 100)
      dValorTotalProdutos = dValorTotalProdutos + (rsNFeItens!Quantidade * rsNFeItens!ValorUnitario)
      dValorTotalDesconto = dValorTotalDesconto + (((rsNFeItens!Quantidade * rsNFeItens!ValorUnitario) * rsNFeItens!Desconto) / 100)
      dValorTotalNFe = dValorTotalNFe + (rsNFeItens!Quantidade * rsNFeItens!ValorUnitario)

      ' PIS
      Print #1, "Q"
      Print #1, "Q04|" & _
                "08"

      ' Cofins
      Print #1, "S"
      Print #1, "S04|" & _
                "08"

      Contador = Contador + 1
      rsNFeItens.MoveNext
   Loop
   ' Totais
   Print #1, "W"

   Dim sValorTotalBC As String
   Dim sValorTotalICMS As String
   Dim sValorTotalBCST As String
   Dim sValorTotalICMSST As String
   Dim sValorTotalProdutos As String
   Dim sValorTotalFrete As String
   Dim sValorTotalSeguro As String
   Dim sValorTotalDesconto As String
   Dim sValorTotalII As String
   Dim sValorTotalIPI As String
   Dim sValorTotalPIS As String
   Dim sValorTotalCofins As String
   Dim sValorTotalOutro As String
   Dim sValorTotalNFe As String

   sValorTotalBC = Substitui(Format(dValorTotalBC, "#########0.00"), ",", ".")
   sValorTotalICMS = Substitui(Format(dValorTotalICMS, "#########0.00"), ",", ".")
   sValorTotalBCST = Substitui(Format(0, "#########0.00"), ",", ".")
   sValorTotalICMSST = Substitui(Format(0, "#########0.00"), ",", ".")
   sValorTotalProdutos = Substitui(Format(dValorTotalProdutos, "#########0.00"), ",", ".")
   sValorTotalFrete = Substitui(Format(0, "#########0.00"), ",", ".")
   sValorTotalSeguro = Substitui(Format(0, "#########0.00"), ",", ".")
   sValorTotalDesconto = Substitui(Format(dValorTotalDesconto, "#########0.00"), ",", ".")
   sValorTotalII = Substitui(Format(0, "#########0.00"), ",", ".")
   sValorTotalIPI = Substitui(Format(0, "#########0.00"), ",", ".")
   sValorTotalPIS = Substitui(Format(0, "#########0.00"), ",", ".")
   sValorTotalCofins = Substitui(Format(0, "#########0.00"), ",", ".")
   sValorTotalOutro = Substitui(Format(0, "#########0.00"), ",", ".")
   sValorTotalNFe = Substitui(Format(dValorTotalNFe, "#########0.00"), ",", ".")

   Print #1, "W02|" & _
             sValorTotalBC & "|" & _
             sValorTotalICMS & "|" & _
             sValorTotalBCST & "|" & _
             sValorTotalICMSST & "|" & _
             sValorTotalProdutos & "|" & _
             sValorTotalFrete & "|" & _
             sValorTotalSeguro & "|" & _
             sValorTotalDesconto & "|" & _
             sValorTotalII & "|" & _
             sValorTotalIPI & "|" & _
             sValorTotalPIS & "|" & _
             sValorTotalCofins & "|" & _
             sValorTotalOutro & "|" & _
             sValorTotalNFe

   ' Frete
   Print #1, "X|" & _
             "0"

   ' Informa�oes Adicionais
   Dim sInfFISCO As String
   Dim sInfEmpresa As String

   sInfFISCO = rsNFe!InformacoesCorpo
   sInfEmpresa = rsNFe!DadosAdicionais
   Print #1, "Z|" & _
             sInfFISCO & "|" & _
             sInfEmpresa

   ' Define como Gerada
   cnSistema.Execute "Update NFe set " & _
            "GeradaNFe = True " & _
            "Where idNFe = " & rsNFe!idNFe

   Close #1

End Function

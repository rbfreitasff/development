VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMDFe 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manifesto Eletrônico de Carga - MDF-e"
   ClientHeight    =   8595
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9615
   Icon            =   "frmMDFe.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
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
            Picture         =   "frmMDFe.frx":030A
            Key             =   "Alterar"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDFe.frx":0466
            Key             =   "Anterior"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDFe.frx":05C2
            Key             =   "Cancelar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDFe.frx":071E
            Key             =   "Gravar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDFe.frx":0882
            Key             =   "Imprimir"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDFe.frx":0E1E
            Key             =   "Excluir"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDFe.frx":13BA
            Key             =   "Localizar"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDFe.frx":1516
            Key             =   "Novo"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDFe.frx":1672
            Key             =   "Inicio"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDFe.frx":17CE
            Key             =   "Proximo"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDFe.frx":192A
            Key             =   "Fim"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDFe.frx":1A86
            Key             =   "Outros"
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7890
      Left            =   60
      TabIndex        =   1
      Top             =   420
      Width           =   9510
      _ExtentX        =   16775
      _ExtentY        =   13917
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "Manifesto"
      TabPicture(0)   =   "frmMDFe.frx":1DA0
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblChaveAcesso"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblProtocolo"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraCabecalho"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtChaveAcesso"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtProtocolo"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "fraLocalCarregamento"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "fraUFsPercurso"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Rodoviário"
      TabPicture(1)   =   "frmMDFe.frx":1DBC
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraTransportadorVolumes"
      Tab(1).Control(1)=   "fraCondutores"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Informações dos Documentos"
      TabPicture(2)   =   "frmMDFe.frx":1DD8
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraNFes"
      Tab(2).Control(1)=   "fraDadosAdicionais"
      Tab(2).ControlCount=   2
      Begin VB.Frame fraDadosAdicionais 
         Caption         =   "Dados Adicionais/Informações Complementares"
         Height          =   1845
         Left            =   -74940
         TabIndex        =   76
         Top             =   4800
         Width           =   9315
         Begin VB.TextBox txtDadosAdicionais 
            Height          =   1515
            Left            =   75
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   77
            Top             =   240
            Width           =   9150
         End
      End
      Begin VB.Frame fraUFsPercurso 
         Caption         =   "UFs do Percurso"
         Height          =   2835
         Left            =   5280
         TabIndex        =   32
         Top             =   2640
         Width           =   4155
         Begin VB.ComboBox cboUFPercurso 
            Height          =   315
            ItemData        =   "frmMDFe.frx":1DF4
            Left            =   120
            List            =   "frmMDFe.frx":1DF6
            Style           =   2  'Dropdown List
            TabIndex        =   34
            Top             =   540
            Width           =   705
         End
         Begin VB.CommandButton cmdInserirPercurso 
            BackColor       =   &H00E0E0E0&
            Height          =   315
            Left            =   3180
            Picture         =   "frmMDFe.frx":1DF8
            Style           =   1  'Graphical
            TabIndex        =   35
            Top             =   540
            Width           =   360
         End
         Begin VB.CommandButton cmdExcluirPercurso 
            BackColor       =   &H80000004&
            Height          =   315
            Left            =   3570
            Picture         =   "frmMDFe.frx":1F42
            Style           =   1  'Graphical
            TabIndex        =   36
            Top             =   540
            Width           =   360
         End
         Begin MSComctlLib.ListView lvwPercurso 
            Height          =   1800
            Left            =   120
            TabIndex        =   37
            Top             =   900
            Width           =   3945
            _ExtentX        =   6959
            _ExtentY        =   3175
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
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "UF"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   180
            TabIndex        =   33
            Top             =   300
            Width           =   210
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "UF"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   -540
            TabIndex        =   81
            Top             =   -60
            Width           =   210
         End
      End
      Begin VB.Frame fraLocalCarregamento 
         Caption         =   "Local do Carregamento / Descarregamento"
         Height          =   2835
         Left            =   60
         TabIndex        =   22
         Top             =   2640
         Width           =   4995
         Begin VB.ComboBox cboUFDescarregamento 
            Height          =   315
            ItemData        =   "frmMDFe.frx":23F8
            Left            =   2160
            List            =   "frmMDFe.frx":23FA
            Style           =   2  'Dropdown List
            TabIndex        =   31
            Top             =   2400
            Width           =   705
         End
         Begin VB.ComboBox cboUFCarregamento 
            Height          =   315
            ItemData        =   "frmMDFe.frx":23FC
            Left            =   60
            List            =   "frmMDFe.frx":23FE
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   540
            Width           =   705
         End
         Begin VB.ComboBox cboMunicipio 
            Height          =   315
            ItemData        =   "frmMDFe.frx":2400
            Left            =   840
            List            =   "frmMDFe.frx":2402
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   540
            Width           =   3225
         End
         Begin VB.CommandButton cmdInserirLocalCarregamento 
            BackColor       =   &H00E0E0E0&
            Height          =   315
            Left            =   4140
            Picture         =   "frmMDFe.frx":2404
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   540
            Width           =   360
         End
         Begin VB.CommandButton cmdExcluirLocalCarregamento 
            BackColor       =   &H80000004&
            Height          =   315
            Left            =   4530
            Picture         =   "frmMDFe.frx":254E
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   540
            Width           =   360
         End
         Begin MSComctlLib.ListView lvwLocalCarregamento 
            Height          =   1440
            Left            =   60
            TabIndex        =   29
            Top             =   900
            Width           =   4845
            _ExtentX        =   8546
            _ExtentY        =   2540
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
         Begin VB.Label lblLocalDescarregamento 
            AutoSize        =   -1  'True
            Caption         =   "Local do Descarregamento"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   30
            Top             =   2460
            Width           =   1920
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "UF"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   23
            Top             =   300
            Width           =   210
         End
         Begin VB.Label lblMunicipio 
            AutoSize        =   -1  'True
            Caption         =   "Municipio"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   840
            TabIndex        =   25
            Top             =   300
            Width           =   675
         End
      End
      Begin VB.Frame fraNFes 
         Height          =   4305
         Left            =   -74940
         TabIndex        =   63
         Top             =   360
         Width           =   9375
         Begin VB.TextBox txtNumeroNota 
            Height          =   315
            Left            =   75
            MaxLength       =   20
            TabIndex        =   65
            Top             =   405
            Width           =   1095
         End
         Begin VB.CommandButton cmdIncluirNFe 
            BackColor       =   &H00E0E0E0&
            Height          =   315
            Left            =   8535
            Picture         =   "frmMDFe.frx":2A04
            Style           =   1  'Graphical
            TabIndex        =   69
            Top             =   405
            Width           =   360
         End
         Begin VB.CommandButton cmdPesquisaNFes 
            BackColor       =   &H00E0E0E0&
            Height          =   315
            Left            =   1200
            Picture         =   "frmMDFe.frx":2B4E
            Style           =   1  'Graphical
            TabIndex        =   66
            Top             =   420
            Width           =   360
         End
         Begin VB.TextBox txtChaveNFe 
            Height          =   315
            Left            =   1590
            MaxLength       =   50
            TabIndex        =   68
            Top             =   405
            Width           =   6900
         End
         Begin VB.CommandButton cmdExcluirNFe 
            BackColor       =   &H80000004&
            Height          =   315
            Left            =   8925
            Picture         =   "frmMDFe.frx":2C98
            Style           =   1  'Graphical
            TabIndex        =   70
            Top             =   405
            Width           =   360
         End
         Begin MSComctlLib.ListView lvwNFes 
            Height          =   2760
            Left            =   60
            TabIndex        =   71
            Top             =   735
            Width           =   9225
            _ExtentX        =   16272
            _ExtentY        =   4868
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
         Begin MSMask.MaskEdBox mskValorTotalProdutos 
            Height          =   285
            Left            =   7815
            TabIndex        =   73
            Top             =   3585
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
            Left            =   7815
            TabIndex        =   75
            Top             =   3885
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   13
            Format          =   "###,##0.00"
            PromptChar      =   "_"
         End
         Begin VB.Label lblNumeroNota 
            AutoSize        =   -1  'True
            Caption         =   "Nota"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   90
            TabIndex        =   64
            Top             =   180
            Width           =   345
         End
         Begin VB.Label lblChaveNFe 
            AutoSize        =   -1  'True
            Caption         =   "Chave da NFe"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   1620
            TabIndex        =   67
            Top             =   180
            Width           =   1035
         End
         Begin VB.Label lblValorTotalProdutos 
            AutoSize        =   -1  'True
            Caption         =   "Total Produtos"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   6720
            TabIndex        =   72
            Top             =   3630
            Width           =   1035
         End
         Begin VB.Label lblTotalNota 
            AutoSize        =   -1  'True
            Caption         =   "Total Nota"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   7005
            TabIndex        =   74
            Top             =   3930
            Width           =   750
         End
      End
      Begin VB.TextBox txtProtocolo 
         Enabled         =   0   'False
         Height          =   315
         Left            =   6960
         TabIndex        =   5
         Top             =   645
         Width           =   2460
      End
      Begin VB.TextBox txtChaveAcesso 
         Enabled         =   0   'False
         Height          =   315
         Left            =   60
         TabIndex        =   3
         Top             =   645
         Width           =   6780
      End
      Begin VB.Frame fraCondutores 
         Caption         =   "Condutores"
         Height          =   2385
         Left            =   -74940
         TabIndex        =   55
         Top             =   1800
         Width           =   9375
         Begin VB.TextBox txtNomeCondutor 
            Height          =   315
            Left            =   1800
            MaxLength       =   50
            TabIndex        =   59
            Top             =   420
            Width           =   6600
         End
         Begin VB.CommandButton cmdIncluirCondutor 
            BackColor       =   &H80000004&
            Height          =   315
            Left            =   8520
            Picture         =   "frmMDFe.frx":314E
            Style           =   1  'Graphical
            TabIndex        =   60
            Top             =   405
            Width           =   360
         End
         Begin VB.CommandButton cmdExcluirCondutor 
            BackColor       =   &H80000004&
            Height          =   315
            Left            =   8880
            Picture         =   "frmMDFe.frx":3604
            Style           =   1  'Graphical
            TabIndex        =   61
            Top             =   405
            Width           =   360
         End
         Begin MSComctlLib.ListView lvwCondutores 
            Height          =   1590
            Left            =   60
            TabIndex        =   62
            Top             =   720
            Width           =   9240
            _ExtentX        =   16298
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
         Begin MSMask.MaskEdBox mskCPFCondutor 
            Height          =   315
            Left            =   75
            TabIndex        =   57
            Top             =   420
            Width           =   1680
            _ExtentX        =   2963
            _ExtentY        =   556
            _Version        =   393216
            PromptChar      =   " "
         End
         Begin VB.Label lblNomeCondutor 
            AutoSize        =   -1  'True
            Caption         =   "Nome"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   1800
            TabIndex        =   58
            Top             =   210
            Width           =   420
         End
         Begin VB.Label lblCPFCondutor 
            AutoSize        =   -1  'True
            Caption         =   "CPF"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   75
            TabIndex        =   56
            Top             =   210
            Width           =   300
         End
      End
      Begin VB.Frame fraTransportadorVolumes 
         Caption         =   "Veículo de Tração"
         Height          =   1335
         Left            =   -74940
         TabIndex        =   38
         Top             =   420
         Width           =   9375
         Begin VB.TextBox txtRenavam 
            Height          =   315
            Left            =   1380
            MaxLength       =   50
            TabIndex        =   54
            Top             =   900
            Width           =   1395
         End
         Begin VB.TextBox txtCapacidadeM3 
            Height          =   315
            Left            =   7620
            MaxLength       =   50
            TabIndex        =   52
            Top             =   570
            Width           =   915
         End
         Begin VB.ComboBox cboUFVeiculo 
            Height          =   315
            ItemData        =   "frmMDFe.frx":3ABA
            Left            =   4800
            List            =   "frmMDFe.frx":3ABC
            Style           =   2  'Dropdown List
            TabIndex        =   42
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox txtCapacidadeKG 
            Height          =   315
            Left            =   6180
            MaxLength       =   50
            TabIndex        =   50
            Top             =   570
            Width           =   915
         End
         Begin VB.TextBox txtTara 
            Height          =   315
            Left            =   7620
            MaxLength       =   50
            TabIndex        =   46
            Top             =   240
            Width           =   1560
         End
         Begin VB.ComboBox cboTipoRodado 
            Height          =   315
            ItemData        =   "frmMDFe.frx":3ABE
            Left            =   1380
            List            =   "frmMDFe.frx":3AD4
            Style           =   2  'Dropdown List
            TabIndex        =   48
            Top             =   570
            Width           =   2895
         End
         Begin VB.ComboBox cboTipoCarroceria 
            Height          =   315
            ItemData        =   "frmMDFe.frx":3B0F
            Left            =   1380
            List            =   "frmMDFe.frx":3B25
            Style           =   2  'Dropdown List
            TabIndex        =   40
            Top             =   240
            Width           =   2895
         End
         Begin MSMask.MaskEdBox mskPlacaVeiculo 
            Height          =   285
            Left            =   6180
            TabIndex        =   44
            Top             =   255
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   8
            Mask            =   "CCC-9999"
            PromptChar      =   " "
         End
         Begin VB.Label lblCapacidadeM3 
            AutoSize        =   -1  'True
            Caption         =   "(M3)"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   7185
            TabIndex        =   51
            Top             =   630
            Width           =   315
         End
         Begin VB.Label lblUFCaminhao 
            AutoSize        =   -1  'True
            Caption         =   "UF"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   4500
            TabIndex        =   41
            Top             =   300
            Width           =   210
         End
         Begin VB.Label lblRenavam 
            AutoSize        =   -1  'True
            Caption         =   "RENAVAM"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   390
            TabIndex        =   53
            Top             =   960
            Width           =   795
         End
         Begin VB.Label lblTara 
            AutoSize        =   -1  'True
            Caption         =   "Tara"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   7200
            TabIndex        =   45
            Top             =   300
            Width           =   330
         End
         Begin VB.Label lblCapacidadeKG 
            AutoSize        =   -1  'True
            Caption         =   "Capacidade (KG)"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   4845
            TabIndex        =   49
            Top             =   630
            Width           =   1215
         End
         Begin VB.Label lblTransportador 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Carroceria"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   39
            Top             =   300
            Width           =   1080
         End
         Begin VB.Label lblPlaca 
            AutoSize        =   -1  'True
            Caption         =   "Placa"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   5700
            TabIndex        =   43
            Top             =   300
            Width           =   405
         End
         Begin VB.Label lblTipoRodado 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Rodado"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   255
            TabIndex        =   47
            Top             =   630
            Width           =   930
         End
      End
      Begin VB.Frame fraCabecalho 
         Height          =   1605
         Left            =   60
         TabIndex        =   78
         Top             =   960
         Width           =   9375
         Begin VB.ComboBox cboFormaEmissao 
            Height          =   315
            ItemData        =   "frmMDFe.frx":3B70
            Left            =   6300
            List            =   "frmMDFe.frx":3B7A
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   840
            Width           =   2445
         End
         Begin VB.ComboBox cboModalidade 
            Height          =   315
            ItemData        =   "frmMDFe.frx":3B93
            Left            =   1560
            List            =   "frmMDFe.frx":3B9D
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   1170
            Width           =   3165
         End
         Begin VB.ComboBox cboUF 
            Height          =   315
            ItemData        =   "frmMDFe.frx":3BB9
            Left            =   8040
            List            =   "frmMDFe.frx":3BBB
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   165
            Width           =   705
         End
         Begin VB.ComboBox cboTipoTransportador 
            Height          =   315
            ItemData        =   "frmMDFe.frx":3BBD
            Left            =   1560
            List            =   "frmMDFe.frx":3BCD
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   825
            Width           =   3165
         End
         Begin VB.ComboBox cboTipoEmitente 
            Height          =   315
            ItemData        =   "frmMDFe.frx":3BE2
            Left            =   1560
            List            =   "frmMDFe.frx":3BEC
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   480
            Width           =   7185
         End
         Begin MSMask.MaskEdBox mskNumero 
            Height          =   285
            Left            =   1560
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
         Begin MSMask.MaskEdBox mskDataEmissao 
            Height          =   285
            Left            =   3600
            TabIndex        =   9
            Top             =   180
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "99/99/9999"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskDataViagem 
            Height          =   285
            Left            =   5760
            TabIndex        =   11
            Top             =   180
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "99/99/9999"
            PromptChar      =   " "
         End
         Begin VB.Label lblFormaEmissao 
            AutoSize        =   -1  'True
            Caption         =   "Forma de Emissão"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   4860
            TabIndex        =   18
            Top             =   900
            Width           =   1290
         End
         Begin VB.Label lblModalidade 
            AutoSize        =   -1  'True
            Caption         =   "Modalidade"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   645
            TabIndex        =   20
            Top             =   1230
            Width           =   825
         End
         Begin VB.Label lblUF 
            AutoSize        =   -1  'True
            Caption         =   "UF"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   7680
            TabIndex        =   12
            Top             =   225
            Width           =   210
         End
         Begin VB.Label lblDataViagem 
            AutoSize        =   -1  'True
            Caption         =   "Viagem"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   5160
            TabIndex        =   10
            Top             =   225
            Width           =   525
         End
         Begin VB.Label lblDataEmissao 
            AutoSize        =   -1  'True
            Caption         =   "Emissão"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   2940
            TabIndex        =   8
            Top             =   225
            Width           =   585
         End
         Begin VB.Label lblTipoTransportador 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Transportador"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   16
            Top             =   885
            Width           =   1350
         End
         Begin VB.Label lblNumero 
            AutoSize        =   -1  'True
            Caption         =   "Número"
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
            Left            =   810
            TabIndex        =   6
            Top             =   225
            Width           =   660
         End
         Begin VB.Label lblTipoEmitente 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Emitente"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   270
            TabIndex        =   14
            Top             =   540
            Width           =   1200
         End
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
         TabIndex        =   79
         Top             =   -240
         Width           =   1110
      End
      Begin VB.Label lblProtocolo 
         AutoSize        =   -1  'True
         Caption         =   "Protocolo de Autorização"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   6960
         TabIndex        =   4
         Top             =   405
         Width           =   1785
      End
      Begin VB.Label lblChaveAcesso 
         AutoSize        =   -1  'True
         Caption         =   "Chave de Acesso"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   60
         TabIndex        =   2
         Top             =   390
         Width           =   1260
      End
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
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
            Object.ToolTipText     =   "Próximo Registro"
            ImageKey        =   "Proximo"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Fim"
            Object.ToolTipText     =   "Último Registro"
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
            Object.ToolTipText     =   "Outras Opções"
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
      TabIndex        =   80
      Top             =   8340
      Visible         =   0   'False
      Width           =   1395
   End
End
Attribute VB_Name = "frmMDFe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ItemList As ListItem

Dim rsMDFe As New ADODB.Recordset
Dim rsNF As New ADODB.Recordset
Dim rsTemp As New ADODB.Recordset
Dim rsDados As New ADODB.Recordset


''''''''''''''''''''''''''''''''''''''''''''''''''
' Criar RecordSets dos MDFes
''''''''''''''''''''''''''''''''''''''''''''''''''
Public rsMDFes As New ADODB.Recordset
Public rsMDFeLocalCarregamento As New ADODB.Recordset
Public rsMDFePercurso  As New ADODB.Recordset
Public rsMDFeCondutores  As New ADODB.Recordset
Public rsMDFeNFes  As New ADODB.Recordset
''''''''''''''''''''''''''''''''''''''''''''''''''


'''''Dim rsMDFes As New ADODB.Recordset
'''''Dim rsMDFeLocalCarregamento As New ADODB.Recordset
'''''Dim rsMDFePercurso  As New ADODB.Recordset
'''''Dim rsMDFeCondutores  As New ADODB.Recordset
'''''Dim rsMDFeNFes  As New ADODB.Recordset

'Dim PagamentoAtual As String
Dim NotaAtual As String
Dim RegistroAtual As Double 'Para posicionar o ponteiro depois de pesquisas
'Dim rsMDFe As New ADODB.Recordset
'Dim rsMDFeItens As New ADODB.Recordset
'Dim rsMDFeBoletos As New ADODB.Recordset
'Dim rsProdutos As New ADODB.Recordset
'Dim rsEmpresa As New ADODB.Recordset
'Dim rsTransportador As New ADODB.Recordset
'Dim rsCFOPs As New ADODB.Recordset
'Dim rsClientes As New ADODB.Recordset
'Dim rsTemp As New ADODB.Recordset
'Dim rsTemp2 As New ADODB.Recordset
'Dim rsNaturezasOperacao As New ADODB.Recordset
'Dim rsContasBancarias As New ADODB.Recordset
'Dim rsSaldoProdutos As New ADODB.Recordset
'Dim rsUnidadesMedida As New ADODB.Recordset
'Dim rsSituacoesTributarias As New ADODB.Recordset
'Dim rsCFOPReferencias As New ADODB.Recordset
'Dim strDescricaoTemp As String
'Dim intItensNota As Integer
'Dim rsLogradouros As New ADODB.Recordset
'Dim rsMunicipios As New ADODB.Recordset
'Dim rsUnidades As New ADODB.Recordset
'Dim rsFormasPagamento As New ADODB.Recordset
'Dim rsFretes As New ADODB.Recordset

Dim Contador As Integer

Private Sub Form_Load()
   I_TituloForm = Me.Caption
   On Error GoTo Erro
   SSTab1.Tab = 0 ' Posiciona no primeiro tab
   Status = 0
   RegistroAtual = 0
   Centraliza frmMDFe
'   rsMDFe.Open "Select * from MDFe Order By Numero", cnSistema, adOpenForwardOnly, adLockOptimistic, 1
   Set rsMDFe = cnSistema.Execute("Select * from MDFe Order By Numero")

   lvwLocalCarregamento.ColumnHeaders.Add , , "UF", 750
   lvwLocalCarregamento.ColumnHeaders.Add , , "Município", 3800

   lvwPercurso.ColumnHeaders.Add , , "UF", 2000

   lvwCondutores.ColumnHeaders.Add , , "CPF", 2000
   lvwCondutores.ColumnHeaders.Add , , "Nome", 6000

   lvwNFes.ColumnHeaders.Add , , "Número", 2000
   lvwNFes.ColumnHeaders.Add , , "Chave", 6000
'   lvwNFes.ColumnHeaders.Add , , "Valor", 1050, lvwColumnRight

''   lvwProdutos.ColumnHeaders.Add , , "Código", 850
''   lvwProdutos.ColumnHeaders.Add , , "Produto", 3200
''   lvwProdutos.ColumnHeaders.Add , , "Quantidade", 1000, lvwColumnRight
''   lvwProdutos.ColumnHeaders.Add , , "Vl. Unitário", 1050, lvwColumnRight
''   lvwProdutos.ColumnHeaders.Add , , "Desc.", 700, lvwColumnRight
''   lvwProdutos.ColumnHeaders.Add , , "Valor Total", 1050, lvwColumnRight
''   lvwProdutos.ColumnHeaders.Add , , "Vl. Líquido", 1050, lvwColumnRight
''
''   lvwBoletos.ColumnHeaders.Add , , "Boleto", 1500
''   lvwBoletos.ColumnHeaders.Add , , "Vencimento", 1300
''   lvwBoletos.ColumnHeaders.Add , , "Valor", 1500, lvwColumnRight

   Carrega_Combos
'   If Registros(cnSistema, "MDFe") = 0 Then
   If Registros2("MDFe") = 0 Then
      Botoes 3, frmMDFe
   Else
      Botoes 1, frmMDFe
      rsMDFe.MoveLast
      Prencher_Campos
   End If

   If I_Acesso = 3 Then ' Controle Níveis de Acesso
      Toolbar.Buttons(2).Visible = False
      Toolbar.Buttons(3).Visible = False
   End If

   Campos False
'   If Registros(cnSistema, "MDFe") = 0 Then
   If Registros2("MDFe") = 0 Then
      Toolbar.Buttons(15).Enabled = False
   End If
''   MDISistema.StatusBar.Panels(1).text = "Cadastro Notas de Saída Eletrônicas"

Exit Sub
Erro:
   If Err.Number = -2147467259 Then
      rsErro = True
      Beep
      MsgBox "Erro na Abertura do Arquivo de Dados" & Chr(13) & "Algum usuário está com o Arquivo em modo Exclusivo", vbExclamation, "Erro"
      Exit Sub
   Else
      rsErro = True
      Beep
      MsgBox "Verificar: " & Err.Number & Chr(13) & Err.Description, vbExclamation, "Sistema"
      Exit Sub
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If Not rsErro Then rsMDFe.Close
End Sub

Private Sub Carrega_Combos()

'  UFs
   Set rsTemp = cnSistema.Execute("Select * from UFs ORDER BY Sigla")
   cboUF.Clear
   cboUFCarregamento.Clear
   cboUFDescarregamento.Clear
   cboUFPercurso.Clear
   cboUFVeiculo.Clear
   Do While Not rsTemp.EOF
      cboUF.AddItem rsTemp!Sigla
      cboUF.ItemData(cboUF.NewIndex) = rsTemp!idUF

      cboUFCarregamento.AddItem rsTemp!Sigla
      cboUFCarregamento.ItemData(cboUFCarregamento.NewIndex) = rsTemp!idUF

      cboUFPercurso.AddItem rsTemp!Sigla
      cboUFPercurso.ItemData(cboUFPercurso.NewIndex) = rsTemp!idUF

      cboUFVeiculo.AddItem rsTemp!Sigla
      cboUFVeiculo.ItemData(cboUFVeiculo.NewIndex) = rsTemp!idUF

      cboUFDescarregamento.AddItem rsTemp!Sigla
      cboUFDescarregamento.ItemData(cboUFDescarregamento.NewIndex) = rsTemp!idUF

      rsTemp.MoveNext
   Loop

''  Clientes
'
''   Set rsTemp = cnSistema.Execute("Select * from Clientes Where Situacao = 0 Order By Nome")
'   Set rsTemp = cnSistema.Execute("Select * from Clientes Order By Nome")
'   cmbCliente.Clear
'   Do While Not rsTemp.EOF
'      cmbCliente.AddItem rsTemp!Nome
'      cmbCliente.ItemData(cmbCliente.NewIndex) = rsTemp!idCliente
'      rsTemp.MoveNext
'   Loop
'
''  Naturezas de Operação
'
'   Set rsTemp = cnSistema.Execute("Select * from NaturezasOperacao ORDER BY Descricao")
'   cmbNaturezaOperacao.Clear
'   Do While Not rsTemp.EOF
'      cmbNaturezaOperacao.AddItem rsTemp!Descricao
'      cmbNaturezaOperacao.ItemData(cmbNaturezaOperacao.NewIndex) = rsTemp!idNaturezaOperacao
'      rsTemp.MoveNext
'   Loop
'
''  Transportadores
'
'   Set rsTemp = cnSistema.Execute("Select * from Transportadores ORDER BY Nome")
'   cmbTransportador.Clear
'   Do While Not rsTemp.EOF
'      cmbTransportador.AddItem rsTemp!Nome
'      cmbTransportador.ItemData(cmbTransportador.NewIndex) = rsTemp!idTransportador
'      rsTemp.MoveNext
'   Loop
'
''  Formas de Pagamento
'
'   Set rsTemp = cnSistema.Execute("Select * from FormasPagamento ORDER BY Descricao")
'   cmbFormaPagamento.Clear
'   Do While Not rsTemp.EOF
'      cmbFormaPagamento.AddItem rsTemp!Descricao
'      cmbFormaPagamento.ItemData(cmbFormaPagamento.NewIndex) = rsTemp!idFormaPagamento
'      rsTemp.MoveNext
'   Loop
'
''  UFs
'   Set rsTemp = cnSistema.Execute("Select * from UFs ORDER BY Sigla")
'   cmbUFPlaca.Clear
'   Do While Not rsTemp.EOF
'      cmbUFPlaca.AddItem rsTemp!Sigla
'      cmbUFPlaca.ItemData(cmbUFPlaca.NewIndex) = rsTemp!idUF
'      rsTemp.MoveNext
'   Loop
'
''  Frete Conta
'   Set rsTemp = cnSistema.Execute("Select * from FreteConta ORDER BY Descricao")
'   cmbFreteConta.Clear
'   Do While Not rsTemp.EOF
'      cmbFreteConta.AddItem rsTemp!Descricao
'      cmbFreteConta.ItemData(cmbFreteConta.NewIndex) = rsTemp!idFreteConta
'      rsTemp.MoveNext
'   Loop
'
'   rsTemp.Close
End Sub

Private Sub mskCPFCondutor_LostFocus()
Dim Verifica_CPF As String
Dim intCliente, Contador As Integer

   If mskCPFCondutor.Text <> Empty Then
      Verifica_CPF = CNPJ_CPF(mskCPFCondutor.Text)
      If Verifica_CPF <> "ERRO" Then
         mskCPFCondutor.Text = CNPJ_CPF(mskCPFCondutor.Text)
      Else
         mskCPFCondutor.SelStart = 0
         mskCPFCondutor.SelLength = Len(mskCPFCondutor.Text)
         mskCPFCondutor.SetFocus
      End If
   End If
End Sub

Private Sub mskDataEmissao_LostFocus()
   rsMDFes!DataEmissao = mskDataEmissao.Text
End Sub

Private Sub mskDataViagem_LostFocus()
   rsMDFes!DataViagem = mskDataViagem.Text
End Sub

Private Sub mskPlacaVeiculo_LostFocus()
   rsMDFes!PlacaVeiculo = mskPlacaVeiculo.Text
End Sub


'Private Sub mskCNPJ_CPF_LostFocus()
'Dim Verifica_CPF As String
'Dim intCliente, Contador As Integer
'
'   If mskCNPJ_CPF.text <> Empty Then
'      Verifica_CPF = CNPJ_CPF(mskCNPJ_CPF.text)
'      If Verifica_CPF <> "ERRO" Then
'          mskCNPJ_CPF.text = CNPJ_CPF(mskCNPJ_CPF.text)
'          Set rsTemp = cnSistema.Execute("Select idCliente From Clientes Where CNPJ_CPF = '" & mskCNPJ_CPF.text & "'")
'          If Not rsTemp.BOF And Not rsTemp.EOF Then
'             intCliente = rsTemp!idCliente
'
'             For Contador = 0 To (cmbCliente.ListCount - 1)
'                 If cmbCliente.ItemData(Contador) = intCliente Then
'                    cmbCliente.ListIndex = Contador
'                    Exit For
'                 End If
'             Next
'          End If
'          If cmbCliente.Enabled Then cmbCliente.SetFocus
'      Else
'          mskCNPJ_CPF.SelStart = 0
'          mskCNPJ_CPF.SelLength = Len(mskCNPJ_CPF.text)
'          mskCNPJ_CPF.SetFocus
'      End If
'   End If
'End Sub
'
'Private Sub mskQuantidade_LostFocus()
'Dim Contador As Integer
'
''''   Set rsTemp = cnSistema.Execute("Select * From TabelaProdutos Where Codigo = '" & txtCodigo.Text & "'")
''''   If Not rsTemp.EOF Then
''''      If rsTemp!Saldo < Val(mskQuantidade.Text) Then
''''         Beep
''''         Set rsTemp2 = cnSistema.Execute("Select * From SaldosConferencias Where idProduto = " & rsTemp!idProduto)
''''         Dim sTruncados As String
''''         If Not rsTemp2.EOF Then
''''            sTruncados = Chr(13) & "Mais: " & rsTemp2!Quantidade & " Truncados"
''''         Else
''''            sTruncados = ""
''''         End If
''''
''''         MsgBox "Saldo Menor que Quantidade." & Chr(13) & "Atual: " & rsTemp!Saldo & sTruncados, vbExclamation, "Erro"
''''         mskDesconto.SetFocus
''''         Exit Sub
''''      End If
''''   End If
'
'   If frmNFe.txtCodigo.text <> "" Then
'    ' Carrega Combos
'      Set rsTemp = cnSistema.Execute("Select * from UnidadesMedida Order By Descricao")
'      frmNFeComplemento.cmbUnidade.Clear
'      Do While Not rsTemp.EOF
'         frmNFeComplemento.cmbUnidade.AddItem rsTemp!Descricao
'         frmNFeComplemento.cmbUnidade.ItemData(frmNFeComplemento.cmbUnidade.NewIndex) = rsTemp!idUnidadeMedida
'         rsTemp.MoveNext
'      Loop
'
'    ' Preencher Campos
'      Set rsTemp = cnSistema.Execute("Select * From Produtos Where Codigo = '" & SQLCheck(frmNFe.txtCodigo.text) & "'")
'      If Not rsTemp.EOF Then
'         ' ICMS
'         frmNFeComplemento.mskICMSProduto.text = rsTemp!ICMS
'         ' Unidade
'         For Contador = 0 To (frmNFeComplemento.cmbUnidade.ListCount - 1)
'            If frmNFeComplemento.cmbUnidade.ItemData(Contador) = rsTemp!idUnidade Then
'               frmNFeComplemento.cmbUnidade.ListIndex = Contador
'               Exit For
'            End If
'         Next
'      End If
'   End If
'
'End Sub
'
'Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
'   If KeyAscii = 13 Then Sendkeys "{TAB}"
'End Sub
'
'Private Sub txtCodigoCliente_LostFocus()
'Dim intCliente, Contador As Integer
'
'   If txtCodigoCliente.text <> Empty Then
'      Set rsTemp = cnSistema.Execute("Select * From Clientes Where Codigo = " & txtCodigoCliente.text)
'      If Not rsTemp.BOF And Not rsTemp.EOF Then
'         mskCNPJ_CPF.text = rsTemp!CNPJ_CPF
'         intCliente = rsTemp!idCliente
'
'         For Contador = 0 To (cmbCliente.ListCount - 1)
'             If cmbCliente.ItemData(Contador) = intCliente Then
'                cmbCliente.ListIndex = Contador
'                Exit For
'             End If
'         Next
'         If cmbCliente.Enabled Then cmbCliente.SetFocus
'      Else
'          MsgBox "Código não encontrado", vbOKOnly, "Visualiza"
'          txtCodigoCliente.text = Empty
'          txtCodigoCliente.SetFocus
'      End If
'   End If
'End Sub
'
Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
   Select Case Button.Key
      Case "Novo"
         Status = 1
         Limpa_Campos
         Botoes 2, frmMDFe
         RegistroAtual = IIf(rsMDFe.EOF, 0, rsMDFe!idMDFe)
         Campos True
         mskDataEmissao.SetFocus

      Case "Cancelar"
         Status = 0
         If rsMDFe.EOF Then
            Botoes 3, frmMDFe
         Else
            Botoes 1, frmMDFe
         End If
         Campos False
'         If Registros(cnSistema, "MDFe") <> 0 Then
         If Registros2("MDFe") = 0 Then
            If RegistroAtual <> 0 Then
               rsMDFe.MoveFirst
               rsMDFe.Find "idMDFe = " & RegistroAtual
            End If
            Prencher_Campos
         Else
            Limpa_Campos
         End If

      Case "Gravar"
         Call Gravar

      Case "Alterar"
         RegistroAtual = rsMDFe!idMDFe
         NotaAtual = mskNumero.Text
         Status = 2
         Call Alterar

      Case "Excluir"
         Status = 3
         Call Excluir
         If rsMDFe.EOF Then
            Botoes 3, frmMDFe
         Else
            Botoes 1, frmMDFe
         End If

      Case "Localizar"
         RegistroAtual = rsMDFe!idMDFe
         Status = 4
         Call Localizar

      Case "Inicio"
         rsMDFe.MoveFirst
         Prencher_Campos

      Case "Fim"
         rsMDFe.MoveLast
         Prencher_Campos

      Case "Proximo"
         rsMDFe.MoveNext
         If rsMDFe.EOF Then rsMDFe.MoveLast
         Prencher_Campos

      Case "Anterior"
         rsMDFe.MovePrevious
         If rsMDFe.BOF Then rsMDFe.MoveFirst
         Prencher_Campos

      Case "Imprimir"
''         Call ImprimirNota
   End Select
End Sub

Private Sub Toolbar_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Dim Contador As Integer
Dim Total_Registros As Integer
Dim PaginaInicial, Paginafinal, NumeroCopias, i

   RegistroAtual = rsMDFe!idMDFe
   Select Case ButtonMenu.Key
      Case "Localiza"
'         Status = 4
'         Call Localizar

         Status = 4
         Registro_Selecionado = False
         Screen.MousePointer = vbDefault
''         frmMDFePesquisa.Show vbModal
         rsMDFe.MoveFirst
         If Registro_Selecionado Then
            rsMDFe.Find "idMDFe = " & Val(Mid(frmMDFePesquisa.lvwDados.SelectedItem.Key, 2, Len(frmMDFePesquisa.lvwDados.SelectedItem.Key)))
         Else
            rsMDFe.Find "idMDFe = " & RegistroAtual
         End If
         Prencher_Campos

      Case "Visualiza"
'         Total_Registros = Registros(cnSistema, "MDFe")
         Total_Registros = Registros2("MDFe")
         If Total_Registros = 0 Then
            MsgBox "Não existe nenhum Registro na Tabela", vbOKOnly, "Visualiza"
            Exit Sub
         End If
         Contador = 1
         Status = 4
         Screen.MousePointer = vbHourglass
         frmVisualiza.lvwDados.ColumnHeaders.Clear
         frmVisualiza.lvwDados.ColumnHeaders.Add , , "Vencimento", 1100
         frmVisualiza.lvwDados.ColumnHeaders.Add , , "Número", 1100
         frmVisualiza.lvwDados.ColumnHeaders.Add , , "Cliente", 3800
         rsMDFe.MoveFirst
         frmVisualiza.lvwDados.ListItems.Clear
         Do While Not rsMDFe.EOF
            frmMDFe.Caption = "Processando " & StrZero(Contador, 8) & " de " & StrZero(Total_Registros, 8)
            Set rsTemp = cnSistema.Execute("Select * From Clientes Where idCliente = " & rsMDFe!idCliente)
            Set ItemList = frmVisualiza.lvwDados.ListItems.Add(, "R" & CStr(rsMDFe!idMDFe), rsMDFe!DataEmissao)
                ItemList.SubItems(1) = rsMDFe!Numero
                ItemList.SubItems(2) = rsTemp!Nome
            rsMDFe.MoveNext
            Contador = Contador + 1
         Loop
         rsMDFe.MoveFirst
         Registro_Selecionado = False
         Screen.MousePointer = vbDefault
         Me.Caption = I_TituloForm
         frmVisualiza.Show vbModal
         rsMDFe.MoveFirst
         If Registro_Selecionado Then
            rsMDFe.Find "idMDFe = " & Val(Mid(frmVisualiza.lvwDados.SelectedItem.Key, 2, Len(frmVisualiza.lvwDados.SelectedItem.Key)))
         Else
            rsMDFe.Find "idMDFe = " & RegistroAtual
         End If
         Prencher_Campos

      Case "Nota"
'''''         Call ImprimirNota

      Case "Boleto"
'''''         Call ImprimirBoleto

      Case "Copiar"
'''''         Call CopiarNota

      Case "AtualizarMDFe"
'''''         lblAtualizarMDFe.Tag = rsMDFe!idMDFe
'''''         frmMDFeAtualizar.Show
'''''         lblAtualizarMDFe.Tag = ""

   End Select
End Sub

Private Sub mskNumero_LostFocus()
   If Status = 1 Then
      Set rsTemp = cnSistema.Execute("SELECT * FROM MDFe WHERE Numero = " & mskNumero.Text)
      If Not rsTemp.EOF Then
         RegistroAtual = rsMDFe!idMDFe
         Status = 2
         rsMDFe.MoveFirst
         rsMDFe.Find "idMDFe = " & rsTemp!idMDFe
         Prencher_Campos
      End If
   End If

   If Status = 4 Then
      RegistroAtual = IIf(rsMDFe.EOF, 0, rsMDFe!idMDFe)
      If mskNumero.Text = Empty Then
         MsgBox "Digite um Número para a Consulta", vbOKOnly, "Localizar"
         If mskNumero.Enabled Then mskNumero.SetFocus
         Exit Sub
      End If
      rsMDFe.MoveFirst
      rsMDFe.Find "Numero Like " & Trim(mskNumero.Text)
      If Not rsMDFe.EOF Then
         Botoes 1, frmMDFe
         Prencher_Campos
         Campos False
      Else
         MsgBox "Número não Encontrado", vbOKOnly + vbExclamation, "Localizar"
         mskNumero.SetFocus
         mskNumero.SelStart = 0
         mskNumero.SelLength = Len(mskNumero.Text)
         rsMDFe.MoveFirst
         If RegistroAtual <> 0 Then rsMDFe.Find "idMDFe = " & RegistroAtual
      End If
   End If

   rsMDFes!Numero = mskNumero.Text
End Sub

Sub Excluir()
On Error GoTo ErroIntegridade
   If MsgBox("Confirma Excluir o registro atual? ", vbYesNo + vbInformation, "Excluir") = vbYes Then
      Atividade "Exclusão: " & mskNumero.Text, Me.Caption
'      cnSistema.Execute "Delete * from MDFeItens Where idMDFe = " & rsMDFe!idMDFe  ' Itens da Nota de Entrada
      cnSistema.Execute "Delete from MDFe Where idMDFe=" & rsMDFe!idMDFe           ' Nota de Entrada
      rsMDFe.Requery
'      If Registros(cnSistema, "MDFe") = 0 Then
      If Registros2("MDFe") = 0 Then
         Limpa_Campos
      Else
         Prencher_Campos
      End If
   End If

On Error GoTo 0
Exit Sub
ErroIntegridade:
   If Err.Number = 0 Then
      ' Operação Ok
   ElseIf Err.Number = -2147467259 Then
      Beep
      MsgBox "Não é possível Excluir este Registro" & Chr(13) & "Existe lançamentos relacionados com este Registro", vbInformation + vbOKOnly, "Excluir"
      Exit Sub
   Else
      Beep
      MsgBox "Verificar: " & Err.Number & Chr(13) & Err.Description, vbExclamation, "Excluir"
      Exit Sub
   End If
End Sub

Sub Alterar()
   Status = 2
   Botoes 2, frmMDFe
   Campos True
   mskDataEmissao.SetFocus
End Sub

Sub Localizar()
   Campos True
   Botoes 4, frmMDFe
   Limpa_Campos
   mskNumero.SetFocus
End Sub

Private Function Verifica_Campos()
Dim strMensagem As String
Verifica_Campos = True

   If mskNumero.Text = Empty Then strMensagem = strMensagem & "Número" & Chr(13)
   If Not IsDate(mskDataEmissao.Text) Or Val(Mid(mskDataEmissao.Text, 7, 4)) < 1900 Then strMensagem = strMensagem & "Data da Emissão" & Chr(13)
   If Not IsDate(mskDataViagem.Text) Or Val(Mid(mskDataViagem.Text, 7, 4)) < 1900 Then strMensagem = strMensagem & "Data da Viagem" & Chr(13)

   If cboUF.ListIndex = -1 Then strMensagem = strMensagem & "UF" & Chr(13)
   If cboTipoEmitente.ListIndex = -1 Then strMensagem = strMensagem & "Tipo de Emitente" & Chr(13)
   If cboTipoTransportador.ListIndex = -1 Then strMensagem = strMensagem & "Tipo de Transporte" & Chr(13)
   If cboFormaEmissao.ListIndex = -1 Then strMensagem = strMensagem & "Forma de Emissão" & Chr(13)
   If cboModalidade.ListIndex = -1 Then strMensagem = strMensagem & "Modalidade" & Chr(13)

   If cboUFDescarregamento.ListIndex = -1 Then strMensagem = strMensagem & "UF de Descarregamento" & Chr(13)

   If cboTipoCarroceria.ListIndex = -1 Then strMensagem = strMensagem & "Tipo de Carroceria" & Chr(13)
   If cboUFVeiculo.ListIndex = -1 Then strMensagem = strMensagem & "UF Veículo" & Chr(13)
   If cboTipoRodado.ListIndex = -1 Then strMensagem = strMensagem & "Tipo de Rodado" & Chr(13)

''''   If cmbNaturezaOperacao.ListIndex >= 0 And cmbCliente.ListIndex >= 0 Then
''''      Set rsClientes = cnSistema.Execute("Select * From Clientes Where idCliente = " & cmbCliente.ItemData(cmbCliente.ListIndex))
''''      Set rsTemp = cnSistema.Execute("Select * From NaturezasOperacao Where idNaturezaOperacao = " & cmbNaturezaOperacao.ItemData(cmbNaturezaOperacao.ListIndex))
''''      If rsClientes!UF = rsEmpresa!UF Then
''''         mskCFOP.Text = rsTemp!CFOPDentroUF
''''      Else
''''         mskCFOP.Text = rsTemp!CFOPForaUF
''''      End If
''''   End If

   If Not strMensagem = Empty Then
      MsgBox "Verifique os Seguintes Campos:" & Chr(13) & strMensagem, vbExclamation + vbOKOnly, "Campos Obrigatórios"
      Verifica_Campos = False
      Exit Function
   End If

End Function

Sub Campos(Parametro As Boolean)
   mskNumero.Enabled = Parametro
   mskDataEmissao.Enabled = Parametro
   mskDataViagem.Enabled = Parametro
   cboUF.Enabled = Parametro
   cboTipoEmitente.Enabled = Parametro
   cboTipoTransportador.Enabled = Parametro
   cboFormaEmissao.Enabled = Parametro
   cboModalidade.Enabled = Parametro

   cboUFCarregamento.Enabled = Parametro
   cboMunicipio.Enabled = Parametro
   cmdInserirLocalCarregamento.Enabled = Parametro
   cmdInserirLocalCarregamento.Enabled = Parametro
   lvwLocalCarregamento.Enabled = Parametro
   cboUFDescarregamento.Enabled = Parametro

   cboUFPercurso.Enabled = Parametro
   cmdInserirPercurso.Enabled = Parametro
   cmdExcluirPercurso.Enabled = Parametro
   lvwPercurso.Enabled = Parametro

   cboTipoCarroceria.Enabled = Parametro
   cboUFVeiculo.Enabled = Parametro
   mskPlacaVeiculo.Enabled = Parametro
   txtTara.Enabled = Parametro
   cboTipoRodado.Enabled = Parametro
   txtCapacidadeKG.Enabled = Parametro
   txtCapacidadeM3.Enabled = Parametro
   txtRenavam.Enabled = Parametro

   mskCPFCondutor.Enabled = Parametro
   txtNomeCondutor.Enabled = Parametro
   cmdIncluirCondutor.Enabled = Parametro
   cmdExcluirCondutor.Enabled = Parametro
   lvwCondutores.Enabled = Parametro

   txtNumeroNota.Enabled = Parametro
   cmdPesquisaNFes.Enabled = Parametro
   txtChaveNFe.Enabled = Parametro
   cmdIncluirNFe.Enabled = Parametro
   cmdExcluirNFe.Enabled = Parametro
   lvwNFes.Enabled = Parametro

   txtDadosAdicionais.Enabled = Parametro

End Sub

Sub Limpa_Campos()

   ' Adicionar MDFes RecorSets Temporários
   Call CriarEstruturaMDFe

   rsMDFes.AddNew
   rsMDFes.Update

   If Status <> 4 Then
      If Not rsMDFe.BOF Or Not rsMDFe.EOF Then
         rsMDFe.MoveLast
         mskNumero.Text = rsMDFe!Numero + 1
      Else
         mskNumero.Text = 1
      End If
   Else
      mskNumero.Text = Empty
   End If

   mskDataEmissao.Text = Date
   mskDataViagem.Text = Date
   cboUF.ListIndex = -1
   cboTipoEmitente.ListIndex = -1
   cboTipoTransportador.ListIndex = -1
   cboFormaEmissao.ListIndex = -1
   cboModalidade.ListIndex = -1

   cboUFCarregamento.ListIndex = -1
   cboMunicipio.ListIndex = -1
   lvwLocalCarregamento.ListItems.Clear
   cboUFDescarregamento.ListIndex = -1

   cboUFPercurso.ListIndex = -1
   lvwPercurso.ListItems.Clear

   cboTipoCarroceria.ListIndex = -1
   cboUFVeiculo.ListIndex = -1
   mskPlacaVeiculo.Text = "   -    "
   txtTara.Text = Empty
   cboTipoRodado.ListIndex = -1
   txtCapacidadeKG.Text = Empty
   txtCapacidadeM3.Text = Empty
   txtRenavam.Text = Empty

   mskCPFCondutor.Text = Empty
   txtNomeCondutor.Text = Empty
   lvwCondutores.ListItems.Clear

   txtNumeroNota.Text = Empty
   txtChaveNFe.Text = Empty
   lvwNFes.ListItems.Clear

   txtDadosAdicionais.Text = Empty

   ' Atualizar Dados para gravação
   rsMDFes!Numero = mskNumero.Text
   rsMDFes!DataEmissao = mskDataEmissao.Text
   rsMDFes!DataViagem = mskDataViagem.Text

End Sub

Sub Prencher_Campos()
Dim Contador As Integer

   ' Adicionar MDFes RecorSets Temporários
   Call CriarEstruturaMDFe

   rsMDFes.AddNew
   rsMDFes.Update

   txtChaveAcesso.Text = IIf(Trim(rsMDFe!ChaveMDFe) = "" Or IsNull(rsMDFe!ChaveMDFe), Empty, rsMDFe!ChaveMDFe)
   rsMDFes!ChaveMDFe = IIf(Trim(rsMDFe!ChaveMDFe) = "" Or IsNull(rsMDFe!ChaveMDFe), Empty, rsMDFe!ChaveMDFe)
   txtProtocolo.Text = IIf(Trim(rsMDFe!Protocolo) = "" Or IsNull(rsMDFe!Protocolo), Empty, rsMDFe!Protocolo)
   rsMDFes!Protocolo = IIf(Trim(rsMDFe!Protocolo) = "" Or IsNull(rsMDFe!Protocolo), Empty, rsMDFe!Protocolo)

   mskNumero.Text = IIf(Trim(rsMDFe!Numero) = "" Or IsNull(rsMDFe!Numero), Empty, rsMDFe!Numero)
   rsMDFes!Numero = IIf(Trim(rsMDFe!Numero) = "" Or IsNull(rsMDFe!Numero), Empty, rsMDFe!Numero)
   mskDataEmissao.Text = IIf(IsNull(rsMDFe!DataEmissao), "  /  /    ", rsMDFe!DataEmissao)
   rsMDFes!DataEmissao = IIf(IsNull(rsMDFe!DataEmissao), "  /  /    ", rsMDFe!DataEmissao)
   mskDataViagem.Text = IIf(IsNull(rsMDFe!DataViagem), "  /  /    ", rsMDFe!DataViagem)
   rsMDFes!DataViagem = IIf(IsNull(rsMDFe!DataViagem), "  /  /    ", rsMDFe!DataViagem)

   Call FUF(rsMDFe!idUF)
   Call FTipoEmitente(rsMDFe!idTipoEmitente)
   Call FTipoTransportador(rsMDFe!idTipoTransportador)
   Call FFormaEmissao(rsMDFe!idFormaEmissao)
   Call FModalidade(rsMDFe!idModalidade)

   ' Locais de Carregamento
   Call FLocalCarregamento(rsMDFe!idMDFe)

   Call FUFDescarregamento(rsMDFe!idUFDescarregamento)

   ' Percurso
   Call FPercurso(rsMDFe!idMDFe)

   Call FTipoCarroceria(rsMDFe!idTipoCarroceria)
   Call FUFVeiculo(rsMDFe!idUFVeiculo)

   mskPlacaVeiculo.Text = IIf(Trim(rsMDFe!PlacaVeiculo) = "" Or IsNull(rsMDFe!PlacaVeiculo), "   -    ", rsMDFe!PlacaVeiculo)
   rsMDFes!PlacaVeiculo = IIf(Trim(rsMDFe!PlacaVeiculo) = "" Or IsNull(rsMDFe!PlacaVeiculo), "   -    ", rsMDFe!PlacaVeiculo)
   txtTara.Text = IIf(Trim(rsMDFe!Tara) = "" Or IsNull(rsMDFe!Tara), Empty, rsMDFe!Tara)
   rsMDFes!Tara = IIf(Trim(rsMDFe!Tara) = "" Or IsNull(rsMDFe!Tara), Empty, rsMDFe!Tara)

   Call FTipoRodado(rsMDFe!idTipoRodado)

   txtCapacidadeKG.Text = IIf(Trim(rsMDFe!CapacidadeKG) = "" Or IsNull(rsMDFe!CapacidadeKG), Empty, rsMDFe!CapacidadeKG)
   rsMDFes!CapacidadeKG = IIf(Trim(rsMDFe!CapacidadeKG) = "" Or IsNull(rsMDFe!CapacidadeKG), Empty, rsMDFe!CapacidadeKG)
   txtCapacidadeM3.Text = IIf(Trim(rsMDFe!CapacidadeM3) = "" Or IsNull(rsMDFe!CapacidadeM3), Empty, rsMDFe!CapacidadeM3)
   rsMDFes!CapacidadeM3 = IIf(Trim(rsMDFe!CapacidadeM3) = "" Or IsNull(rsMDFe!CapacidadeM3), Empty, rsMDFe!CapacidadeM3)
   txtRenavam.Text = IIf(Trim(rsMDFe!Renavam) = "" Or IsNull(rsMDFe!Renavam), Empty, rsMDFe!Renavam)
   rsMDFes!Renavam = IIf(Trim(rsMDFe!Renavam) = "" Or IsNull(rsMDFe!Renavam), Empty, rsMDFe!Renavam)

   ' Percurso
   Call FCondutores(rsMDFe!idMDFe)

   ' Percurso
   Call FNFes(rsMDFe!idMDFe)

   txtDadosAdicionais.Text = IIf(Trim(rsMDFe!DadosAdicionais) = "" Or IsNull(rsMDFe!DadosAdicionais), Empty, rsMDFe!DadosAdicionais)
   rsMDFes!DadosAdicionais = IIf(Trim(rsMDFe!DadosAdicionais) = "" Or IsNull(rsMDFe!DadosAdicionais), Empty, rsMDFe!DadosAdicionais)


'   Set rsTemp = cnSistema.Execute("SELECT * FROM NFeItens WHERE NFeItens.idNFe = " & rsNFe!idNFe & " ORDER BY idNFeItem")
'
'   lvwProdutos.ListItems.Clear
'   Do While Not rsTemp.EOF
'      Set rsProdutos = cnSistema.Execute("SELECT * FROM Produtos WHERE Produtos.idProduto = " & rsTemp!idProduto)
'      If Not rsProdutos.EOF Then
'         cValorBruto = (rsTemp!quantidade * rsTemp!ValorUnitario)
'         cValorDesconto = (((rsTemp!quantidade * rsTemp!ValorUnitario) * rsTemp!Desconto) / 100)
'         cValorBonificacao = (((cValorBruto - cValorDesconto) * rsNFe!Bonificacao) / 100)
'         cValorLiquido = (cValorBruto - cValorDesconto - cValorBonificacao)
'
''         Set ItemList = lvwProdutos.ListItems.Add(, "R" & CStr(rsTemp!idProduto), rsProdutos!Codigo)
'         Set ItemList = lvwProdutos.ListItems.Add(, "R" & CStr(rsTemp!idNFeItem), rsProdutos!Codigo)
'         ItemList.SubItems(1) = Trim(rsProdutos!Descricao) & " " & Trim(rsTemp!DescricaoComplementar)
'         ItemList.SubItems(2) = Format(rsTemp!quantidade, mskQuantidade.Format)
'         ItemList.SubItems(3) = Format(rsTemp!ValorUnitario, mskUnitario.Format)
'         ItemList.SubItems(4) = Format(rsTemp!Desconto, "##,##0.00")
'         ItemList.SubItems(5) = Format((rsTemp!quantidade * rsTemp!ValorUnitario), "##,##0.00")
''         ItemList.SubItems(6) = Format((rsTemp!Quantidade * rsTemp!ValorUnitario) * (1 - ((rsTemp!Desconto) / 100)), "##,##0.00")
'         ItemList.SubItems(6) = Format(cValorLiquido, "##,##0.00")
'      End If
'
'      rsTemp.MoveNext
'   Loop
'
''  Boletos
'   Set rsTemp = cnSistema.Execute("SELECT * FROM NFeBoletos WHERE NFeBoletos.idNFe = " & rsNFe!idNFe)
'
'   lvwBoletos.ListItems.Clear
'   Do While Not rsTemp.EOF
'      Set ItemList = lvwBoletos.ListItems.Add(, "R" & rsTemp!Numero, rsTemp!Numero)
'      ItemList.SubItems(1) = rsTemp!Vencimento
'      ItemList.SubItems(2) = Format(rsTemp!Valor, "##,##0.00")
'
'      rsTemp.MoveNext
'   Loop
'
''  Total da Nota
'   Set rsTemp = cnSistema.Execute("Select * From TotalNFe Where Numero = " & mskNumero.text)
'   If Not rsTemp.EOF Then
'      mskValorTotalProdutos.text = Format(rsTemp!Total, "###,##0.00")
'      mskValorFrete.text = Format(rsTemp!TotalFrete, "###,##0.00")
'      mskValorTotalNota.text = Format(rsTemp!Total + IIf(Not IsNull(rsTemp!TotalFrete), rsTemp!TotalFrete, 0) + rsNFe!OutrasDespesas, "###,##0.00")
'      mskBaseCalculoICMS.text = Format(IIf(rsTemp!ValorICMS > 0, rsTemp!BaseCalculo, 0), "###,##0.00")
'      mskValorICMS.text = Format(rsTemp!ValorICMS, "###,##0.00")
''      mskVolumePesoBruto.Text = Format(rsNFe!VolumePesoBruto, "###,##0.00")
''      mskVolumePesoLiquido.Text = Format(rsNFe!VolumePesoLiquido, "###,##0.00")
'
'''''      mskVolumePesoBruto.Text = Format(rsTemp!PesoBruto, "###,##0.00")
'''''      mskVolumePesoLiquido.Text = Format(rsTemp!PesoLiquido, "###,##0.00")
'   Else
'      mskValorTotalProdutos.text = Format(0, "###,##0.00")
'      mskValorFrete.text = Format(0, "###,##0.00")
'      mskValorTotalNota.text = Format(0, "###,##0.00")
'      mskBaseCalculoICMS.text = Format(0, "###,##0.00")
'      mskValorICMS.text = Format(0, "###,##0.00")
''      mskVolumePesoBruto.Text = Format(0, "###,##0.00")
''      mskVolumePesoLiquido.Text = Format(0, "###,##0.00")
'   End If

End Sub

Sub Gravar()
Dim sSql As String

   If Not Verifica_Campos() Then Exit Sub

   Select Case Status
      Case 1 'Inclusão
         If MsgBox("Confirma Incluir o registro atual", vbYesNo + vbQuestion, "Inclusão") = vbYes Then
            sSql = Montar_Insert
            cnSistema.Execute sSql

            Atividade "Inclusão: " & mskNumero.Text, Me.Caption
            rsMDFe.Requery
            rsMDFe.Find "Numero = '" & mskNumero.Text & "'"
            
            ' Gravar itens
            GravarItens (rsMDFe!idMDFe)
         End If

      Case 2 'Alteracão
         If MsgBox("Confirma Alterar o registro atual", vbYesNo + vbQuestion, "Alteração") = vbYes Then

            sSql = Montar_Update
            cnSistema.Execute sSql

'            Atividade "Alterar: " & mskNumero.Text, Me.Caption
            rsMDFe.Requery
            rsMDFe.Find "Numero = '" & mskNumero.Text & "'"
            
            ' Gravar itens
            GravarItens (rsMDFe!idMDFe)
         End If
   End Select
   Prencher_Campos
   Botoes 1, frmMDFe
   Campos False
   Status = 0
   SSTab1.Tab = 0 ' Posiciona no primeiro tab
''   txtCodigo.SetFocus
End Sub

Private Function Montar_Insert() As String
Dim sSql As String

   sSql = ""
   sSql = sSql & "Insert Into MDFe ("
   sSql = sSql & vbCrLf & "  Numero, "
   sSql = sSql & vbCrLf & "  DataEmissao, "
   sSql = sSql & vbCrLf & "  DataViagem, "
   sSql = sSql & vbCrLf & "  idUF, "
   sSql = sSql & vbCrLf & "  idTipoEmitente, "
   sSql = sSql & vbCrLf & "  idTipoTransportador, "
   sSql = sSql & vbCrLf & "  idFormaEmissao, "
   sSql = sSql & vbCrLf & "  idModalidade, "
   sSql = sSql & vbCrLf & "  idUFDescarregamento, "
   sSql = sSql & vbCrLf & "  idTipoCarroceria, "
   sSql = sSql & vbCrLf & "  idUFVeiculo, "
   sSql = sSql & vbCrLf & "  PlacaVeiculo, "
   sSql = sSql & vbCrLf & "  Tara, "
   sSql = sSql & vbCrLf & "  idTipoRodado, "
   sSql = sSql & vbCrLf & "  CapacidadeKG, "
   sSql = sSql & vbCrLf & "  CapacidadeM3, "
   sSql = sSql & vbCrLf & "  Renavam, "
   sSql = sSql & vbCrLf & "  DadosAdicionais"
   sSql = sSql & vbCrLf & "         )"

   sSql = sSql & vbCrLf & "Values ("
   sSql = sSql & vbCrLf & rsMDFes!Numero & ", "
   sSql = sSql & vbCrLf & "'" & rsMDFes!DataEmissao & "', "
   sSql = sSql & vbCrLf & "'" & rsMDFes!DataViagem & "', "
   sSql = sSql & vbCrLf & rsMDFes!idUF & ", "
   sSql = sSql & vbCrLf & rsMDFes!idTipoEmitente & ", "
   sSql = sSql & vbCrLf & rsMDFes!idTipoTransportador & ", "
   sSql = sSql & vbCrLf & rsMDFes!idFormaEmissao & ", "
   sSql = sSql & vbCrLf & rsMDFes!idModalidade & ", "
   sSql = sSql & vbCrLf & rsMDFes!idUFDescarregamento & ", "
   sSql = sSql & vbCrLf & rsMDFes!idTipoCarroceria & ", "
   sSql = sSql & vbCrLf & rsMDFes!idUFVeiculo & ", "
   sSql = sSql & vbCrLf & "'" & rsMDFes!PlacaVeiculo & "', "
   sSql = sSql & vbCrLf & "'" & rsMDFes!Tara & "', "
   sSql = sSql & vbCrLf & rsMDFes!idTipoRodado & ", "
   sSql = sSql & vbCrLf & "'" & rsMDFes!CapacidadeKG & "', "
   sSql = sSql & vbCrLf & "'" & rsMDFes!CapacidadeM3 & "', "
   sSql = sSql & vbCrLf & "'" & rsMDFes!Renavam & "', "
   sSql = sSql & vbCrLf & "'" & rsMDFes!DadosAdicionais & "'"
   sSql = sSql & vbCrLf & ")"

   Montar_Insert = sSql

End Function

Private Function Montar_Update() As String
Dim sSql As String

   sSql = ""
   sSql = sSql & "Update MDFe set "
   sSql = sSql & vbCrLf & "  Numero = " & rsMDFes!Numero & ", "
   sSql = sSql & vbCrLf & "  DataEmissao = '" & rsMDFes!DataEmissao & "', "
   sSql = sSql & vbCrLf & "  DataViagem = '" & rsMDFes!DataViagem & "', "
   sSql = sSql & vbCrLf & "  idUF = " & rsMDFes!idUF & ", "
   sSql = sSql & vbCrLf & "  idTipoEmitente = " & rsMDFes!idTipoEmitente & ", "
   sSql = sSql & vbCrLf & "  idTipoTransportador = " & rsMDFes!idTipoTransportador & ", "
   sSql = sSql & vbCrLf & "  idFormaEmissao = " & rsMDFes!idFormaEmissao & ", "
   sSql = sSql & vbCrLf & "  idModalidade = " & rsMDFes!idModalidade & ", "
   sSql = sSql & vbCrLf & "  idUFDescarregamento = " & rsMDFes!idUFDescarregamento & ", "
   sSql = sSql & vbCrLf & "  idTipoCarroceria = " & rsMDFes!idTipoCarroceria & ", "
   sSql = sSql & vbCrLf & "  idUFVeiculo = " & rsMDFes!idUFVeiculo & ", "
   sSql = sSql & vbCrLf & "  PlacaVeiculo = '" & rsMDFes!PlacaVeiculo & "', "
   sSql = sSql & vbCrLf & "  Tara = '" & rsMDFes!Tara & "', "
   sSql = sSql & vbCrLf & "  idTipoRodado = " & rsMDFes!idTipoRodado & ", "
   sSql = sSql & vbCrLf & "  CapacidadeKG = '" & rsMDFes!CapacidadeKG & "', "
   sSql = sSql & vbCrLf & "  CapacidadeM3 = '" & rsMDFes!CapacidadeM3 & "', "
   sSql = sSql & vbCrLf & "  Renavam = '" & rsMDFes!Renavam & "', "
   sSql = sSql & vbCrLf & "  DadosAdicionais = '" & rsMDFes!DadosAdicionais & "'"
   sSql = sSql & vbCrLf & " Where idMDFe = " & rsMDFe!idMDFe

   Montar_Update = sSql

End Function

Private Sub GravarItens(id As Long)

   ' Locais de Carregamento
   cnSistema.Execute "DELETE FROM MDFeLocalCarregamento WHERE idMDFe=" & id
   rsMDFeLocalCarregamento.MoveFirst
   Do While Not rsMDFeLocalCarregamento.EOF
      cnSistema.Execute "INSERT INTO MDFeLocalCarregamento (idMDFe,idUF,idMunicipio) " & _
                        "VALUES (" & id & "," & rsMDFeLocalCarregamento!idUF & "," & rsMDFeLocalCarregamento!idMunicipio & ")"
   
      rsMDFeLocalCarregamento.MoveNext
   Loop
   
   ' Percursos
   cnSistema.Execute "DELETE FROM MDFePercurso WHERE idMDFe=" & id
   rsMDFePercurso.MoveFirst
   Do While Not rsMDFePercurso.EOF
      cnSistema.Execute "INSERT INTO MDFePercurso (idMDFe,idUF) " & _
                        "VALUES (" & id & "," & rsMDFePercurso!idUF & ")"
   
      rsMDFePercurso.MoveNext
   Loop
   
   ' Condutores
   cnSistema.Execute "DELETE FROM MDFeCondutores WHERE idMDFe=" & id
   rsMDFeCondutores.MoveFirst
   Do While Not rsMDFeCondutores.EOF
      cnSistema.Execute "INSERT INTO MDFeCondutores (idMDFe,CPF,Nome) " & _
                        "VALUES (" & id & ",'" & rsMDFeCondutores!CPF & "','" & rsMDFeCondutores!Nome & "')"
   
      rsMDFeCondutores.MoveNext
   Loop
   
   ' NFes
   cnSistema.Execute "DELETE FROM MDFeNFes WHERE idMDFe=" & id
   rsMDFeNFes.MoveFirst
   Do While Not rsMDFeNFes.EOF
      cnSistema.Execute "INSERT INTO MDFeNFes (idMDFe,Numero,ChaveNFe) " & _
                        "VALUES (" & id & ",'" & rsMDFeNFes!Numero & "','" & rsMDFeNFes!ChaveNFe & "')"
   
      rsMDFeNFes.MoveNext
   Loop
                              
''   Set rsSistema = Nothing
End Sub


'Private Sub Carrega_Combos()
'
''  Clientes
'
''   Set rsTemp = cnSistema.Execute("Select * from Clientes Where Situacao = 0 Order By Nome")
'   Set rsTemp = cnSistema.Execute("Select * from Clientes Order By Nome")
'   cmbCliente.Clear
'   Do While Not rsTemp.EOF
'      cmbCliente.AddItem rsTemp!Nome
'      cmbCliente.ItemData(cmbCliente.NewIndex) = rsTemp!idCliente
'      rsTemp.MoveNext
'   Loop
'
''  Naturezas de Operação
'
'   Set rsTemp = cnSistema.Execute("Select * from NaturezasOperacao ORDER BY Descricao")
'   cmbNaturezaOperacao.Clear
'   Do While Not rsTemp.EOF
'      cmbNaturezaOperacao.AddItem rsTemp!Descricao
'      cmbNaturezaOperacao.ItemData(cmbNaturezaOperacao.NewIndex) = rsTemp!idNaturezaOperacao
'      rsTemp.MoveNext
'   Loop
'
''  Transportadores
'
'   Set rsTemp = cnSistema.Execute("Select * from Transportadores ORDER BY Nome")
'   cmbTransportador.Clear
'   Do While Not rsTemp.EOF
'      cmbTransportador.AddItem rsTemp!Nome
'      cmbTransportador.ItemData(cmbTransportador.NewIndex) = rsTemp!idTransportador
'      rsTemp.MoveNext
'   Loop
'
''  Formas de Pagamento
'
'   Set rsTemp = cnSistema.Execute("Select * from FormasPagamento ORDER BY Descricao")
'   cmbFormaPagamento.Clear
'   Do While Not rsTemp.EOF
'      cmbFormaPagamento.AddItem rsTemp!Descricao
'      cmbFormaPagamento.ItemData(cmbFormaPagamento.NewIndex) = rsTemp!idFormaPagamento
'      rsTemp.MoveNext
'   Loop
'
''  UFs
'   Set rsTemp = cnSistema.Execute("Select * from UFs ORDER BY Sigla")
'   cmbUFPlaca.Clear
'   Do While Not rsTemp.EOF
'      cmbUFPlaca.AddItem rsTemp!Sigla
'      cmbUFPlaca.ItemData(cmbUFPlaca.NewIndex) = rsTemp!idUF
'      rsTemp.MoveNext
'   Loop
'
''  Frete Conta
'   Set rsTemp = cnSistema.Execute("Select * from FreteConta ORDER BY Descricao")
'   cmbFreteConta.Clear
'   Do While Not rsTemp.EOF
'      cmbFreteConta.AddItem rsTemp!Descricao
'      cmbFreteConta.ItemData(cmbFreteConta.NewIndex) = rsTemp!idFreteConta
'      rsTemp.MoveNext
'   Loop
'
'   rsTemp.Close
'End Sub
'
'Private Sub cmdExcluir_Click()
'   Beep
'   If MsgBox("Deseja excluir este item", vbYesNo + vbQuestion, "Excluir") = vbYes Then
'      cnSistema.Execute "Update NFe set " & _
'            "ValorTotalProdutos = '" & mskValorTotalProdutos.text & "', " & _
'            "ValorTotalNota = '" & mskValorTotalNota.text & "' " & _
'            "Where idNFe = " & rsNFe!idNFe
'
'''      cnSistema.Execute "Delete from NFeItens Where idNFe = " & rsNFe!idNFe & " And idProduto = " & Val(Mid(lvwProdutos.SelectedItem.Key, 2, Len(lvwProdutos.SelectedItem.Key)))
'      cnSistema.Execute "Delete from NFeItens Where idNFe = " & rsNFe!idNFe & " And idNFeItem = " & Val(Mid(lvwProdutos.SelectedItem.Key, 2, Len(lvwProdutos.SelectedItem.Key)))
'      lvwProdutos.ListItems.Remove (lvwProdutos.SelectedItem.Index)
'
'      Set rsTemp = cnSistema.Execute("Select * From TotalNFe Where Numero = " & mskNumero.text)
'      If Not rsTemp.EOF Then
'         mskValorTotalProdutos.text = Format(rsTemp!Total, "###,##0.00")
'         mskValorTotalNota.text = Format(rsTemp!Total, "###,##0.00")
'         mskBaseCalculoICMS.text = Format(rsTemp!BaseCalculo, "###,##0.00")
'         mskValorICMS.text = Format(rsTemp!ValorICMS, "###,##0.00")
'         mskVolumePesoBruto.text = Format(rsTemp!PesoBruto, "###,##0.00")
'         mskVolumePesoLiquido.text = Format(rsTemp!PesoLiquido, "###,##0.00")
'      Else
'         mskValorTotalProdutos.text = Format(0, "###,##0.00")
'         mskValorTotalNota.text = Format(0, "###,##0.00")
'         mskBaseCalculoICMS.text = Format(0, "###,##0.00")
'         mskValorICMS.text = Format(0, "###,##0.00")
'         mskVolumePesoBruto.text = Format(0, "###,##0.00")
'         mskVolumePesoLiquido.text = Format(0, "###,##0.00")
'      End If
'   End If
'End Sub

'Private Sub cmdIncluir_Click()
'   If lvwProdutos.ListItems.Count = intItensNota Then
'      Beep
'      MsgBox "Total de Itens da Nota Excedido", vbExclamation, "Erro"
'
'      txtCodigo.text = Empty
'      txtProduto.text = Empty
'      mskQuantidade.text = Empty
'      mskDesconto.text = Empty
'      mskUnitario.text = Empty
'      txtCodigo.SetFocus
'
'      Exit Sub
'   End If
'
'   If Verifica_Campos_Produtos() Then
'      Set rsTemp = cnSistema.Execute("Select * From Produtos Where Codigo = '" & txtCodigo.text & "'")
'      If Not rsTemp!Marca Then
'         If mskDesconto.text <> "" Then
'            mskDesconto.text = Val(Substitui(mskDesconto.text, ",", "."))
'         Else
'            mskDesconto.text = Val(Substitui(mskDescontoGeral.text, ",", "."))
'         End If
'      Else
'         mskDesconto.text = 0
'      End If
'
''      Set ItemList = lvwProdutos.ListItems.Add(, "R" & CStr(rsTemp!idProduto), txtCodigo.text)
''      ItemList.SubItems(1) = Trim(txtProduto.text) & " " & Trim(frmNFeComplemento.txtDescricaoComplementar.text)
''      ItemList.SubItems(2) = Format(mskQuantidade.text, mskQuantidade.Format)
''      ItemList.SubItems(3) = Format(mskUnitario.text, mskUnitario.Format)
''      ItemList.SubItems(4) = Format(mskDesconto.text, "##0.00")
''      ItemList.SubItems(5) = Format((mskQuantidade.text * mskUnitario.text), "##,##0.00")
''      ItemList.SubItems(6) = Format((mskQuantidade.text * mskUnitario.text) * (1 - ((Val(Substitui(mskDesconto.text, ",", ".")) + Val(Substitui(mskBonificacao.text, ",", "."))) / 100)), "##0.00")
'
'      mskQuantidade.text = Val(Substitui(mskQuantidade.text, ",", "."))
'      mskDesconto.text = Val(Substitui(mskDesconto.text, ",", "."))
'      mskUnitario.text = Val(Substitui(mskUnitario.text, ",", "."))
'      frmNFeComplemento.mskBaseReduzidaICMS.text = Val(Substitui(frmNFeComplemento.mskBaseReduzidaICMS.text, ",", "."))
'
'      Dim iUnidade As Integer, iSituacaoTributaria As Integer
'
'      If frmNFeComplemento.cmbUnidade.ListIndex = -1 Then
'         iUnidade = 1
'      Else
'         iUnidade = frmNFeComplemento.cmbUnidade.ItemData(frmNFeComplemento.cmbUnidade.ListIndex)
'      End If
'
'      If frmNFeComplemento.cmbSituacaoTributaria.ListIndex = -1 Then
'         iSituacaoTributaria = rsTemp!idSituacaoTributaria
'      Else
'         iSituacaoTributaria = frmNFeComplemento.cmbSituacaoTributaria.ItemData(frmNFeComplemento.cmbSituacaoTributaria.ListIndex)
'      End If
'
'      If frmNFeComplemento.mskCFOP.text = " .   " Then frmNFeComplemento.mskCFOP.text = mskCFOP.text
'
'      If Val(Substitui(frmNFeComplemento.mskBaseReduzidaICMS.text, ",", ".")) = 0 Then
'         If Not rsEmpresa.EOF Then
'            If rsClientes!UF = rsEmpresa!UF Then
'               frmNFeComplemento.mskBaseReduzidaICMS.text = rsTemp!BaseReduzidaICMSdUF
'            Else
'               frmNFeComplemento.mskBaseReduzidaICMS.text = rsTemp!BaseReduzidaICMSfUF
'            End If
'         End If
'      End If
'
'      cnSistema.Execute "Insert Into NFeItens (idNFe,idProduto,Data,Quantidade,Desconto,ValorUnitario,ICMS,BaseReduzida,DescricaoComplementar,idUnidade,idSituacaoTributaria,DiscriminacaoProduto,IPI,BaseReduzidaIPI,ClassificacaoFiscal,ValorFrete,CFOP) " & _
'                        "Values (" & rsNFe!idNFe & "," & rsTemp!idProduto & ",'" & mskDataEmissao.text & _
'                        "','" & CStrValor(mskQuantidade.text) & "','" & CStrValor(mskDesconto.text) & "','" & CStrValor(mskUnitario.text) & "','" & _
'                        CStrValor(frmNFeComplemento.mskICMSProduto.text) & "','" & CStrValor(frmNFeComplemento.mskBaseReduzidaICMS.text) & "','" & SQLCheck(frmNFeComplemento.txtDescricaoComplementar.text) & "'," & iUnidade & "," & iSituacaoTributaria & ",'" & _
'                        SQLCheck(frmNFeComplemento.txtDiscriminacaoProduto.text) & "','" & CStrValor(frmNFeComplemento.mskIPIProduto.text) & "','" & CStrValor(frmNFeComplemento.mskBaseReduzidaIPI.text) & "','" & SQLCheck(frmNFeComplemento.txtClassificacaoFiscal.text) & "','" & CStrValor(frmNFeComplemento.mskValorFrete.text) & "','" & frmNFeComplemento.mskCFOP.text & "')"
'
'      cnSistema.Execute "Update NFe set " & _
'            "ValorTotalProdutos = '" & mskValorTotalProdutos.text & "', " & _
'            "ValorTotalNota = '" & mskValorTotalNota.text & "' " & _
'            "Where idNFe = " & rsNFe!idNFe
'
'      Set rsTemp = cnSistema.Execute("Select * From TotalNFe Where Numero = " & mskNumero.text)
'      If Not rsTemp.EOF Then
'         mskValorTotalProdutos.text = Format(rsTemp!Total, "###,##0.00")
'         mskValorFrete.text = Format(rsTemp!TotalFrete, "###,##0.00")
'         mskValorTotalNota.text = Format(rsTemp!Total + rsTemp!TotalFrete, "###,##0.00")
'         mskBaseCalculoICMS.text = Format(rsTemp!BaseCalculo, "###,##0.00")
'         mskValorICMS.text = Format(rsTemp!ValorICMS, "###,##0.00")
'         mskVolumePesoBruto.text = Format(rsTemp!PesoTotal, "###,##0.00")
'         mskVolumePesoLiquido.text = Format(rsTemp!PesoTotal, "###,##0.00")
'      Else
'         mskValorTotalProdutos.text = Format(0, "###,##0.00")
'         mskValorFrete.text = Format(0, "###,##0.00")
'         mskValorTotalNota.text = Format(0, "###,##0.00")
'         mskBaseCalculoICMS.text = Format(0, "###,##0.00")
'         mskValorICMS.text = Format(0, "###,##0.00")
'         mskVolumePesoBruto.text = Format(0, "###,##0.00")
'         mskVolumePesoLiquido.text = Format(0, "###,##0.00")
'      End If
'
'      ' Atualiza Corpo da Nota
'      Set rsCFOPs = cnSistema.Execute("Select * From CFOPs Where CFOP = '" & mskCFOP.text & "'")
'      Set rsProdutos = cnSistema.Execute("Select * From Produtos Where Codigo = '" & txtCodigo.text & "'")
'
'      If Not rsCFOPs.EOF Then
'         Set rsCFOPReferencias = cnSistema.Execute("Select * From CFOPReferencias Where idCFOP = " & rsCFOPs!idCFOP & " AND idProduto = " & rsProdutos!idProduto)
'         If Not rsCFOPReferencias.EOF Then
'            cnSistema.Execute "Update NFe set " & _
'                  "InformacoesCorpo = '" & rsCFOPReferencias!InformacoesCorpo & "' " & _
'                  "Where idNFe = " & rsNFe!idNFe
'         End If
'
''         Prencher_Campos
'      End If
'
'      Prencher_Campos
'
'      txtCodigo.text = Empty
'      txtProduto.text = Empty
'      mskQuantidade.text = Empty
'      mskDesconto.text = Empty
'      mskUnitario.text = Empty
'
'      frmNFeComplemento.mskICMSProduto.text = Empty
'      frmNFeComplemento.mskBaseReduzidaICMS.text = Empty
'      frmNFeComplemento.txtDescricaoComplementar.text = Empty
'      frmNFeComplemento.cmbUnidade.ListIndex = -1
'      frmNFeComplemento.cmbSituacaoTributaria.ListIndex = -1
'      frmNFeComplemento.txtDiscriminacaoProduto.text = Empty
'      frmNFeComplemento.mskIPIProduto.text = Empty
'      frmNFeComplemento.mskBaseReduzidaIPI.text = Empty
'      frmNFeComplemento.mskValorFrete.text = Empty
'      frmNFeComplemento.txtClassificacaoFiscal.text = Empty
'
'      txtCodigo.SetFocus
'   End If
'End Sub
'
'Private Function Verifica_Campos_Produtos()
'Dim strMensagem As String
'Dim ProcuraItem As ListItem
'Verifica_Campos_Produtos = True
'
'   If txtCodigo.text = Empty Then strMensagem = strMensagem & "Código" & Chr(13)
'   If txtProduto.text = Empty Then strMensagem = strMensagem & "Produto" & Chr(13)
'   If Val(Substitui(mskQuantidade.text, ",", ".")) = 0 Then strMensagem = strMensagem & "Quantidade" & Chr(13)
'
'   Set rsProdutos = cnSistema.Execute("SELECT * FROM Produtos WHERE Codigo = '" & txtCodigo.text & "'")
'   If Not rsProdutos.EOF Then
''      Set rsTemp = cnSistema.Execute("SELECT * FROM SaldoProdutos WHERE idProduto = " & rsProdutos!idProduto)
''      If rsTemp!Saldo < Val(mskQuantidade.Text) Then strMensagem = strMensagem & "Saldo de Produtos inferior a quantidade. Saldo Atual " & Round(rsTemp!Saldo, 2) & Chr(13)
'   End If
'
''   Set rsTemp = cnSistema.Execute("Select * From Clientes Where idCliente = " & cboCliente.ItemData(cmbCliente.ListIndex))
''   If Not rsTemp.EOF Then
''      If Len(Trim(rsTemp!Endereco)) = 0 Then strMensagem = strMensagem & "Endereço do cliente não cadastrado" & Chr(13)
''      If Len(Trim(rsTemp!Bairro)) = 0 Then strMensagem = strMensagem & "Bairro do cliente não cadastrado" & Chr(13)
''      If Len(Trim(RemoveCaracteres(rsTemp!CNPJ_CPF))) = 0 Then strMensagem = strMensagem & "CPF/CNPJ do cliente não cadastrado" & Chr(13)
''   End If
'
'   If Not strMensagem = Empty Then
'      Beep
'      MsgBox "Verifique os Seguintes Campos:" & Chr(13) & strMensagem, vbExclamation + vbOKOnly, "Campos Obrigatórios"
'      Verifica_Campos_Produtos = False
'      Exit Function
'   End If
'
'   Set ProcuraItem = lvwProdutos.FindItem(txtCodigo.text)
'   If Not ProcuraItem Is Nothing Then
'      Beep
'
'      If MsgBox("Produto já existe na Relação" & Chr(13) & "Deseja substituí-lo", vbYesNo + vbQuestion, "Produtos") = vbNo Then
'         Verifica_Campos_Produtos = False
'         txtCodigo.text = Empty
'         txtProduto.text = Empty
'         mskQuantidade.text = Empty
'         mskDesconto.text = Empty
'         mskUnitario.text = Empty
'         txtCodigo.SetFocus
'
'         Exit Function
'      Else
'         mskValorTotalProdutos.text = Format(rsNFe!ValorTotalProdutos - Val(Substitui(lvwProdutos.ListItems(lvwProdutos.SelectedItem.Index).ListSubItems(6), ",", ".")), "###,##0.00")
'         mskValorTotalNota.text = Format(rsNFe!ValorTotalNota - Val(Substitui(lvwProdutos.ListItems(lvwProdutos.SelectedItem.Index).ListSubItems(6), ",", ".")), "###,##0.00")
'
'         cnSistema.Execute "Update NFe set " & _
'               "ValorTotalProdutos = '" & mskValorTotalProdutos.text & "', " & _
'               "ValorTotalNota = '" & mskValorTotalNota.text & "' " & _
'               "Where idNFe = " & rsNFe!idNFe
'
'         Set rsTemp = cnSistema.Execute("Select * From Produtos Where Codigo = '" & txtCodigo.text & "'")
'         cnSistema.Execute "Delete from NFeItens Where idNFe = " & rsNFe!idNFe & " and idProduto = " & rsTemp!idProduto
'         lvwProdutos.ListItems.Remove ProcuraItem.Index
'      End If
'   End If
'End Function
'
'Private Sub cmdProduto_Click()
''   frmPesquisaProduto.Show vbModal
'
'   Registro_Selecionado = False
'   Screen.MousePointer = vbDefault
'   frmPesquisaProduto.Show vbModal
'   If Registro_Selecionado Then
'      Set rsTemp = cnSistema.Execute("Select * From Produtos Where idProduto = " & Val(Mid(frmPesquisaProduto.lvwDados.SelectedItem.Key, 2, Len(frmPesquisaProduto.lvwDados.SelectedItem.Key))))
'   End If
'
''   rsProdutos.MoveFirst
''   If Registro_Selecionado Then
''      rsProdutos.Find "idProduto = " & Val(Mid(frmPesquisaProduto.lvwDados.SelectedItem.Key, 2, Len(frmPesquisaProduto.lvwDados.SelectedItem.Key)))
''   End If
''
'   If frmPesquisaProduto.lvwDados.ListItems.Count <> 0 Then
''      txtCodigo.Text = Mid(frmPesquisaProduto.lvwDados.SelectedItem.Key, 2, Len(frmPesquisaProduto.lvwDados.SelectedItem.Key))
'      txtCodigo.text = rsTemp!Codigo
'      txtCodigo.SetFocus
'      Sendkeys "{TAB}"
'   End If
''
'End Sub
'
'Private Sub txtCodigo_LostFocus()
'   If txtCodigo.text <> Empty Then
'      Set rsNaturezasOperacao = cnSistema.Execute("Select * From NaturezasOperacao Where Descricao = '" & cmbNaturezaOperacao.text & "'")
'      Set rsTemp = cnSistema.Execute("Select * From Produtos Where Codigo = '" & SQLCheck(txtCodigo.text) & "'")
'      If Not rsTemp.EOF Then
'         If rsTemp!Situacao = 0 Then
'            Set rsTemp2 = cnSistema.Execute("Select * From NFeItens Where idNFe = " & rsNFe!idNFe & " And idProduto = " & rsTemp!idProduto)
'            If Not rsTemp2.EOF Then
'               mskUnitario.text = rsTemp2!ValorUnitario
'               mskQuantidade.text = rsTemp2!quantidade
'               mskDesconto.text = rsTemp2!Desconto
'            End If
'
'            txtCodigo.text = txtCodigo.text
'            txtProduto.text = rsTemp!Descricao
'            If rsNaturezasOperacao!Tipo <> 1 Then
'               mskUnitario.text = rsTemp!preco
'            Else
'               mskUnitario.text = rsTemp!ValorCusto
'            End If
'            mskQuantidade.SetFocus
'         Else
'            Beep
'            MsgBox "Este Produto está Inativo", vbOKOnly + vbInformation, "Produtos"
'            txtCodigo.SelStart = 0
'            txtCodigo.SelLength = Len(txtCodigo.text)
'            txtCodigo.SetFocus
'         End If
'      Else
'         Beep
'         MsgBox "Não existe Produto com este Código", vbOKOnly + vbInformation, "Produtos"
'         txtCodigo.SelStart = 0
'         txtCodigo.SelLength = Len(txtCodigo.text)
'         txtCodigo.SetFocus
'      End If
'   End If
'End Sub
'
'Private Sub cmbCliente_Click()
'   If cmbCliente.ListIndex <> -1 Then
'      Set rsTemp = cnSistema.Execute("Select * From Clientes Where idCliente = " & cmbCliente.ItemData(cmbCliente.ListIndex))
'      lblDescricaoEndereco.Caption = Trim(rsTemp!NomeFantasia) & Chr(13) & Trim(rsTemp!Endereco) & ", " & Trim(rsTemp!Bairro) & Chr(13) & Trim(rsTemp!Cidade) & " - " & rsTemp!UF & Chr(13) & "CEP: " & rsTemp!CEP & " Fone: " & rsTemp!Telefone1
'   End If
'End Sub
'
'Private Sub cmbCliente_LostFocus()
'   If Status = 1 Then
'      If mskNumero.text <> "" And cmbCliente.ListIndex >= 0 Then
'         Set rsTemp = cnSistema.Execute("Select * From NFe Where Numero = " & mskNumero.text & " and idCliente = " & cmbCliente.ItemData(cmbCliente.ListIndex))
'         If Not rsTemp.EOF Then
'            Beep
'            MsgBox "Nota fiscal já lançada", vbExclamation + vbOKOnly, "Localização"
'            mskNumero.text = "      "
'            mskNumero.SetFocus
'            Exit Sub
'         End If
'      End If
'
'      Set rsClientes = cnSistema.Execute("Select * From Clientes Where idCliente = " & cmbCliente.ItemData(cmbCliente.ListIndex))
'
'    ' Preencher Campos
'      If rsClientes!DescontoMaximo <> 0 Then
'         mskDescontoGeral.text = rsClientes!DescontoMaximo
'      End If
'
'    ' Pesquisa Clientes
'      Dim Contador As Integer
'      For Contador = 0 To (cmbFormaPagamento.ListCount - 1)
'         If cmbFormaPagamento.ItemData(Contador) = rsClientes!idFormaPagamento Then
'            cmbFormaPagamento.ListIndex = Contador
'            Exit For
'         End If
'      Next
'
'      Dim strMensagem As String
'      strMensagem = "Forma de Pagamento: " & cmbFormaPagamento.text & Chr(13)
'      strMensagem = strMensagem & "Desconto Máximo: " & rsClientes!DescontoMaximo
'
''      MsgBox "Verifique os Seguintes Campos:" & Chr(13) & strMensagem, vbExclamation + vbOKOnly, "Campos Obrigatórios"
'
'    ' Situação Pendente
'      If rsClientes!Situacao <> 0 Then
'         strMensagem = rsClientes!MotivoInativacao
'         MsgBox "Cliente Inativo e com situação Pendente: " & Chr(13) & strMensagem, vbExclamation + vbOKOnly, "Campos Obrigatórios"
'      End If
'   End If
'End Sub
'
'Private Sub mskValorFrete_LostFocus()
'   mskValorTotalNota.text = 0
'   mskValorTotalNota.text = Format(CStrValor(mskValorTotalNota.text) + CStrValor(mskValorTotalProdutos.text), "###,##0.00")
'   mskValorTotalNota.text = Format(CStrValor(mskValorTotalNota.text) + CStrValor(mskValorFrete.text), "###,##0.00")
'   mskValorTotalNota.text = Format(CStrValor(mskValorTotalNota.text) + CStrValor(mskOutrasDespesas.text), "###,##0.00")
'End Sub
'
'Private Sub mskOutrasDespesas_LostFocus()
'   mskValorTotalNota.text = 0
'   mskValorTotalNota.text = Format(CStrValor(mskValorTotalNota.text) + CStrValor(mskValorTotalProdutos.text), "###,##0.00")
'   mskValorTotalNota.text = Format(CStrValor(mskValorTotalNota.text) + CStrValor(mskValorFrete.text), "###,##0.00")
'   mskValorTotalNota.text = Format(CStrValor(mskValorTotalNota.text) + CStrValor(mskOutrasDespesas.text), "###,##0.00")
'End Sub
'
'Private Sub mskValorTotalProdutos_LostFocus()
'   mskValorTotalNota.text = 0
'   mskValorTotalNota.text = Format(CStrValor(mskValorTotalNota.text) + CStrValor(mskValorTotalProdutos.text), "###,##0.00")
'   mskValorTotalNota.text = Format(CStrValor(mskValorTotalNota.text) + CStrValor(mskValorFrete.text), "###,##0.00")
'   mskValorTotalNota.text = Format(CStrValor(mskValorTotalNota.text) + CStrValor(mskOutrasDespesas.text), "###,##0.00")
'End Sub
'
'Private Sub cmbNaturezaOperacao_LostFocus()
'   If cmbNaturezaOperacao.ListIndex >= 0 And cmbCliente.ListIndex >= 0 Then
'      Set rsClientes = cnSistema.Execute("Select * From Clientes Where idCliente = " & cmbCliente.ItemData(cmbCliente.ListIndex))
'      Set rsTemp = cnSistema.Execute("Select * From NaturezasOperacao Where idNaturezaOperacao = " & cmbNaturezaOperacao.ItemData(cmbNaturezaOperacao.ListIndex))
'      If Not rsEmpresa.EOF Then
'         If rsClientes!UF = rsEmpresa!UF Then
'            mskCFOP.text = rsTemp!CFOPDentroUF
'         Else
'            mskCFOP.text = rsTemp!CFOPForaUF
'         End If
'      End If
'   End If
'End Sub
'
'Private Sub txtProduto_KeyPress(KeyAscii As Integer)
'   If KeyAscii = 13 Then Sendkeys "{TAB}"
'End Sub
'
'Private Sub txtProduto_LostFocus()
'   If mskNumero.Enabled Then Exit Sub
'
'   If txtCodigo.text <> Empty And strDescricaoTemp = txtProduto.text Then Exit Sub
'   If txtCodigo.text = Empty And txtProduto.text = Empty Then
''      Beep
''      MsgBox "Digite o Código ou a Descrição do Produto", vbOKOnly + vbInformation, "Produtos"
''      txtCodigo.SetFocus
''      Exit Sub
''      Set rsTemp = cnSistema.Execute("Select * From Produtos Order By Descricao")
'      Set rsTemp = cnSistema.Execute("Select * From TabelaProdutos Where Situacao = 0 Order By Descricao")
'   Else
'      Set rsTemp = cnSistema.Execute("Select * From TabelaProdutos Where Situacao = 0 And Descricao Like '%" & txtProduto.text & "%' Order By Descricao")
'   End If
'
'   If rsTemp.EOF Then
'      Beep
'      MsgBox "Não existem Produtos com esta Descrição", vbOKOnly + vbInformation, "Produtos"
'      txtProduto.SelStart = 0
'      txtProduto.SelLength = Len(txtProduto.text)
'      txtProduto.SetFocus
'      Exit Sub
'   End If
'
'   Screen.MousePointer = vbHourglass
'   frmVisualiza.lvwDados.ColumnHeaders.Clear
'   frmVisualiza.lvwDados.ColumnHeaders.Add , , "Produto", 3680
'   frmVisualiza.lvwDados.ColumnHeaders.Add , , "Código", 970
'   frmVisualiza.lvwDados.ColumnHeaders.Add , , "Valor", 700, lvwColumnRight
'   frmVisualiza.lvwDados.ColumnHeaders.Add , , "Saldo", 700, lvwColumnRight
'   frmVisualiza.lvwDados.ListItems.Clear
'   Do While Not rsTemp.EOF
'      Set ItemList = frmVisualiza.lvwDados.ListItems.Add(, "R" & CStr(rsTemp!idProduto), rsTemp!Descricao)
'      ItemList.SubItems(1) = rsTemp!Codigo
'      ItemList.SubItems(2) = Format(rsTemp!preco, "###,##0.00")
'      ItemList.SubItems(3) = Format(rsTemp!Saldo, "###,##0.00")
'      rsTemp.MoveNext
'   Loop
'   Registro_Selecionado = False
'   Screen.MousePointer = vbDefault
'   frmVisualiza.Show vbModal
'   If Registro_Selecionado Then
'      Set rsNaturezasOperacao = cnSistema.Execute("Select * From NaturezasOperacao Where Descricao = '" & cmbNaturezaOperacao.text & "'")
'      Set rsTemp = cnSistema.Execute("Select * From Produtos Where idProduto = " & Val(Mid(frmVisualiza.lvwDados.SelectedItem.Key, 2, Len(frmVisualiza.lvwDados.SelectedItem.Key))), 5)
'      txtCodigo.text = rsTemp!Codigo
'      txtProduto.text = rsTemp!Descricao
'      If rsNaturezasOperacao!Tipo <> 1 Then
'         mskUnitario.text = rsTemp!preco
'      Else
'         mskUnitario.text = rsTemp!ValorCusto
'      End If
'
'    ' Carrega Combos
'      Set rsTemp = cnSistema.Execute("Select * from UnidadesMedida Order By Descricao")
'      frmNFeComplemento.cmbUnidade.Clear
'      Do While Not rsTemp.EOF
'         frmNFeComplemento.cmbUnidade.AddItem rsTemp!Descricao
'         frmNFeComplemento.cmbUnidade.ItemData(frmNFeComplemento.cmbUnidade.NewIndex) = rsTemp!idUnidadeMedida
'         rsTemp.MoveNext
'      Loop
'
'''    ' Preencher Campos
'''      Set rsTemp = cnSistema.Execute("Select * From Produtos Where Codigo = '" & SQLCheck(frmNFe.txtCodigo.Text) & "'")
'''      If Not rsTemp.EOF Then
'''         ' ICMS
'''         frmNFeComplemento.mskICMSProduto.Text = rsTemp!ICMS
'''         ' Unidade
'''         For Contador = 0 To (frmNFeComplemento.cmbUnidade.ListCount - 1)
'''            If frmNFeComplemento.cmbUnidade.ItemData(Contador) = rsTemp!idUnidade Then
'''               frmNFeComplemento.cmbUnidade.ListIndex = Contador
'''               Exit For
'''            End If
'''         Next
'''      End If
'   End If
'End Sub
'
'Private Sub mskCFOP_LostFocus()
'   Set rsTemp = cnSistema.Execute("Select * From CFOPs Where CFOP = '" & mskCFOP.text & "'")
'End Sub

'Private Function Verifica_Campos_Boletos()
'Dim strMensagem As String
'Dim ProcuraItem As ListItem
'Verifica_Campos_Boletos = True
'
'   If mskNumeroBoleto.text = Empty Then strMensagem = strMensagem & "Boleto" & Chr(13)
'   If mskVencimentoBoleto.text = Empty Then strMensagem = strMensagem & "Vencimento" & Chr(13)
'   If Val(Substitui(mskValorBoleto.text, ",", ".")) = 0 Then strMensagem = strMensagem & "Valor" & Chr(13)
'
'   If Not strMensagem = Empty Then
'      Beep
'      MsgBox "Verifique os Seguintes Campos:" & Chr(13) & strMensagem, vbExclamation + vbOKOnly, "Campos Obrigatórios"
'      Verifica_Campos_Boletos = False
'      Exit Function
'   End If
'
'   Set ProcuraItem = lvwBoletos.FindItem(mskNumeroBoleto.text)
'End Function
'
'Private Sub cmdPesquisaCliente_Click()
'   Registro_Selecionado = False
'   Screen.MousePointer = vbDefault
'   frmPesquisaClientes.Show vbModal
'   If Registro_Selecionado Then
'      Set rsTemp = cnSistema.Execute("Select * From Clientes Where idCliente = " & Val(Mid(frmPesquisaClientes.lvwDados.SelectedItem.Key, 2, Len(frmPesquisaClientes.lvwDados.SelectedItem.Key))))
'   End If
'
'   If frmPesquisaClientes.lvwDados.ListItems.Count <> 0 Then
'      mskCNPJ_CPF.text = rsTemp!CNPJ_CPF
'      mskCNPJ_CPF.SetFocus
'      Sendkeys "{TAB}"
'   End If
'   cmbCliente.SetFocus
'End Sub
'
'Private Function SaltarLinha(iParametro As Integer) As Integer
'Dim Contador As Integer
'
'   If iParametro > 0 Then
'      For Contador = 1 To (iParametro - 1)
'          Print #1, ""
'      Next
'   End If
'End Function
'
'Private Sub mskQuantidade_KeyPress(KeyAscii As Integer)
'   If KeyAscii = 13 Then Sendkeys "{TAB}"
'   If KeyAscii = 46 Then KeyAscii = 44
'   If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 44 And KeyAscii <> 8 Then
'      KeyAscii = 0
'   End If
'   If KeyAscii = 44 Then
'      If InStr(mskQuantidade.ClipText, ",") <> 0 Then
'         KeyAscii = 0
'      End If
'   End If
'   If Len(mskQuantidade.text) > 6 And KeyAscii <> 8 And KeyAscii <> 44 Then
'      If InStr(mskQuantidade.ClipText, ",") = 0 Then
'         KeyAscii = 0
'      End If
'   End If
'End Sub
'
'Private Sub mskDesconto_KeyPress(KeyAscii As Integer)
'   If KeyAscii = 13 Then Sendkeys "{TAB}"
'   If KeyAscii = 46 Then KeyAscii = 44
'   If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 44 And KeyAscii <> 8 Then
'      KeyAscii = 0
'   End If
'   If KeyAscii = 44 Then
'      If InStr(mskDesconto.ClipText, ",") <> 0 Then
'         KeyAscii = 0
'      End If
'   End If
'   If Len(mskDesconto.text) > 1 And KeyAscii <> 8 And KeyAscii <> 44 Then
'      If InStr(mskDesconto.ClipText, ",") = 0 Then
'         KeyAscii = 0
'      End If
'   End If
'End Sub
'
'Private Sub mskUnitario_KeyPress(KeyAscii As Integer)
'   If KeyAscii = 13 Then Sendkeys "{TAB}"
'   If KeyAscii = 46 Then KeyAscii = 44
'   If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 44 And KeyAscii <> 8 Then
'      KeyAscii = 0
'   End If
'   If KeyAscii = 44 Then
'      If InStr(mskUnitario.ClipText, ",") <> 0 Then
'         KeyAscii = 0
'      End If
'   End If
'End Sub
'
'Private Function FData(dData As String) As Boolean
'   If Not IsDate(dData) Or Val(Mid(dData, 7, 4)) < 1900 Then
'      MsgBox "Data Inválida", vbOKOnly + vbInformation, "Validação"
'      FData = False
'      Exit Function
'   Else
'      FData = True
'   End If
'End Function
'
'Private Sub cmdAnotacoes_Click()
'Dim Contador As Integer
'
'   If frmNFe.txtCodigo.text <> "" Then
'    ' Carrega Combos
'      ' Unidades de Medida
'      Set rsTemp = cnSistema.Execute("Select * from UnidadesMedida Order By Descricao")
'      frmNFeComplemento.cmbUnidade.Clear
'      Do While Not rsTemp.EOF
'         frmNFeComplemento.cmbUnidade.AddItem rsTemp!Descricao
'         frmNFeComplemento.cmbUnidade.ItemData(frmNFeComplemento.cmbUnidade.NewIndex) = rsTemp!idUnidadeMedida
'         rsTemp.MoveNext
'      Loop
'
'      ' Situacoes Tributarias
'      Set rsTemp = cnSistema.Execute("Select * from SituacoesTributarias Order By Descricao")
'      frmNFeComplemento.cmbSituacaoTributaria.Clear
'      Do While Not rsTemp.EOF
'         frmNFeComplemento.cmbSituacaoTributaria.AddItem rsTemp!Descricao
'         frmNFeComplemento.cmbSituacaoTributaria.ItemData(frmNFeComplemento.cmbSituacaoTributaria.NewIndex) = rsTemp!idSituacaoTributaria
'         rsTemp.MoveNext
'      Loop
'
'    ' Preencher Campos
'      Set rsTemp = cnSistema.Execute("Select * From Produtos Where Codigo = '" & SQLCheck(frmNFe.txtCodigo.text) & "'")
'      If Not rsTemp.EOF Then
'         ' ICMS
'         frmNFeComplemento.mskICMSProduto.text = rsTemp!ICMS
'         If Not rsEmpresa.EOF Then
'            If rsClientes!UF = rsEmpresa!UF Then
'               frmNFeComplemento.mskBaseReduzidaICMS.text = rsTemp!BaseReduzidaICMSdUF
'            Else
'               frmNFeComplemento.mskBaseReduzidaICMS.text = rsTemp!BaseReduzidaICMSfUF
'            End If
'         End If
'         frmNFeComplemento.mskCFOP.text = mskCFOP.text
'
'         ' Unidade
'         For Contador = 0 To (frmNFeComplemento.cmbUnidade.ListCount - 1)
'            If frmNFeComplemento.cmbUnidade.ItemData(Contador) = rsTemp!idUnidade Then
'               frmNFeComplemento.cmbUnidade.ListIndex = Contador
'               Exit For
'            End If
'         Next
'
'         ' Situacao Tributaria
'         For Contador = 0 To (frmNFeComplemento.cmbSituacaoTributaria.ListCount - 1)
'            If frmNFeComplemento.cmbSituacaoTributaria.ItemData(Contador) = rsTemp!idSituacaoTributaria Then
'               frmNFeComplemento.cmbSituacaoTributaria.ListIndex = Contador
'               Exit For
'            End If
'         Next
'      End If
'
'      frmNFeComplemento.Show vbModal
'      cmdIncluir.SetFocus
'   End If
'End Sub
'
'Sub ImprimirNota()
'
'   If LerArquivoINI("SEPD", "TipoImpressao", CaminhoINI & "\System.ini") = 1 Then
'      Call ImprimirDanfe
'   ElseIf LerArquivoINI("SEPD", "TipoImpressao", CaminhoINI & "\System.ini") = 2 Then
'      Call FolhaSolta
'   ElseIf LerArquivoINI("SEPD", "TipoImpressao", CaminhoINI & "\System.ini") = 3 Then
'      Call NotaServico
'   End If
'End Sub
'
'Sub FolhaSolta()
'Dim iItensNota As Integer
'Dim iPosLinNumero As Integer
'Dim iPosColMarcaEntrada As Integer
'Dim iPosColMarcaSaida As Integer
'Dim iPosColNumero As Integer
'Dim iPosLinNatureza As Integer
'Dim iPosColNatureza As Integer
'Dim iPosColCBO As Integer
'Dim iPosLinCliente As Integer
'Dim iPosColCliente As Integer
'Dim iPosColEmissao As Integer
'Dim iPosLinEndereco As Integer
'Dim iPosColEndereco As Integer
'Dim iPosColBairro As Integer
'Dim iPosColSaida As Integer
'Dim iPosLinCidade As Integer
'Dim iPosColCidade As Integer
'Dim iPosColTelefone As Integer
'Dim iPosColUF As Integer
'Dim iPosColIE_CI As Integer
'Dim iPosLinBairro As Integer
'Dim iPosColCEP As Integer
'Dim iPosColCNPJ As Integer
'Dim iPosLinCobranca As Integer
'Dim iPosLinProdutos As Integer
'Dim iPosColCodigo As Integer
'Dim iPosColDescricao As Integer
'Dim iPosColUnidade As Integer
'Dim iPosColQuantidade As Integer
'Dim iPosColVlUnitario As Integer
'Dim iPosColVlLiquido As Integer
'Dim iPosColICMS As Integer
'Dim iPosLinInfoCN As Integer
'Dim iPosColInfoCN As Integer
'Dim iPosLinBase As Integer
'Dim iPosColBase As Integer
'Dim iPosColValorICMS As Integer
'Dim iPosColTotalProdutos As Integer
'Dim iPosLinTotalNota As Integer
'Dim iPosColTotalNota As Integer
'Dim iPosLinTransportador As Integer
'Dim iPosColTransportador As Integer
'Dim iPosColFreteConta As Integer
'Dim iPosColPlacaVeiculo As Integer
'Dim iPosColUFPlaca As Integer
'Dim iPosColCNPJ_CPFTrans As Integer
'Dim iPosLinEndTrans As Integer
'Dim iPosColEndTrans As Integer
'Dim iPosColCidTrans As Integer
'Dim iPosColUFTrans As Integer
'Dim iPosColIE_CITrans As Integer
'Dim iPosLinVolQuant As Integer
'Dim iPosColVolQuant As Integer
'Dim iPosColVolMarca As Integer
'Dim iPosColVolNumero As Integer
'Dim iPosColPesoBruto As Integer
'Dim iPosColPesoLiquido As Integer
'Dim iPosLinDadosAdic As Integer
'Dim iPosColDadosAdic As Integer
'Dim iPosLinNumeroFim As Integer
'Dim iPosColNumeroFim As Integer
'Dim iPosLinProximaNota As Integer
'
'   Set rsNFeItens = cnSistema.Execute("SELECT * FROM NFeItens WHERE idNFe = " & rsNFe!idNFe)
'   If rsNFeItens.EOF Then
'      MsgBox "Nota não pode ser impressa sem Produtos ", vbExclamation + vbOKOnly, "Campos Obrigatórios"
'      Exit Sub
'   End If
'
' ' Marca e Numero
'   iPosLinNumero = LerArquivoINI("Notas Fiscais", "PosLinNumero", CaminhoINI & "\NotasManuais.ini")
'   iPosColMarcaEntrada = LerArquivoINI("Notas Fiscais", "PosColMarcaEntrada", CaminhoINI & "\NotasManuais.ini")
'   iPosColMarcaSaida = LerArquivoINI("Notas Fiscais", "PosColMarcaSaida", CaminhoINI & "\NotasManuais.ini")
'   iPosColNumero = LerArquivoINI("Notas Fiscais", "PosColNumero", CaminhoINI & "\NotasManuais.ini")
'
' ' Natureza e CBO
'   iPosLinNatureza = LerArquivoINI("Notas Fiscais", "PosLinNatureza", CaminhoINI & "\NotasManuais.ini")
'   iPosColNatureza = LerArquivoINI("Notas Fiscais", "PosColNatureza", CaminhoINI & "\NotasManuais.ini")
'   iPosColCBO = LerArquivoINI("Notas Fiscais", "PosColCBO", CaminhoINI & "\NotasManuais.ini")
'
' ' Cliente, CNPJ e Emissão
'   iPosLinCliente = LerArquivoINI("Notas Fiscais", "PosLinCliente", CaminhoINI & "\NotasManuais.ini")
'   iPosColCliente = LerArquivoINI("Notas Fiscais", "PosColCliente", CaminhoINI & "\NotasManuais.ini")
'   iPosColEmissao = LerArquivoINI("Notas Fiscais", "PosColEmissao", CaminhoINI & "\NotasManuais.ini")
'
' ' Endereço e Data de Saida
'   iPosLinEndereco = LerArquivoINI("Notas Fiscais", "PosLinEndereco", CaminhoINI & "\NotasManuais.ini")
'   iPosColEndereco = LerArquivoINI("Notas Fiscais", "PosColEndereco", CaminhoINI & "\NotasManuais.ini")
'   iPosColSaida = LerArquivoINI("Notas Fiscais", "PosColSaida", CaminhoINI & "\NotasManuais.ini")
'
' ' Cidade, Telefone, UF e Inscrição Estadual
'   iPosLinCidade = LerArquivoINI("Notas Fiscais", "PosLinCidade", CaminhoINI & "\NotasManuais.ini")
'   iPosColCidade = LerArquivoINI("Notas Fiscais", "PosColCidade", CaminhoINI & "\NotasManuais.ini")
'   iPosColTelefone = LerArquivoINI("Notas Fiscais", "PosColTelefone", CaminhoINI & "\NotasManuais.ini")
'   iPosColUF = LerArquivoINI("Notas Fiscais", "PosColUF", CaminhoINI & "\NotasManuais.ini")
'   iPosColIE_CI = LerArquivoINI("Notas Fiscais", "PosColIE_CI", CaminhoINI & "\NotasManuais.ini")
'
' ' Bairro, CEP e CNPJ
'   iPosLinBairro = LerArquivoINI("Notas Fiscais", "PosLinBairro", CaminhoINI & "\NotasManuais.ini")
'   iPosColBairro = LerArquivoINI("Notas Fiscais", "PosColBairro", CaminhoINI & "\NotasManuais.ini")
'   iPosColCEP = LerArquivoINI("Notas Fiscais", "PosColCEP", CaminhoINI & "\NotasManuais.ini")
'   iPosColCNPJ = LerArquivoINI("Notas Fiscais", "PosColCNPJ", CaminhoINI & "\NotasManuais.ini")
'
' ' Cobranca
'   iPosLinCobranca = LerArquivoINI("Notas Fiscais", "PosLinCobranca", CaminhoINI & "\NotasManuais.ini")
'
' ' Produtos
'   iPosLinProdutos = LerArquivoINI("Notas Fiscais", "PosLinProdutos", CaminhoINI & "\NotasManuais.ini")
'   iPosColCodigo = LerArquivoINI("Notas Fiscais", "PosColCodigo", CaminhoINI & "\NotasManuais.ini")
'   iPosColDescricao = LerArquivoINI("Notas Fiscais", "PosColDescricao", CaminhoINI & "\NotasManuais.ini")
'   iPosColUnidade = LerArquivoINI("Notas Fiscais", "PosColUnidade", CaminhoINI & "\NotasManuais.ini")
'   iPosColQuantidade = LerArquivoINI("Notas Fiscais", "PosColQuantidade", CaminhoINI & "\NotasManuais.ini")
'   iPosColVlUnitario = LerArquivoINI("Notas Fiscais", "PosColVlUnitario", CaminhoINI & "\NotasManuais.ini")
'   iPosColVlLiquido = LerArquivoINI("Notas Fiscais", "PosColVlLiquido", CaminhoINI & "\NotasManuais.ini")
'   iPosColICMS = LerArquivoINI("Notas Fiscais", "PosColICMS", CaminhoINI & "\NotasManuais.ini")
'
' ' Informações do Corpo da Nota
'   iPosLinInfoCN = LerArquivoINI("Notas Fiscais", "PosLinInfoCN", CaminhoINI & "\NotasManuais.ini")
'   iPosColInfoCN = LerArquivoINI("Notas Fiscais", "PosColInfoCN", CaminhoINI & "\NotasManuais.ini")
'
' ' Base de Calculo, Valor do ICMS e Valor Total dos Produtos
'   iPosLinBase = LerArquivoINI("Notas Fiscais", "PosLinBase", CaminhoINI & "\NotasManuais.ini")
'   iPosColBase = LerArquivoINI("Notas Fiscais", "PosColBase", CaminhoINI & "\NotasManuais.ini")
'   iPosColValorICMS = LerArquivoINI("Notas Fiscais", "PosColValorICMS", CaminhoINI & "\NotasManuais.ini")
'   iPosColTotalProdutos = LerArquivoINI("Notas Fiscais", "PosColTotalProdutos", CaminhoINI & "\NotasManuais.ini")
'
' ' Valor Total da Nota
'   iPosLinTotalNota = LerArquivoINI("Notas Fiscais", "PosLinTotalNota", CaminhoINI & "\NotasManuais.ini")
'   iPosColTotalNota = LerArquivoINI("Notas Fiscais", "PosColTotalNota", CaminhoINI & "\NotasManuais.ini")
'
' ' Transportador
'   iPosLinTransportador = LerArquivoINI("Notas Fiscais", "PosLinTransportador", CaminhoINI & "\NotasManuais.ini")
'   iPosColTransportador = LerArquivoINI("Notas Fiscais", "PosColTransportador", CaminhoINI & "\NotasManuais.ini")
'   iPosColFreteConta = LerArquivoINI("Notas Fiscais", "PosColFreteConta", CaminhoINI & "\NotasManuais.ini")
'   iPosColPlacaVeiculo = LerArquivoINI("Notas Fiscais", "PosColPlacaVeiculo", CaminhoINI & "\NotasManuais.ini")
'   iPosColUFPlaca = LerArquivoINI("Notas Fiscais", "PosColUFPlaca", CaminhoINI & "\NotasManuais.ini")
'   iPosColCNPJ_CPFTrans = LerArquivoINI("Notas Fiscais", "PosColCNPJ_CPFTrans", CaminhoINI & "\NotasManuais.ini")
'
'   iPosLinEndTrans = LerArquivoINI("Notas Fiscais", "PosLinEndTrans", CaminhoINI & "\NotasManuais.ini")
'   iPosColEndTrans = LerArquivoINI("Notas Fiscais", "PosColEndTrans", CaminhoINI & "\NotasManuais.ini")
'   iPosColCidTrans = LerArquivoINI("Notas Fiscais", "PosColCidTrans", CaminhoINI & "\NotasManuais.ini")
'   iPosColUFTrans = LerArquivoINI("Notas Fiscais", "PosColUFTrans", CaminhoINI & "\NotasManuais.ini")
'   iPosColIE_CITrans = LerArquivoINI("Notas Fiscais", "PosColIE_CITrans", CaminhoINI & "\NotasManuais.ini")
'
' ' Volume
'   iPosLinVolQuant = LerArquivoINI("Notas Fiscais", "PosLinVolQuant", CaminhoINI & "\NotasManuais.ini")
'   iPosColVolQuant = LerArquivoINI("Notas Fiscais", "PosColVolQuant", CaminhoINI & "\NotasManuais.ini")
'   iPosColVolMarca = LerArquivoINI("Notas Fiscais", "PosColVolMarca", CaminhoINI & "\NotasManuais.ini")
'   iPosColVolNumero = LerArquivoINI("Notas Fiscais", "PosColVolNumero", CaminhoINI & "\NotasManuais.ini")
'   iPosColPesoBruto = LerArquivoINI("Notas Fiscais", "PosColPesoBruto", CaminhoINI & "\NotasManuais.ini")
'   iPosColPesoLiquido = LerArquivoINI("Notas Fiscais", "PosColPesoLiquido", CaminhoINI & "\NotasManuais.ini")
'
' ' Dados Adicionais
'   iPosLinDadosAdic = LerArquivoINI("Notas Fiscais", "PosLinDadosAdic", CaminhoINI & "\Notas.ini")
'   iPosColDadosAdic = LerArquivoINI("Notas Fiscais", "PosColDadosAdic", CaminhoINI & "\Notas.ini")
'
' ' Numero Final
'   iPosLinNumeroFim = LerArquivoINI("Notas Fiscais", "PosLinNumeroFim", CaminhoINI & "\NotasManuais.ini")
'   iPosColNumeroFim = LerArquivoINI("Notas Fiscais", "PosColNumeroFim", CaminhoINI & "\NotasManuais.ini")
'   iPosLinProximaNota = LerArquivoINI("Notas Fiscais", "PosLinProximaNota", CaminhoINI & "\NotasManuais.ini")
'
'   If rsNFe!Impressa Then
'      Beep
'      MsgBox "Nota Fiscal já Impressa", vbExclamation, "Aviso"
'   End If
'
'   cdlgImprimirNota.CancelError = True
'
'   Registro_Selecionado = False
'   VisualizarImpressao
'   If Registro_Selecionado Then
'      Open LerArquivoINI("Impressoras", "Notas", CaminhoINI & "\System.ini") For Output As #1
''      Open caminhoini & "\teste.txt" For Output As #1
'
'      Set rsCFOPs = cnSistema.Execute("Select * From CFOPs Where idCFOP = " & rsNFe!idCFOP)
'      Set rsClientes = cnSistema.Execute("Select * From Clientes Where idCliente = " & rsNFe!idCliente)
'      Set rsNaturezasOperacao = cnSistema.Execute("Select * From NaturezasOperacao Where idNaturezaOperacao = " & rsNFe!idNaturezaOperacao)
'
'    ' Numero da Nota
'      Print #1, Chr(27) & "x0"; Chr$(27) & Chr(69); Chr(15)
'      SaltarLinha (iPosLinNumero)
'      Print #1, Tab(iPosColMarcaSaida); "X"
'      SaltarLinha (iPosLinNatureza)
'
'    ' Codigo Fiscal de Operação
'      Print #1, Tab(iPosColNatureza); rsNaturezasOperacao!Descricao; Tab(iPosColCBO); rsCFOPs!CFOP
'      SaltarLinha (iPosLinCliente)
'
'    ' Cliente e Emissão
''''      Print #1, Tab(iPosColCliente); RemoveAcentos(rsClientes!Nome); Tab(iPosColCNPJ); rsClientes!CNPJ_CPF; Tab(iPosColEmissao); rsNFe!DataEmissao
'      Print #1, Tab(iPosColCliente); RemoveAcentos(rsClientes!Nome); Tab(iPosColEmissao); rsNFe!DataEmissao
'      SaltarLinha (iPosLinEndereco)
'
'    ' Endereço e Saida
''''      Print #1, Tab(iPosColEndereco); RemoveAcentos(rsClientes!Endereco); Tab(iPosColBairro); RemoveAcentos(rsClientes!Bairro); Tab(iPosColCEP); rsClientes!CEP; Tab(iPosColSaida); rsNFe!DataEmissao
'      Print #1, Tab(iPosColEndereco); RemoveAcentos(IIf(Not IsNull(rsClientes!Endereco), rsClientes!Endereco, "")); Tab(iPosColSaida); rsNFe!DataEmissao
'      SaltarLinha (iPosLinCidade)
'
'    ' Cidade, Telefone, UF e Inscrição Estadual
''''      Print #1, Tab(iPosColCidade); RemoveAcentos(rsClientes!Cidade); Tab(iPosColTelefone); rsClientes!Telefone1; Tab(iPosColUF); rsClientes!UF; Tab(iPosColIE_CI); rsClientes!IE_CI
'      Print #1, Tab(iPosColCidade); RemoveAcentos(IIf(Not IsNull(rsClientes!Cidade), rsClientes!Cidade, "")); Tab(iPosColTelefone); rsClientes!Telefone1; Tab(iPosColUF); IIf(Not IsNull(rsClientes!UF), rsClientes!UF, ""); Tab(iPosColIE_CI); IIf(Not IsNull(rsClientes!IE_CI), rsClientes!IE_CI, "")
'      SaltarLinha (iPosLinBairro)
'
'    ' Bairro, CEP e CNPJ/CPF
''''      Print #1, Tab(iPosColCidade); RemoveAcentos(rsClientes!Cidade); Tab(iPosColTelefone); rsClientes!Telefone1; Tab(iPosColUF); rsClientes!UF; Tab(iPosColIE_CI); rsClientes!IE_CI
'      Print #1, Tab(iPosColBairro); RemoveAcentos(IIf(Not IsNull(rsClientes!Bairro), rsClientes!Bairro, "")); Tab(iPosColCEP); IIf(Not IsNull(rsClientes!CEP), rsClientes!CEP, ""); Tab(iPosColCNPJ); IIf(Not IsNull(rsClientes!CNPJ_CPF), rsClientes!CNPJ_CPF, "")
'      SaltarLinha (iPosLinCobranca)
'
'    ' Cobranca
''            Print #1, Tab(5); rsNFe!Documento; Spc(5); cmbFormaPagamento.Text; Spc(5); rsNFe!DataVencimento
'      SaltarLinha (iPosLinProdutos)
'
'    ' Produtos
'
'      Dim cValorBruto As Currency
'      Dim cValorDesconto As Currency
'      Dim cValorBonificacao As Currency
'      Dim cValorLiquido As Currency
'
'      Dim iLinhasProdutos As Integer
'      iLinhasProdutos = LerArquivoINI("Notas Fiscais", "ItensNota", CaminhoINI & "\NotasManuais.ini")
'      Set rsNFeItens = cnSistema.Execute("SELECT NFeItens.idNFe, NFeItens.idProduto, NFeItens.idUnidade, Produtos.Descricao, NFeItens.idClassificacaoFiscal, NFeItens.idSituacaoTributaria, NFeItens.Data, NFeItens.Quantidade, NFeItens.ValorUnitario, NFeItens.Desconto, NFeItens.IPI, NFeItens.ICMS, NFeItens.DescricaoComplementar " & _
'                                                  "FROM NFeItens INNER JOIN Produtos ON NFeItens.idProduto = Produtos.idProduto " & _
'                                                  "Where NFeItens.idNFe = " & rsNFe!idNFe & " ORDER BY Produtos.Descricao")
'
'      Do While Not rsNFeItens.EOF
'         Set rsProdutos = cnSistema.Execute("SELECT * FROM Produtos WHERE idProduto = " & rsNFeItens!idProduto)
'         Set rsUnidadesMedida = cnSistema.Execute("SELECT * FROM UnidadesMedida WHERE idUnidadeMedida = " & rsNFeItens!idUnidade)
'         Dim sDescricao As String, sUnidade As String
'         sDescricao = RemoveAcentos(Mid(Trim(rsProdutos!Descricao), 1, 50) & " " & Trim(rsNFeItens!DescricaoComplementar))
'         If Not rsUnidadesMedida.EOF Then
'            sUnidade = rsUnidadesMedida!Sigla
'         Else
'            sUnidade = " "
'         End If
'
'         cValorBruto = (rsNFeItens!quantidade * rsNFeItens!ValorUnitario)
'         cValorDesconto = (((rsNFeItens!quantidade * rsNFeItens!ValorUnitario) * rsNFeItens!Desconto) / 100)
'         cValorBonificacao = (((cValorBruto - cValorDesconto) * rsNFe!Bonificacao) / 100)
'         cValorLiquido = (cValorBruto - cValorDesconto - cValorBonificacao)
'
'         Print #1, Tab(iPosColCodigo); rsProdutos!Codigo; _
'                   Tab(iPosColDescricao); sDescricao; _
'                   Tab(iPosColUnidade); sUnidade; _
'                   Tab(iPosColQuantidade); FormataTXT(Format(rsNFeItens!quantidade, "##,##0.00"), 2.1, 10); _
'                   Tab(iPosColVlUnitario); FormataTXT(Format(rsNFeItens!ValorUnitario, "##,##0.00"), 2.1, 10); _
'                   Tab(iPosColVlLiquido); FormataTXT(Format(cValorLiquido, "##,##0.00"), 2.1, 12); _
'                   Tab(iPosColICMS); Format(rsNFeItens!ICMS, "##0")
'
'         iLinhasProdutos = iLinhasProdutos - 1
'         rsNFeItens.MoveNext
'      Loop
'
'    ' Informações do Corpo da Nota
'      SaltarLinha (iLinhasProdutos + iPosLinInfoCN)
'      Dim sInformacoesCorpo1 As String, sInformacoesCorpo2 As String, sInformacoesCorpo3 As String
'      If Len(rsNFe!InformacoesCorpo) >= 240 Then
'         sInformacoesCorpo1 = Mid(rsNFe!InformacoesCorpo, 1, 120)
'         sInformacoesCorpo2 = Mid(rsNFe!InformacoesCorpo, 121, 120)
'         sInformacoesCorpo3 = Mid(rsNFe!InformacoesCorpo, 241, 120)
'
'         Print #1, Tab(iPosColInfoCN); sInformacoesCorpo1
'         Print #1, Tab(iPosColInfoCN); sInformacoesCorpo2
'         Print #1, Tab(iPosColInfoCN); sInformacoesCorpo3
'         SaltarLinha (iPosLinBase - 3)
'      Else
'         If Len(rsNFe!InformacoesCorpo) >= 120 Then
'            sInformacoesCorpo1 = Mid(rsNFe!InformacoesCorpo, 1, 120)
'            sInformacoesCorpo2 = Mid(rsNFe!InformacoesCorpo, 121, 120)
'
'            Print #1, Tab(iPosColInfoCN); sInformacoesCorpo1
'            Print #1, Tab(iPosColInfoCN); sInformacoesCorpo2
'            SaltarLinha (iPosLinBase - 2)
'         Else
'            If Len(rsNFe!InformacoesCorpo) >= 1 Then
'               sInformacoesCorpo1 = rsNFe!InformacoesCorpo
'
'               Print #1, Tab(iPosColInfoCN); sInformacoesCorpo1
'               SaltarLinha (iPosLinBase - 1)
'            Else
'               SaltarLinha (iPosLinBase)
'            End If
'         End If
'      End If
'
''      Print #1, Tab(iPosColInfoCN); rsNFe!InformacoesCorpo
'
'    ' Total
'      Set rsTemp = cnSistema.Execute("Select * From TotalNFe Where Numero = " & mskNumero.text)
'      Dim dBaseCalculo As Double, dValorICMS As Double
'      If rsTemp!ValorICMS > 0 Then
'         dBaseCalculo = rsTemp!BaseCalculo
'         dValorICMS = rsTemp!ValorICMS
'      Else
'         dBaseCalculo = 0
'         dValorICMS = 0
'      End If
'
'      If Not rsTemp.EOF Then
''         Print #1, Tab(10); Format(rsNFe!BaseCalculoICMS, "#0.00"); Tab(40); Format(rsNFe!ValorICMS, "#0.00"); Tab(130); Format(rsTemp!Total, "#0.00")
'         Print #1, Tab(iPosColBase); Format(dBaseCalculo, "##,##0.00"); Tab(iPosColValorICMS); Format(dValorICMS, "##,##0.00"); Tab(iPosColTotalProdutos); Format(rsTemp!Total, "##,##0.00")
'         SaltarLinha (iPosLinTotalNota)
'         Print #1, Tab(iPosColTotalNota); Format(rsTemp!Total, "##,##0.00")
'         SaltarLinha (iPosLinTransportador)
'      Else
''         Print #1, Tab(10); Format(rsNFe!BaseCalculoICMS, "#0.00"); Tab(40); Format(rsNFe!ValorICMS, "#0.00"); Tab(130); Format(0, "#0.00")
'         Print #1, Tab(iPosColBase); Format(dBaseCalculo, "##,##0.00"); Tab(iPosColValorICMS); Format(dValorICMS, "##,##0.00"); Tab(iPosColTotalProdutos); Format(0, "##,##0.00")
'         SaltarLinha (iPosLinTotalNota)
'         Print #1, Tab(iPosColTotalNota); Format(0, "##,##0.00")
'         SaltarLinha (iPosLinTransportador)
'      End If
'
'    ' Transportador
'      Set rsTransportador = cnSistema.Execute("SELECT * FROM Transportadores WHERE idTransportador = " & rsNFe!idTransportador)
'      If Not rsTransportador.EOF Then
'         Dim sFreteConta As String
'         sFreteConta = IIf(rsNFe!FreteConta > 0, rsNFe!FreteConta, "")
'
'         Print #1, Tab(iPosColTransportador); rsTransportador!Nome; Tab(iPosColFreteConta); sFreteConta; Tab(iPosColPlacaVeiculo); rsNFe!PlacaVeiculo; Tab(iPosColUFPlaca); rsTransportador!UFPlaca; Tab(iPosColCNPJ_CPFTrans); ; IIf(Len(rsTransportador!CNPJ_CPF) > 4, rsTransportador!CNPJ_CPF, "")
'         SaltarLinha (iPosLinEndTrans)
'         Print #1, Tab(iPosColEndTrans); rsTransportador!Endereco; Tab(iPosColCidTrans); rsTransportador!Cidade; Tab(iPosColUFTrans); rsTransportador!UF; Tab(iPosColIE_CITrans); rsTransportador!IE_CI
'         SaltarLinha (iPosLinVolQuant)
'      Else
'         SaltarLinha (6)
'      End If
'      Print #1, Tab(iPosColVolQuant); rsNFe!VolumeQuantidade; Tab(iPosColVolMarca); rsNFe!VolumeMarca; Tab(iPosColVolNumero); rsNFe!VolumeNumero; Tab(iPosColPesoBruto); Format(rsTemp!PesoTotal, "##,##0.00"); Tab(iPosColPesoLiquido); Format(rsTemp!PesoTotal, "##,##0.00")
'
'    ' Dados Adicionais
'      SaltarLinha (iPosLinDadosAdic)
'      Dim sDadosAdicionais1 As String, sDadosAdicionais2 As String, sDadosAdicionais3 As String
'      If Len(rsNFe!DadosAdicionais) >= 120 Then
'         sDadosAdicionais1 = RemoveAcentos(Mid(rsNFe!DadosAdicionais, 1, 60))
'         sDadosAdicionais2 = RemoveAcentos(Mid(rsNFe!DadosAdicionais, 61, 60))
'         sDadosAdicionais3 = RemoveAcentos(Mid(rsNFe!DadosAdicionais, 121, 60))
'
'         Print #1, Tab(iPosColDadosAdic); sDadosAdicionais1
'         Print #1, Tab(iPosColDadosAdic); sDadosAdicionais2
'         Print #1, Tab(iPosColDadosAdic); sDadosAdicionais3
'         SaltarLinha (iPosLinNumeroFim - 3)
'      Else
'         If Len(rsNFe!DadosAdicionais) >= 60 Then
'            sDadosAdicionais1 = RemoveAcentos(Mid(rsNFe!DadosAdicionais, 1, 60))
'            sDadosAdicionais2 = RemoveAcentos(Mid(rsNFe!DadosAdicionais, 61, 60))
'
'            Print #1, Tab(iPosColDadosAdic); sDadosAdicionais1
'            Print #1, Tab(iPosColDadosAdic); sDadosAdicionais2
'            SaltarLinha (iPosLinNumeroFim - 2)
'         Else
'            If Len(rsNFe!DadosAdicionais) >= 1 Then
'               sDadosAdicionais1 = RemoveAcentos(rsNFe!DadosAdicionais)
'
'               Print #1, Tab(iPosColDadosAdic); sDadosAdicionais1
'               SaltarLinha (iPosLinNumeroFim - 1)
'            Else
'               SaltarLinha (iPosLinNumeroFim)
'            End If
'         End If
'      End If
'
'    ' Numero da Nota
''      SaltarLinha (iPosLinNumeroFim)
'      Print #1, Tab(iPosColNumeroFim); Chr$(27) & Chr(70); " " & Chr(18)
'      SaltarLinha (iPosLinProximaNota)
'
'      Close #1
'      cnSistema.Execute "Update NFe set " & _
'            "Impressa = " & True & " " & _
'            "Where idNFe = " & rsNFe!idNFe
'   End If
'
'End Sub
'
'Sub NotaServico()
'Dim iItensNota As Integer
'
'Dim iPosLinEmissao As Integer
'Dim iPosColEmissao As Integer
'Dim iPosColNumero  As Integer
'
'Dim iPosLinCliente As Integer
'Dim iPosColCliente As Integer
'
'Dim iPosLinEndereco As Integer
'Dim iPosColEndereco As Integer
'
'Dim iPosLinCidade As Integer
'Dim iPosColCidade As Integer
'Dim iPosColUF As Integer
'Dim iPosColTelefone As Integer
'
'Dim iPosLinCNPJ As Integer
'Dim iPosColCNPJ As Integer
'Dim iPosColIE As Integer
'
'Dim iPosLinCobranca As Integer
'Dim iPosColDocBol1 As Integer
'Dim iPosColVenBol1 As Integer
'Dim iPosColValBol1 As Integer
'Dim iPosColDocBol2 As Integer
'Dim iPosColVenBol2 As Integer
'Dim iPosColValBol2 As Integer
'Dim iPosColDocBol3 As Integer
'Dim iPosColVenBol3 As Integer
'Dim iPosColValBol3 As Integer
'Dim iPosColDocBol4 As Integer
'Dim iPosColVenBol4 As Integer
'Dim iPosColValBol4 As Integer
'
'Dim iPosLinProdutos As Integer
'Dim iPosColQuantidade As Integer
'Dim iPosColUnidade As Integer
'Dim iPosColDescricao As Integer
'Dim iPosColVlUnitario As Integer
'Dim iPosColVlLiquido As Integer
'
'Dim iPosLinTotalNota As Integer
'Dim iPosColTotalNota As Integer
'
'Dim iPosLinNumeroFim As Integer
'Dim iPosColNumeroFim As Integer
'Dim iPosLinProximaNota As Integer
'
'   Set rsNFeItens = cnSistema.Execute("SELECT * FROM NFeItens WHERE idNFe = " & rsNFe!idNFe)
'   If rsNFeItens.EOF Then
'      MsgBox "Nota não pode ser impressa sem Produtos ", vbExclamation + vbOKOnly, "Campos Obrigatórios"
'      Exit Sub
'   End If
'
' ' Emissao e Numero
'   iPosLinEmissao = LerArquivoINI("Notas Fiscais", "PosLinEmissao", CaminhoINI & "\NotasServico.ini")
'   iPosColEmissao = LerArquivoINI("Notas Fiscais", "PosColEmissao", CaminhoINI & "\NotasServico.ini")
'   iPosColNumero = LerArquivoINI("Notas Fiscais", "PosColNumero", CaminhoINI & "\NotasServico.ini")
'
' ' Cliente
'   iPosLinCliente = LerArquivoINI("Notas Fiscais", "PosLinCliente", CaminhoINI & "\NotasServico.ini")
'   iPosColCliente = LerArquivoINI("Notas Fiscais", "PosColCliente", CaminhoINI & "\NotasServico.ini")
'
' ' Endereco
'   iPosLinEndereco = LerArquivoINI("Notas Fiscais", "PosLinEndereco", CaminhoINI & "\NotasServico.ini")
'   iPosColEndereco = LerArquivoINI("Notas Fiscais", "PosColEndereco", CaminhoINI & "\NotasServico.ini")
'
' ' Cidade, UF e Telefone
'   iPosLinCidade = LerArquivoINI("Notas Fiscais", "PosLinCidade", CaminhoINI & "\NotasServico.ini")
'   iPosColCidade = LerArquivoINI("Notas Fiscais", "PosColCidade", CaminhoINI & "\NotasServico.ini")
'   iPosColUF = LerArquivoINI("Notas Fiscais", "PosColUF", CaminhoINI & "\NotasServico.ini")
'   iPosColTelefone = LerArquivoINI("Notas Fiscais", "PosColTelefone", CaminhoINI & "\NotasServico.ini")
'
' ' CNPJ e IE
'   iPosLinCNPJ = LerArquivoINI("Notas Fiscais", "PosLinCNPJ", CaminhoINI & "\NotasServico.ini")
'   iPosColCNPJ = LerArquivoINI("Notas Fiscais", "PosColCNPJ", CaminhoINI & "\NotasServico.ini")
'   iPosColIE = LerArquivoINI("Notas Fiscais", "PosColIE", CaminhoINI & "\NotasServico.ini")
'
' ' Cobranca
'   iPosLinCobranca = LerArquivoINI("Notas Fiscais", "PosLinCobranca", CaminhoINI & "\NotasServico.ini")
'   iPosColDocBol1 = LerArquivoINI("Notas Fiscais", "PosColDocBol1", CaminhoINI & "\NotasServico.ini")
'   iPosColVenBol1 = LerArquivoINI("Notas Fiscais", "PosColVenBol1", CaminhoINI & "\NotasServico.ini")
'   iPosColValBol1 = LerArquivoINI("Notas Fiscais", "PosColValBol1", CaminhoINI & "\NotasServico.ini")
'   iPosColDocBol2 = LerArquivoINI("Notas Fiscais", "PosColDocBol2", CaminhoINI & "\NotasServico.ini")
'   iPosColVenBol2 = LerArquivoINI("Notas Fiscais", "PosColVenBol2", CaminhoINI & "\NotasServico.ini")
'   iPosColValBol2 = LerArquivoINI("Notas Fiscais", "PosColValBol2", CaminhoINI & "\NotasServico.ini")
'   iPosColDocBol3 = LerArquivoINI("Notas Fiscais", "PosColDocBol3", CaminhoINI & "\NotasServico.ini")
'   iPosColVenBol3 = LerArquivoINI("Notas Fiscais", "PosColVenBol3", CaminhoINI & "\NotasServico.ini")
'   iPosColValBol3 = LerArquivoINI("Notas Fiscais", "PosColValBol3", CaminhoINI & "\NotasServico.ini")
'   iPosColDocBol4 = LerArquivoINI("Notas Fiscais", "PosColDocBol4", CaminhoINI & "\NotasServico.ini")
'   iPosColVenBol4 = LerArquivoINI("Notas Fiscais", "PosColVenBol4", CaminhoINI & "\NotasServico.ini")
'   iPosColValBol4 = LerArquivoINI("Notas Fiscais", "PosColValBol4", CaminhoINI & "\NotasServico.ini")
'
' ' Produtos
'   iPosLinProdutos = LerArquivoINI("Notas Fiscais", "PosLinProdutos", CaminhoINI & "\NotasServico.ini")
'   iPosColQuantidade = LerArquivoINI("Notas Fiscais", "PosColQuantidade", CaminhoINI & "\NotasServico.ini")
'   iPosColUnidade = LerArquivoINI("Notas Fiscais", "PosColUnidade", CaminhoINI & "\NotasServico.ini")
'   iPosColDescricao = LerArquivoINI("Notas Fiscais", "PosColDescricao", CaminhoINI & "\NotasServico.ini")
'   iPosColVlUnitario = LerArquivoINI("Notas Fiscais", "PosColVlUnitario", CaminhoINI & "\NotasServico.ini")
'   iPosColVlLiquido = LerArquivoINI("Notas Fiscais", "PosColVlLiquido", CaminhoINI & "\NotasServico.ini")
'
' ' Valor Total da Nota
'   iPosLinTotalNota = LerArquivoINI("Notas Fiscais", "PosLinTotalNota", CaminhoINI & "\NotasServico.ini")
'   iPosColTotalNota = LerArquivoINI("Notas Fiscais", "PosColTotalNota", CaminhoINI & "\NotasServico.ini")
'
' ' Numero Final
'   iPosLinNumeroFim = LerArquivoINI("Notas Fiscais", "PosLinNumeroFim", CaminhoINI & "\NotasServico.ini")
'   iPosColNumeroFim = LerArquivoINI("Notas Fiscais", "PosColNumeroFim", CaminhoINI & "\NotasServico.ini")
'   iPosLinProximaNota = LerArquivoINI("Notas Fiscais", "PosLinProximaNota", CaminhoINI & "\NotasServico.ini")
'
'   If rsNFe!Impressa Then
'      Beep
'      MsgBox "Nota Fiscal já Impressa", vbExclamation, "Aviso"
'   End If
'
'   cdlgImprimirNota.CancelError = True
'
'   Registro_Selecionado = False
'''*   VisualizarImpressao
'''*   If Registro_Selecionado Then
'
'      Open LerArquivoINI("Impressoras", "Notas", CaminhoINI & "\System.ini") For Output As #1
''      Open CaminhoINI & "\teste.txt" For Output As #1
'
'      Set rsClientes = cnSistema.Execute("Select * From Clientes Where idCliente = " & rsNFe!idCliente)
'
'    ' Emissao e Numero
'      Print #1, Chr(27) & "x0"; Chr(15)
'      SaltarLinha (iPosLinEmissao)
'      Print #1, Tab(iPosColEmissao); rsNFe!DataEmissao; Tab(iPosColNumero); StrZero(rsNFe!Numero, 6)
'      SaltarLinha (iPosLinCliente)
'
'    ' Cliente
'      Print #1, Tab(iPosColCliente); RemoveAcentos(IIf(Not IsNull(rsClientes!Nome), rsClientes!Nome, ""))
'      SaltarLinha (iPosLinEndereco)
'
'    ' Endereço
'      Print #1, Tab(iPosColEndereco); RemoveAcentos(IIf(Not IsNull(rsClientes!Endereco), rsClientes!Endereco, ""))
'      SaltarLinha (iPosLinCidade)
'
'    ' Cidade, UF e Telefone
'      Print #1, Tab(iPosColCidade); RemoveAcentos(IIf(Not IsNull(rsClientes!Cidade), rsClientes!Cidade, "")); Tab(iPosColUF); IIf(Not IsNull(rsClientes!UF), rsClientes!UF, ""); Tab(iPosColTelefone); IIf(Not IsNull(rsClientes!Telefone1), rsClientes!Telefone1, "")
'      SaltarLinha (iPosLinCNPJ)
'
'    ' CNPJ e IE
'      Print #1, Tab(iPosColCNPJ); IIf(Not IsNull(rsClientes!CNPJ_CPF), rsClientes!CNPJ_CPF, ""); Tab(iPosColIE); IIf(Not IsNull(rsClientes!IE_CI), rsClientes!IE_CI, "")
'      SaltarLinha (iPosLinCobranca)
'
'    ' Cobranca
'      Dim ContCob As Integer
'      ContCob = 1
'
'      Dim DocBol1 As String
'      Dim VenBol1 As String
'      Dim ValBol1 As String
'      Dim DocBol2 As String
'      Dim VenBol2 As String
'      Dim ValBol2 As String
'      Dim DocBol3 As String
'      Dim VenBol3 As String
'      Dim ValBol3 As String
'      Dim DocBol4 As String
'      Dim VenBol4 As String
'      Dim ValBol4 As String
'
'      Set rsNFeBoletos = cnSistema.Execute("SELECT * FROM NFeBoletos WHERE idNFe = " & rsNFe!idNFe)
'      Do While Not rsNFeBoletos.EOF
'         If ContCob = 1 Then
'            DocBol1 = rsNFeBoletos!Numero
'            VenBol1 = FormataTXT(Format(rsNFeBoletos!Valor, "##,##0.00"), 2.1, 10)
'            ValBol1 = rsNFeBoletos!Vencimento
'
'         ElseIf ContCob = 2 Then
'            DocBol2 = rsNFeBoletos!Numero
'            VenBol2 = FormataTXT(Format(rsNFeBoletos!Valor, "##,##0.00"), 2.1, 10)
'            ValBol2 = rsNFeBoletos!Vencimento
'
'         ElseIf ContCob = 3 Then
'            DocBol3 = rsNFeBoletos!Numero
'            VenBol3 = FormataTXT(Format(rsNFeBoletos!Valor, "##,##0.00"), 2.1, 10)
'            ValBol3 = rsNFeBoletos!Vencimento
'
'         ElseIf ContCob = 4 Then
'            DocBol4 = rsNFeBoletos!Numero
'            VenBol4 = FormataTXT(Format(rsNFeBoletos!Valor, "##,##0.00"), 2.1, 10)
'            ValBol4 = rsNFeBoletos!Vencimento
'         End If
'
'         rsNFeBoletos.MoveNext
'         ContCob = ContCob + 1
'      Loop
'
'      Print #1, Tab(iPosColDocBol1); DocBol1; _
'                Tab(iPosColVenBol1); VenBol1; _
'                Tab(iPosColValBol1); ValBol1; _
'                Tab(iPosColDocBol2); DocBol2; _
'                Tab(iPosColVenBol2); VenBol2; _
'                Tab(iPosColValBol2); ValBol2; _
'                Tab(iPosColDocBol3); DocBol3; _
'                Tab(iPosColVenBol3); VenBol3; _
'                Tab(iPosColValBol3); ValBol3; _
'                Tab(iPosColDocBol4); DocBol4; _
'                Tab(iPosColVenBol4); VenBol4; _
'                Tab(iPosColValBol4); ValBol4
'
'      SaltarLinha (iPosLinProdutos)
'
'    ' Produtos
'      Dim cValorBruto As Currency
'      Dim cValorDesconto As Currency
'      Dim cValorBonificacao As Currency
'      Dim cValorLiquido As Currency
'
'      Dim iLinhasProdutos As Integer
'      iLinhasProdutos = LerArquivoINI("Notas Fiscais", "ItensNota", CaminhoINI & "\NotasServico.ini")
'      Set rsNFeItens = cnSistema.Execute("SELECT NFeItens.idNFe, NFeItens.idProduto, NFeItens.idUnidade, Produtos.Descricao, NFeItens.idClassificacaoFiscal, NFeItens.idSituacaoTributaria, NFeItens.Data, NFeItens.Quantidade, NFeItens.ValorUnitario, NFeItens.Desconto, NFeItens.IPI, NFeItens.ICMS, NFeItens.DescricaoComplementar " & _
'                                                  "FROM NFeItens INNER JOIN Produtos ON NFeItens.idProduto = Produtos.idProduto " & _
'                                                  "Where NFeItens.idNFe = " & rsNFe!idNFe & " ORDER BY Produtos.Descricao")
'
'      Do While Not rsNFeItens.EOF
'         Set rsProdutos = cnSistema.Execute("SELECT * FROM Produtos WHERE idProduto = " & rsNFeItens!idProduto)
'         Set rsUnidadesMedida = cnSistema.Execute("SELECT * FROM UnidadesMedida WHERE idUnidadeMedida = " & rsNFeItens!idUnidade)
'         Dim sDescricao As String, sUnidade As String
'         sDescricao = RemoveAcentos(Mid(Trim(rsProdutos!Descricao), 1, 50) & " " & Trim(rsNFeItens!DescricaoComplementar))
'         If Not rsUnidadesMedida.EOF Then
'            sUnidade = rsUnidadesMedida!Sigla
'         Else
'            sUnidade = " "
'         End If
'
'         cValorBruto = (rsNFeItens!quantidade * rsNFeItens!ValorUnitario)
'         cValorDesconto = (((rsNFeItens!quantidade * rsNFeItens!ValorUnitario) * rsNFeItens!Desconto) / 100)
'         cValorBonificacao = (((cValorBruto - cValorDesconto) * rsNFe!Bonificacao) / 100)
'         cValorLiquido = (cValorBruto - cValorDesconto - cValorBonificacao)
'
'         Print #1, Tab(iPosColQuantidade); FormataTXT(Format(rsNFeItens!quantidade, "##,##0.00"), 2.1, 10); _
'                   Tab(iPosColUnidade); sUnidade; _
'                   Tab(iPosColDescricao); sDescricao; _
'                   Tab(iPosColVlUnitario); FormataTXT(Format(rsNFeItens!ValorUnitario, "##,##0.00"), 2.1, 10); _
'                   Tab(iPosColVlLiquido); FormataTXT(Format(cValorLiquido, "##,##0.00"), 2.1, 12)
'
'         iLinhasProdutos = iLinhasProdutos - 1
'         rsNFeItens.MoveNext
'      Loop
'
'      ' Informações do Corpo da Nota
'      SaltarLinha (iLinhasProdutos + iPosLinTotalNota)
'
''      Print #1, Tab(iPosColInfoCN); rsNFe!InformacoesCorpo
''      SaltarLinha (iPosLinBase)
'
'    ' Total
'      Set rsTemp = cnSistema.Execute("Select * From TotalNFe Where Numero = " & mskNumero.text)
'      If Not rsTemp.EOF Then
'         Print #1, Tab(iPosColTotalNota); Format(rsTemp!Total, "##,##0.00")
'         SaltarLinha (iPosLinNumeroFim)
'      Else
'         Print #1, Tab(iPosColTotalNota); Format(0, "##,##0.00")
'         SaltarLinha (iPosLinNumeroFim)
'      End If
'
'    ' Numero da Nota
'      Print #1, Tab(iPosColNumeroFim); StrZero(rsNFe!Numero, 6) & Chr(18)
'      SaltarLinha (iPosLinProximaNota)
'
'      Close #1
'      cnSistema.Execute "Update NFe set " & _
'            "Impressa = " & True & " " & _
'            "Where idNFe = " & rsNFe!idNFe
'''*   End If
'
'End Sub
'
'Private Sub VisualizarImpressao()
'Dim Contador As Integer
'   Screen.MousePointer = vbHourglass
'   frmVisualizaImpressao.lvwDados.ColumnHeaders.Clear
'   frmVisualizaImpressao.lvwDados.ColumnHeaders.Add , , "Nota Eletrônica", 11450
'   frmVisualizaImpressao.lvwDados.ListItems.Clear
'
'   Set rsCFOPs = cnSistema.Execute("Select * From CFOPs Where idCFOP = " & rsNFe!idCFOP)
'   Set rsClientes = cnSistema.Execute("Select * From Clientes Where idCliente = " & rsNFe!idCliente)
'   Set rsNaturezasOperacao = cnSistema.Execute("Select * From NaturezasOperacao Where idNaturezaOperacao = " & rsNFe!idNaturezaOperacao)
'
''  Numero da Nota
'   Contador = 1
'
'   Set ItemList = frmVisualizaImpressao.lvwDados.ListItems.Add(, "R" & CStr(Contador), Space(99) & StrZero(rsNFe!Numero, 6))
'
''  Codigo Fiscal de Operação
'   Contador = Contador + 1
'   Set ItemList = frmVisualizaImpressao.lvwDados.ListItems.Add(, "R" & CStr(Contador), rsNaturezasOperacao!Descricao & Space(20) & rsCFOPs!CFOP)
'
''  Cliente, CNPJ e Emissão
'   Contador = Contador + 1
'   Set ItemList = frmVisualizaImpressao.lvwDados.ListItems.Add(, "R" & CStr(Contador), rsClientes!Codigo & "-" & rsClientes!Nome & Space(10) & rsClientes!CNPJ_CPF & Space(26) & rsNFe!DataEmissao)
'
''  Endereço, Bairro e CEP
'   Contador = Contador + 1
'   Set ItemList = frmVisualizaImpressao.lvwDados.ListItems.Add(, "R" & CStr(Contador), rsClientes!Endereco & Space(10) & rsClientes!Bairro & Space(10) & rsClientes!CEP)
'
''  Cidade, Telefone, UF e Inscrição Estadual
'   Contador = Contador + 1
'   Set ItemList = frmVisualizaImpressao.lvwDados.ListItems.Add(, "R" & CStr(Contador), rsClientes!Cidade & Space(10) & rsClientes!Telefone1 & Space(10) & rsClientes!UF & Space(10) & rsClientes!IE_CI)
'
''  Documento
'   Contador = Contador + 1
'   Set ItemList = frmVisualizaImpressao.lvwDados.ListItems.Add(, "R" & CStr(Contador), rsNFe!Documento & Space(10) & cmbFormaPagamento.text & Space(56) & rsNFe!DataVencimento)
'
'   Contador = Contador + 1
'   Set ItemList = frmVisualizaImpressao.lvwDados.ListItems.Add(, "R" & CStr(Contador), "")
'
''  Produtos
'
'   Dim cValorBruto As Currency
'   Dim cValorDesconto As Currency
'   Dim cValorBonificacao As Currency
'   Dim cValorLiquido As Currency
'
'   Dim ContadorProdutos As Integer
'   ContadorProdutos = 0
'   Set rsNFeItens = cnSistema.Execute("SELECT NFeItens.idNFe, NFeItens.idProduto,NFeItens.idUnidade, Produtos.Descricao, NFeItens.idClassificacaoFiscal, NFeItens.idSituacaoTributaria, NFeItens.Data, NFeItens.Quantidade, NFeItens.ValorUnitario, NFeItens.Desconto, NFeItens.IPI, NFeItens.ICMS, NFeItens.DescricaoComplementar " & _
'                                               "FROM NFeItens INNER JOIN Produtos ON NFeItens.idProduto = Produtos.idProduto " & _
'                                               "Where NFeItens.idNFe = " & rsNFe!idNFe & " ORDER BY Produtos.Descricao")
'
'   Do While Not rsNFeItens.EOF
'      Contador = Contador + 1
'      ContadorProdutos = ContadorProdutos + 1
'
'      Set rsProdutos = cnSistema.Execute("SELECT * FROM Produtos WHERE idProduto = " & rsNFeItens!idProduto)
'      Set rsUnidadesMedida = cnSistema.Execute("SELECT * FROM UnidadesMedida WHERE idUnidadeMedida = " & rsNFeItens!idUnidade)
'      Set rsSituacoesTributarias = cnSistema.Execute("SELECT * FROM SituacoesTributarias WHERE idSituacaoTributaria = " & rsNFeItens!idSituacaoTributaria)
'      Dim sDescricao As String, sUnidade As String, sCST As String
'      sDescricao = Mid(Trim(rsProdutos!Descricao), 1, 50) & " " & Trim(rsNFeItens!DescricaoComplementar)
'      If Not rsUnidadesMedida.EOF Then
'         sUnidade = rsUnidadesMedida!Sigla
'      Else
'         sUnidade = " "
'      End If
'
'      If Not rsSituacoesTributarias.EOF Then
'         sCST = rsSituacoesTributarias!Codigo
'      Else
'         sCST = " "
'      End If
'
'      cValorBruto = (rsNFeItens!quantidade * rsNFeItens!ValorUnitario)
'      cValorDesconto = (((rsNFeItens!quantidade * rsNFeItens!ValorUnitario) * rsNFeItens!Desconto) / 100)
'      cValorBonificacao = (((cValorBruto - cValorDesconto) * rsNFe!Bonificacao) / 100)
'      cValorLiquido = (cValorBruto - cValorDesconto - cValorBonificacao)
'
'      Set ItemList = frmVisualizaImpressao.lvwDados.ListItems.Add(, "R" & CStr(Contador), rsProdutos!Codigo & _
'                Space(2) & FormataTXT(sDescricao, 1, 44) & _
'                Space(3) & sCST & _
'                Space(3) & sUnidade & _
'                Space(3) & FormataTXT(Format(rsNFeItens!quantidade, "##,##0.00"), 2.1, 10) & _
'                Space(3) & FormataTXT(Format(rsNFeItens!ValorUnitario, "##,##0.00"), 2.1, 10) & _
'                Space(3) & FormataTXT(Format(rsNFeItens!Desconto, "#0.00"), 2.1, 4) & IIf(rsNFe!Bonificacao = 0, "", "/" & FormataTXT(Format(rsNFe!Bonificacao, "#0.00"), 2.1, 4)) & _
'                Space(3) & FormataTXT(Format(cValorLiquido, "##,##0.00"), 2.1, 10) & _
'                Space(3) & Format(rsNFeItens!ICMS, "##0"))
'
'      rsNFeItens.MoveNext
'   Loop
'
'   Contador = Contador + 1
'   Set ItemList = frmVisualizaImpressao.lvwDados.ListItems.Add(, "R" & CStr(Contador), "")
'
'   Contador = Contador + 1
''   Set ItemList = frmVisualizaImpressao.lvwDados.ListItems.Add(, "R" & CStr(Contador), rsNFe!InformacoesCorpo)
'
'   ' Informações do Corpo da Nota
'   Dim sInformacoesCorpo1 As String, sInformacoesCorpo2 As String, sInformacoesCorpo3 As String
'   If Len(rsNFe!InformacoesCorpo) >= 240 Then
'      sInformacoesCorpo1 = RemoveAcentos(Mid(rsNFe!InformacoesCorpo, 1, 120))
'      sInformacoesCorpo2 = RemoveAcentos(Mid(rsNFe!InformacoesCorpo, 121, 120))
'      sInformacoesCorpo3 = RemoveAcentos(Mid(rsNFe!InformacoesCorpo, 241, 120))
'
'      Contador = Contador + 1
'      Set ItemList = frmVisualizaImpressao.lvwDados.ListItems.Add(, "R" & CStr(Contador), sInformacoesCorpo1)
'      Contador = Contador + 1
'      Set ItemList = frmVisualizaImpressao.lvwDados.ListItems.Add(, "R" & CStr(Contador), sInformacoesCorpo2)
'      Contador = Contador + 1
'      Set ItemList = frmVisualizaImpressao.lvwDados.ListItems.Add(, "R" & CStr(Contador), sInformacoesCorpo3)
'   Else
'      If Len(rsNFe!InformacoesCorpo) >= 120 Then
'         sInformacoesCorpo1 = RemoveAcentos(Mid(rsNFe!InformacoesCorpo, 1, 120))
'         sInformacoesCorpo2 = RemoveAcentos(Mid(rsNFe!InformacoesCorpo, 121, 120))
'
'         Contador = Contador + 1
'         Set ItemList = frmVisualizaImpressao.lvwDados.ListItems.Add(, "R" & CStr(Contador), sInformacoesCorpo1)
'         Contador = Contador + 1
'         Set ItemList = frmVisualizaImpressao.lvwDados.ListItems.Add(, "R" & CStr(Contador), sInformacoesCorpo2)
'      Else
'         If Len(rsNFe!InformacoesCorpo) >= 1 Then
'            sInformacoesCorpo1 = RemoveAcentos(rsNFe!InformacoesCorpo)
'
'            Contador = Contador + 1
'            Set ItemList = frmVisualizaImpressao.lvwDados.ListItems.Add(, "R" & CStr(Contador), sInformacoesCorpo1)
'         End If
'      End If
'   End If
'
'   Contador = Contador + 1
'   Set ItemList = frmVisualizaImpressao.lvwDados.ListItems.Add(, "R" & CStr(Contador), "")
'
''  Total
'
'   Set rsTemp = cnSistema.Execute("Select * From TotalNFe Where Numero = " & mskNumero.text)
'   Dim dBaseCalculo As Double, dValorICMS As Double
'   If rsTemp!ValorICMS > 0 Then
'      dBaseCalculo = rsTemp!BaseCalculo
'      dValorICMS = rsTemp!ValorICMS
'   Else
'      dBaseCalculo = 0
'      dValorICMS = 0
'   End If
'
'   If Not rsTemp.EOF Then
'      Contador = Contador + 1
'      Set ItemList = frmVisualizaImpressao.lvwDados.ListItems.Add(, "R" & CStr(Contador), "Total de Produtos: " & StrZero(ContadorProdutos, 3) & Space(31) & Format(dBaseCalculo, "##,##0.00") & Space(10) & Format(dValorICMS, "##,##0.00") & Space(10) & Format(rsTemp!Total, "##,##0.00"))
'
'      Contador = Contador + 1
'      Set ItemList = frmVisualizaImpressao.lvwDados.ListItems.Add(, "R" & CStr(Contador), Space(90) & Format(rsTemp!Total, "##,##0.00"))
'   Else
'      Contador = Contador + 1
''      Set ItemList = frmVisualizaImpressao.lvwDados.ListItems.Add(, "R" & CStr(Contador), Space(71) & Format(rsNFe!BaseCalculoICMS, "#0.00") & Space(10) & Format(rsNFe!ValorICMS, "#0.00") & Space(10) & Format(0, "#0.00"))
'      Set ItemList = frmVisualizaImpressao.lvwDados.ListItems.Add(, "R" & CStr(Contador), Space(71) & Format(dBaseCalculo, "##,##0.00") & Space(10) & Format(dValorICMS, "##,##0.00") & Space(10) & Format(0, "##,##0.00"))
'
'      Contador = Contador + 1
'      Set ItemList = frmVisualizaImpressao.lvwDados.ListItems.Add(, "R" & CStr(Contador), Space(99) & Format(0, "##,##0.00"))
'   End If
'
''  Transportador
'   Set rsTransportador = cnSistema.Execute("SELECT * FROM Transportadores WHERE idTransportador = " & rsNFe!idTransportador)
'   If Not rsTransportador.EOF Then
'      Contador = Contador + 1
'      Set ItemList = frmVisualizaImpressao.lvwDados.ListItems.Add(, "R" & CStr(Contador), rsTransportador!Nome & Space(10) & rsNFe!PlacaVeiculo & Space(10) & rsTransportador!UFPlaca & Space(10) & rsTransportador!CNPJ_CPF)
'
'      Contador = Contador + 1
'      Set ItemList = frmVisualizaImpressao.lvwDados.ListItems.Add(, "R" & CStr(Contador), rsTransportador!Endereco & Space(10) & rsTransportador!Cidade & Space(10) & rsTransportador!UF & Space(10) & rsTransportador!IE_CI)
'   Else
'      Contador = Contador + 2
'   End If
'   Contador = Contador + 1
'   Set ItemList = frmVisualizaImpressao.lvwDados.ListItems.Add(, "R" & CStr(Contador), rsNFe!VolumeQuantidade & Space(10) & rsNFe!VolumeMarca & Space(10) & rsNFe!VolumeNumero & Space(10) & Format(rsNFe!VolumePesoBruto, "##,##0") & Space(10) & Format(rsNFe!VolumePesoLiquido, "##,##0"))
'
'   ' Volume
'   Contador = Contador + 1
'   Set ItemList = frmVisualizaImpressao.lvwDados.ListItems.Add(, "R" & CStr(Contador), Space(5) & rsNFe!VolumeQuantidade & Space(5) & rsNFe!VolumeEspecie & Space(5) & rsNFe!VolumeMarca & Space(5) & rsNFe!VolumeNumero & Space(5) & Format(rsTemp!PesoBruto, "##,##0.00") & Space(5) & Format(rsTemp!PesoLiquido, "##,##0.00"))
'   Contador = Contador + 1
'   Set ItemList = frmVisualizaImpressao.lvwDados.ListItems.Add(, "R" & CStr(Contador), "")
'
'   ' Dados Adicionais
'   Dim sDadosAdicionais1 As String, sDadosAdicionais2 As String, sDadosAdicionais3 As String
'   If Len(rsNFe!DadosAdicionais) >= 120 Then
'      sDadosAdicionais1 = RemoveAcentos(Mid(rsNFe!DadosAdicionais, 1, 60))
'      sDadosAdicionais2 = RemoveAcentos(Mid(rsNFe!DadosAdicionais, 61, 60))
'      sDadosAdicionais3 = RemoveAcentos(Mid(rsNFe!DadosAdicionais, 121, 60))
'
'      Contador = Contador + 1
'      Set ItemList = frmVisualizaImpressao.lvwDados.ListItems.Add(, "R" & CStr(Contador), sDadosAdicionais1)
'      Contador = Contador + 1
'      Set ItemList = frmVisualizaImpressao.lvwDados.ListItems.Add(, "R" & CStr(Contador), sDadosAdicionais2)
'      Contador = Contador + 1
'      Set ItemList = frmVisualizaImpressao.lvwDados.ListItems.Add(, "R" & CStr(Contador), sDadosAdicionais3)
'      Contador = Contador + 1
'      Set ItemList = frmVisualizaImpressao.lvwDados.ListItems.Add(, "R" & CStr(Contador), "")
'   Else
'      If Len(rsNFe!DadosAdicionais) >= 60 Then
'         sDadosAdicionais1 = RemoveAcentos(Mid(rsNFe!DadosAdicionais, 1, 60))
'         sDadosAdicionais2 = RemoveAcentos(Mid(rsNFe!DadosAdicionais, 61, 60))
'
'         Contador = Contador + 1
'         Set ItemList = frmVisualizaImpressao.lvwDados.ListItems.Add(, "R" & CStr(Contador), sDadosAdicionais1)
'         Contador = Contador + 1
'         Set ItemList = frmVisualizaImpressao.lvwDados.ListItems.Add(, "R" & CStr(Contador), sDadosAdicionais2)
'         Contador = Contador + 1
'         Set ItemList = frmVisualizaImpressao.lvwDados.ListItems.Add(, "R" & CStr(Contador), "")
'      Else
'         If Len(rsNFe!DadosAdicionais) >= 1 Then
'            sDadosAdicionais1 = RemoveAcentos(rsNFe!DadosAdicionais)
'
'            Contador = Contador + 1
'            Set ItemList = frmVisualizaImpressao.lvwDados.ListItems.Add(, "R" & CStr(Contador), sDadosAdicionais1)
'            Contador = Contador + 1
'            Set ItemList = frmVisualizaImpressao.lvwDados.ListItems.Add(, "R" & CStr(Contador), "")
'         End If
'      End If
'   End If
'
''  Numero da Nota
'   Contador = Contador + 1
'   Set ItemList = frmVisualizaImpressao.lvwDados.ListItems.Add(, "R" & CStr(Contador), Space(99) & StrZero(rsNFe!Numero, 6))
'
'   Screen.MousePointer = vbDefault
'   frmNFe.Caption = "Visualiza Impressão de Notas Eletrônicas"
'   frmVisualizaImpressao.Show vbModal
'End Sub
'
'Sub ImprimirBoleto()
'Dim Contador As Integer
'Dim Total_Registros As Integer
'Dim PaginaInicial, Paginafinal, NumeroCopias, i
'Dim B_Local As String
'Dim B_Vencimento As String
'Dim B_Emissao As Date
'Dim B_Documento As String
'Dim B_Valor As Double
'Dim B_Taxa As Double
'Dim B_Inst1 As String
'Dim B_Inst2 As String
'Dim B_Inst3 As String
'Dim B_Inst4 As String
'Dim B_Inst5 As String
'Dim B_Inst6 As String
'Dim B_CliCNPJ As String
'Dim B_Endereco As String
'
'Dim iPosLinLocal As Integer
'Dim iPosColLocal As Integer
'Dim iPosColVencimento As Integer
'Dim iPosLinEmissao As Integer
'Dim iPosColEmissao As Integer
'Dim iPosColDocumento As Integer
'Dim iPosLinValor As Integer
'Dim iPosColValor As Integer
'Dim iPosLinInstrucoes As Integer
'Dim iPosColInstrucoes As Integer
'Dim iPosLinCliente As Integer
'Dim iPosColCliente As Integer
'Dim iPosLinProximo As Integer
'
'     Total_Registros = Registros(cnSistema, "ContasBancarias")
'     If Total_Registros = 0 Then
'        MsgBox "Não existe nenhum Registro na Tabela", vbOKOnly, "Visualiza"
'        Exit Sub
'     End If
'     Contador = 1
'     Status = 4
'     Screen.MousePointer = vbHourglass
'     frmVisualiza.lvwDados.ColumnHeaders.Clear
'     frmVisualiza.lvwDados.ColumnHeaders.Add , , "Banco", 6000
'     rsContasBancarias.MoveFirst
'     frmVisualiza.lvwDados.ListItems.Clear
'     Do While Not rsContasBancarias.EOF
'        frmNFe.Caption = "Processando " & StrZero(Contador, 8) & " de " & StrZero(Total_Registros, 8)
'        Set ItemList = frmVisualiza.lvwDados.ListItems.Add(, "R" & CStr(rsContasBancarias!idContaBancaria), rsContasBancarias!Descricao)
'        rsContasBancarias.MoveNext
'        Contador = Contador + 1
'     Loop
'     rsContasBancarias.MoveFirst
'     Registro_Selecionado = False
'     Screen.MousePointer = vbDefault
'     frmNFe.Caption = "Notas de Saida Eletrônicas"
'     frmVisualiza.Show vbModal
'     If Registro_Selecionado Then
'        rsContasBancarias.Find "idContaBancaria = " & Val(Mid(frmVisualiza.lvwDados.SelectedItem.Key, 2, Len(frmVisualiza.lvwDados.SelectedItem.Key)))
'     End If
'
'   ' Mostra a janela para impressora
'
'   ' Captura os valores definidos pelo usuário na janela
'
'     iPosLinLocal = rsContasBancarias!PosicaoLocalLinha
'     iPosColLocal = rsContasBancarias!PosicaoLocalColuna
'     iPosColVencimento = rsContasBancarias!PosicaoVencimentoColuna
'     iPosLinEmissao = rsContasBancarias!PosicaoEmissaoLinha
'     iPosColEmissao = rsContasBancarias!PosicaoEmissaoColuna
'     iPosColDocumento = rsContasBancarias!PosicaoDocumentoColuna
'     iPosLinValor = rsContasBancarias!PosicaoValorLinha
'     iPosColValor = rsContasBancarias!PosicaoValorColuna
'     iPosLinInstrucoes = rsContasBancarias!PosicaoInstrucoesLinha
'     iPosColInstrucoes = rsContasBancarias!PosicaoInstrucoesColuna
'     iPosLinCliente = rsContasBancarias!PosicaoClienteLinha
'     iPosColCliente = rsContasBancarias!PosicaoClienteColuna
'     iPosLinProximo = rsContasBancarias!PosicaoProximoLinha
'
'     Set rsTemp = cnSistema.Execute("SELECT * From NFe WHERE NFe.Documento = '" & rsNFe!Documento & "'")
'
'     Do While Not rsTemp.EOF
'        B_Documento = B_Documento & StrZero(rsTemp!Numero, 6) & "/"
'
'        Set rsTemp2 = cnSistema.Execute("Select * From TotalNFe Where Numero = " & rsTemp!Numero)
'        If Not rsTemp2.EOF Then
'           B_Valor = B_Valor + rsTemp2!Total
'        Else
'           B_Valor = B_Valor + 0
'        End If
'        rsTemp.MoveNext
'     Loop
'
'     NumeroCopias = InputBox("Digite a Quantidade de Parcelas", "Quantidade", 1)
'     If Val(NumeroCopias) = 0 Then Exit Sub
'
'     B_Local = rsContasBancarias!LocalPagamento
'     B_Emissao = rsNFe!DataEmissao
'     B_Valor = B_Valor / NumeroCopias
'     B_Taxa = rsContasBancarias!TaxaBancaria
'     B_Inst1 = rsContasBancarias!Instrucoes1
'     B_Inst2 = rsContasBancarias!Instrucoes2
'     B_Inst3 = rsContasBancarias!Instrucoes3
'     B_Inst4 = rsContasBancarias!Instrucoes4
'     B_Inst5 = rsContasBancarias!Instrucoes5
'     B_CliCNPJ = Trim(rsClientes!Nome) & " CPF/CNPJ: " & rsClientes!CNPJ_CPF
'     B_Endereco = Trim(rsClientes!Endereco) & " " & Trim(rsClientes!Bairro) & " " & Trim(rsClientes!Cidade) & " " & Trim(rsClientes!CEP)
'
'     Dim sMsg As String, sVencimento As String
'     sMsg = "Digite o Vencimento"
'
'     Open LerArquivoINI("Impressoras", "Boletos1", CaminhoINI & "\System.ini") For Output As #1
'
'     For i = 1 To NumeroCopias
'       ' Digitar o Vencimento
'         Dim bValidaData As Boolean
'         bValidaData = False
'         Do While Not bValidaData
'            B_Vencimento = InputBox("Digite o Vencimento", "Vencimento", "")
'            If Val(B_Vencimento) = 0 Then
'               Close #1
'               Exit Sub
'            End If
'
'            If FData(B_Vencimento) Then
'               bValidaData = True
'            End If
'         Loop
'
'       ' Local de Pagamento e Vencimento
'         SaltarLinha (iPosLinLocal)
'         Print #1, Chr(27) & Chr(15); Tab(iPosColLocal); B_Local; Tab(iPosColVencimento); B_Vencimento
'
'       ' Emissão e Documento
'         SaltarLinha (iPosLinEmissao)
'         Print #1, Tab(iPosColEmissao); B_Emissao; Tab(iPosColDocumento); B_Documento; Chr(27) & Chr(8)
'
'       ' Valor
'         SaltarLinha (iPosLinValor)
'         Print #1, Tab(iPosColValor); Format(B_Valor, "Standard")
'
'       ' Instruções
'         SaltarLinha (iPosLinInstrucoes)
'         Print #1, Tab(iPosColInstrucoes); B_Inst1
'         Print #1, Tab(iPosColInstrucoes); B_Inst2
'         Print #1, Tab(iPosColInstrucoes); B_Inst3
'         Print #1, Tab(iPosColInstrucoes); B_Inst4
'         Print #1, Tab(iPosColInstrucoes); B_Inst5; Chr(27) & Chr(15)
'
'       ' Instruções
'         SaltarLinha (iPosLinCliente)
'         Print #1, Tab(iPosColCliente); B_CliCNPJ
'         Print #1, Tab(iPosColCliente); B_Endereco; Chr(27) & Chr(8)
'
'       ' Próximo
'         SaltarLinha (iPosLinProximo + 1)
'     Next
'     Close #1
'
'End Sub
'
'Sub CopiarNota()
'
'   If MsgBox("Confirma Cópia", vbYesNo + vbQuestion, "Inclusão") = vbYes Then
'      RegistroAtual = IIf(rsNFe.EOF, 0, rsNFe!idNFe)
'      Dim iNumero As Long
'
'   '  Notas Eletronicas
'      rsNFe.MoveLast
'      iNumero = rsNFe!Numero + 1
'      If RegistroAtual <> 0 Then
'         rsNFe.MoveFirst
'         rsNFe.Find "idNFe = " & RegistroAtual
'      End If
'
'      cnSistema.Execute "Insert Into NFe (Numero,Cupom,idCliente,idNaturezaOperacao,idCFOP,DadosAdicionais,DataEmissao,DataVencimento,Hora,BaseCalculoICMS,ValorICMS,ValorFrete,ValorTotalProdutos,BaseICMSSubstituicao,ValorICMSSubstituicao,OutrasDespesas,ValorTotalNota,idTransportador,FreteConta,PlacaVeiculo,VolumeQuantidade,VolumeMarca,VolumeNumero,VolumePesoBruto,VolumePesoLiquido,InformacoesCorpo,idFormaPagamento,DescontoGeral,Bonificacao,Documento,Observacao) " & _
'               "Values (" & iNumero & "," & rsNFe!Cupom & "," & rsNFe!idCliente & "," & rsNFe!idNaturezaOperacao & "," & rsNFe!idCFOP & ",'" & rsNFe!DadosAdicionais & "','" & rsNFe!DataEmissao & "','" & rsNFe!DataVencimento & "','" & rsNFe!Hora & "','" & rsNFe!BaseCalculoICMS & "','" & rsNFe!ValorICMS & "'," & _
'                       "'" & rsNFe!valorfrete & "','" & rsNFe!ValorTotalProdutos & "','" & rsNFe!BaseICMSSubstituicao & "','" & rsNFe!ValorICMSSubstituicao & "','" & rsNFe!OutrasDespesas & "'," & _
'                       "'" & rsNFe!ValorTotalNota & "'," & rsNFe!idTransportador & "," & rsNFe!FreteConta & ",'" & rsNFe!PlacaVeiculo & "','" & rsNFe!VolumeQuantidade & "','" & rsNFe!VolumeMarca & "','" & rsNFe!VolumeNumero & "','" & rsNFe!VolumePesoBruto & "','" & rsNFe!VolumePesoLiquido & "','" & rsNFe!InformacoesCorpo & "'," & 0 & _
'                       "," & rsNFe!idFormaPagamento & ",'" & rsNFe!DescontoGeral & "','" & rsNFe!Bonificacao & "','" & rsNFe!Documento & "','" & rsNFe!Observacao & "')"
'
'      rsNFe.Requery
'
'   '  Notas Eletronicas Itens
'      rsNFe.MoveLast
'      iNumero = rsNFe!idNFe
'      If RegistroAtual <> 0 Then
'         rsNFe.MoveFirst
'         rsNFe.Find "idNFe = " & RegistroAtual
'      End If
'
'      Set rsNFeItens = cnSistema.Execute("Select * From NFeItens Where idNFe = " & rsNFe!idNFe)
'      Do While Not rsNFeItens.EOF
'         cnSistema.Execute "Insert Into NFeItens (idNFe,idProduto,Data,Quantidade,Desconto,ValorUnitario,ICMS,BaseReduzida,DescricaoComplementar) " & _
'                  "Values (" & iNumero & "," & rsNFeItens!idProduto & ",'" & rsNFeItens!Data & _
'                  "','" & rsNFeItens!quantidade & "','" & rsNFeItens!Desconto & "','" & rsNFeItens!ValorUnitario & "','" & rsNFeItens!ICMS & _
'                  "','" & rsNFeItens!BaseReduzida & "','" & rsNFeItens!DescricaoComplementar & "')"
'
'         rsNFeItens.MoveNext
'      Loop
'
'      rsNFe.MoveLast
'      Prencher_Campos
'   End If
'End Sub
'
'Private Function ImprimirDanfe()
'Dim sPercMargAdICMSST As String
'Dim sUF As String
'Dim sChaveAcesso As String
'Dim sNaturezaOperacao As String
'Dim sFormaPagamento As String
'Dim sModelo As String
'Dim sSerie As String
'Dim sNumero As String
'Dim sDataEmissao As String
'Dim sDataSaida As String
'Dim sTipoNF As String
'Dim sCodigoMunicipio As String
'Dim sFormatoDANFE As String
'Dim sTipoEmissao As String
'Dim sDVChaveAcesso As String
'Dim sidAmbiente As String
'Dim sFinalidade As String
'Dim sProcessoEmissao As String
'Dim sVersaoAplicativo As String
'Dim sNomeArquivo As String
'
'   If MsgBox("Confirma Emissão da Nota", vbYesNo + vbInformation, "Confirmação") = vbNo Then
'      Exit Function
'   End If
'
''  Identificação do Arquivo
'   Set rsEmpresa = cnSistema.Execute("Select * From Empresa")
'   sNomeArquivo = StrZero(Val(mskNumero.text), 6) & "_" & RemoveCaracteres(rsEmpresa!CNPJ_CPF) & "_55_" & Mid(mskDataEmissao.text, 1, 2) & "_" & Mid(mskDataEmissao.text, 4, 2) & "_" & Mid(mskDataEmissao.text, 7, 4) & "-nfe.txt"
'
'   Open "C:\NFe\XML\" & sNomeArquivo For Output As #1
'
''   Set rsNFe = cnSistema.Execute("Select * From NFe WHERE GeradaNFe=False")
'   Print #1, "NOTAFISCAL|1"
'
'   ' Cabecalho
'   Print #1, "A|1.10|NFe"
'
'   ' Identificadores
'   Set rsNaturezasOperacao = cnSistema.Execute("Select * From NaturezasOperacao WHERE idNaturezaOperacao=" & rsNFe!idNaturezaOperacao)
'   Set rsFormasPagamento = cnSistema.Execute("Select * From FormasPagamento WHERE idFormaPagamento=" & rsNFe!idFormaPagamento)
'   Set rsCFOPs = cnSistema.Execute("Select * From CFOPs WHERE idCFOP=" & rsNFe!idCFOP)
'
'   Set rsTemp = cnSistema.Execute("Select * From UFs WHERE Sigla=" & rsEmpresa!UF)
'   If Not rsTemp.EOF Then
'      sUF = rsTemp!Codigo ' Minas Gerais
'   Else
'      sUF = "  " ' Minas Gerais
'   End If
'
'   sChaveAcesso = ""
'   sNaturezaOperacao = IIf(Not rsNaturezasOperacao.EOF, rsNaturezasOperacao!Descricao, "VENDA")
'   If Not rsFormasPagamento.EOF Then
'      If rsFormasPagamento!TipoPagamento <= 1 Then
'         sFormaPagamento = "0"
'      Else
'         sFormaPagamento = "1"
'      End If
'   Else
'      sFormaPagamento = "0"
'   End If
'   sModelo = "55"
'   sSerie = "1"
'   sNumero = rsNFe!Numero
'   sDataEmissao = Format(rsNFe!DataEmissao, "yyyy-mm-dd")
'   sDataSaida = Format(rsNFe!DataVencimento, "yyyy-mm-dd")
'   sTipoNF = IIf(rsCFOPs!Tipo = 0, 0, 1)
'
'   Set rsTemp = cnSistema.Execute("Select * From Municipios WHERE Nome=" & rsEmpresa!Nome)
'   If Not rsTemp.EOF Then
'      sCodigoMunicipio = RemoveCaracteres(rsTemp!Codigo)
'   Else
'      sCodigoMunicipio = "5212501"
'   End If
'
'   sFormatoDANFE = "1"
'   sTipoEmissao = LerArquivoINI("Notas Fiscais", "TipoEmissaoNFe", CaminhoINI & "\System.ini")
'   sDVChaveAcesso = ""
'   sidAmbiente = LerArquivoINI("Notas Fiscais", "idAmbienteNFe", CaminhoINI & "\System.ini") ' Trocar pra 1 no Oficial
'   sFinalidade = "1" ' NFe Normal
''   sProcessoEmissao = "3" ' Utilizando Software do Fisco
'   sProcessoEmissao = "0" ' Utilizando Aplicativo do contribuinte
'   sVersaoAplicativo = "1.4.1"
'   '      sVersaoAplicativo = "TESTE 1.4.0"
'
'   Print #1, "B|" & _
'             sUF & "|" & _
'             sChaveAcesso & "|" & _
'             sNaturezaOperacao & "|" & _
'             sFormaPagamento & "|" & _
'             sModelo & "|" & _
'             sSerie & "|" & _
'             sNumero & "|" & _
'             sDataEmissao & "|" & _
'             sDataSaida & "|" & _
'             sTipoNF & "|" & _
'             sCodigoMunicipio & "|" & _
'             sFormatoDANFE & "|" & _
'             sTipoEmissao & "|" & _
'             sDVChaveAcesso & "|" & _
'             sidAmbiente & "|" & _
'             sFinalidade & "|" & _
'             sProcessoEmissao & "|" & _
'             sVersaoAplicativo
'
'   ' Emitente
'   Dim sERazaoSocial As String
'   Dim sEFantasia As String
'   Dim sEIE As String
'   Dim sEIEST As String
'   Dim sEIM As String
'   Dim sECNAE As String
'
'   sERazaoSocial = rsEmpresa!Nome
'   sEFantasia = ""
'   sEIE = IIf(rsEmpresa!IE_CI <> "ISENTO", Trim(RemoveCaracteres(rsEmpresa!IE_CI)), "ISENTO")
'   sEIEST = ""
'   sEIM = ""
'   sECNAE = ""
'
'   Print #1, "C|" & _
'             sERazaoSocial & "|" & _
'             sEFantasia & "|" & _
'             sEIE & "|" & _
'             sEIEST & "|" & _
'             sEIM & "|" & _
'             sECNAE
'
'   Dim sECNPJ As String
'   sECNPJ = RemoveCaracteres(rsEmpresa!CNPJ_CPF)
'
'   Print #1, "C02|" & _
'             sECNPJ
'
'   Dim sELogradouro As String
'   Dim sENumero As String
'   Dim sEComplemento As String
'   Dim sEBairro As String
'   Dim sECodigoMunicipio As String
'   Dim sEMunicipio As String
'   Dim sEUF As String
'   Dim sECEP As String
'   Dim sECodigoPais As String
'   Dim sEPais As String
'   Dim sETelefone As String
'
'   sELogradouro = "RUA"
'   sENumero = "25"
'   sEComplemento = rsEmpresa!Endereco
'   sEBairro = rsEmpresa!Bairro
'   sECodigoMunicipio = "5212501"
'   sEMunicipio = rsEmpresa!Cidade
'   sEUF = rsEmpresa!UF
'   sECEP = RemoveCaracteres(rsEmpresa!CEP)
'   sECodigoPais = "1058"
'   sEPais = "BRASIL"
'   sETelefone = RemoveCaracteres(rsEmpresa!Telefone1)
'
'   Print #1, "C05|" & _
'             sELogradouro & "|" & _
'             sENumero & "|" & _
'             sEComplemento & "|" & _
'             sEBairro & "|" & _
'             sECodigoMunicipio & "|" & _
'             sEMunicipio & "|" & _
'             sEUF & "|" & _
'             sECEP & "|" & _
'             sECodigoPais & "|" & _
'             sEPais & "|" & _
'             sETelefone
'
'   ' Destinatario
'   Set rsClientes = cnSistema.Execute("Select * From Clientes WHERE idCliente=" & rsNFe!idCliente)
'
'   Dim sDRazaoSocial As String
'   Dim sDIE As String
'   Dim sDISUF As String
'
'   sDRazaoSocial = rsClientes!Nome
'   sDIE = IIf(rsClientes!IE_CI <> "ISENTO", Trim(RemoveCaracteres(rsClientes!IE_CI)), "ISENTO")
'   sDISUF = ""
'
'   Print #1, "E|" & _
'             sDRazaoSocial & "|" & _
'             sDIE & "|" & _
'             sDISUF
'
'   Dim sDCNPJ As String
'   sDCNPJ = RemoveCaracteres(rsClientes!CNPJ_CPF)
'
'   If Len(Trim(sDCNPJ)) > 11 Then
'      Print #1, "E02|" & _
'                sDCNPJ
'   Else
'      Print #1, "E03|" & _
'                sDCNPJ
'   End If
'
'   ' Endereco
'   Set rsLogradouros = cnSistema.Execute("Select * From Logradouros WHERE idLogradouro=" & rsClientes!idLogradouro)
'   Set rsMunicipios = cnSistema.Execute("Select * From Municipios WHERE idMunicipio=" & rsClientes!idMunicipio)
'
'   Dim sDLogradouro As String
'   Dim sDNumero As String
'   Dim sDComplemento As String
'   Dim sDBairro As String
'   Dim sDCodigoMunicipio As String
'   Dim sDMunicipio As String
'   Dim sDUF As String
'   Dim sDCEP As String
'   Dim sDCodigoPais As String
'   Dim sDPais As String
'   Dim sDTelefone As String
'
'   sDLogradouro = IIf(Not rsLogradouros.EOF, rsLogradouros!Abreviacao, ".")
'   sDNumero = Trim(rsClientes!Numero)
'   sDComplemento = Trim(rsClientes!Endereco)
'   sDBairro = rsClientes!Bairro
'   sDCodigoMunicipio = Trim(rsClientes!CodigoMunicipio)
'   sDMunicipio = rsMunicipios!Nome
'   sDUF = rsClientes!UF
'   sDCEP = RemoveCaracteres(rsClientes!CEP)
'   sDCodigoPais = "1058"
'   sDPais = "BRASIL"
'   sDTelefone = StrZero(Val(rsClientes!PrefixoFone1), 2) & Trim(FormataTXT(RemoveCaracteres(rsClientes!Telefone1), 1, 10))
'
'   Print #1, "E05|" & _
'             sDLogradouro & "|" & _
'             sDNumero & "|" & _
'             sDComplemento & "|" & _
'             sDBairro & "|" & _
'             sDCodigoMunicipio & "|" & _
'             sDMunicipio & "|" & _
'             sDUF & "|" & _
'             sDCEP & "|" & _
'             sDCodigoPais & "|" & _
'             sDPais & "|" & _
'             sDTelefone
'
'   ' Itens
'   Dim Contador As Integer
'   Contador = 1
'
'   Dim dValorTotalBC As Double
'   Dim dValorTotalICMS As Double
'   Dim dValorTotalBCST As Double
'   Dim dValorTotalICMSST As Double
'   Dim dValorTotalProdutos As Double
'   Dim dValorTotalFrete As Double
'   Dim dValorTotalSeguro As Double
'   Dim dValorTotalDesconto As Double
'   Dim dValorTotalII As Double
'   Dim dValorTotalIPI As Double
'   Dim dValorTotalPIS As Double
'   Dim dValorTotalCofins As Double
'   Dim dValorTotalOutro As Double
'   Dim dValorTotalNFe As Double
'
'   dValorTotalBC = 0
'   dValorTotalICMS = 0
'   dValorTotalBCST = 0
'   dValorTotalICMSST = 0
'   dValorTotalProdutos = 0
'   dValorTotalFrete = 0
'   dValorTotalSeguro = 0
'   dValorTotalDesconto = 0
'   dValorTotalII = 0
'   dValorTotalIPI = 0
'   dValorTotalPIS = 0
'   dValorTotalCofins = 0
'   dValorTotalOutro = 0
'   dValorTotalNFe = 0
'
'   Set rsNFeItens = cnSistema.Execute("SELECT * FROM NFeItens WHERE NFeItens.idNFe = " & rsNFe!idNFe)
'   Do While Not rsNFeItens.EOF
'      Set rsProdutos = cnSistema.Execute("SELECT * FROM Produtos WHERE idProduto = " & rsNFeItens!idProduto)
'      Set rsUnidades = cnSistema.Execute("Select * From UnidadesMedida WHERE idUnidadeMedida=" & rsProdutos!idUnidade)
'      Set rsSituacoesTributarias = cnSistema.Execute("Select * from SituacoesTributarias WHERE idSituacaoTributaria=" & rsNFeItens!idSituacaoTributaria)
'
'      Print #1, "H|" & _
'                Contador & "|" & _
'                ""
'
'      Dim sCodigoProduto As String
'      Dim sCodigoBarras As String
'      Dim sDescricaoProduto As String
'      Dim sCodigoNCM As String
'      Dim sEXTIPI As String
'      Dim sGenero As String
'      Dim sCFOP As String
'      Dim sUnidComercial As String
'      Dim sQuantidadeComercial As String
'      Dim sVlUnitarioComercial As String
'      Dim sVlTotalBruto As String
'      Dim sCodigoBarrasTrib As String
'      Dim sUnidTrib As String
'      Dim sQuantidadeTrib As String
'      Dim sVlUnitarioTrib As String
'      Dim sVlFrete As String
'      Dim sVlSeguro As String
'      Dim sVlDesconto As String
'
'      sCodigoProduto = StrZero(Val(rsProdutos!Codigo), 5)
'      sCodigoBarras = ""
'      sDescricaoProduto = IIf(Trim(rsProdutos!DiscriminacaoProduto) = "", RemoveAcentos(rsProdutos!Descricao), RemoveAcentos(rsProdutos!DiscriminacaoProduto))
'      sCodigoNCM = ""
'      sEXTIPI = ""
'      sGenero = ""
'      sCFOP = RemoveCaracteres(rsNFeItens!CFOP)
'      sUnidComercial = IIf(Not rsUnidades.EOF, rsUnidades!Sigla, "UN")
'      sQuantidadeComercial = Substitui(Format(rsNFeItens!quantidade, "#######0.0000"), ",", ".")
'      sVlUnitarioComercial = Substitui(Format(rsNFeItens!ValorUnitario, "#######0.0000"), ",", ".")
'      sVlTotalBruto = Substitui(Format(rsNFeItens!quantidade * rsNFeItens!ValorUnitario, "#######0.00"), ",", ".")
'      sCodigoBarrasTrib = ""
'      sUnidTrib = IIf(Not rsUnidades.EOF, rsUnidades!Sigla, "UN")
'      sQuantidadeTrib = Substitui(Format(rsNFeItens!quantidade, "#######0.0000"), ",", ".")
'      sVlUnitarioTrib = Substitui(Format(rsNFeItens!ValorUnitario, "#######0.0000"), ",", ".")
'      sVlFrete = ""
'      sVlSeguro = ""
'      sVlDesconto = IIf(rsNFeItens!Desconto > 0, Substitui(Format((((rsNFeItens!quantidade * rsNFeItens!ValorUnitario) * rsNFeItens!Desconto) / 100), "#######0.00"), ",", "."), "")
'
'      Print #1, "I|" & _
'                sCodigoProduto & "|" & _
'                sCodigoBarras & "|" & _
'                sDescricaoProduto & "|" & _
'                sCodigoNCM & "|" & _
'                sEXTIPI & "|" & _
'                sGenero & "|" & _
'                sCFOP & "|" & _
'                sUnidComercial & "|" & _
'                sQuantidadeComercial & "|" & _
'                sVlUnitarioComercial & "|" & _
'                sVlTotalBruto & "|" & _
'                sCodigoBarrasTrib & "|" & _
'                sUnidTrib & "|" & _
'                sQuantidadeTrib & "|" & _
'                sVlUnitarioTrib & "|" & _
'                sVlFrete & "|" & _
'                sVlSeguro & "|" & _
'                sVlDesconto
'
'      ' Tributos Incidentes
'      Print #1, "M"
'      Print #1, "N"
'
'      Dim sCST As String
'      sCST = Mid(rsSituacoesTributarias!Codigo, 2, 2)
'
'      Dim sOrigem As String
'      Dim sModalidadeBC As String
'      Dim sPercRedBC As String
'      Dim sValorBC As String
'      Dim sICMS As String
'      Dim sValorICMS As String
'      Dim dCalculoBC As Double
'
'      Dim Nx As String
'      Select Case sCST
'             Case "00" ' Tributada Integralmente
'                  dCalculoBC = (rsNFeItens!quantidade * rsNFeItens!ValorUnitario) - (((rsNFeItens!quantidade * rsNFeItens!ValorUnitario) * rsNFeItens!Desconto) / 100)
'
'                  sOrigem = "0"
'                  sModalidadeBC = "3"
'                  sValorBC = Substitui(Format(dCalculoBC, "#######0.00"), ",", ".")
'                  sICMS = Substitui(Format(rsNFeItens!ICMS, "###0.00"), ",", ".")
'                  sValorICMS = Substitui(Format(((dCalculoBC * rsNFeItens!ICMS) / 100), "#######0.00"), ",", ".")
'
'                  Print #1, "N02|" & _
'                            sOrigem & "|" & _
'                            sCST & "|" & _
'                            sModalidadeBC & "|" & _
'                            sValorBC & "|" & _
'                            sICMS & "|" & _
'                            sValorICMS & "|"
'
'            Case "10"
'                 Nx = "03"
'            Case "20"
'                 dCalculoBC = Round(((((rsNFeItens!quantidade * rsNFeItens!ValorUnitario) - (((rsNFeItens!quantidade * rsNFeItens!ValorUnitario) * rsNFeItens!Desconto) / 100)) * rsNFeItens!BaseReduzida) / 100), 2)
'
'                  sOrigem = "0"
'                  sModalidadeBC = "3"
'                  sPercRedBC = Substitui(Format(rsNFeItens!BaseReduzida, "#########0.00"), ",", ".")
'                  sValorBC = Substitui(Format(dCalculoBC, "#######0.00"), ",", ".")
'                  sICMS = Substitui(Format(rsNFeItens!ICMS, "###0.00"), ",", ".")
'                  sValorICMS = Substitui(Format(((dCalculoBC * rsNFeItens!ICMS) / 100), "#######0.00"), ",", ".")
'
'                  Print #1, "N04|" & _
'                            sOrigem & "|" & _
'                            sCST & "|" & _
'                            sModalidadeBC & "|" & _
'                            sPercRedBC & "|" & _
'                            sValorBC & "|" & _
'                            sICMS & "|" & _
'                            sValorICMS
'
'             Case "30"
'                  Nx = "05"
'             Case "40"
'                  Nx = "06"
'             Case "51"
'                  Nx = "07"
'             Case "60"
'                  Nx = "08"
'             Case "70"
'                  Nx = "09"
'             Case "90"
'                  Nx = "10"
'      End Select
'
'      ' Totalizar
'      dValorTotalBC = dValorTotalBC + dCalculoBC
'      dValorTotalICMS = dValorTotalICMS + ((dCalculoBC * rsNFeItens!ICMS) / 100)
'      dValorTotalProdutos = dValorTotalProdutos + (rsNFeItens!quantidade * rsNFeItens!ValorUnitario)
'      dValorTotalDesconto = dValorTotalDesconto + (((rsNFeItens!quantidade * rsNFeItens!ValorUnitario) * rsNFeItens!Desconto) / 100)
'      dValorTotalNFe = dValorTotalNFe + (rsNFeItens!quantidade * rsNFeItens!ValorUnitario)
'
'      ' PIS
'      Print #1, "Q"
'      Print #1, "Q04|" & _
'                "08"
'
'      ' Cofins
'      Print #1, "S"
'      Print #1, "S04|" & _
'                "08"
'
'      Contador = Contador + 1
'      rsNFeItens.MoveNext
'   Loop
'   ' Totais
'   Print #1, "W"
'
'   Dim sValorTotalBC As String
'   Dim sValorTotalICMS As String
'   Dim sValorTotalBCST As String
'   Dim sValorTotalICMSST As String
'   Dim sValorTotalProdutos As String
'   Dim sValorTotalFrete As String
'   Dim sValorTotalSeguro As String
'   Dim sValorTotalDesconto As String
'   Dim sValorTotalII As String
'   Dim sValorTotalIPI As String
'   Dim sValorTotalPIS As String
'   Dim sValorTotalCofins As String
'   Dim sValorTotalOutro As String
'   Dim sValorTotalNFe As String
'
'   sValorTotalBC = Substitui(Format(dValorTotalBC, "#########0.00"), ",", ".")
'   sValorTotalICMS = Substitui(Format(dValorTotalICMS, "#########0.00"), ",", ".")
'   sValorTotalBCST = Substitui(Format(0, "#########0.00"), ",", ".")
'   sValorTotalICMSST = Substitui(Format(0, "#########0.00"), ",", ".")
'   sValorTotalProdutos = Substitui(Format(dValorTotalProdutos, "#########0.00"), ",", ".")
'   sValorTotalFrete = Substitui(Format(0, "#########0.00"), ",", ".")
'   sValorTotalSeguro = Substitui(Format(0, "#########0.00"), ",", ".")
'   sValorTotalDesconto = Substitui(Format(dValorTotalDesconto, "#########0.00"), ",", ".")
'   sValorTotalII = Substitui(Format(0, "#########0.00"), ",", ".")
'   sValorTotalIPI = Substitui(Format(0, "#########0.00"), ",", ".")
'   sValorTotalPIS = Substitui(Format(0, "#########0.00"), ",", ".")
'   sValorTotalCofins = Substitui(Format(0, "#########0.00"), ",", ".")
'   sValorTotalOutro = Substitui(Format(0, "#########0.00"), ",", ".")
'   sValorTotalNFe = Substitui(Format(dValorTotalNFe, "#########0.00"), ",", ".")
'
'   Print #1, "W02|" & _
'             sValorTotalBC & "|" & _
'             sValorTotalICMS & "|" & _
'             sValorTotalBCST & "|" & _
'             sValorTotalICMSST & "|" & _
'             sValorTotalProdutos & "|" & _
'             sValorTotalFrete & "|" & _
'             sValorTotalSeguro & "|" & _
'             sValorTotalDesconto & "|" & _
'             sValorTotalII & "|" & _
'             sValorTotalIPI & "|" & _
'             sValorTotalPIS & "|" & _
'             sValorTotalCofins & "|" & _
'             sValorTotalOutro & "|" & _
'             sValorTotalNFe
'
'   ' Frete
'   Print #1, "X|" & _
'             "0"
'
'   ' Informaçoes Adicionais
'   Dim sInfFISCO As String
'   Dim sInfEmpresa As String
'
'   sInfFISCO = rsNFe!InformacoesCorpo
'   sInfEmpresa = rsNFe!DadosAdicionais
'   Print #1, "Z|" & _
'             sInfFISCO & "|" & _
'             sInfEmpresa
'
'   ' Define como Gerada
'   cnSistema.Execute "Update NFe set " & _
'            "GeradaNFe = True " & _
'            "Where idNFe = " & rsNFe!idNFe
'
'   Close #1
'
'End Function

Private Sub cboUFCarregamento_LostFocus()

   '  Municipios
   Set rsTemp = cnSistema.Execute("Select * from Municipios WHERE UF = '" & cboUFCarregamento.Text & "' ORDER BY Nome")
   cboMunicipio.Clear
   Do While Not rsTemp.EOF
      cboMunicipio.AddItem rsTemp!Nome
      cboMunicipio.ItemData(cboMunicipio.NewIndex) = rsTemp!idMunicipio
      rsTemp.MoveNext
   Loop

End Sub

Private Sub FLocalCarregamento(ByRef iIdMDFe As Integer)
Dim iIdLocalCarregamento As Integer

'''''   Set rsMDFeLocalCarregamento = cnSistema.Execute("SELECT LC.idMDFeLocalCarregamento, LC.idMDFe, M.Nome AS Municipio, UF.Sigla AS UF " & _
'''''                                                   "FROM MDFeLocalCarregamento AS LC, Municipios AS M, UFs AS UF " & _
'''''                                                   "WHERE LC.idMunicipio = M.idMunicipio AND LC.idUF = UF.idUF AND LC.idMDFe = " & iIdMDFe)
'''''
'''''   lvwLocalCarregamento.ListItems.Clear
'''''   Do While Not rsMDFeLocalCarregamento.EOF
'''''         Set ItemList = lvwLocalCarregamento.ListItems.Add(, "R" & CStr(rsMDFeLocalCarregamento!idMDFeLocalCarregamento), rsMDFeLocalCarregamento!UF)
'''''         ItemList.SubItems(1) = rsMDFeLocalCarregamento!Municipio
'''''
'''''      rsMDFeLocalCarregamento.MoveNext
'''''   Loop


   Set rsDados = cnSistema.Execute("SELECT LC.idMDFeLocalCarregamento, LC.idMDFe, LC.idMunicipio, M.Nome AS Municipio, LC.idUF, UF.Sigla AS UF " & _
                                  "FROM MDFeLocalCarregamento AS LC, Municipios AS M, UFs AS UF " & _
                                  "WHERE LC.idMunicipio = M.idMunicipio AND LC.idUF = UF.idUF AND LC.idMDFe = " & iIdMDFe)

   lvwLocalCarregamento.ListItems.Clear
   Do While Not rsDados.EOF
'      Set ItemList = lvwLocalCarregamento.ListItems.Add(, "R" & CStr(rsDados!idMDFeLocalCarregamento), rsDados!UF)

      iIdLocalCarregamento = lvwLocalCarregamento.ListItems.Count + 1

      Set ItemList = lvwLocalCarregamento.ListItems.Add(, "R" & CStr(lvwLocalCarregamento.ListItems.Count + 1), rsDados!UF)
      ItemList.SubItems(1) = rsDados!Municipio

      rsMDFeLocalCarregamento.AddNew
      rsMDFeLocalCarregamento!idMDFeLocalCarregamento = iIdLocalCarregamento
      rsMDFeLocalCarregamento!idMDFe = iIdMDFe
      rsMDFeLocalCarregamento!idMunicipio = rsDados!idMunicipio
      rsMDFeLocalCarregamento!Municipio = rsDados!Municipio
      rsMDFeLocalCarregamento!idUF = rsDados!idUF
      rsMDFeLocalCarregamento!UF = rsDados!UF
      rsMDFeLocalCarregamento.Update
      
      rsDados.MoveNext
   Loop

End Sub

Private Sub FUF(ByRef iIdUF As Integer)

   For Contador = 0 To (cboUF.ListCount - 1)
      If cboUF.ItemData(Contador) = iIdUF Then
         cboUF.ListIndex = Contador
         rsMDFes!idUF = cboUF.ItemData(Contador)
         Exit For
      End If
   Next

End Sub

Private Sub FTipoEmitente(ByRef iIdTipoEmitente As Integer)

   For Contador = 0 To (cboTipoEmitente.ListCount - 1)
      If cboTipoEmitente.ItemData(Contador) = iIdTipoEmitente Then
         cboTipoEmitente.ListIndex = Contador
         rsMDFes!idTipoEmitente = cboTipoEmitente.ItemData(Contador)
         Exit For
      End If
   Next

End Sub

Private Sub FTipoTransportador(ByRef iIdTipoTransportador As Integer)

   For Contador = 0 To (cboTipoTransportador.ListCount - 1)
      If cboTipoTransportador.ItemData(Contador) = iIdTipoTransportador Then
         cboTipoTransportador.ListIndex = Contador
         rsMDFes!idTipoTransportador = cboTipoTransportador.ItemData(Contador)
         Exit For
      End If
   Next

End Sub

Private Sub FFormaEmissao(ByRef iIdFormaEmissao As Integer)

   For Contador = 0 To (cboFormaEmissao.ListCount - 1)
      If cboFormaEmissao.ItemData(Contador) = iIdFormaEmissao Then
         cboFormaEmissao.ListIndex = Contador
         rsMDFes!idFormaEmissao = cboFormaEmissao.ItemData(Contador)
         Exit For
      End If
   Next

End Sub

Private Sub FModalidade(ByRef iIdModalidade As Integer)

   For Contador = 0 To (cboModalidade.ListCount - 1)
      If cboModalidade.ItemData(Contador) = iIdModalidade Then
         cboModalidade.ListIndex = Contador
         rsMDFes!idModalidade = cboModalidade.ItemData(Contador)
         Exit For
      End If
   Next

End Sub

Private Sub FTipoCarroceria(ByRef iIdTipoCarroceria As Integer)

   For Contador = 0 To (cboTipoCarroceria.ListCount - 1)
      If cboTipoCarroceria.ItemData(Contador) = iIdTipoCarroceria Then
         cboTipoCarroceria.ListIndex = Contador
         rsMDFes!idTipoCarroceria = cboTipoCarroceria.ItemData(Contador)
         Exit For
      End If
   Next

End Sub

Private Sub FUFDescarregamento(ByRef iIdUFDescarregamento As Integer)

   For Contador = 0 To (cboUFDescarregamento.ListCount - 1)
      If cboUFDescarregamento.ItemData(Contador) = iIdUFDescarregamento Then
         cboUFDescarregamento.ListIndex = Contador
         rsMDFes!idUFDescarregamento = cboUFDescarregamento.ItemData(Contador)
         Exit For
      End If
   Next

End Sub

Private Sub FUFVeiculo(ByRef iIdUFVeiculo As Integer)

   For Contador = 0 To (cboUFVeiculo.ListCount - 1)
      If cboUFVeiculo.ItemData(Contador) = iIdUFVeiculo Then
         cboUFVeiculo.ListIndex = Contador
         rsMDFes!idUFVeiculo = cboUFVeiculo.ItemData(Contador)
         Exit For
      End If
   Next

End Sub

Private Sub FTipoRodado(ByRef iIdTipoRodado As Integer)

   For Contador = 0 To (cboTipoRodado.ListCount - 1)
      If cboTipoRodado.ItemData(Contador) = iIdTipoRodado Then
         cboTipoRodado.ListIndex = Contador
         rsMDFes!idTipoRodado = cboTipoRodado.ItemData(Contador)
         Exit For
      End If
   Next

End Sub

Private Sub FPercurso(ByRef iIdMDFe As Integer)

   Set rsDados = cnSistema.Execute("SELECT P.idMDFePercurso, P.idMDFe, P.idUF, UF.Sigla AS UF " & _
                                  "FROM MDFePercurso AS P, UFs AS UF " & _
                                  "WHERE P.idUF = UF.idUF AND P.idMDFe = " & iIdMDFe)

   lvwPercurso.ListItems.Clear
   Do While Not rsDados.EOF
      Set ItemList = lvwPercurso.ListItems.Add(, "R" & CStr(lvwPercurso.ListItems.Count + 1), rsDados!UF)

      rsMDFePercurso.AddNew
      rsMDFePercurso!idMDFe = iIdMDFe
      rsMDFePercurso!idUF = rsDados!idUF
      rsMDFePercurso.Update

      rsDados.MoveNext
   Loop

End Sub

Private Sub FCondutores(ByRef iIdMDFe As Integer)

   Set rsDados = cnSistema.Execute("SELECT C.idMDFeCondutor, C.idMDFe, C.CPF AS CPF, C.Nome AS Nome " & _
                                  "FROM MDFeCondutores AS C " & _
                                  "WHERE C.idMDFe = " & iIdMDFe)

   lvwCondutores.ListItems.Clear
   Do While Not rsDados.EOF
      Set ItemList = lvwCondutores.ListItems.Add(, "R" & CStr(lvwCondutores.ListItems.Count + 1), rsDados!CPF)
      ItemList.SubItems(1) = rsDados!Nome

      rsMDFeCondutores.AddNew
      rsMDFeCondutores!idMDFe = iIdMDFe
      rsMDFeCondutores!CPF = rsDados!CPF
      rsMDFeCondutores!Nome = rsDados!Nome
      rsMDFePercurso.Update

      rsDados.MoveNext
   Loop

End Sub

Private Sub FNFes(ByRef iIdMDFe As Integer)

   Set rsDados = cnSistema.Execute("SELECT NF.idMDFeNFe, NF.idMDFe, NF.Numero AS Numero, NF.ChaveNFe AS ChaveNFe " & _
                                  "FROM MDFeNFes AS NF " & _
                                  "WHERE NF.idMDFe = " & iIdMDFe)

   lvwNFes.ListItems.Clear
   Do While Not rsDados.EOF
      Set ItemList = lvwNFes.ListItems.Add(, "R" & CStr(lvwNFes.ListItems.Count + 1), rsDados!Numero)
      ItemList.SubItems(1) = FFormataChaveNF(IIf(IsNull(rsDados!ChaveNFe), "", rsDados!ChaveNFe))
'      ItemList.SubItems(2) = Format(mskValorTotalProdutos.Text, "###,##0.00")

      rsMDFeNFes.AddNew
      rsMDFeNFes!idMDFe = iIdMDFe
      rsMDFeNFes!Numero = rsDados!Numero
      rsMDFeNFes!ChaveNFe = rsDados!ChaveNFe
      rsMDFeNFes.Update

      rsDados.MoveNext
   Loop

End Sub

Public Function CriarEstruturaMDFe()

   Set rsMDFes = Nothing
   Set rsMDFeLocalCarregamento = Nothing
   Set rsMDFePercurso = Nothing
   Set rsMDFeCondutores = Nothing
   Set rsMDFeNFes = Nothing

   ' Criar Estrutura rsMDFes
   '------------------------------------------------------------------------------------------
    With rsMDFes
      .Fields.Append "idMDFe", adInteger
      .Fields.Append "idUF", adInteger
      .Fields.Append "idUFDescarregamento", adInteger
      .Fields.Append "idTipoEmitente", adInteger
      .Fields.Append "idTipoTransportador", adInteger
      .Fields.Append "idFormaEmissao", adInteger
      .Fields.Append "idModalidade", adInteger
      .Fields.Append "idTipoCarroceria", adInteger
      .Fields.Append "idUFVeiculo", adInteger
      .Fields.Append "idTipoRodado", adInteger

      .Fields.Append "ChaveMDFe", adVarChar, 50
      .Fields.Append "Protocolo", adVarChar, 20
      .Fields.Append "Numero", adInteger
      .Fields.Append "DataEmissao", adDate
      .Fields.Append "DataViagem", adDate
      .Fields.Append "PlacaVeiculo", adVarChar, 10
      .Fields.Append "Tara", adVarChar, 20
      .Fields.Append "CapacidadeKG", adVarChar, 20
      .Fields.Append "CapacidadeM3", adVarChar, 20
      .Fields.Append "Renavam", adVarChar, 20
      .Fields.Append "DadosAdicionais", adVarChar, 2000

      .Fields.Refresh
      .Open
    End With

   ' Criar Estrutura rsMDFeLocalCarregamento
   '------------------------------------------------------------------------------------------
    With rsMDFeLocalCarregamento
      .Fields.Append "idMDFeLocalCarregamento", adVarChar, 20
      .Fields.Append "idMDFe", adInteger
      .Fields.Append "idUF", adInteger
      .Fields.Append "UF", adVarChar, 2
      .Fields.Append "idMunicipio", adInteger
      .Fields.Append "Municipio", adVarChar, 50

      .Fields.Refresh
      .Open
    End With

   ' Criar Estrutura rsMDFePercurso
   '------------------------------------------------------------------------------------------
    With rsMDFePercurso
      .Fields.Append "idMDFePercurso", adVarChar, 20
      .Fields.Append "idMDFe", adInteger
      .Fields.Append "idUF", adInteger

      .Fields.Refresh
      .Open
    End With

   ' Criar Estrutura rsMDFeCondutores
   '------------------------------------------------------------------------------------------
    With rsMDFeCondutores
      .Fields.Append "idMDFeCondutor", adVarChar, 20
      .Fields.Append "idMDFe", adInteger
      .Fields.Append "CPF", adVarChar, 20
      .Fields.Append "Nome", adVarChar, 50

      .Fields.Refresh
      .Open
    End With

   ' Criar Estrutura rsMDFeNFes
   '------------------------------------------------------------------------------------------
    With rsMDFeNFes
      .Fields.Append "idMDFeNFe", adVarChar, 20
      .Fields.Append "idMDFe", adInteger
      .Fields.Append "Numero", adInteger
      .Fields.Append "ChaveNFe", adVarChar, 50

      .Fields.Refresh
      .Open
    End With

End Function

Private Sub txtCapacidadeKG_LostFocus()
   rsMDFes!CapacidadeKG = txtCapacidadeKG.Text
End Sub

Private Sub txtCapacidadeM3_LostFocus()
   rsMDFes!CapacidadeM3 = txtCapacidadeM3.Text
End Sub

Private Sub txtDadosAdicionais_LostFocus()
   rsMDFes!DadosAdicionais = txtDadosAdicionais.Text
End Sub

Private Sub txtNumeroNota_LostFocus()
   
   Set rsNF = cnSistema.Execute("SELECT * FROM " & I_TabelasNF & " WHERE Numero = " & txtNumeroNota.Text)
   If I_ModeloNF = "55" Then
      If Not IsNull(rsNF!ChaveNFe) Or IsEmpty(rsNF!ChaveNFe) Then
         txtChaveNFe.Text = rsNF!ChaveNFe                                                           ' Chave da NF
      End If
   ElseIf I_ModeloNF = "65" Then
      If Not IsNull(rsNF!ChaveNFCe) Or IsEmpty(rsNF!ChaveNFCe) Then
         txtChaveNFe.Text = rsNF!ChaveNFCe                                                          ' Chave da NF
      End If
   End If
   
End Sub

Private Sub txtRenavam_LostFocus()
   rsMDFes!Renavam = txtRenavam.Text
End Sub

Private Sub txtTara_LostFocus()
   rsMDFes!Tara = txtTara.Text
End Sub

Private Sub cboFormaEmissao_LostFocus()
   rsMDFes!idFormaEmissao = FVerificaCombo(cboFormaEmissao.ItemData(cboFormaEmissao.ListIndex))
End Sub

Private Sub cboModalidade_LostFocus()
   rsMDFes!idModalidade = FVerificaCombo(cboModalidade.ItemData(cboModalidade.ListIndex))
End Sub

Private Sub cboTipoCarroceria_LostFocus()
   rsMDFes!idTipoCarroceria = FVerificaCombo(cboTipoCarroceria.ItemData(IIf(cboTipoCarroceria.ListIndex <> -1, cboTipoCarroceria.ListIndex, 0)))
End Sub

Private Sub cboTipoEmitente_LostFocus()
   rsMDFes!idTipoEmitente = FVerificaCombo(cboTipoEmitente.ItemData(cboTipoEmitente.ListIndex))
End Sub

Private Sub cboTipoRodado_LostFocus()
   rsMDFes!idTipoRodado = FVerificaCombo(cboTipoRodado.ItemData(cboTipoRodado.ListIndex))
End Sub

Private Sub cboTipoTransportador_LostFocus()
   rsMDFes!idTipoTransportador = FVerificaCombo(cboTipoTransportador.ItemData(cboTipoTransportador.ListIndex))
End Sub

Private Sub cboUF_LostFocus()
   rsMDFes!idUF = FVerificaCombo(cboUF.ItemData(cboUF.ListIndex))
End Sub

Private Sub cboUFDescarregamento_LostFocus()
   rsMDFes!idUFDescarregamento = FVerificaCombo(cboUFDescarregamento.ItemData(cboUFDescarregamento.ListIndex))
End Sub

Private Sub cboUFVeiculo_LostFocus()
   rsMDFes!idUFVeiculo = FVerificaCombo(cboUFVeiculo.ItemData(cboUFVeiculo.ListIndex))
End Sub

Private Function FVerificaCombo(ByRef iCampo As Integer) As Integer

   If iCampo = -1 Then
      FVerificaCombo = 0
   Else
      FVerificaCombo = iCampo
   End If

End Function

Private Sub cmdExcluirLocalCarregamento_Click()
Dim iId As Integer
'   Beep
   If MsgBox("Deseja excluir este item", vbYesNo + vbQuestion, "Excluir") = vbYes Then
   
'      cnSistema.Execute "Delete from NFeBoletos Where Numero = '" & Mid(lvwBoletos.SelectedItem.Key, 2, Len(lvwBoletos.SelectedItem.Key)) & "'"
'      lvwBoletos.ListItems.Remove (lvwBoletos.SelectedItem.Index)
   
'''      lvwProdutos.ListItems.Remove (lvwProdutos.SelectedItem.Index)
   
''''      iId = Mid(lvwLocalCarregamento.ListItems(Contador).Key, 2, Len(lvwLocalCarregamento.ListItems(Contador).Key))

''Parei aqui pegando a chave


'      iId = lvwLocalCarregamento.SelectedItem.Index
      iId = Mid(lvwLocalCarregamento.SelectedItem.Key, 2, Len(lvwLocalCarregamento.SelectedItem.Key))
'      rsMDFeLocalCarregamento.Requery
      rsMDFeLocalCarregamento.MoveFirst
      rsMDFeLocalCarregamento.Find "idMDFeLocalCarregamento = " & iId
      rsMDFeLocalCarregamento.Delete adAffectCurrent
      
      ' Recarregar View
      lvwLocalCarregamento.ListItems.Clear
      rsMDFeLocalCarregamento.MoveFirst
      Do While Not rsMDFeLocalCarregamento.EOF
'         Set ItemList = lvwLocalCarregamento.ListItems.Add(, "R" & CStr(lvwLocalCarregamento.ListItems.Count + 1), rsMDFeLocalCarregamento!UF)
         Set ItemList = lvwLocalCarregamento.ListItems.Add(, "R" & CStr(rsMDFeLocalCarregamento!idMDFeLocalCarregamento), rsMDFeLocalCarregamento!UF)
         ItemList.SubItems(1) = rsMDFeLocalCarregamento!Municipio
   
         rsMDFeLocalCarregamento.MoveNext
      Loop
      
   End If
End Sub


Private Sub cmdInserirLocalCarregamento_Click()
Dim iIdLocalCarregamento As Integer

   If cboUFCarregamento.ListIndex = -1 Or cboMunicipio.ListIndex = -1 Then
      MsgBox "Município e UF são obrigatórios", vbExclamation, "Sistema"
      Exit Sub
   End If
      
'   If MsgBox("Confirma Incluir o registro atual", vbYesNo + vbQuestion, "Inclusão") = vbYes Then

      If rsMDFeLocalCarregamento.EOF Then
         iIdLocalCarregamento = lvwLocalCarregamento.ListItems.Count + 1
      Else
         rsMDFeLocalCarregamento.MoveLast
         iIdLocalCarregamento = rsMDFeLocalCarregamento!idMDFeLocalCarregamento + 1
      End If
   
      Set ItemList = lvwLocalCarregamento.ListItems.Add(, "R" & CStr(lvwLocalCarregamento.ListItems.Count + 1), cboUFCarregamento.Text)
      ItemList.SubItems(1) = cboMunicipio.Text

      rsMDFeLocalCarregamento.AddNew
      rsMDFeLocalCarregamento!idMDFeLocalCarregamento = iIdLocalCarregamento
      rsMDFeLocalCarregamento!idMDFe = rsMDFe!idMDFe
      rsMDFeLocalCarregamento!idUF = cboUFCarregamento.ItemData(cboUFCarregamento.ListIndex)
      rsMDFeLocalCarregamento!UF = cboUFCarregamento.Text
      rsMDFeLocalCarregamento!idMunicipio = cboMunicipio.ItemData(cboMunicipio.ListIndex)
      rsMDFeLocalCarregamento!Municipio = cboMunicipio.Text
      rsMDFeLocalCarregamento.Update
   
   
      cboUFCarregamento.ListIndex = -1
      cboMunicipio.ListIndex = -1
'   End If
   
'   If Verifica_Campos_Boletos() Then
'      Set ItemList = lvwBoletos.ListItems.Add(, "R" & mskNumeroBoleto.Text, mskNumeroBoleto.Text)
'      ItemList.SubItems(1) = mskVencimentoBoleto.Text
'      ItemList.SubItems(2) = Format(mskValorBoleto.Text, "##,##0.00")
'
'      cnSistema.Execute "Insert Into NFeBoletos (idNFe,Numero,Vencimento,Valor) " & _
'                        "Values (" & rsNFe!idNFe & ",'" & mskNumeroBoleto.Text & "','" & mskVencimentoBoleto.Text & _
'                        "'," & Substitui(mskValorBoleto.Text, ",", ".") & ")"
'
'      mskNumeroBoleto.Text = Empty
'      mskVencimentoBoleto.Text = "  /  /    "
'      mskValorBoleto.Text = Empty
'      mskNumeroBoleto.SetFocus
'   End If

End Sub

Private Sub cmdInserirPercurso_Click()
   If cboUFPercurso.ListIndex = -1 Then
      MsgBox "UF é obrigatório", vbExclamation, "Sistema"
      Exit Sub
   End If
      
'   If MsgBox("Confirma Incluir o registro atual", vbYesNo + vbQuestion, "Inclusão") = vbYes Then
      rsMDFePercurso.AddNew
      rsMDFePercurso!idMDFe = rsMDFe!idMDFe
      rsMDFePercurso!idUF = cboUFPercurso.ItemData(cboUFPercurso.ListIndex)
      rsMDFePercurso.Update
   
      Set ItemList = lvwPercurso.ListItems.Add(, "R" & CStr(lvwPercurso.ListItems.Count + 1), cboUFPercurso.Text)
   
      cboUFPercurso.ListIndex = -1
'   End If
End Sub


Private Sub cmdIncluirCondutor_Click()

   If IsEmpty(mskCPFCondutor.Text) Or IsEmpty(txtNomeCondutor.Text) Then
      MsgBox "CPF e Nome são obrigatórios", vbExclamation, "Sistema"
      Exit Sub
   End If
      
'   If MsgBox("Confirma Incluir o registro atual", vbYesNo + vbQuestion, "Inclusão") = vbYes Then
      rsMDFeCondutores.AddNew
      rsMDFeCondutores!idMDFe = rsMDFe!idMDFe
      rsMDFeCondutores!CPF = mskCPFCondutor.Text
      rsMDFeCondutores!Nome = txtNomeCondutor.Text
      rsMDFeCondutores.Update
   
      Set ItemList = lvwCondutores.ListItems.Add(, "R" & CStr(lvwCondutores.ListItems.Count + 1), mskCPFCondutor.Text)
      ItemList.SubItems(1) = txtNomeCondutor.Text
   
      mskCPFCondutor.Text = Empty
      txtNomeCondutor.Text = Empty
'   End If

End Sub

Private Sub cmdIncluirNFe_Click()
   If IsEmpty(txtNumeroNota.Text) Or IsEmpty(txtChaveNFe.Text) Then
      MsgBox "Número e Chave da Nota são obrigatórios", vbExclamation, "Sistema"
      Exit Sub
   End If
      
'   If MsgBox("Confirma Incluir o registro atual", vbYesNo + vbQuestion, "Inclusão") = vbYes Then
      rsMDFeNFes.AddNew
      rsMDFeNFes!idMDFe = rsMDFe!idMDFe
      rsMDFeNFes!Numero = txtNumeroNota.Text
      rsMDFeNFes!ChaveNFe = txtChaveNFe.Text
      rsMDFeNFes.Update
   
      Set ItemList = lvwNFes.ListItems.Add(, "R" & CStr(lvwNFes.ListItems.Count + 1), txtNumeroNota.Text)
      ItemList.SubItems(1) = FFormataChaveNF(IIf(IsNull(txtChaveNFe.Text), "", txtChaveNFe.Text))
   
      txtNumeroNota.Text = Empty
      txtChaveNFe.Text = Empty
'   End If
End Sub


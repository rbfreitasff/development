VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMDFe 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manifesto Eletr�nico de Carga - MDF-e"
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
      TabCaption(1)   =   "Rodovi�rio"
      TabPicture(1)   =   "frmMDFe.frx":1DBC
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraTransportadorVolumes"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "fraCondutores"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Informa��es dos Documentos"
      TabPicture(2)   =   "frmMDFe.frx":1DD8
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraNFes"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "fraDadosAdicionais"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).ControlCount=   2
      Begin VB.Frame fraDadosAdicionais 
         Caption         =   "Dados Adicionais/Informa��es Complementares"
         Height          =   1845
         Left            =   -74940
         TabIndex        =   75
         Top             =   4800
         Width           =   9315
         Begin VB.TextBox txtDadosAdicionais 
            Height          =   1515
            Left            =   75
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   76
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
            TabIndex        =   80
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
            Width           =   1455
         End
         Begin VB.CommandButton cmdIncluirNFe 
            BackColor       =   &H00E0E0E0&
            Height          =   315
            Left            =   8535
            Picture         =   "frmMDFe.frx":2A04
            Style           =   1  'Graphical
            TabIndex        =   68
            Top             =   405
            Width           =   360
         End
         Begin VB.TextBox txtChaveNFe 
            Height          =   315
            Left            =   1590
            MaxLength       =   44
            TabIndex        =   67
            Top             =   405
            Width           =   6900
         End
         Begin VB.CommandButton cmdExcluirNFe 
            BackColor       =   &H80000004&
            Height          =   315
            Left            =   8925
            Picture         =   "frmMDFe.frx":2B4E
            Style           =   1  'Graphical
            TabIndex        =   69
            Top             =   405
            Width           =   360
         End
         Begin MSComctlLib.ListView lvwNFes 
            Height          =   2760
            Left            =   60
            TabIndex        =   70
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
            TabIndex        =   72
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
            TabIndex        =   74
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
            TabIndex        =   66
            Top             =   180
            Width           =   1035
         End
         Begin VB.Label lblValorTotalProdutos 
            AutoSize        =   -1  'True
            Caption         =   "Total Produtos"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   6720
            TabIndex        =   71
            Top             =   3630
            Width           =   1035
         End
         Begin VB.Label lblTotalNota 
            AutoSize        =   -1  'True
            Caption         =   "Total Nota"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   7005
            TabIndex        =   73
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
            Picture         =   "frmMDFe.frx":3004
            Style           =   1  'Graphical
            TabIndex        =   60
            Top             =   405
            Width           =   360
         End
         Begin VB.CommandButton cmdExcluirCondutor 
            BackColor       =   &H80000004&
            Height          =   315
            Left            =   8880
            Picture         =   "frmMDFe.frx":34BA
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
         Caption         =   "Ve�culo de Tra��o"
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
            ItemData        =   "frmMDFe.frx":3970
            Left            =   4800
            List            =   "frmMDFe.frx":3972
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
            ItemData        =   "frmMDFe.frx":3974
            Left            =   1380
            List            =   "frmMDFe.frx":3976
            Style           =   2  'Dropdown List
            TabIndex        =   48
            Top             =   570
            Width           =   2895
         End
         Begin VB.ComboBox cboTipoCarroceria 
            Height          =   315
            ItemData        =   "frmMDFe.frx":3978
            Left            =   1380
            List            =   "frmMDFe.frx":397A
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
         TabIndex        =   77
         Top             =   960
         Width           =   9375
         Begin VB.ComboBox cboFormaEmissao 
            Height          =   315
            ItemData        =   "frmMDFe.frx":397C
            Left            =   6300
            List            =   "frmMDFe.frx":397E
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   840
            Width           =   2445
         End
         Begin VB.ComboBox cboModalidade 
            Height          =   315
            ItemData        =   "frmMDFe.frx":3980
            Left            =   1560
            List            =   "frmMDFe.frx":3982
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   1170
            Width           =   3165
         End
         Begin VB.ComboBox cboUF 
            Height          =   315
            ItemData        =   "frmMDFe.frx":3984
            Left            =   8040
            List            =   "frmMDFe.frx":3986
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   165
            Width           =   705
         End
         Begin VB.ComboBox cboTipoTransportador 
            Height          =   315
            ItemData        =   "frmMDFe.frx":3988
            Left            =   1560
            List            =   "frmMDFe.frx":398A
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   825
            Width           =   3165
         End
         Begin VB.ComboBox cboTipoEmitente 
            Height          =   315
            ItemData        =   "frmMDFe.frx":398C
            Left            =   1560
            List            =   "frmMDFe.frx":398E
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
            Caption         =   "Forma de Emiss�o"
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
            Caption         =   "Emiss�o"
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
         TabIndex        =   78
         Top             =   -240
         Width           =   1110
      End
      Begin VB.Label lblProtocolo 
         AutoSize        =   -1  'True
         Caption         =   "Protocolo de Autoriza��o"
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
      TabIndex        =   79
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
Dim rsNFe As New ADODB.Recordset
Dim rsTemp As New ADODB.Recordset
Dim rsDados As New ADODB.Recordset

Dim NotaAtual As String
Dim RegistroAtual As Double 'Para posicionar o ponteiro depois de pesquisas

Dim Contador As Integer

Private Sub Form_Load()
Dim oPreencherRs As New PreencherRS

   ' Adicionar MDFes RecorSets Tempor�rios
   
   I_TituloForm = Me.Caption
   On Error GoTo Erro
   SSTab1.Tab = 0 ' Posiciona no primeiro tab
   Status = 0
   RegistroAtual = 0
   Centraliza frmMDFe
'   rsMDFe.Open "Select * from MDFe Order By Numero", cnSistema, adOpenForwardOnly, adLockOptimistic, 1
   Set rsMDFe = cnSistema.Execute("Select * from MDFe Order By Numero")

   lvwLocalCarregamento.ColumnHeaders.Add , , "UF", 500
   lvwLocalCarregamento.ColumnHeaders.Add , , "Munic�pio", 4000

   lvwPercurso.ColumnHeaders.Add , , "UF", 3600

   lvwCondutores.ColumnHeaders.Add , , "CPF", 1400
   lvwCondutores.ColumnHeaders.Add , , "Nome", 7500

   lvwNFes.ColumnHeaders.Add , , "N�mero", 1000
   lvwNFes.ColumnHeaders.Add , , "Chave", 7900

   Carrega_Combos
'   If Registros(cnSistema, "MDFe") = 0 Then
   If Registros2("MDFe") = 0 Then
      Botoes 3, frmMDFe
   Else
      Botoes 1, frmMDFe
      rsMDFe.MoveLast
      Prencher_Campos
   End If

   If I_Acesso = 3 Then ' Controle N�veis de Acesso
      Toolbar.Buttons(2).Visible = False
      Toolbar.Buttons(3).Visible = False
   End If

   Campos False
'   If Registros(cnSistema, "MDFe") = 0 Then
   If Registros2("MDFe") = 0 Then
      Toolbar.Buttons(15).Enabled = False
   End If
''   MDISistema.StatusBar.Panels(1).text = "Cadastro Notas de Sa�da Eletr�nicas"

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

   ' Tipo de Emitente
   cboTipoEmitente.AddItem "1. Prestador de servi�o de transporte"
   cboTipoEmitente.ItemData(cboTipoEmitente.NewIndex) = 1

   cboTipoEmitente.AddItem "2. N�o prestador de servi�o de transporte"
   cboTipoEmitente.ItemData(cboTipoEmitente.NewIndex) = 2

   cboTipoEmitente.AddItem "3. Prestador de servi�o de transporte que emitir� CT-e Globalizado"
   cboTipoEmitente.ItemData(cboTipoEmitente.NewIndex) = 3

   ' Tipo de Transportador
   cboTipoTransportador.AddItem "1. ETC"
   cboTipoTransportador.ItemData(cboTipoTransportador.NewIndex) = 1

   cboTipoTransportador.AddItem "2. TAC"
   cboTipoTransportador.ItemData(cboTipoTransportador.NewIndex) = 2

   cboTipoTransportador.AddItem "3. CTC"
   cboTipoTransportador.ItemData(cboTipoTransportador.NewIndex) = 3

   ' Forma de Emiss�o
   cboFormaEmissao.AddItem "1. Normal"
   cboFormaEmissao.ItemData(cboFormaEmissao.NewIndex) = 1

   cboFormaEmissao.AddItem "2. Contig�ncia"
   cboFormaEmissao.ItemData(cboFormaEmissao.NewIndex) = 2

   ' Modalidade
   cboModalidade.AddItem "1. Rodovi�rio"
   cboModalidade.ItemData(cboModalidade.NewIndex) = 1

   cboModalidade.AddItem "2. A�reo"
   cboModalidade.ItemData(cboModalidade.NewIndex) = 2

   cboModalidade.AddItem "3. Aquavi�rio"
   cboModalidade.ItemData(cboModalidade.NewIndex) = 3

   cboModalidade.AddItem "4. Ferrovi�rio"
   cboModalidade.ItemData(cboModalidade.NewIndex) = 4

   ' Tipo Rodado
   cboTipoRodado.AddItem "01. Truck"
   cboTipoRodado.ItemData(cboTipoRodado.NewIndex) = 1

   cboTipoRodado.AddItem "02. Toco"
   cboTipoRodado.ItemData(cboTipoRodado.NewIndex) = 2

   cboTipoRodado.AddItem "03. Cavalo Mec�nico"
   cboTipoRodado.ItemData(cboTipoRodado.NewIndex) = 3

   cboTipoRodado.AddItem "04. VAN"
   cboTipoRodado.ItemData(cboTipoRodado.NewIndex) = 4

   cboTipoRodado.AddItem "05. Utilit�rio"
   cboTipoRodado.ItemData(cboTipoRodado.NewIndex) = 5

   cboTipoRodado.AddItem "06. Outros"
   cboTipoRodado.ItemData(cboTipoRodado.NewIndex) = 6

   ' Tipo de Carroceria
   cboTipoCarroceria.AddItem "00. N�o aplic�vel"
   cboTipoCarroceria.ItemData(cboTipoCarroceria.NewIndex) = 0

   cboTipoCarroceria.AddItem "01. Aberta"
   cboTipoCarroceria.ItemData(cboTipoCarroceria.NewIndex) = 1

   cboTipoCarroceria.AddItem "02. Fechada/Ba�"
   cboTipoCarroceria.ItemData(cboTipoCarroceria.NewIndex) = 2

   cboTipoCarroceria.AddItem "03. Granelera"
   cboTipoCarroceria.ItemData(cboTipoCarroceria.NewIndex) = 3

   cboTipoCarroceria.AddItem "04. Porta Container"
   cboTipoCarroceria.ItemData(cboTipoCarroceria.NewIndex) = 4

   cboTipoCarroceria.AddItem "05. Sider"
   cboTipoCarroceria.ItemData(cboTipoCarroceria.NewIndex) = 5
   
End Sub

Private Sub mskCPFCondutor_LostFocus()
Dim Verifica_CPF As String
Dim intCliente, Contador As Integer

   If mskCPFCondutor.Text <> Empty Then
      Verifica_CPF = CNPJ_CPF(mskCPFCondutor.Text)
      If Verifica_CPF <> "ERRO" Then
         mskCPFCondutor.Text = CNPJ_CPF(mskCPFCondutor.Text)
         
         ' Busca condutor cadastrado
         Set rsTemp = cnSistema.Execute("SELECT TOP 1 Nome FROM MDFeCondutores WHERE CPF = '" & mskCPFCondutor.Text & "'")
         If Not rsTemp.EOF Then
            txtNomeCondutor.Text = rsTemp!Nome
         End If
         Set rsTemp = Nothing
      Else
         mskCPFCondutor.SelStart = 0
         mskCPFCondutor.SelLength = Len(mskCPFCondutor.Text)
         mskCPFCondutor.SetFocus
      End If
   End If
End Sub

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
         If Registros2("MDFe") <> 0 Then
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
         MsgBox "Digite um N�mero para a Consulta", vbOKOnly, "Localizar"
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
         MsgBox "N�mero n�o Encontrado", vbOKOnly + vbExclamation, "Localizar"
         mskNumero.SetFocus
         mskNumero.SelStart = 0
         mskNumero.SelLength = Len(mskNumero.Text)
         rsMDFe.MoveFirst
         If RegistroAtual <> 0 Then rsMDFe.Find "idMDFe = " & RegistroAtual
      End If
   End If
End Sub

Sub Excluir()
On Error GoTo ErroIntegridade
   If MsgBox("Confirma Excluir o registro atual? ", vbYesNo + vbInformation, "Excluir") = vbYes Then
'      Atividade "Exclus�o: " & mskNumero.Text, Me.Caption
'      cnSistema.Execute "Delete * from MDFeItens Where idMDFe = " & rsMDFe!idMDFe  ' Itens da Nota de Entrada
      cnSistema.Execute "Delete from MDFeCondutores Where idMDFe=" & rsMDFe!idMDFe           ' Condutores
      cnSistema.Execute "Delete from MDFeLocalCarregamento Where idMDFe=" & rsMDFe!idMDFe           ' Condutores
      cnSistema.Execute "Delete from MDFeNFes Where idMDFe=" & rsMDFe!idMDFe           ' Condutores
      cnSistema.Execute "Delete from MDFePercurso Where idMDFe=" & rsMDFe!idMDFe           ' Condutores
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

Private Function Preencher_RsMDFe()
   
   rsMDFes!Numero = mskNumero.Text
   rsMDFes!DataEmissao = mskDataEmissao.Text
   rsMDFes!DataViagem = mskDataViagem.Text
   rsMDFes!idUF = FVerificaCombo(cboUF.ItemData(cboUF.ListIndex))
   rsMDFes!idTipoEmitente = FVerificaCombo(cboTipoEmitente.ItemData(cboTipoEmitente.ListIndex))
   rsMDFes!idTipoTransportador = FVerificaCombo(cboTipoTransportador.ItemData(cboTipoTransportador.ListIndex))
   rsMDFes!idFormaEmissao = FVerificaCombo(cboFormaEmissao.ItemData(cboFormaEmissao.ListIndex))
   rsMDFes!idModalidade = FVerificaCombo(cboModalidade.ItemData(cboModalidade.ListIndex))
   rsMDFes!idUFDescarregamento = FVerificaCombo(cboUFDescarregamento.ItemData(cboUFDescarregamento.ListIndex))
   rsMDFes!idTipoCarroceria = FVerificaCombo(cboTipoCarroceria.ItemData(IIf(cboTipoCarroceria.ListIndex <> -1, cboTipoCarroceria.ListIndex, 0)))
   rsMDFes!idUFVeiculo = FVerificaCombo(cboUFVeiculo.ItemData(cboUFVeiculo.ListIndex))
   rsMDFes!PlacaVeiculo = mskPlacaVeiculo.Text
   rsMDFes!tara = txtTara.Text
   rsMDFes!idTipoRodado = FVerificaCombo(cboTipoRodado.ItemData(cboTipoRodado.ListIndex))
   rsMDFes!CapacidadeKG = txtCapacidadeKG.Text
   rsMDFes!CapacidadeM3 = txtCapacidadeM3.Text
   rsMDFes!Renavam = txtRenavam.Text


End Function

Private Function Verifica_Campos()
Dim strMensagem As String
Verifica_Campos = True

   If mskNumero.Text = Empty Then strMensagem = strMensagem & "N�mero" & Chr(13)
   If Not IsDate(mskDataEmissao.Text) Or Val(Mid(mskDataEmissao.Text, 7, 4)) < 1900 Then strMensagem = strMensagem & "Data da Emiss�o" & Chr(13)
   If Not IsDate(mskDataViagem.Text) Or Val(Mid(mskDataViagem.Text, 7, 4)) < 1900 Then strMensagem = strMensagem & "Data da Viagem" & Chr(13)

   If cboUF.ListIndex = -1 Then strMensagem = strMensagem & "UF" & Chr(13)
   If cboTipoEmitente.ListIndex = -1 Then strMensagem = strMensagem & "Tipo de Emitente" & Chr(13)
   If cboTipoTransportador.ListIndex = -1 Then strMensagem = strMensagem & "Tipo de Transporte" & Chr(13)
   If cboFormaEmissao.ListIndex = -1 Then strMensagem = strMensagem & "Forma de Emiss�o" & Chr(13)
   If cboModalidade.ListIndex = -1 Then strMensagem = strMensagem & "Modalidade" & Chr(13)

   If cboUFDescarregamento.ListIndex = -1 Then strMensagem = strMensagem & "UF de Descarregamento" & Chr(13)

   If cboTipoCarroceria.ListIndex = -1 Then strMensagem = strMensagem & "Tipo de Carroceria" & Chr(13)
   If cboUFVeiculo.ListIndex = -1 Then strMensagem = strMensagem & "UF Ve�culo" & Chr(13)
   If cboTipoRodado.ListIndex = -1 Then strMensagem = strMensagem & "Tipo de Rodado" & Chr(13)

   If Not strMensagem = Empty Then
      MsgBox "Verifique os Seguintes Campos:" & Chr(13) & strMensagem, vbExclamation + vbOKOnly, "Campos Obrigat�rios"
      Verifica_Campos = False
      Exit Function
   Else
      Preencher_RsMDFe
   End If

End Function

Private Function Verifica_Campos_NFes()
Dim strMensagem As String
Verifica_Campos_NFes = True

   If txtNumeroNota.Text = Empty Then strMensagem = strMensagem & "N�mero" & Chr(13)
   If txtChaveNFe.Text = Empty Then strMensagem = strMensagem & "Chave da NFe" & Chr(13)
   If Len(txtChaveNFe.Text) <> 44 Then strMensagem = strMensagem & "Tamanho da Chave est� incorreto" & Chr(13)

   If Not strMensagem = Empty Then
      MsgBox "Verifique os Seguintes Campos:" & Chr(13) & strMensagem, vbExclamation + vbOKOnly, "Campos Obrigat�rios"
      Verifica_Campos_NFes = False
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
   cmdExcluirLocalCarregamento.Enabled = Parametro
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
'   cmdPesquisaNFes.Enabled = Parametro
   txtChaveNFe.Enabled = Parametro
   cmdIncluirNFe.Enabled = Parametro
   cmdExcluirNFe.Enabled = Parametro
   lvwNFes.Enabled = Parametro

   txtDadosAdicionais.Enabled = Parametro

End Sub

Sub Limpa_Campos()
Dim oPreencherRs As New PreencherRS

   ' Adicionar MDFes RecorSets Tempor�rios
   Call oPreencherRs.CriarEstruturaMDFe
   
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

   ' Atualizar Dados para grava��o
   rsMDFes!Numero = mskNumero.Text
   rsMDFes!idMDFe = mskNumero.Text
   rsMDFes!DataEmissao = mskDataEmissao.Text
   rsMDFes!DataViagem = mskDataViagem.Text

End Sub

Sub Prencher_Campos()
Dim Contador As Integer
Dim oPreencherRs As New PreencherRS

   ' Adicionar MDFes RecorSets Tempor�rios
'   Call oPreencherRs.PreencherRsDadosMDFe(rsMDFe!Numero)
   Call oPreencherRs.PreencherRsMDFe(rsMDFe!Numero)
   
   ' Adicionar MDFes RecorSets Tempor�rios
''   Call CriarEstruturaMDFe--
''
''   rsMDFes.AddNew
''   rsMDFes.Update

   txtChaveAcesso.Text = rsMDFes!ChaveMDFe
   txtProtocolo.Text = rsMDFes!Protocolo

   mskNumero.Text = rsMDFes!Numero
   mskDataEmissao.Text = rsMDFes!DataEmissao
   mskDataViagem.Text = rsMDFes!DataViagem

   ' Combos
   Call FUF(rsMDFes!idUF)
   Call FTipoEmitente(rsMDFes!idTipoEmitente)
   Call FTipoTransportador(rsMDFes!idTipoTransportador)
   Call FFormaEmissao(rsMDFes!idFormaEmissao)
   Call FModalidade(rsMDFes!idModalidade)

   ' Locais de Carregamento
   Call FWLocalCarregamento(rsMDFe!idMDFe)
'   Call FWLocalCarregamento(rsMDFe!Numero)

   Call FUFDescarregamento(rsMDFe!idUFDescarregamento)

   ' Percurso
   Call FWPercurso(rsMDFe!idMDFe)
'   Call FWPercurso(rsMDFe!Numero)

   Call FTipoCarroceria(rsMDFe!idTipoCarroceria)
   Call FUFVeiculo(rsMDFe!idUFVeiculo)

   mskPlacaVeiculo.Text = rsMDFes!PlacaVeiculo
   txtTara.Text = rsMDFes!tara

   Call FTipoRodado(rsMDFe!idTipoRodado)

   txtCapacidadeKG.Text = rsMDFes!CapacidadeKG
   txtCapacidadeM3.Text = rsMDFes!CapacidadeM3
   txtRenavam.Text = rsMDFes!Renavam

   ' Percurso
   Call FWCondutores(rsMDFe!idMDFe)
'   Call FWCondutores(rsMDFe!Numero)

   ' Percurso
   Call FWNFes(rsMDFe!idMDFe)
'   Call FWNFes(rsMDFe!Numero)

   txtDadosAdicionais.Text = rsMDFes!DadosAdicionais

End Sub

Sub Gravar()
Dim sSql As String
Dim rsInclusao As New ADODB.Recordset

   RegistroAtual = IIf(rsMDFe.EOF, 0, rsMDFe!idMDFe)
   If Not Verifica_Campos() Then Exit Sub

   Select Case Status
      Case 1 'Inclus�o
         If MsgBox("Confirma Incluir o registro atual", vbYesNo + vbQuestion, "Inclus�o") = vbYes Then
            sSql = Montar_Insert
            cnSistema.Execute sSql

            Set rsMDFe = cnSistema.Execute("Select * from MDFe Order By Numero")
            rsMDFe.MoveLast
            
            ' Gravar parcelas da nota
            Set rsInclusao = cnSistema.Execute("SELECT * FROM MDFe WHERE Numero = " & mskNumero.Text)
            If Not rsInclusao.EOF Then
               GravarItens (rsInclusao!idMDFe)
            End If
            
'            Atividade "Inclus�o: " & mskNumero.Text, Me.Caption
'''''            rsMDFe.Requery
'''''            rsMDFe.Find "Numero = '" & mskNumero.Text & "'"
'''''            rsMDFe.MoveLast
            
            ' Gravar itens
'            GravarItens (rsMDFes!idMDFe)
'''''            Set rsInclusao = cnSistema.Execute("SELECT idMDFe FROM MDFe WHERE Numero = " & mskNumero.Text)
'''''            GravarItens (rsInclusao!idmdfe)
            
'''''            rsMDFe.Requery
'''''            rsMDFe.MoveLast
'''''            rsMDFe.Find "idMDFe = " & rsInclusao!idmdfe
'''''            Prencher_Campos
            
'            rsMDFe.Requery
'            rsMDFe.Find "Numero = " & mskNumero.Text
            
            
            Set rsInclusao = Nothing
         End If

      Case 2 'Alterac�o
         If MsgBox("Confirma Alterar o registro atual", vbYesNo + vbQuestion, "Altera��o") = vbYes Then

            sSql = Montar_Update
            cnSistema.Execute sSql

'            Atividade "Alterar: " & mskNumero.Text, Me.Caption
'''''            rsMDFe.Requery
'''''            rsMDFe.Find "Numero = '" & mskNumero.Text & "'"
'            rsMDFe.Requery
            Set rsMDFe = cnSistema.Execute("Select * from MDFe Order By Numero")
            rsMDFe.Find "idMDFe = " & RegistroAtual
            
            ' Gravar itens
            GravarItens (rsMDFe!idMDFe)
            
'''            Prencher_Campos
         End If
   End Select
'   rsMDFe.Requery
'   rsMDFe.Find "Numero = '" & mskNumero.Text & "'"
   
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
   sSql = sSql & vbCrLf & "'" & rsMDFes!tara & "', "
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
   sSql = sSql & vbCrLf & "  Tara = '" & rsMDFes!tara & "', "
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
   
   If rsMDFeLocalCarregamento.RecordCount > 0 Then rsMDFeLocalCarregamento.MoveFirst
'   If Not rsMDFeLocalCarregamento.EOF Then rsMDFeLocalCarregamento.MoveFirst
   Do While Not rsMDFeLocalCarregamento.EOF
      cnSistema.Execute "INSERT INTO MDFeLocalCarregamento (idMDFe,idUF,idMunicipio) " & _
                        "VALUES (" & id & "," & rsMDFeLocalCarregamento!idUF & "," & rsMDFeLocalCarregamento!idMunicipio & ")"
   
      rsMDFeLocalCarregamento.MoveNext
   Loop
   
   ' Percursos
   cnSistema.Execute "DELETE FROM MDFePercurso WHERE idMDFe=" & id
'   If Not rsMDFePercurso.EOF Then rsMDFePercurso.MoveFirst
   If rsMDFePercurso.RecordCount > 0 Then rsMDFePercurso.MoveFirst
   Do While Not rsMDFePercurso.EOF
      cnSistema.Execute "INSERT INTO MDFePercurso (idMDFe,idUF) " & _
                        "VALUES (" & id & "," & rsMDFePercurso!idUF & ")"
   
      rsMDFePercurso.MoveNext
   Loop
   
   ' Condutores
   cnSistema.Execute "DELETE FROM MDFeCondutores WHERE idMDFe=" & id
'   If Not rsMDFeCondutores.EOF Then rsMDFeCondutores.MoveFirst
   If rsMDFeCondutores.RecordCount > 0 Then rsMDFeCondutores.MoveFirst
   Do While Not rsMDFeCondutores.EOF
      cnSistema.Execute "INSERT INTO MDFeCondutores (idMDFe,CPF,Nome) " & _
                        "VALUES (" & id & ",'" & rsMDFeCondutores!CPF & "','" & rsMDFeCondutores!Nome & "')"
   
      rsMDFeCondutores.MoveNext
   Loop
   
   ' NFes
   cnSistema.Execute "DELETE FROM MDFeNFes WHERE idMDFe=" & id
'   If Not rsMDFeNFes.EOF Then rsMDFeNFes.MoveFirst
   If rsMDFeNFes.RecordCount > 0 Then rsMDFeNFes.MoveFirst
   Do While Not rsMDFeNFes.EOF
      cnSistema.Execute "INSERT INTO MDFeNFes (idMDFe,Numero,ChaveNFe) " & _
                        "VALUES (" & id & ",'" & rsMDFeNFes!Numero & "','" & rsMDFeNFes!ChaveNFe & "')"
   
      rsMDFeNFes.MoveNext
   Loop
                              
''   Set rsSistema = Nothing
End Sub

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

Private Sub FWLocalCarregamento(ByRef iIdMDFe As Integer)
Dim iIdLocalCarregamento As Integer

   Set rsDados = cnSistema.Execute("SELECT LC.idMDFeLocalCarregamento, LC.idMDFe, LC.idMunicipio, M.Nome AS Municipio, LC.idUF, UF.Sigla AS UF " & _
                                  "FROM MDFeLocalCarregamento AS LC, Municipios AS M, UFs AS UF " & _
                                  "WHERE LC.idMunicipio = M.idMunicipio AND LC.idUF = UF.idUF AND LC.idMDFe = " & iIdMDFe)


   lvwLocalCarregamento.ListItems.Clear
   Do While Not rsDados.EOF
'      Set ItemList = lvwLocalCarregamento.ListItems.Add(, "R" & CStr(rsDados!idMDFeLocalCarregamento), rsDados!UF)

      iIdLocalCarregamento = rsDados!idMDFeLocalCarregamento

      Set ItemList = lvwLocalCarregamento.ListItems.Add(, "R" & CStr(rsDados!idMDFeLocalCarregamento), rsDados!UF)
      ItemList.SubItems(1) = rsDados!Municipio

'''''      iIdLocalCarregamento = lvwLocalCarregamento.ListItems.Count + 1

'''''      Set ItemList = lvwLocalCarregamento.ListItems.Add(, "R" & CStr(lvwLocalCarregamento.ListItems.Count + 1), rsDados!UF)
'''''      ItemList.SubItems(1) = rsDados!Municipio

'''''      rsMDFeLocalCarregamento.AddNew
'''''      rsMDFeLocalCarregamento!idMDFeLocalCarregamento = iIdLocalCarregamento
'''''      rsMDFeLocalCarregamento!idMDFe = iIdMDFe
'''''      rsMDFeLocalCarregamento!idMunicipio = rsDados!idMunicipio
'''''      rsMDFeLocalCarregamento!Municipio = rsDados!Municipio
'''''      rsMDFeLocalCarregamento!idUF = rsDados!idUF
'''''      rsMDFeLocalCarregamento!UF = rsDados!UF
'''''      rsMDFeLocalCarregamento.Update
      
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

Private Sub FWPercurso(ByRef iIdMDFe As Integer)
Dim iIdPercurso As Integer

   Set rsDados = cnSistema.Execute("SELECT P.idMDFePercurso, P.idMDFe, P.idUF, UF.Sigla AS UF " & _
                                  "FROM MDFePercurso AS P, UFs AS UF " & _
                                  "WHERE P.idUF = UF.idUF AND P.idMDFe = " & iIdMDFe)

   lvwPercurso.ListItems.Clear
   Do While Not rsDados.EOF
   
'      iIdPercurso = lvwPercurso.ListItems.Count + 1
      iIdPercurso = rsDados!idMDFePercurso

      Set ItemList = lvwPercurso.ListItems.Add(, "R" & CStr(rsDados!idMDFePercurso), rsDados!UF)
'      Set ItemList = lvwPercurso.ListItems.Add(, "R" & CStr(lvwPercurso.ListItems.Count + 1), rsDados!UF)

'''''      rsMDFePercurso.AddNew
'''''      rsMDFePercurso!idMDFePercurso = iIdPercurso
'''''      rsMDFePercurso!idMDFe = iIdMDFe
'''''      rsMDFePercurso!idUF = rsDados!idUF
'''''      rsMDFePercurso!UF = rsDados!UF
'''''      rsMDFePercurso.Update

      rsDados.MoveNext
   Loop

End Sub

Private Sub FWCondutores(ByRef iIdMDFe As Integer)
Dim iIdMDFeCondutor As Integer

   Set rsDados = cnSistema.Execute("SELECT C.idMDFeCondutor, C.idMDFe, C.CPF AS CPF, C.Nome AS Nome " & _
                                  "FROM MDFeCondutores AS C " & _
                                  "WHERE C.idMDFe = " & iIdMDFe)

   lvwCondutores.ListItems.Clear
   Do While Not rsDados.EOF
   
      iIdMDFeCondutor = rsDados!idMDFeCondutor
'      iIdMDFeCondutor = lvwCondutores.ListItems.Count + 1
   
      Set ItemList = lvwCondutores.ListItems.Add(, "R" & CStr(rsDados!idMDFeCondutor), rsDados!CPF)
'      Set ItemList = lvwCondutores.ListItems.Add(, "R" & CStr(lvwCondutores.ListItems.Count + 1), rsDados!CPF)
      ItemList.SubItems(1) = rsDados!Nome

      rsMDFeCondutores.AddNew
      rsMDFeCondutores!idMDFeCondutor = iIdMDFeCondutor
      rsMDFeCondutores!idMDFe = iIdMDFe
      rsMDFeCondutores!CPF = rsDados!CPF
      rsMDFeCondutores!Nome = rsDados!Nome
      rsMDFeCondutores.Update

      rsDados.MoveNext
   Loop

End Sub

Private Sub FWNFes(ByRef iIdMDFe As Integer)
Dim iIdMDFeNFe As Integer

   Set rsDados = cnSistema.Execute("SELECT NF.idMDFeNFe, NF.idMDFe, NF.Numero AS Numero, NF.ChaveNFe AS ChaveNFe " & _
                                  "FROM MDFeNFes AS NF " & _
                                  "WHERE NF.idMDFe = " & iIdMDFe)

   lvwNFes.ListItems.Clear
   Do While Not rsDados.EOF
   
      iIdMDFeNFe = rsDados!idMDFeNFe
'      iIdMDFeNFe = lvwNFes.ListItems.Count + 1
   
'      Set ItemList = lvwNFes.ListItems.Add(, "R" & CStr(lvwNFes.ListItems.Count + 1), rsDados!Numero)
      Set ItemList = lvwNFes.ListItems.Add(, "R" & CStr(rsDados!idMDFeNFe), rsDados!Numero)
      ItemList.SubItems(1) = FFormataChaveNF(IIf(IsNull(rsDados!ChaveNFe), "", rsDados!ChaveNFe))
'      ItemList.SubItems(2) = Format(mskValorTotalProdutos.Text, "###,##0.00")

      rsMDFeNFes.AddNew
      rsMDFeNFes!idMDFeNFe = iIdMDFeNFe
      rsMDFeNFes!idMDFe = iIdMDFe
      rsMDFeNFes!Numero = rsDados!Numero
      rsMDFeNFes!ChaveNFe = rsDados!ChaveNFe
      rsMDFeNFes.Update

      rsDados.MoveNext
   Loop

End Sub

Private Sub txtDadosAdicionais_LostFocus()
   rsMDFes!DadosAdicionais = txtDadosAdicionais.Text
End Sub

Private Sub txtNumeroNota_LostFocus()
   
'   Set rsNF = cnSistema.Execute("SELECT * FROM " & I_TabelasNF & " WHERE Numero = " & txtNumeroNota.Text)
'   If I_ModeloNF = "55" Then
'      If Not IsNull(rsNF!ChaveNFe) Or IsEmpty(rsNF!ChaveNFe) Then
'         txtChaveNFe.Text = rsNF!ChaveNFe                                                           ' Chave da NF
'      End If
'   ElseIf I_ModeloNF = "65" Then
'      If Not IsNull(rsNF!ChaveNFCe) Or IsEmpty(rsNF!ChaveNFCe) Then
'         txtChaveNFe.Text = rsNF!ChaveNFCe                                                          ' Chave da NF
'      End If
'   End If
   
'   Set rsMDFe = cnSistema.Execute("SELECT * FROM MDFe WHERE Numero = " & mskNumero.Text)
'   If Not rsMDFe.EOF And (Not IsNull(rsMDFe!ChaveMDFe) Or IsEmpty(rsMDFe!ChaveMDFe)) Then
'      txtChaveAcesso.Text = rsMDFe!ChaveMDFe                                                           ' Chave da NF
'   End If
   
   If Not IsEmpty(txtNumeroNota.Text) And txtNumeroNota.Text <> "" Then
      Set rsNFe = cnSistema.Execute("SELECT * FROM NFe WHERE Numero = " & txtNumeroNota.Text)
      If Not rsNFe.EOF And (Not IsNull(rsNFe!ChaveNFe) Or IsEmpty(rsNFe!ChaveNFe)) Then
         txtChaveNFe.Text = rsNFe!ChaveNFe                                                           ' Chave da NF
      End If
      Set rsNFe = Nothing
   End If
   
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
   
   If MsgBox("Deseja excluir este item", vbYesNo + vbQuestion, "Excluir") = vbYes Then

      iId = Val(Mid(lvwLocalCarregamento.SelectedItem.Key, 2, Len(lvwLocalCarregamento.SelectedItem.Key)))
      rsMDFeLocalCarregamento.MoveFirst
      rsMDFeLocalCarregamento.Find "idMDFeLocalCarregamento = " & iId
      rsMDFeLocalCarregamento.Delete adAffectCurrent
      
      ' Recarregar View
      lvwLocalCarregamento.ListItems.Clear
      rsMDFeLocalCarregamento.MoveFirst
      Do While Not rsMDFeLocalCarregamento.EOF
         Set ItemList = lvwLocalCarregamento.ListItems.Add(, "R" & CStr(rsMDFeLocalCarregamento!idMDFeLocalCarregamento), rsMDFeLocalCarregamento!UF)
         ItemList.SubItems(1) = rsMDFeLocalCarregamento!Municipio
   
         rsMDFeLocalCarregamento.MoveNext
      Loop
   End If
End Sub


Private Sub cmdInserirLocalCarregamento_Click()
Dim iIdLocalCarregamento As Integer

   If cboUFCarregamento.ListIndex = -1 Or cboMunicipio.ListIndex = -1 Then
      MsgBox "Munic�pio e UF s�o obrigat�rios", vbExclamation, "Sistema"
      Exit Sub
   End If
      
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
   rsMDFeLocalCarregamento!idMDFe = rsMDFes!idMDFe
   rsMDFeLocalCarregamento!idUF = cboUFCarregamento.ItemData(cboUFCarregamento.ListIndex)
   rsMDFeLocalCarregamento!UF = cboUFCarregamento.Text
   rsMDFeLocalCarregamento!idMunicipio = cboMunicipio.ItemData(cboMunicipio.ListIndex)
   rsMDFeLocalCarregamento!Municipio = cboMunicipio.Text
   rsMDFeLocalCarregamento.Update

   cboUFCarregamento.ListIndex = -1
   cboMunicipio.ListIndex = -1

End Sub

Private Sub cmdIncluirCondutor_Click()
Dim iIdCondutor As Integer

   If IsEmpty(mskCPFCondutor.Text) Or IsEmpty(txtNomeCondutor.Text) Then
      MsgBox "CPF e Nome s�o obrigat�rios", vbExclamation, "Sistema"
      Exit Sub
   End If
      
   If rsMDFeCondutores.EOF Then
      iIdCondutor = lvwCondutores.ListItems.Count + 1
   Else
      rsMDFeCondutores.MoveLast
      iIdCondutor = rsMDFeCondutores!idMDFeCondutor + 1
   End If

   rsMDFeCondutores.AddNew
   rsMDFeCondutores!idMDFeCondutor = iIdCondutor
   rsMDFeCondutores!idMDFe = rsMDFes!idMDFe
   rsMDFeCondutores!CPF = mskCPFCondutor.Text
   rsMDFeCondutores!Nome = txtNomeCondutor.Text
   rsMDFeCondutores.Update

   Set ItemList = lvwCondutores.ListItems.Add(, "R" & CStr(lvwCondutores.ListItems.Count + 1), mskCPFCondutor.Text)
   ItemList.SubItems(1) = txtNomeCondutor.Text

   mskCPFCondutor.Text = Empty
   txtNomeCondutor.Text = Empty

End Sub

Private Sub cmdIncluirNFe_Click()

   If Not Verifica_Campos_NFes Then Exit Sub

'   If MsgBox("Confirma Incluir o registro atual", vbYesNo + vbQuestion, "Inclus�o") = vbYes Then
      rsMDFeNFes.AddNew
      rsMDFeNFes!idMDFe = rsMDFes!idMDFe
      rsMDFeNFes!Numero = Val(txtNumeroNota.Text)
      rsMDFeNFes!ChaveNFe = txtChaveNFe.Text
      rsMDFeNFes.Update
   
      Set ItemList = lvwNFes.ListItems.Add(, "R" & CStr(lvwNFes.ListItems.Count + 1), txtNumeroNota.Text)
      ItemList.SubItems(1) = FFormataChaveNF(IIf(IsNull(txtChaveNFe.Text), "", txtChaveNFe.Text))
   
      txtNumeroNota.Text = Empty
      txtChaveNFe.Text = Empty
'   End If
End Sub

Private Sub cmdInserirPercurso_Click()
Dim iIdPercurso As Integer

   If cboUFPercurso.ListIndex = -1 Then
      MsgBox "UF � obrigat�rio", vbExclamation, "Sistema"
      Exit Sub
   End If
      
   If rsMDFePercurso.EOF Then
      iIdPercurso = lvwPercurso.ListItems.Count + 1
   Else
      rsMDFePercurso.MoveLast
      iIdPercurso = rsMDFePercurso!idMDFePercurso + 1
   End If
      
   rsMDFePercurso.AddNew
   rsMDFePercurso!idMDFePercurso = iIdPercurso
   rsMDFePercurso!idMDFe = rsMDFes!idMDFe
   rsMDFePercurso!idUF = cboUFPercurso.ItemData(cboUFPercurso.ListIndex)
   rsMDFePercurso.Update

   Set ItemList = lvwPercurso.ListItems.Add(, "R" & CStr(lvwPercurso.ListItems.Count + 1), cboUFPercurso.Text)

   cboUFPercurso.ListIndex = -1

End Sub

Private Sub cmdExcluirPercurso_Click()
Dim iId As Integer
   
   If MsgBox("Deseja excluir este item", vbYesNo + vbQuestion, "Excluir") = vbYes Then
      iId = Val(Mid(lvwPercurso.SelectedItem.Key, 2, Len(lvwPercurso.SelectedItem.Key)))
'      iId = lvwPercurso.SelectedItem.Index
      rsMDFePercurso.MoveFirst
      rsMDFePercurso.Find "idMDFePercurso = " & iId
      rsMDFePercurso.Delete adAffectCurrent
      
      ' Recarregar View
      lvwPercurso.ListItems.Clear
      rsMDFePercurso.MoveFirst
      Do While Not rsMDFePercurso.EOF
'         Set ItemList = lvwPercurso.ListItems.Add(, "R" & CStr(lvwPercurso.ListItems.Count + 1), rsMDFePercurso!UF)
         Set ItemList = lvwPercurso.ListItems.Add(, "R" & CStr(rsMDFePercurso!idMDFePercurso), rsMDFePercurso!UF)
   
         rsMDFePercurso.MoveNext
      Loop
   End If
   
End Sub

Private Sub cmdExcluirCondutor_Click()
Dim iId As Integer
   
   If MsgBox("Deseja excluir este item", vbYesNo + vbQuestion, "Excluir") = vbYes Then
      iId = Val(Mid(lvwCondutores.SelectedItem.Key, 2, Len(lvwCondutores.SelectedItem.Key)))
'      iId = lvwCondutores.SelectedItem.Index
      rsMDFeCondutores.MoveFirst
      rsMDFeCondutores.Find "idMDFeCondutor = " & iId
      rsMDFeCondutores.Delete adAffectCurrent
      
      ' Recarregar View
      lvwCondutores.ListItems.Clear
      rsMDFeCondutores.MoveFirst
      Do While Not rsMDFeCondutores.EOF
         Set ItemList = lvwCondutores.ListItems.Add(, "R" & CStr(rsMDFeCondutores!idMDFeCondutor), rsMDFeCondutores!CPF)
'         Set ItemList = lvwCondutores.ListItems.Add(, "R" & CStr(lvwCondutores.ListItems.Count + 1), rsMDFeCondutores!CPF)
         ItemList.SubItems(1) = rsMDFeCondutores!Nome
   
         rsMDFeCondutores.MoveNext
      Loop
   End If
End Sub

Private Sub cmdExcluirNFe_Click()
Dim iId As Integer
   
   If MsgBox("Deseja excluir este item", vbYesNo + vbQuestion, "Excluir") = vbYes Then
      iId = Val(Mid(lvwNFes.SelectedItem.Key, 2, Len(lvwNFes.SelectedItem.Key)))
'      iId = lvwNFes.SelectedItem.Index
      rsMDFeNFes.MoveFirst
      rsMDFeNFes.Find "idMDFeNFe = " & iId
      rsMDFeNFes.Delete adAffectCurrent
      
      ' Recarregar View
      lvwNFes.ListItems.Clear
      rsMDFeNFes.MoveFirst
      Do While Not rsMDFeNFes.EOF
         Set ItemList = lvwNFes.ListItems.Add(, "R" & CStr(rsMDFeNFes!idMDFeNFe), rsMDFeNFes!Numero)
'         Set ItemList = lvwNFes.ListItems.Add(, "R" & CStr(lvwNFes.ListItems.Count + 1), rsMDFeNFes!Numero)
         ItemList.SubItems(1) = FFormataChaveNF(IIf(IsNull(rsMDFeNFes!ChaveNFe), "", rsMDFeNFes!ChaveNFe))
   
         rsMDFeNFes.MoveNext
      Loop
   End If
End Sub

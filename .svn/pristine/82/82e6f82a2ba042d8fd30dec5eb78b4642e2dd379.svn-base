VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmNFe 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Emitir NFe"
   ClientHeight    =   8880
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9630
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8880
   ScaleWidth      =   9630
   Begin VB.Frame fraDadosProduto 
      Height          =   1410
      Left            =   9780
      TabIndex        =   107
      Top             =   660
      Visible         =   0   'False
      Width           =   4170
      Begin VB.Label lblDadosProduto 
         BackColor       =   &H80000018&
         ForeColor       =   &H80000008&
         Height          =   1230
         Left            =   45
         TabIndex        =   108
         Top             =   135
         Width           =   4080
      End
   End
   Begin VB.ComboBox cboDados 
      Height          =   960
      Left            =   1260
      Style           =   1  'Simple Combo
      TabIndex        =   3
      Top             =   660
      Width           =   8235
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7110
      Left            =   60
      TabIndex        =   5
      Top             =   1680
      Width           =   9510
      _ExtentX        =   16775
      _ExtentY        =   12541
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Nota Fiscal"
      TabPicture(0)   =   "frmNFe.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblCancelada"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblProtocolo"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblChaveAcesso"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtProtocolo"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtChaveAcesso"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "fraEndereco"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "fraDatas"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "fraCabecalho"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "fraCliente"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "fraProdutos"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "Informações Adicionais"
      TabPicture(1)   =   "frmNFe.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtChaveAcessoDevolucao"
      Tab(1).Control(1)=   "fraDadosAdicionais"
      Tab(1).Control(2)=   "fraTransportadorVolumes"
      Tab(1).Control(3)=   "fraFatura"
      Tab(1).Control(4)=   "fraInformacoesCorpoNota"
      Tab(1).Control(5)=   "lblChaveAcessoDevolucao"
      Tab(1).ControlCount=   6
      Begin VB.TextBox txtChaveAcessoDevolucao 
         Height          =   315
         Left            =   -74940
         TabIndex        =   109
         Top             =   5895
         Width           =   6780
      End
      Begin VB.Frame fraProdutos 
         Height          =   3045
         Left            =   60
         TabIndex        =   41
         Top             =   3990
         Width           =   9375
         Begin VB.CommandButton cmdAnotacoes 
            BackColor       =   &H00E0E0E0&
            Height          =   315
            Left            =   8175
            Picture         =   "frmNFe.frx":0038
            Style           =   1  'Graphical
            TabIndex        =   52
            ToolTipText     =   "Informações Adicionais"
            Top             =   405
            Width           =   360
         End
         Begin VB.CommandButton cmdInserir 
            BackColor       =   &H00E0E0E0&
            Height          =   315
            Left            =   8535
            Picture         =   "frmNFe.frx":040A
            Style           =   1  'Graphical
            TabIndex        =   53
            ToolTipText     =   "Confirma inclusão de produtos"
            Top             =   405
            Width           =   360
         End
         Begin VB.CommandButton cmdRemover 
            BackColor       =   &H80000004&
            Height          =   315
            Left            =   8925
            Picture         =   "frmNFe.frx":0554
            Style           =   1  'Graphical
            TabIndex        =   54
            ToolTipText     =   "Exclui produto selecionado"
            Top             =   405
            Width           =   360
         End
         Begin MSMask.MaskEdBox mskQuantidade 
            Height          =   315
            Left            =   5610
            TabIndex        =   47
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
            Left            =   7500
            TabIndex        =   51
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
         Begin MSMask.MaskEdBox mskValorUnitario 
            Height          =   315
            Left            =   6480
            TabIndex        =   49
            Top             =   405
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   556
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
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSComctlLib.ListView lvwProdutos 
            Height          =   1560
            Left            =   90
            TabIndex        =   55
            Top             =   720
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
            TabIndex        =   57
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
            TabIndex        =   59
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
            TabIndex        =   67
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
            TabIndex        =   61
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
            TabIndex        =   65
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
            TabIndex        =   63
            Top             =   2625
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   13
            Format          =   "###,##0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskCodigoBarra 
            Height          =   315
            Left            =   90
            TabIndex        =   43
            Top             =   405
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   20
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   " "
         End
         Begin VB.ComboBox cboProduto 
            Height          =   1350
            Left            =   1380
            Style           =   1  'Simple Combo
            TabIndex        =   45
            Top             =   405
            Width           =   4185
         End
         Begin VB.Label lblTotalNota 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   8055
            TabIndex        =   71
            Top             =   2625
            Width           =   1215
         End
         Begin VB.Label lblTotalProdutos 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   8055
            TabIndex        =   69
            Top             =   2325
            Width           =   1215
         End
         Begin VB.Label lblCodigo 
            AutoSize        =   -1  'True
            Caption         =   "Código"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   90
            TabIndex        =   42
            Top             =   180
            Width           =   495
         End
         Begin VB.Label lblProduto 
            AutoSize        =   -1  'True
            Caption         =   "Produto"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   1380
            TabIndex        =   44
            Top             =   180
            Width           =   555
         End
         Begin VB.Label lblQuantidade 
            AutoSize        =   -1  'True
            Caption         =   "Quantidade"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   5610
            TabIndex        =   46
            Top             =   180
            Width           =   825
         End
         Begin VB.Label lblDesconto 
            AutoSize        =   -1  'True
            Caption         =   "Desc."
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   7725
            TabIndex        =   50
            Top             =   180
            Width           =   420
         End
         Begin VB.Label lblUnitario 
            AutoSize        =   -1  'True
            Caption         =   "Vl. Unitário"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   6660
            TabIndex        =   48
            Top             =   180
            Width           =   765
         End
         Begin VB.Label lblValorFrete 
            AutoSize        =   -1  'True
            Caption         =   "Valor do Frete"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   4785
            TabIndex        =   64
            Top             =   2370
            Width           =   990
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Base Subst."
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   225
            TabIndex        =   58
            Top             =   2670
            Width           =   855
         End
         Begin VB.Label lblValorICMS 
            AutoSize        =   -1  'True
            Caption         =   "Valor do ICMS"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   2430
            TabIndex        =   60
            Top             =   2370
            Width           =   1020
         End
         Begin VB.Label lblBaseCalculoICMS 
            AutoSize        =   -1  'True
            Caption         =   "Base do ICMS"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   60
            TabIndex        =   56
            Top             =   2370
            Width           =   1020
         End
         Begin VB.Label lblValorTotalProdutos 
            AutoSize        =   -1  'True
            Caption         =   "Total Produtos"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   6960
            TabIndex        =   68
            Top             =   2370
            Width           =   1035
         End
         Begin VB.Label lblValorTotalNota 
            AutoSize        =   -1  'True
            Caption         =   "Total Nota"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   7245
            TabIndex        =   70
            Top             =   2670
            Width           =   750
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Vl. ICMS Sub."
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   2460
            TabIndex        =   62
            Top             =   2670
            Width           =   990
         End
         Begin VB.Label lblDespesas 
            AutoSize        =   -1  'True
            Caption         =   "Despesas"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   5040
            TabIndex        =   66
            Top             =   2700
            Width           =   705
         End
      End
      Begin VB.Frame fraCliente 
         Caption         =   "Cliente"
         Height          =   1425
         Left            =   60
         TabIndex        =   10
         Top             =   1050
         Width           =   5475
         Begin VB.ComboBox cboCliente 
            Height          =   960
            Left            =   105
            Style           =   1  'Simple Combo
            TabIndex        =   11
            Top             =   240
            Width           =   5265
         End
      End
      Begin VB.Frame fraCabecalho 
         Height          =   1560
         Left            =   60
         TabIndex        =   14
         Top             =   2475
         Width           =   6975
         Begin VB.CommandButton cmdPesquisaCFOP 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1500
            Picture         =   "frmNFe.frx":0ADE
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   480
            Width           =   360
         End
         Begin VB.ComboBox cboNaturezaOperacao 
            Height          =   315
            Left            =   4380
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   165
            Width           =   2505
         End
         Begin VB.ComboBox cboFormaPagamento 
            Height          =   315
            Left            =   4380
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   480
            Width           =   2505
         End
         Begin VB.TextBox txtDocumento 
            Height          =   315
            Left            =   4380
            MaxLength       =   20
            TabIndex        =   31
            Top             =   795
            Width           =   2490
         End
         Begin VB.TextBox txtObservacao 
            Height          =   315
            Left            =   810
            MaxLength       =   20
            TabIndex        =   33
            Top             =   1125
            Width           =   6060
         End
         Begin MSMask.MaskEdBox mskNumero 
            Height          =   285
            Left            =   810
            TabIndex        =   16
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
            TabIndex        =   22
            Top             =   495
            Width           =   645
            _ExtentX        =   1138
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   5
            Mask            =   "9.999"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskCupom 
            Height          =   285
            Left            =   2550
            TabIndex        =   18
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
            TabIndex        =   27
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
            TabIndex        =   29
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
         Begin VB.Label lblCupom 
            AutoSize        =   -1  'True
            Caption         =   "Cupom"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   1980
            TabIndex        =   17
            Top             =   225
            Width           =   495
         End
         Begin VB.Label lblCFOP 
            AutoSize        =   -1  'True
            Caption         =   "CFOP"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   315
            TabIndex        =   21
            Top             =   585
            Width           =   420
         End
         Begin VB.Label lblNaturezaOperacao 
            AutoSize        =   -1  'True
            Caption         =   "Natureza"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   3645
            TabIndex        =   19
            Top             =   225
            Width           =   645
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
            Left            =   75
            TabIndex        =   15
            Top             =   225
            Width           =   660
         End
         Begin VB.Label lblTipoPagamento 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Pg."
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   3690
            TabIndex        =   24
            Top             =   540
            Width           =   600
         End
         Begin VB.Label lblDocumento 
            AutoSize        =   -1  'True
            Caption         =   "Documento"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   3465
            TabIndex        =   30
            Top             =   855
            Width           =   825
         End
         Begin VB.Label lblObservacao 
            AutoSize        =   -1  'True
            Caption         =   "Obs."
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   405
            TabIndex        =   32
            Top             =   1185
            Width           =   330
         End
         Begin VB.Label lblDescontoGeral 
            AutoSize        =   -1  'True
            Caption         =   "Desconto"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   45
            TabIndex        =   26
            Top             =   840
            Width           =   690
         End
         Begin VB.Label lblBonificacao 
            AutoSize        =   -1  'True
            Caption         =   "Bonificação"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   1635
            TabIndex        =   28
            Top             =   855
            Visible         =   0   'False
            Width           =   840
         End
      End
      Begin VB.Frame fraDatas 
         Height          =   1560
         Left            =   7080
         TabIndex        =   34
         Top             =   2475
         Width           =   2355
         Begin MSMask.MaskEdBox mskDataEmissao 
            Height          =   285
            Left            =   1275
            TabIndex        =   36
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
            TabIndex        =   38
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
            TabIndex        =   40
            Top             =   795
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   8
            Mask            =   "99:99:99"
            PromptChar      =   " "
         End
         Begin VB.Label lblHora 
            AutoSize        =   -1  'True
            Caption         =   "Hora"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   855
            TabIndex        =   39
            Top             =   840
            Width           =   345
         End
         Begin VB.Label lblDataVencimento 
            AutoSize        =   -1  'True
            Caption         =   "Saída"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   765
            TabIndex        =   37
            Top             =   540
            Width           =   435
         End
         Begin VB.Label lblDataEmissao 
            AutoSize        =   -1  'True
            Caption         =   "Emissão"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   615
            TabIndex        =   35
            Top             =   240
            Width           =   585
         End
      End
      Begin VB.Frame fraDadosAdicionais 
         Caption         =   "Dados Adicionais/Informações Complementares"
         Height          =   2385
         Left            =   -74940
         TabIndex        =   92
         Top             =   1800
         Width           =   4575
         Begin VB.TextBox txtDadosAdicionais 
            Height          =   2055
            Left            =   75
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   93
            Top             =   240
            Width           =   4410
         End
      End
      Begin VB.Frame fraTransportadorVolumes 
         Caption         =   "Transportador/Volumes"
         Height          =   1335
         Left            =   -74940
         TabIndex        =   72
         Top             =   420
         Width           =   9375
         Begin VB.ComboBox cmbTransportador 
            Height          =   315
            Left            =   1020
            Style           =   2  'Dropdown List
            TabIndex        =   74
            Top             =   240
            Width           =   3735
         End
         Begin VB.ComboBox cmbFreteConta 
            Height          =   315
            ItemData        =   "frmNFe.frx":0C28
            Left            =   5400
            List            =   "frmNFe.frx":0C38
            Style           =   2  'Dropdown List
            TabIndex        =   83
            Top             =   240
            Width           =   1935
         End
         Begin VB.TextBox txtVolumeMarca 
            Height          =   315
            Left            =   2880
            MaxLength       =   50
            TabIndex        =   80
            Top             =   570
            Width           =   1860
         End
         Begin VB.TextBox txtVolumeNumero 
            Height          =   315
            Left            =   5400
            MaxLength       =   50
            TabIndex        =   85
            Top             =   570
            Width           =   1920
         End
         Begin VB.TextBox txtVolumeQuantidade 
            Height          =   315
            Left            =   1020
            MaxLength       =   50
            TabIndex        =   76
            Top             =   570
            Width           =   1215
         End
         Begin VB.TextBox txtVolumeEspecie 
            Height          =   315
            Left            =   5400
            MaxLength       =   50
            TabIndex        =   87
            Top             =   900
            Width           =   1920
         End
         Begin VB.ComboBox cmbUFPlaca 
            Height          =   315
            Left            =   7920
            Style           =   2  'Dropdown List
            TabIndex        =   91
            Top             =   570
            Width           =   735
         End
         Begin MSMask.MaskEdBox mskPlaca 
            Height          =   285
            Left            =   7920
            TabIndex        =   89
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
            TabIndex        =   78
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
            TabIndex        =   106
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
         Begin VB.Label lblFreteConta 
            AutoSize        =   -1  'True
            Caption         =   "Frete"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   4995
            TabIndex        =   82
            Top             =   300
            Width           =   360
         End
         Begin VB.Label lblPlaca 
            AutoSize        =   -1  'True
            Caption         =   "Placa"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   7440
            TabIndex        =   88
            Top             =   300
            Width           =   405
         End
         Begin VB.Label lblTransportador 
            AutoSize        =   -1  'True
            Caption         =   "Nome"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   525
            TabIndex        =   73
            Top             =   300
            Width           =   420
         End
         Begin VB.Label lblVolumeQuantidade 
            AutoSize        =   -1  'True
            Caption         =   "Quantidade"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   75
            Top             =   630
            Width           =   825
         End
         Begin VB.Label lblVolumeMarca 
            AutoSize        =   -1  'True
            Caption         =   "Marca"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   2340
            TabIndex        =   79
            Top             =   630
            Width           =   450
         End
         Begin VB.Label lblVolumeNumero 
            AutoSize        =   -1  'True
            Caption         =   "Número"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   4800
            TabIndex        =   84
            Top             =   630
            Width           =   555
         End
         Begin VB.Label lblVolumePesoBruto 
            AutoSize        =   -1  'True
            Caption         =   "Peso Bruto"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   165
            TabIndex        =   77
            Top             =   960
            Width           =   780
         End
         Begin VB.Label lblVolumePesoLiquido 
            AutoSize        =   -1  'True
            Caption         =   "Peso Líquido"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   2340
            TabIndex        =   81
            Top             =   960
            Width           =   945
         End
         Begin VB.Label lblVolumeEspecie 
            AutoSize        =   -1  'True
            Caption         =   "Espécie"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   4785
            TabIndex        =   86
            Top             =   960
            Width           =   570
         End
         Begin VB.Label lblUFCaminhao 
            AutoSize        =   -1  'True
            Caption         =   "UF"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   7620
            TabIndex        =   90
            Top             =   600
            Width           =   210
         End
      End
      Begin VB.Frame fraFatura 
         Caption         =   "Faturamento"
         Height          =   2385
         Left            =   -70320
         TabIndex        =   94
         Top             =   1800
         Width           =   4755
         Begin VB.CommandButton cmdExcluirFatura 
            BackColor       =   &H80000004&
            Height          =   315
            Left            =   4320
            Style           =   1  'Graphical
            TabIndex        =   102
            Top             =   405
            Width           =   360
         End
         Begin VB.CommandButton cmdIncluirFatura 
            BackColor       =   &H80000004&
            Height          =   315
            Left            =   3960
            Style           =   1  'Graphical
            TabIndex        =   101
            Top             =   405
            Width           =   360
         End
         Begin MSMask.MaskEdBox mskNumeroBoleto 
            Height          =   285
            Left            =   60
            TabIndex        =   96
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
            TabIndex        =   98
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
            TabIndex        =   100
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
            TabIndex        =   103
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
         Begin VB.Label lblBoleto 
            AutoSize        =   -1  'True
            Caption         =   "Boleto"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   60
            TabIndex        =   95
            Top             =   210
            Width           =   450
         End
         Begin VB.Label lblVencimentoBoleto 
            AutoSize        =   -1  'True
            Caption         =   "Vencimento"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   1710
            TabIndex        =   97
            Top             =   210
            Width           =   840
         End
         Begin VB.Label lblValorBoleto 
            AutoSize        =   -1  'True
            Caption         =   "Valor"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   2700
            TabIndex        =   99
            Top             =   210
            Width           =   360
         End
      End
      Begin VB.Frame fraInformacoesCorpoNota 
         Caption         =   "Informações no Corpo da Nota"
         Height          =   1275
         Left            =   -74940
         TabIndex        =   104
         Top             =   4200
         Width           =   9375
         Begin VB.TextBox txtInformacoesCorpo 
            Height          =   915
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   105
            Top             =   240
            Width           =   9135
         End
      End
      Begin VB.Frame fraEndereco 
         Caption         =   "Nome Fantasia/Endereço"
         Height          =   1425
         Left            =   5580
         TabIndex        =   12
         Top             =   1050
         Width           =   3855
         Begin VB.Label lblDescricaoEndereco 
            ForeColor       =   &H80000008&
            Height          =   1080
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Width           =   3615
         End
      End
      Begin VB.TextBox txtChaveAcesso 
         Enabled         =   0   'False
         Height          =   315
         Left            =   75
         TabIndex        =   7
         Top             =   660
         Width           =   6780
      End
      Begin VB.TextBox txtProtocolo 
         Enabled         =   0   'False
         Height          =   315
         Left            =   6900
         TabIndex        =   9
         Top             =   660
         Width           =   2460
      End
      Begin VB.Label lblChaveAcessoDevolucao 
         AutoSize        =   -1  'True
         Caption         =   "Chave de Acesso - Devolução"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   -74940
         TabIndex        =   110
         Top             =   5640
         Width           =   2175
      End
      Begin VB.Label lblChaveAcesso 
         AutoSize        =   -1  'True
         Caption         =   "Chave de Acesso"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   75
         TabIndex        =   6
         Top             =   405
         Width           =   1260
      End
      Begin VB.Label lblProtocolo 
         AutoSize        =   -1  'True
         Caption         =   "Protocolo de Autorização"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   6900
         TabIndex        =   8
         Top             =   420
         Width           =   1785
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
         TabIndex        =   4
         Top             =   60
         Width           =   1110
      End
   End
   Begin MSComctlLib.ImageList imgToolBar 
      Left            =   5820
      Top             =   60
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNFe.frx":0C76
            Key             =   "Pesquisar"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNFe.frx":16CA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNFe.frx":19F82
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNFe.frx":1D264
            Key             =   "Imprimir"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNFe.frx":3328E
            Key             =   "Excluir"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9630
      _ExtentX        =   16986
      _ExtentY        =   1058
      ButtonWidth     =   2355
      ButtonHeight    =   1005
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imgToolBar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Novo"
            Key             =   "Novo"
            Object.ToolTipText     =   "Novo cadastro"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Excluir"
            Key             =   "Excluir"
            Object.ToolTipText     =   "Exclui cadastro"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Gravar"
            Key             =   "Gravar"
            Object.ToolTipText     =   "Grava alterações"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Imprimir"
            Key             =   "Imprimir"
            Object.ToolTipText     =   "Imprime cadastro"
            ImageIndex      =   4
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Label lblRegistros 
      Alignment       =   2  'Center
      Caption         =   "registros"
      Height          =   435
      Left            =   60
      TabIndex        =   1
      Top             =   1140
      Width           =   1185
   End
   Begin VB.Label lblPesquisa 
      Alignment       =   2  'Center
      Caption         =   "» Saida, Cliente ou Data"
      Height          =   375
      Left            =   60
      TabIndex        =   2
      Top             =   660
      Width           =   1155
   End
End
Attribute VB_Name = "frmNFe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public localErro As String

Dim rsDados As New ADODB.Recordset
Dim ItemList As ListItem
Dim bytCaixa As Byte
Dim bytControleCredito As Byte
Dim curDesconto As Currency
Dim curComissao As Currency
Dim curValorTotal As Currency
Dim curValorPagar As Currency
Dim curValorCredito As Currency
Dim curValorDesconto As Currency
Dim curValorFrete As Currency
Dim ProdutoPromocional As Boolean
Dim bolAplicarDesconto As Boolean
Dim PermitirDescontoVendedor As Boolean
Dim BloquearDescontoPromocional As Boolean

Dim dblValorTotal As Double 'Valor total da Compra
Dim rsNFe As New ADODB.Recordset
Dim rsCFOPs As New ADODB.Recordset
Dim rsUFs As New ADODB.Recordset
Dim rsTotalNFe As New ADODB.Recordset
Dim rsEmpresa As New ADODB.Recordset
Dim rsClientes As New ADODB.Recordset
Dim rsTemp As New ADODB.Recordset
Dim bolDesconto As Boolean
Dim RegistroAtual As Double 'Para posicionar o ponteiro depois de pesquisas

Private Sub Form_Load()
   Limpa_Campos
   
   SSTab1.Tab = 0 ' Posiciona no primeiro tab
   
''''   lvwProdutos.ColumnHeaders.Add , , "Cód. de Barra", 1280
''   lvwProdutos.ColumnHeaders.Add , , "Produto", 4400
''   lvwProdutos.ColumnHeaders.Add , , "Quantidade", 1000, lvwColumnRight
''   lvwProdutos.ColumnHeaders.Add , , "Valor Unitário", 1300, lvwColumnRight
''   lvwProdutos.ColumnHeaders.Add , , "Desc.", 800, lvwColumnRight
''   lvwProdutos.ColumnHeaders.Add , , "Valor Total", 1300, lvwColumnRight

'''   lvwProdutos.ColumnHeaders.Add , , "Produto", 4400
'''   lvwProdutos.ColumnHeaders.Add , , "Quant.", 1000, lvwColumnRight
'''   lvwProdutos.ColumnHeaders.Add , , "Valor unitário", 1300, lvwColumnRight
'''   lvwProdutos.ColumnHeaders.Add , , "Desc.", 800, lvwColumnRight
'''   lvwProdutos.ColumnHeaders.Add , , "Valor total", 1300, lvwColumnRight
'''   lvwProdutos.ColumnHeaders.Add , , "Valor a pagar", 1300, lvwColumnRight
'''   lvwProdutos.ColumnHeaders.Add , , "Comissão", 0
'''   lvwProdutos.ColumnHeaders.Add , , "DataRecebido", 0
'''   lvwProdutos.ColumnHeaders.Add , , "Caixa", 0
   
   lvwProdutos.ColumnHeaders.Add , , "Produto", 3300
   lvwProdutos.ColumnHeaders.Add , , "Quant.", 700, lvwColumnRight
   lvwProdutos.ColumnHeaders.Add , , "Valor unitário", 1200, lvwColumnRight
   lvwProdutos.ColumnHeaders.Add , , "Valor total", 1200, lvwColumnRight
   lvwProdutos.ColumnHeaders.Add , , "Desc.", 650, lvwColumnRight
   lvwProdutos.ColumnHeaders.Add , , "Valor a pagar", 1200, lvwColumnRight
   lvwProdutos.ColumnHeaders.Add , , "Comissão", 0
   lvwProdutos.ColumnHeaders.Add , , "DataRecebido", 0
   lvwProdutos.ColumnHeaders.Add , , "Caixa", 0
   lvwProdutos.ColumnHeaders.Add , , "ICMS", 0
   lvwProdutos.ColumnHeaders.Add , , "BaseReduzida", 0
   lvwProdutos.ColumnHeaders.Add , , "DescricaoComplementar", 0
   lvwProdutos.ColumnHeaders.Add , , "Unidade", 0
   lvwProdutos.ColumnHeaders.Add , , "SituacaoTributaria", 0
   lvwProdutos.ColumnHeaders.Add , , "DiscriminacaoProduto", 0
   lvwProdutos.ColumnHeaders.Add , , "IPI", 0
   lvwProdutos.ColumnHeaders.Add , , "BaseReduzidaIPI", 0
   lvwProdutos.ColumnHeaders.Add , , "ClassificacaoFiscal", 0
   lvwProdutos.ColumnHeaders.Add , , "CFOP", 650
   lvwProdutos.ColumnHeaders.Add , , "ValorFrete", 0

   Centraliza frmNFe
'   MDISistema.StatusBar.Panels(1).text = "Emitir NFe"
End Sub

Private Sub cboCliente_Click()
   If cboCliente.ListIndex <> -1 Then
'      Set rsTemp = cnSistema.Execute("Select * From Clientes Where idCliente = " & cboCliente.ItemData(cboCliente.ListIndex))
      Set rsTemp = cnSistema.Execute("Select * From ClientesInfFiscais Where idCliente = " & cboCliente.ItemData(cboCliente.ListIndex))
'      Set rsUFs = cnSistema.Execute("Select * From UFs Where idUF = " & rsTemp!idUF)
      Set rsUFs = cnSistema.Execute("Select * From UFs Where Sigla = '" & rsTemp!SiglaUF & "'")
      If Not rsTemp.EOF And Not rsUFs.EOF Then
         lblDescricaoEndereco.Caption = Trim(rsTemp!Endereco) & ", " & Trim(rsTemp!Bairro) & Chr(13) & Trim(rsTemp!NomeMunicipio) & " - " & rsUFs!Sigla & Chr(13) & "CEP: " & rsTemp!CEP & " Fone: " & rsTemp!Telefone
      End If
   End If
End Sub

Private Sub mskCodigoBarra_GotFocus()
   mskCodigoBarra.SelStart = 0
   mskCodigoBarra.SelLength = Len(mskCodigoBarra.Text)
End Sub

Private Sub mskCodigoBarra_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then cboProduto.SetFocus
End Sub

Private Sub mskCodigoBarra_LostFocus()
Dim rsTemp As New ADODB.Recordset

   If Len(Trim(mskCodigoBarra.Text)) <> 0 Then
      Set rsTemp = cnSistema.Execute("SELECT idProduto, Descricao, ValorVenda FROM Produtos WHERE CodigoBarra='" & mskCodigoBarra.Text & "'")
      If rsTemp.EOF Then
         MsgBox "Código de barras não cadastrado", vbOKOnly, "Localizar"
         mskCodigoBarra.SetFocus
         mskCodigoBarra.SelStart = 0
         mskCodigoBarra.SelLength = Len(mskCodigoBarra.Text)
         Exit Sub
      Else
         mskCodigoBarra.Tag = rsTemp!idProduto
         cboProduto.Text = rsTemp!Descricao
         mskValorUnitario.Text = rsTemp!ValorVenda
         mskQuantidade.SetFocus
      End If
   End If
End Sub

Private Sub mskDescontoGeral_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then Sendkeys "{TAB}"
   If KeyAscii = 27 Then
      If mskDescontoGeral.Enabled Then
         mskDescontoGeral.SetFocus
         mskDescontoGeral.SelStart = 0
         mskDescontoGeral.SelLength = Len(mskDescontoGeral.Text)
      Else
         Sendkeys "{TAB}"
      End If
   End If
   If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 32 And KeyAscii <> 40 And KeyAscii <> 41 And KeyAscii <> 44 And KeyAscii <> 8 Then
      KeyAscii = 0
   End If
   If KeyAscii = 44 Then
      If InStr(mskDescontoGeral.Text, ",") <> 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim rsDados As ADODB.Recordset

   Select Case Button.Key
      Case "Novo"
         If Validar_Permissao(1, "mnu" & Mid(Me.Name, 4, Len(Me.Name))) Then
            Limpa_Campos
            Botoes 3, frmNFe
            Botoes_Extras 3
            Set rsDados = cnSistema.Execute("Select Numero From NFe Order by Numero DESC")
            If Not rsDados.EOF Then
               mskNumero.Text = rsDados!Numero + 1
            Else
               mskNumero.Text = 1
            End If
            Set rsDados = Nothing
            cboCliente.SetFocus
         End If
         
      Case "Gravar"
        Call Gravar
            
      Case "Excluir"
         If Validar_Permissao(3, "mnu" & Mid(Me.Name, 4, Len(Me.Name))) Then Call Excluir
         
      Case "Imprimir"
'''         Set rsDados = cnSistema.Execute("SELECT idGrupoAcesso FROM Usuarios WHERE idUsuario = " & idUser)
'''         If rsDados!idGrupoAcesso <> 1 Then
'''            Set rsDados = cnSistema.Execute("SELECT dbo.Usuarios.idUsuario, dbo.GrupoAcessoItens.Name FROM dbo.Usuarios INNER JOIN dbo.GrupoAcesso ON dbo.Usuarios.idGrupoAcesso = dbo.GrupoAcesso.idGrupoAcesso INNER JOIN dbo.GrupoAcessoItens ON dbo.GrupoAcesso.idGrupoAcesso = dbo.GrupoAcessoItens.idGrupoAcesso " & _
'''                                            "WHERE (dbo.Usuarios.idUsuario = " & idUser & ") AND (dbo.GrupoAcessoItens.Name = 'mnuRelNFe')")
'''            If Not rsDados.EOF() Then frmRelNFe.Show Else MsgBox "Usuário não autorizado para esta operação", vbInformation, "Acesso Negado"
'''         Else
'''            frmRelNFe.Show
'''         End If
'''         Set rsDados = Nothing
   End Select
End Sub

Private Function Verifica_Campos()
Dim strMensagem As String
Verifica_Campos = True

   If Not IsDate(mskDataEmissao.Text) Or Year(mskDataEmissao.Text) < 2000 Then strMensagem = strMensagem & "Data de Emissão" & Chr(13)
   If cboCliente.ListIndex = -1 Then strMensagem = strMensagem & "Cliente" & Chr(13)
   
   Set rsTemp = cnSistema.Execute("Select * From ClientesInfFiscais Where idCliente = " & cboCliente.ItemData(cboCliente.ListIndex))
   If Not rsTemp.EOF Then
      If Len(Trim(rsTemp!Endereco)) = 0 Then strMensagem = strMensagem & "Endereço do cliente não cadastrado" & Chr(13)
      If Len(Trim(rsTemp!Bairro)) = 0 Then strMensagem = strMensagem & "Bairro do cliente não cadastrado" & Chr(13)
      If Len(Trim(RemoveCaracteres(rsTemp!CN))) = 0 Then strMensagem = strMensagem & "CPF/CNPJ do cliente não cadastrado" & Chr(13)
   End If
   
   If Not strMensagem = Empty Then
      MsgBox "Verifique os seguintes campos:" & Chr(13) & strMensagem, vbExclamation + vbOKOnly, "Campos Obrigatórios"
      Verifica_Campos = False
      Exit Function
   End If
End Function

Private Sub Prencher_Campos()
Dim intContador As Integer
Dim rsNFeItens As New ADODB.Recordset
Dim rsClientes As New ADODB.Recordset
Dim rsCFOPs As New ADODB.Recordset
Dim rsTemp As New ADODB.Recordset
Dim Contador As Integer

   Set rsDados = cnSistema.Execute("SELECT * FROM NFe WHERE idNFe=" & cboDados.ItemData(cboDados.ListIndex))
   Set rsClientes = cnSistema.Execute("Select * From ClientesInfFiscais Where idCliente = " & rsDados!idCliente)
   
   txtChaveAcesso.Text = IIf(Trim(rsDados!ChaveNFe) = "" Or IsNull(rsDados!ChaveNFe), Empty, rsDados!ChaveNFe)
   txtChaveAcessoDevolucao.Text = IIf(Trim(rsDados!ChaveAcessoDevolucao) = "" Or IsNull(rsDados!ChaveAcessoDevolucao), Empty, rsDados!ChaveAcessoDevolucao)
   txtProtocolo.Text = IIf(Trim(rsDados!Protocolo) = "" Or IsNull(rsDados!Protocolo), Empty, rsDados!Protocolo)
   
   mskNumero.Text = IIf(Trim(rsDados!Numero) = "" Or IsNull(rsDados!Numero), Empty, rsDados!Numero)
   mskCupom.Text = IIf(Trim(rsDados!Cupom) = "" Or IsNull(rsDados!Cupom), Empty, rsDados!Cupom)
'   Set rsCFOPs = cnSistema.Execute("Select * From CFOPs Where idCFOP = " & IIf(Not IsNull(rsDados!idCFOP), rsDados!idCFOP, 0))
'   If Not rsCFOPs.EOF Then
'      mskCFOP.Text = rsCFOPs!CFOP
'   End If
   mskDataEmissao.Text = IIf(IsNull(rsDados!DataEmissao), "  /  /    ", Format(rsDados!DataEmissao, "dd/mm/yyyy"))
   mskDataVencimento.Text = IIf(IsNull(rsDados!DataVencimento), "  /  /    ", Format(rsDados!DataVencimento, "dd/mm/yyyy"))
   mskHora.Text = IIf(Trim(rsDados!Hora) = "" Or IsNull(rsDados!Hora), "  :  :  ", Format(rsDados!Hora, "HH:MM:SS"))
   mskBaseCalculoICMS.Text = IIf(Trim(rsDados!BaseCalculoICMS) = "" Or IsNull(rsDados!BaseCalculoICMS), Empty, rsDados!BaseCalculoICMS)
   mskValorICMS.Text = IIf(Trim(rsDados!ValorICMS) = "" Or IsNull(rsDados!ValorICMS), Empty, rsDados!ValorICMS)
'   mskValorFrete.Text = IIf(Trim(rsDados!ValorFrete) = "" Or IsNull(rsDados!ValorFrete), Empty, rsDados!ValorFrete)
   mskBaseICMSSubstituicao.Text = IIf(Trim(rsDados!BaseICMSSubstituicao) = "" Or IsNull(rsDados!BaseICMSSubstituicao), Empty, rsDados!BaseICMSSubstituicao)
   mskValorICMSSubstituicao.Text = IIf(Trim(rsDados!ValorICMSSubstituicao) = "" Or IsNull(rsDados!ValorICMSSubstituicao), Empty, rsDados!ValorICMSSubstituicao)
   mskOutrasDespesas.Text = IIf(Trim(rsDados!OutrasDespesas) = "" Or IsNull(rsDados!OutrasDespesas), Empty, rsDados!OutrasDespesas)
   txtDadosAdicionais.Text = IIf(Trim(rsDados!DadosAdicionais) = "" Or IsNull(rsDados!DadosAdicionais), Empty, rsDados!DadosAdicionais)
   mskDescontoGeral.Text = IIf(Trim(rsDados!DescontoGeral) = "" Or IsNull(rsDados!DescontoGeral), Empty, rsDados!DescontoGeral)
   mskBonificacao.Text = IIf(Trim(rsDados!Bonificacao) = "" Or IsNull(rsDados!Bonificacao), Empty, rsDados!Bonificacao)
   txtDocumento.Text = IIf(Trim(rsDados!Documento) = "" Or IsNull(rsDados!Documento), Empty, rsDados!Documento)
   txtObservacao.Text = IIf(Trim(rsDados!Observacao) = "" Or IsNull(rsDados!Observacao), Empty, rsDados!Observacao)

   mskPlaca.Text = IIf(Trim(rsDados!PlacaVeiculo) = "" Or IsNull(rsDados!PlacaVeiculo), "   -    ", rsDados!PlacaVeiculo)
   txtVolumeQuantidade.Text = IIf(Trim(rsDados!VolumeQuantidade) = "" Or IsNull(rsDados!VolumeQuantidade), Empty, rsDados!VolumeQuantidade)
   txtVolumeMarca.Text = IIf(Trim(rsDados!VolumeMarca) = "" Or IsNull(rsDados!VolumeMarca), Empty, rsDados!VolumeMarca)
   txtVolumeEspecie.Text = IIf(Trim(rsDados!VolumeEspecie) = "" Or IsNull(rsDados!VolumeEspecie), Empty, rsDados!VolumeEspecie)
   txtVolumeNumero.Text = IIf(Trim(rsDados!VolumeNumero) = "" Or IsNull(rsDados!VolumeNumero), Empty, rsDados!VolumeNumero)
   mskVolumePesoBruto.Text = IIf(Trim(rsDados!VolumePesoBruto) = "" Or IsNull(rsDados!VolumePesoBruto), Empty, rsDados!VolumePesoBruto)
   mskVolumePesoLiquido.Text = IIf(Trim(rsDados!VolumePesoLiquido) = "" Or IsNull(rsDados!VolumePesoLiquido), Empty, rsDados!VolumePesoLiquido)
   txtInformacoesCorpo.Text = IIf(Trim(rsDados!InformacoesCorpo) = "" Or IsNull(rsDados!InformacoesCorpo), Empty, rsDados!InformacoesCorpo)
   'cmbFreteConta.ListIndex = (rsDados!FreteConta - 1)
   If rsDados!UFCaminhao <> "" Then cmbUFPlaca.Text = rsDados!UFCaminhao
   If rsDados!Situacao = 3 Then
      lblCancelada.Visible = True
   Else
      lblCancelada.Visible = False
   End If
   
   For Contador = 0 To (cboNaturezaOperacao.ListCount - 1)
      If cboNaturezaOperacao.ItemData(Contador) = rsDados!idNaturezaOperacao Then
         cboNaturezaOperacao.ListIndex = Contador
         Exit For
      End If
   Next

   For Contador = 0 To (cmbFreteConta.ListCount - 1)
      If cmbFreteConta.ItemData(Contador) = rsDados!FreteConta Then
         cmbFreteConta.ListIndex = Contador
         Exit For
      End If
   Next

   For Contador = 0 To (cmbTransportador.ListCount - 1)
      If cmbTransportador.ItemData(Contador) = rsDados!idTransportador Then
         cmbTransportador.ListIndex = Contador
         Exit For
      End If
   Next

   For Contador = 0 To (cboFormaPagamento.ListCount - 1)
      If cboFormaPagamento.ItemData(Contador) = rsDados!idFormaPagamento Then
         cboFormaPagamento.ListIndex = Contador
         Exit For
      End If
   Next

'   cboFormaPagamento.ListIndex = rsdados!idFormaPagamento
   
   For Contador = 0 To (cboCliente.ListCount - 1)
       If cboCliente.ItemData(Contador) = rsDados!idCliente Then
          cboCliente.ListIndex = Contador
          Exit For
       End If
   Next

   If cboCliente.ListIndex <> -1 Then
      Set rsTemp = cnSistema.Execute("Select * From ClientesInfFiscais Where idCliente = " & cboCliente.ItemData(cboCliente.ListIndex))
      Set rsUFs = cnSistema.Execute("Select * From UFs Where Sigla = '" & rsTemp!SiglaUF & "'")
      If Not rsTemp.EOF And Not rsUFs.EOF Then
         lblDescricaoEndereco.Caption = Trim(rsTemp!Endereco) & ", " & Trim(rsTemp!Bairro) & Chr(13) & Trim(rsTemp!NomeMunicipio) & " - " & rsUFs!Sigla & Chr(13) & "CEP: " & rsTemp!CEP & " Fone: " & rsTemp!Telefone
      End If
   End If

''  Produtos
'
'   Dim cValorBruto As Currency
'   Dim cValorDesconto As Currency
'   Dim cValorBonificacao As Currency
'   Dim cValorLiquido As Currency
'
'   Set rsTemp = cnSistema.Execute("SELECT * FROM NFeItens WHERE NFeItens.idNFe = " & rsDados!idNFe)
'
'   lvwProdutos.ListItems.Clear
'   Do While Not rsTemp.EOF
'      Set rsProdutos = cnSistema.Execute("SELECT * FROM Produtos WHERE Produtos.idProduto = " & rsTemp!idProduto)
'      If Not rsProdutos.EOF Then
'         cValorBruto = (rsTemp!Quantidade * rsTemp!ValorUnitario)
'         cValorDesconto = (((rsTemp!Quantidade * rsTemp!ValorUnitario) * rsTemp!Desconto) / 100)
'         cValorBonificacao = (((cValorBruto - cValorDesconto) * rsDados!Bonificacao) / 100)
'         cValorLiquido = (cValorBruto - cValorDesconto - cValorBonificacao)
'
'         Set ItemList = lvwProdutos.ListItems.Add(, "R" & CStr(rsTemp!idProduto), rsProdutos!Codigo)
'         ItemList.SubItems(1) = Trim(rsProdutos!Descricao) & " " & Trim(rsTemp!DescricaoComplementar)
'         ItemList.SubItems(2) = Format(rsTemp!Quantidade, mskQuantidade.Format)
'         ItemList.SubItems(3) = Format(rsTemp!ValorUnitario, mskUnitario.Format)
'         ItemList.SubItems(4) = Format(rsTemp!Desconto, "##,##0.00")
'         ItemList.SubItems(5) = Format((rsTemp!Quantidade * rsTemp!ValorUnitario), "##,##0.00")
'         ItemList.SubItems(6) = Format(cValorLiquido, "##,##0.00")
'      End If
'
'      rsTemp.MoveNext
'   Loop

   curValorFrete = 0

'   chkNotaCancelada.value = IIf(rsDados!Cancelada, 1, 0)
   lvwProdutos.ListItems.Clear
   Set rsNFeItens = cnSistema.Execute("SELECT dbo.Produtos.idProduto, CASE dbo.Produtos.EmbalagemMultipla WHEN 1 THEN dbo.Produtos.DescricaoEmbalagemMultipla ELSE dbo.Produtos.Descricao END AS Descricao, " & _
                                      "dbo.NFeItens.ICMS, dbo.NFeItens.BaseReduzida, dbo.NFeItens.DescricaoComplementar, dbo.NFeItens.Unidade, dbo.NFeItens.idSituacaoTributaria, dbo.NFeItens.DiscriminacaoProduto, dbo.NFeItens.IPI, dbo.NFeItens.BaseReduzidaIPI, dbo.NFeItens.ClassificacaoFiscal, dbo.NFeItens.CFOP, " & _
                                      "dbo.NFeItens.Quantidade , dbo.NFeItens.ValorUnitario, dbo.NFeItens.Desconto, dbo.NFeItens.ValorFrete FROM dbo.Produtos INNER JOIN dbo.NFeItens ON dbo.Produtos.idProduto = dbo.NFeItens.idProduto " & _
                                      "WHERE dbo.NFeItens.idNFe = " & rsDados!idNFe)
   Do While Not rsNFeItens.EOF
      Set ItemList = lvwProdutos.ListItems.Add(, "R" & rsNFeItens!idProduto, rsNFeItens!Descricao)
      ItemList.SubItems(1) = rsNFeItens!Quantidade
      ItemList.SubItems(2) = Format(rsNFeItens!ValorUnitario, "R$ #,###,##0.00")
      ItemList.SubItems(3) = Format(rsNFeItens!ValorUnitario * rsNFeItens!Quantidade, "R$ #,###,##0.00")
      ItemList.SubItems(4) = Round((rsNFeItens!Desconto * 100) / rsNFeItens!ValorUnitario, 2)
      ItemList.SubItems(5) = Format(Round((rsNFeItens!ValorUnitario * rsNFeItens!Quantidade) * (1 - (rsNFeItens!Desconto / 100)), 2), "R$ #,###,##0.00")
'      ItemList.SubItems(7) = IIf(IsNull(rsNFeItens!Financeiro), "NULL", "'" & Format(rsNFeItens!Financeiro, "mm/dd/yyyy hh:mm:ss") & "'")
'      ItemList.SubItems(8) = IIf(rsNFeItens!Caixa, 1, 0)

      ItemList.SubItems(9) = rsNFeItens!ICMS
      ItemList.SubItems(10) = rsNFeItens!BaseReduzida
      ItemList.SubItems(11) = rsNFeItens!DescricaoComplementar
      ItemList.SubItems(12) = rsNFeItens!Unidade
      ItemList.SubItems(13) = rsNFeItens!idSituacaoTributaria
      ItemList.SubItems(14) = rsNFeItens!DiscriminacaoProduto
      ItemList.SubItems(15) = rsNFeItens!IPI
      ItemList.SubItems(16) = rsNFeItens!BaseReduzidaIPI
      ItemList.SubItems(17) = rsNFeItens!ClassificacaoFiscal
      ItemList.SubItems(18) = rsNFeItens!CFOP

      curValorTotal = curValorTotal + Round((rsNFeItens!ValorUnitario * rsNFeItens!Quantidade), 2)
      curValorDesconto = curValorDesconto + (rsNFeItens!Desconto * rsNFeItens!Quantidade)
      curValorFrete = curValorFrete + IIf(Not IsNull(rsNFeItens!ValorFrete), rsNFeItens!ValorFrete, 0)
      
'      lblValorTotal.Caption = "Valor Total: " & Format(curValorTotal, "R$ #,###,##0.00")
'      lblValorPagar.Caption = "Valor a pagar: " & Format(curValorTotal - curValorDesconto - rsDados!Arredondamento, "R$ #,###,##0.00")
'      lblValorDesconto.Caption = "Desconto: " & Format(curValorDesconto, "R$ #,###,##0.00")
      rsNFeItens.MoveNext
   Loop
   rsNFeItens.Close
   
   ' Total da Nota
   Set rsTotalNFe = cnSistema.Execute("SELECT dbo.NFe.idNFe, dbo.NFe.idNFe, SUM(ROUND(dbo.NFeItens.Quantidade * dbo.NFeItens.ValorUnitario, 2))- SUM(ROUND(dbo.NFeItens.Quantidade * dbo.NFeItens.ValorUnitario * dbo.NFeItens.Desconto / 100, 2))- SUM(ROUND((dbo.NFeItens.Quantidade * dbo.NFeItens.ValorUnitario - dbo.NFeItens.Quantidade * dbo.NFeItens.ValorUnitario * dbo.NFeItens.Desconto / 100) * dbo.NFe.Bonificacao / 100, 2)) AS Total " & _
                                      "FROM dbo.NFe INNER JOIN dbo.NFeItens ON dbo.NFe.idNFe = dbo.NFeItens.idNFe INNER JOIN dbo.Produtos ON dbo.NFeItens.idProduto = dbo.Produtos.idProduto " & _
                                      "Where Numero = " & rsDados!Numero & _
                                      "GROUP BY dbo.NFe.idNFe, dbo.NFe.Numero")
   
   If Not rsTotalNFe.EOF Then
      mskValorFrete.Text = Format(curValorFrete, "##,##0.00") & " "
      lblTotalProdutos.Caption = Format(rsTotalNFe!Total, "##,##0.00") & " "
      lblTotalNota.Caption = Format(rsTotalNFe!Total + curValorFrete, "##,##0.00") & " "
   Else
      lblTotalProdutos.Caption = Format(0, "##,##0.00") & " "
      lblTotalNota.Caption = Format(0, "##,##0.00") & " "
   End If
   
''   Me.Refresh
''   Set rsDados = Nothing
''   Botoes 2, frmNFe
''   Botoes_Extras 2

'  Boletos
   Set rsTemp = cnSistema.Execute("SELECT * FROM NFeBoletos WHERE NFeBoletos.idNFe = " & rsDados!idNFe)

   lvwBoletos.ListItems.Clear
   Do While Not rsTemp.EOF
      Set ItemList = lvwBoletos.ListItems.Add(, "R" & rsTemp!Numero, rsTemp!Numero)
      ItemList.SubItems(1) = rsTemp!Vencimento
      ItemList.SubItems(2) = Format(rsTemp!Valor, "##,##0.00")
   
      rsTemp.MoveNext
   Loop


'''''  Total da Nota
''''   Set rsTemp = cnSistema.Execute("Select * From TotalNFe Where Numero = " & mskNumero.Text)
''''   If Not rsTemp.EOF Then
''''      lbltotalprodutos.caption = Format(rsTemp!Total, "###,##0.00")
''''      mskValorTotalNota.Text = Format(rsTemp!Total + rsdados!ValorFrete + rsdados!OutrasDespesas, "###,##0.00")
''''      mskBaseCalculoICMS.Text = Format(IIf(rsTemp!ValorICMS > 0, rsTemp!BaseCalculo, 0), "###,##0.00")
''''      mskValorICMS.Text = Format(rsTemp!ValorICMS, "###,##0.00")
''''   Else
''''      lbltotalprodutos.caption = Format(0, "###,##0.00")
''''      mskValorTotalNota.Text = Format(0, "###,##0.00")
''''      mskBaseCalculoICMS.Text = Format(0, "###,##0.00")
''''      mskValorICMS.Text = Format(0, "###,##0.00")
''''   End If
   
''   dblValorTotal = 0
''   mskDataEmissao.Text = rsDados!DataEmissao
''   chkCancelada.value = IIf(rsDados!Cancelada, 1, 0)
''   mskNumero.Text = rsDados!Numero
''   txtObservacao.Text = rsDados!Observacao
''   mskDescontoGeral.Text = rsDados!DescontoGeral
''   mskCodigoBarra.Text = Space(13)
''   mskCodigoBarra.Tag = Empty
''   cboProduto.Text = Empty
''   cboProduto.ListIndex = -1
''   mskQuantidade.Text = Space(6)
''   mskValorUnitario.Text = 0
''   'Carregar Fornecedor
''   cboCliente.Clear
''   Set rsClientes = cnSistema.Execute("SELECT idFornecedor, RazaoSocial From ClientesInfFiscais WHERE idFornecedor = " & rsDados!idFornecedor)
''   cboCliente.AddItem rsClientes!RazaoSocial
''   cboCliente.ItemData(cboCliente.NewIndex) = rsClientes!idFornecedor
''   cboCliente.ListIndex = 0
''   Set rsClientes = Nothing
''   lvwProdutos.ListItems.Clear
''   Set rsdadosItens = cnSistema.Execute("SELECT Produtos.CodigoBarra, Produtos.Descricao, NFeItens.idProduto, NFeItens.Quantidade, NFeItens.ValorUnitario, NFeItens.Desconto " & _
''                                           "FROM Produtos INNER JOIN NFeItens ON Produtos.idProduto = NFeItens.idProduto " & _
''                                           "WHERE NFeItens.idNFe = " & cboDados.ItemData(cboDados.ListIndex))
''   Do While Not rsnfeItens.EOF
''      curDesconto = rsNFeItens!Desconto
''      Set ItemList = lvwProdutos.ListItems.Add(, "R" & rsNFeItens!idProduto, rsNFeItens!CodigoBarra)
''      ItemList.SubItems(1) = rsNFeItens!Descricao
''      ItemList.SubItems(2) = rsNFeItens!Quantidade
''      ItemList.SubItems(3) = Format(rsNFeItens!ValorUnitario, "R$ ###,##0.00")
''      ItemList.SubItems(4) = Format(rsNFeItens!Desconto, "##0.00")
''      ItemList.SubItems(5) = Format(Round(rsNFeItens!ValorUnitario * (1 - (curDesconto / 100)), 2) * rsNFeItens!Quantidade, "R$ #,###,##0.00")
''      dblValorTotal = dblValorTotal + (Round(rsNFeItens!ValorUnitario * (1 - (curDesconto / 100)), 2) * rsNFeItens!Quantidade)
''
''      rsNFeItens.MoveNext
''   Loop
''   rsNFeItens.Close
''   lblValorTotal.Caption = "Valor Total: " & Format(dblValorTotal, "R$ #,###,##0.00") & "  Desconto Geral: " & Format(rsDados!DescontoGeral, "R$ #,###,##0.00") & "  Valor Liquido: " & Format(dblValorTotal - rsDados!DescontoGeral, "R$ #,###,##0.00")
   Set rsDados = Nothing
   Botoes 2, frmNFe
   Botoes_Extras 2
End Sub

Private Sub Limpa_Campos()
   Set rsDados = cnSistema.Execute("SELECT COUNT(*) AS Registros FROM NFe")
   lblRegistros.Caption = IIf(rsDados!Registros > 1, rsDados!Registros & " registros", IIf(rsDados!Registros = 1, "1 registro", "Nenhum registro"))
   Set rsDados = Nothing
   cboDados.Clear
   Botoes 1, frmNFe
   Botoes_Extras 1

   txtChaveAcesso.Text = Empty
   txtChaveAcessoDevolucao.Text = Empty
   txtProtocolo.Text = Empty
   mskCupom.Text = Empty
   cboNaturezaOperacao.ListIndex = -1
   mskCFOP.Text = " .   "
   mskDataEmissao.Text = Date
   mskDataVencimento.Text = Date
   mskHora.Text = Time
   cboCliente.ListIndex = -1
   mskBaseCalculoICMS.Text = Empty
   mskValorICMS.Text = Empty
   mskValorFrete.Text = 0
   lblTotalProdutos.Caption = Empty
   lblTotalNota.Caption = Empty
   mskBaseICMSSubstituicao.Text = Empty
   mskValorICMSSubstituicao.Text = Empty
   mskOutrasDespesas.Text = Empty
   txtDadosAdicionais.Text = LerArquivoINI("NFe", "DadosAdicionais", App.Path & "\System.ini")
   txtInformacoesCorpo.Text = LerArquivoINI("NFe", "InformacoesCorpo", App.Path & "\System.ini")
   cboFormaPagamento.ListIndex = -1
   mskDescontoGeral.Text = Empty
   mskBonificacao.Text = Empty
   txtDocumento.Text = Empty
   txtObservacao.Text = Empty
   
   lblCancelada.Visible = False
   
   cmbTransportador.ListIndex = -1
   cmbFreteConta.ListIndex = -1
   mskPlaca.Text = "   -    "
   cmbUFPlaca.ListIndex = -1
   txtVolumeQuantidade.Text = Empty
   txtVolumeMarca.Text = Empty
   txtVolumeEspecie.Text = Empty
   txtVolumeNumero.Text = Empty
   mskVolumePesoBruto.Text = Empty
   mskVolumePesoLiquido.Text = Empty

   lvwProdutos.ListItems.Clear
   lvwBoletos.ListItems.Clear
   
   lblDescricaoEndereco.Caption = Empty

'''   lblNFeNumero.Caption = Empty
'''   mskData.Text = Date
'''   chkCancelada.value = 0
'''   mskNotaNumero.Text = Space(7)
'''   mskDescontoGeral.Text = 0
'''   txtDesconto.Text = Space(5) & "0"
'''   txtDesconto.BackColor = &H80000005
'''   mskAlteraDesconto.BackColor = &H80000005
'''   txtObservacao.Text = Empty
'''   cboCliente.Clear
'''   lblDadosFornecedor.Caption = Empty
'''   lblDadosProduto.Caption = Empty
'''   lblTotalItens.Caption = "0" & " Itens"
'''   mskCodigoBarra.Text = Space(13)
'''   mskCodigoBarra.Tag = Empty
'''   cboProduto.Clear
'''   mskQuantidade.Text = "1" & Space(5)
'''   mskValorUnitario.Text = 0
'''   lvwProdutos.ListItems.Clear
'''   dblValorTotal = 0
''''   lblValorTotal.Caption = "Valor Total: " & Format(dblValorTotal, "R$ #,###,##0.00")
'''   lblValorTotal.Caption = "Valor Total: " & Format(dblValorTotal, "R$ #,###,##0.00") & "  Desconto Geral: " & Format(mskDescontoGeral.Text, "R$ #,###,##0.00") & "  Valor Liquido: " & Format(dblValorTotal - mskDescontoGeral.Text, "R$ #,###,##0.00")
   
   
   Carrega_Combos_Extras
End Sub

Private Sub Excluir()
On Error GoTo ErroIntegridade

   If cboDados.ListIndex <> -1 Then
      If MsgBox("Confirma Excluir o registro atual? ", vbYesNo + vbInformation, "Excluir") = vbYes Then
         Atividade "Exclusão: " & Trim(SQLCheck(mskDataEmissao.Text)), Me.Caption
         cnSistema.Execute "DELETE FROM NFeItens WHERE idNFe=" & cboDados.ItemData(cboDados.ListIndex)
         cnSistema.Execute "DELETE FROM NFe WHERE idNFe=" & cboDados.ItemData(cboDados.ListIndex)
         Limpa_Campos
      End If
   Else
      MsgBox "Selecione primeiro um registro", vbInformation + vbOKOnly, "Excluir"
   End If
   
On Error GoTo 0
Exit Sub
ErroIntegridade:
   If Err.Number = 0 Then
      ' Operação Ok
   ElseIf Err.Number = -2147217873 Then
      MsgBox "Não é possível Excluir este Registro" & Chr(13) & "Existe lançamentos relacionados com este Registro", vbInformation + vbOKOnly, "Excluir"
      Exit Sub
   Else
      MsgBox "Verificar: " & Err.Number & Chr(13) & Err.Description, vbExclamation, "Excluir"
      Exit Sub
   End If
End Sub

Private Sub Gravar()
On Error GoTo Erro
Dim rsSistema As New ADODB.Recordset
Dim rsInclusao As New ADODB.Recordset
Dim CNPJ_CPF As String
Dim dDataEmissao As String
Dim dDataSaida As String
Dim dHora As String
Dim strCFOP As Integer
Dim iTransportador As Integer, iFreteConta As Integer

   localErro = "Transportador"
   If cmbTransportador.ListIndex = -1 Then
      iTransportador = 0
   Else
      iTransportador = cmbTransportador.ItemData(cmbTransportador.ListIndex)
   End If

'   If cmbFreteConta.ListIndex = -1 Then
'      iFreteConta = 0
'   Else
'      iFreteConta = IIf(cmbFreteConta.ListIndex = 0, 1, 2)
'   End If

   localErro = "Frete Conta"
   If cmbFreteConta.ListIndex = -1 Then
      iFreteConta = 9
   Else
      iFreteConta = Mid(cmbFreteConta.Text, 1, 1)
   End If

   localErro = "Hora"
   If mskHora.Text = "  :  :  " Then
      dHora = "00:00:00"
   Else
      dHora = mskHora.Text
   End If
   
   localErro = "Data da Emissão"
   dDataEmissao = mskDataEmissao.Text & " " & dHora
   
   localErro = "Data da Saida"
   dDataSaida = mskDataVencimento.Text & " " & dHora
   
   If IsDate(dDataSaida) < IsDate(dDataEmissao) Then
      dDataSaida = dDataEmissao
   End If
   
'   localErro = "CFOP"
'   Set rsCFOPs = cnSistema.Execute("Select * From CFOPs Where CFOP = '" & mskCFOP.Text & "'")
'   If Not rsCFOPs.EOF Then strCFOP = rsCFOPs!idCFOP

   localErro = "Gravação"
   If cboDados.ListIndex = -1 Then 'Inclusão
      localErro = "Gravar Inclusão"
      If Validar_Permissao(1, "mnu" & Mid(Me.Name, 4, Len(Me.Name))) Then
         If Not Verifica_Campos() Then Exit Sub
         If MsgBox("Confirma Incluir o registro atual", vbYesNo + vbQuestion, "Inclusão") = vbYes Then
         
            cnSistema.Execute "Insert Into NFe (Numero,Cupom,idCliente,idNaturezaOperacao,idCFOP,DadosAdicionais,DataEmissao,DataCaixa,DataVencimento,Hora,BaseCalculoICMS,ValorICMS,ValorFrete,ValorTotalProdutos,BaseICMSSubstituicao,ValorICMSSubstituicao,OutrasDespesas,ValorTotalNota,idTransportador,FreteConta,PlacaVeiculo,UFCaminhao,VolumeQuantidade,VolumeMarca,VolumeEspecie,VolumeNumero,VolumePesoBruto,VolumePesoLiquido,InformacoesCorpo,idFormaPagamento,DescontoGeral,Bonificacao,Documento,Observacao,Situacao,ChaveAcessoDevolucao) " & _
                              "Values (" & Val(mskNumero.Text) & "," & Val(mskCupom.Text) & "," & cboCliente.ItemData(cboCliente.ListIndex) & "," & cboNaturezaOperacao.ItemData(cboNaturezaOperacao.ListIndex) & "," & strCFOP & ",'" & txtDadosAdicionais.Text & "','" & Format(dDataEmissao, "mm/dd/yyyy hh:mm:ss") & "','" & Format(dDataEmissao, "mm/dd/yyyy hh:mm:ss") & "','" & Format(dDataSaida, "mm/dd/yyyy hh:mm:ss") & "','" & dHora & "','" & CStrValor(mskBaseCalculoICMS.ClipText) & "','" & CStrValor(mskValorICMS.ClipText) & "'," & _
                                      "'" & CStrValor(mskValorFrete.ClipText) & "','" & CStrValor(lblTotalProdutos.Caption) & "','" & CStrValor(mskBaseICMSSubstituicao.ClipText) & "','" & CStrValor(mskValorICMSSubstituicao.ClipText) & "','" & CStrValor(mskOutrasDespesas.ClipText) & "'," & _
                                      "'" & CStrValor(lblTotalNota.Caption) & "'," & iTransportador & "," & iFreteConta & ",'" & UCase(mskPlaca.Text) & "','" & cmbUFPlaca.Text & "','" & txtVolumeQuantidade.Text & "','" & txtVolumeMarca.Text & "','" & txtVolumeEspecie.Text & "','" & txtVolumeNumero.Text & "','" & CStrValor(mskVolumePesoBruto.ClipText) & "','" & CStrValor(mskVolumePesoLiquido.ClipText) & "','" & txtInformacoesCorpo.Text & "'" & _
                                      "," & cboFormaPagamento.ItemData(cboFormaPagamento.ListIndex) & ",'" & CStrValor(mskDescontoGeral.ClipText) & "','" & CStrValor(mskBonificacao.ClipText) & "','" & txtDocumento.Text & "','" & SQLCheck(txtObservacao.Text) & "',0,'" & SQLCheck(txtChaveAcessoDevolucao.Text) & "')"
         
''            cnSistema.Execute "Insert Into NFe (Numero,Cupom,idCliente,idNaturezaOperacao,idCFOP,DadosAdicionais,DataEmissao,DataCaixa,DataVencimento,Hora,BaseCalculoICMS,ValorICMS,ValorFrete,ValorTotalProdutos,BaseICMSSubstituicao,ValorICMSSubstituicao,OutrasDespesas,ValorTotalNota,idTransportador,FreteConta,PlacaVeiculo,UFCaminhao,VolumeQuantidade,VolumeMarca,VolumeEspecie,VolumeNumero,VolumePesoBruto,VolumePesoLiquido,InformacoesCorpo,idFormaPagamento,DescontoGeral,Bonificacao,Documento,Observacao,Situacao) " & _
''                              "Values (" & Val(mskNumero.Text) & "," & Val(mskCupom.Text) & "," & cboCliente.ItemData(cboCliente.ListIndex) & "," & cboNaturezaOperacao.ItemData(cboNaturezaOperacao.ListIndex) & "," & strCFOP & ",'" & txtDadosAdicionais.Text & "','" & mskDataEmissao.Text & "','" & mskDataEmissao.Text & "','" & mskDataVencimento.Text & "','" & dHora & "','" & CStrValor(Val(mskBaseCalculoICMS.ClipText)) & "','" & CStrValor(Val(mskValorICMS.ClipText)) & "'," & _
''                                      "'" & CStrValor(Val(mskValorFrete.ClipText)) & "','" & CStrValor(Val(lblTotalProdutos.Caption)) & "','" & CStrValor(Val(mskBaseICMSSubstituicao.ClipText)) & "','" & CStrValor(Val(mskValorICMSSubstituicao.ClipText)) & "','" & CStrValor(Val(mskOutrasDespesas.ClipText)) & "'," & _
''                                      "'" & CStrValor(Val(mskValorTotalNota.ClipText)) & "'," & iTransportador & "," & iFreteConta & ",'" & UCase(mskPlaca.Text) & "','" & cmbUFPlaca.Text & "','" & txtVolumeQuantidade.Text & "','" & txtVolumeMarca.Text & "','" & txtVolumeEspecie.Text & "','" & txtVolumeNumero.Text & "','" & CStrValor(Val(mskVolumePesoBruto.ClipText)) & "','" & CStrValor(Val(mskVolumePesoLiquido.ClipText)) & "','" & txtInformacoesCorpo.Text & "'" & _
''                                      "," & cboFormaPagamento.ItemData(cboFormaPagamento.ListIndex) & ",'" & CStrValor(Val(mskDescontoGeral.ClipText)) & "','" & CStrValor(Val(mskBonificacao.ClipText)) & "','" & txtDocumento.Text & "','" & SQLCheck(txtObservacao.Text) & "',0)"
         
'' Format(Now, "mm/dd/yyyy hh:mm:ss")
         
''            cnSistema.Execute "INSERT INTO Entradas (Data,NotaNumero,Observacao,idFornecedor,DescontoGeral,Cancelada) " & _
''                              "VALUES ('" & Format(mskData.Text, "mm/dd/yyyy") & "','" & mskNotaNumero.Text & "','" & SQLCheck(txtObservacao.Text) & "'," & cboFornecedor.ItemData(cboFornecedor.ListIndex) & "," & Replace(mskDescontoGeral.Text, ",", ".") & "," & chkCancelada.value & ")"
                              
            Set rsInclusao = cnSistema.Execute("SELECT IDENT_CURRENT('NFe') AS 'Identity'")
            localErro = "Gravar Alteração Produtos"
            Gravar_Produtos rsInclusao!Identity

            Set rsInclusao = Nothing
            Set rsSistema = Nothing
'            localErro = "Gravar Alteração Atividade"
'            Atividade "Inclusão: " & Trim(SQLCheck(mskDataEmissao.Text)), Me.Caption
         End If
      End If
   Else 'Alteracão
      localErro = "Gravar Alteração"
'      If Validar_Permissao(2, "mnu" & Mid(Me.Name, 4, Len(Me.Name))) Then
         If Not Verifica_Campos() Then Exit Sub
         If MsgBox("Confirma Alterar o registro atual", vbYesNo + vbQuestion, "Alteração") = vbYes Then
         
            cnSistema.Execute "Update NFe set " & _
                  "Numero = " & Val(mskNumero.Text) & ", " & "Cupom = " & Val(mskCupom.Text) & ", " & _
                  "idCliente = " & cboCliente.ItemData(cboCliente.ListIndex) & ", " & "idNaturezaOperacao = " & cboNaturezaOperacao.ItemData(cboNaturezaOperacao.ListIndex) & ", " & _
                  "idCFOP = " & strCFOP & ", " & _
                  "DataEmissao = '" & Format(mskDataEmissao.Text, "mm-dd-yyyy hh:mm:ss") & "', " & "DataCaixa = '" & Format(mskDataEmissao.Text, "mm-dd-yyyy hh:mm:ss") & "', " & "DataVencimento = '" & Format(mskDataVencimento.Text, "mm-dd-yyyy hh:mm:ss") & "', " & "Hora = '" & dHora & "', " & "DadosAdicionais = '" & SQLCheck(txtDadosAdicionais.Text) & "', " & _
                  "BaseCalculoICMS = '" & CStrValor(mskBaseCalculoICMS.ClipText) & "', " & "ValorICMS = '" & CStrValor(mskValorICMS.ClipText) & "', " & _
                  "ValorFrete = '" & CStrValor(mskValorFrete.ClipText) & "', " & "ValorTotalProdutos = '" & CStrValor(lblTotalProdutos.Caption) & "', " & _
                  "BaseICMSSubstituicao = '" & CStrValor(mskBaseICMSSubstituicao.ClipText) & "', " & _
                  "ValorICMSSubstituicao = '" & CStrValor(mskValorICMSSubstituicao.ClipText) & "', " & _
                  "OutrasDespesas = '" & CStrValor(mskOutrasDespesas.ClipText) & "', " & _
                  "ValorTotalNota = '" & CStrValor(lblTotalNota.Caption) & "', " & _
                  "idTransportador = " & iTransportador & ", " & "FreteConta = " & iFreteConta & ", " & _
                  "PlacaVeiculo = '" & UCase(mskPlaca.Text) & "', " & "UFCaminhao = '" & cmbUFPlaca.Text & "', " & _
                  "VolumeQuantidade = '" & txtVolumeQuantidade.Text & "', " & _
                  "VolumeMarca = '" & txtVolumeMarca.Text & "', " & "VolumeNumero = '" & txtVolumeNumero.Text & "', " & "VolumeEspecie = '" & txtVolumeEspecie.Text & "', " & _
                  "VolumePesoBruto = '" & CStrValor(mskVolumePesoBruto.ClipText) & "', " & _
                  "VolumePesoLiquido = '" & CStrValor(mskVolumePesoLiquido.ClipText) & "', " & _
                  "InformacoesCorpo = '" & txtInformacoesCorpo.Text & "', " & _
                  "idFormaPagamento = " & cboFormaPagamento.ItemData(cboFormaPagamento.ListIndex) & ", " & _
                  "DescontoGeral = '" & CStrValor(mskDescontoGeral.ClipText) & "', " & _
                  "Bonificacao = '" & CStrValor(mskBonificacao.ClipText) & "', " & _
                  "Documento = '" & txtDocumento.Text & "', " & _
                  "Observacao = '" & SQLCheck(txtObservacao.Text) & "', " & _
                  "ChaveAcessoDevolucao = '" & SQLCheck(txtChaveAcessoDevolucao.Text) & "' " & _
                  "Where idNFe = " & cboDados.ItemData(cboDados.ListIndex)
         
'            cnSistema.Execute "Update NFe set " & _
'                  "Numero = " & Val(mskNumero.Text) & ", " & "Cupom = " & Val(mskCupom.Text) & ", " & _
'                  "idCliente = " & cboCliente.ItemData(cboCliente.ListIndex) & ", " & "idNaturezaOperacao = " & cboNaturezaOperacao.ItemData(cboNaturezaOperacao.ListIndex) & ", " & _
'                  "idCFOP = " & strCFOP & ", " & _
'                  "DataEmissao = '" & Format(mskDataEmissao.Text, "mm/dd/yyyy hh:mm:ss") & "', " & "DataCaixa = '" & Format(mskDataEmissao.Text, "mm/dd/yyyy hh:mm:ss") & "', " & "DataVencimento = '" & Format(mskDataVencimento.Text, "mm/dd/yyyy hh:mm:ss") & "', " & "Hora = '" & dHora & "', " & "DadosAdicionais = '" & SQLCheck(txtDadosAdicionais.Text) & "', " & _
'                  "BaseCalculoICMS = '" & CStrValor(mskBaseCalculoICMS.ClipText) & "', " & "ValorICMS = '" & CStrValor(mskValorICMS.ClipText) & "', " & _
'                  "ValorFrete = '" & CStrValor(mskValorFrete.ClipText) & "', " & "ValorTotalProdutos = '" & CStrValor(lblTotalProdutos.Caption) & "', " & _
'                  "BaseICMSSubstituicao = '" & CStrValor(mskBaseICMSSubstituicao.ClipText) & "', " & _
'                  "ValorICMSSubstituicao = '" & CStrValor(mskValorICMSSubstituicao.ClipText) & "', " & _
'                  "OutrasDespesas = '" & CStrValor(mskOutrasDespesas.ClipText) & "', " & _
'                  "ValorTotalNota = '" & CStrValor(lblTotalNota.Caption) & "', " & _
'                  "idTransportador = " & iTransportador & ", " & "FreteConta = " & iFreteConta & ", " & _
'                  "PlacaVeiculo = '" & UCase(mskPlaca.Text) & "', " & "UFCaminhao = '" & cmbUFPlaca.Text & "', " & _
'                  "VolumeQuantidade = '" & txtVolumeQuantidade.Text & "', " & _
'                  "VolumeMarca = '" & txtVolumeMarca.Text & "', " & "VolumeNumero = '" & txtVolumeNumero.Text & "', " & "VolumeEspecie = '" & txtVolumeEspecie.Text & "', " & _
'                  "VolumePesoBruto = '" & CStrValor(mskVolumePesoBruto.ClipText) & "', " & _
'                  "VolumePesoLiquido = '" & CStrValor(mskVolumePesoLiquido.ClipText) & "', " & _
'                  "InformacoesCorpo = '" & txtInformacoesCorpo.Text & "', " & _
'                  "idFormaPagamento = " & cboFormaPagamento.ItemData(cboFormaPagamento.ListIndex) & ", " & _
'                  "DescontoGeral = '" & CStrValor(mskDescontoGeral.ClipText) & "', " & _
'                  "Bonificacao = '" & CStrValor(mskBonificacao.ClipText) & "', " & _
'                  "Documento = '" & txtDocumento.Text & "', " & _
'                  "Observacao = '" & SQLCheck(txtObservacao.Text) & "', " & _
'                  "ChaveAcessoDevolucao = '" & SQLCheck(txtChaveAcessoDevolucao.Text) & "' " & _
'                  "Where idNFe = " & cboDados.ItemData(cboDados.ListIndex)
         
         
         
         
         
         
         
         
         
         
         
         
         
         
         
         
''            cnSistema.Execute "UPDATE Entradas SET " & _
''                              "Data = '" & Format(mskData.Text, "mm/dd/yyyy") & "', " & _
''                              "Cancelada = " & chkCancelada.value & ", " & _
''                              "NotaNumero = '" & mskNotaNumero.Text & "', " & _
''                              "Observacao = '" & SQLCheck(txtObservacao.Text) & "', " & _
''                              "DescontoGeral = " & Replace(mskDescontoGeral.Text, ",", ".") & ", " & _
''                              "idFornecedor = " & cboFornecedor.ItemData(cboFornecedor.ListIndex) & " " & _
''                              "WHERE idEntrada = " & cboDados.ItemData(cboDados.ListIndex)
                              
            localErro = "Gravar Alteração Produtos"
            Gravar_Produtos cboDados.ItemData(cboDados.ListIndex)
'            localErro = "Gravar Alteração Atividade"
'            Atividade "Alterar: " & Trim(SQLCheck(mskDataEmissao.Text)), Me.Caption
         End If
'      End If
   End If
   Limpa_Campos
   cboDados.SetFocus
   Exit Sub
Erro:
   MsgBox Err.Number & " - " & localErro & " - " & Err.Description & " - " & TypeName(Me)
End Sub

Private Sub cboDados_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then Carrega_Combos
End Sub

Private Sub cboDados_Click()
   Prencher_Campos
End Sub

Private Sub Carrega_Combos()
Dim rsCombo As New ADODB.Recordset
Dim ContadorRegistros As Boolean

   ContadorRegistros = False
   If cboDados.Text = Empty Then
      Set rsCombo = cnSistema.Execute("SELECT COUNT(idNFe) AS Registros FROM NFe")
      ContadorRegistros = True
   Else
      If IsDate(cboDados.Text) Then
''         Set rsCombo = cnSistema.Execute("SELECT COUNT(idNFe) AS Registros FROM NFe WHERE (DataEmissao >= '" & Format(cboDados.Text, "mm/dd/yyyy") & " 00:00:00') AND (DataEmissao <= '" & Format(cboDados.Text, "mm/dd/yyyy") & " 23:59:59') ")
         Set rsCombo = cnSistema.Execute("SELECT COUNT(idNFe) AS Registros FROM NFe WHERE (DataEmissao >= '" & Format(cboDados.Text, "mm/dd/yyyy") & " 00:00:00') AND (DataEmissao <= '" & Format(cboDados.Text, "mm/dd/yyyy") & " 23:59:59') ")
         ContadorRegistros = True
      Else
         If IsNumeric(cboDados.Text) Then
            Set rsCombo = cnSistema.Execute("SELECT COUNT(idNFe) AS Registros FROM NFe WHERE (idNFe = " & cboDados.Text & ")")
            ContadorRegistros = True
         Else
            Set rsCombo = cnSistema.Execute("SELECT COUNT(dbo.NFe.idNFe) AS Registros FROM dbo.NFe LEFT OUTER JOIN dbo.ClientesInfFiscais ON dbo.NFe.idCliente = dbo.ClientesInfFiscais.idCliente " & _
                                            "WHERE dbo.ClientesInfFiscais.Nome LIKE '%" & cboDados.Text & "%'")
            ContadorRegistros = True
         End If
      End If
   End If

   If ContadorRegistros Then
      If rsCombo!Registros > 5000 Then
         MsgBox "A consulta possui mais de 5.000 resultados" & vbCrLf & "Redefina sua pesquisa", vbOKOnly + vbInformation, "Pesquisa"
         Exit Sub
      End If
   End If
   
   Screen.MousePointer = vbHourglass
   If cboDados.Text = Empty Then
      Set rsCombo = cnSistema.Execute("SELECT dbo.NFe.idNFe, dbo.NFe.Numero, dbo.NFe.DataEmissao, dbo.ClientesInfFiscais.Nome FROM dbo.NFe INNER JOIN dbo.ClientesInfFiscais ON dbo.NFe.idCliente = dbo.ClientesInfFiscais.idCliente ORDER BY dbo.ClientesInfFiscais.Nome, dbo.NFe.DataEmissao DESC")
   Else
      If IsDate(cboDados.Text) Then
''         Set rsCombo = cnSistema.Execute("SELECT dbo.NFe.idNFe, dbo.NFe.Numero, dbo.NFe.DataEmissao, dbo.ClientesInfFiscais.Nome FROM dbo.NFe LEFT OUTER JOIN dbo.ClientesInfFiscais ON dbo.NFe.idCliente = dbo.ClientesInfFiscais.idCliente " & _
''                                         "WHERE (dbo.NFe.DataEmissao >= '" & Format(cboDados.Text, "mm/dd/yyyy") & " 00:00:00') AND (dbo.NFe.DataEmissao <= '" & Format(cboDados.Text, "mm/dd/yyyy") & " 23:59:59') ORDER BY dbo.ClientesInfFiscais.Nome, dbo.NFe.DataEmissao DESC")
      
         Set rsCombo = cnSistema.Execute("SELECT dbo.NFe.idNFe, dbo.NFe.Numero, dbo.NFe.DataEmissao, dbo.ClientesInfFiscais.Nome FROM dbo.NFe LEFT OUTER JOIN dbo.ClientesInfFiscais ON dbo.NFe.idCliente = dbo.ClientesInfFiscais.idCliente " & _
                                         "WHERE (dbo.NFe.DataEmissao >= '" & Format(cboDados.Text, "mm/dd/yyyy") & " 00:00:00') AND (dbo.NFe.DataEmissao <= '" & Format(cboDados.Text, "mm/dd/yyyy") & " 23:59:59') ORDER BY dbo.ClientesInfFiscais.Nome, dbo.NFe.DataEmissao DESC")
      Else
         If IsNumeric(cboDados.Text) Then
            Set rsCombo = cnSistema.Execute("SELECT dbo.NFe.idNFe, dbo.NFe.Numero, dbo.NFe.DataEmissao, dbo.ClientesInfFiscais.Nome FROM dbo.NFe LEFT OUTER JOIN dbo.ClientesInfFiscais ON dbo.NFe.idCliente = dbo.ClientesInfFiscais.idCliente " & _
                                            "WHERE (dbo.NFe.Numero = " & cboDados.Text & ")")
         Else
            Set rsCombo = cnSistema.Execute("SELECT dbo.NFe.idNFe, dbo.NFe.Numero, dbo.NFe.DataEmissao, dbo.ClientesInfFiscais.Nome FROM dbo.NFe LEFT OUTER JOIN dbo.ClientesInfFiscais ON dbo.NFe.idCliente = dbo.ClientesInfFiscais.idCliente " & _
                                            "WHERE dbo.ClientesInfFiscais.Nome LIKE '%" & cboDados.Text & "%' ORDER BY dbo.ClientesInfFiscais.Nome, dbo.NFe.DataEmissao DESC")
         End If
      End If
   End If
   Limpa_Campos
   Do While Not rsCombo.EOF
      cboDados.AddItem rsCombo!Numero & " - " & rsCombo!Nome & " - " & rsCombo!DataEmissao
      cboDados.ItemData(cboDados.NewIndex) = rsCombo!idNFe
      rsCombo.MoveNext
   Loop
   Set rsCombo = Nothing
   Screen.MousePointer = vbDefault
End Sub

Private Sub Carrega_Combos_Extras()
Dim rsCombo As New ADODB.Recordset
   Set rsCombo = cnSistema.Execute("SELECT idFormaPagamento, Descricao FROM FormaPagamento ORDER BY Descricao")
   cboFormaPagamento.Clear
   Do While Not rsCombo.EOF
      cboFormaPagamento.AddItem rsCombo!Descricao
      cboFormaPagamento.ItemData(cboFormaPagamento.NewIndex) = rsCombo!idFormaPagamento
      rsCombo.MoveNext
   Loop
   
   Set rsCombo = cnSistema.Execute("SELECT idCliente, Nome From ClientesInfFiscais ORDER BY Nome")
   cboCliente.Clear
   Do While Not rsCombo.EOF
      cboCliente.AddItem rsCombo!Nome
      cboCliente.ItemData(cboCliente.NewIndex) = rsCombo!idCliente
      rsCombo.MoveNext
   Loop
   
   Set rsCombo = cnSistema.Execute("SELECT idNaturezaOperacao, Descricao FROM NaturezasOperacao ORDER BY Descricao")
   cboNaturezaOperacao.Clear
   Do While Not rsCombo.EOF
      cboNaturezaOperacao.AddItem rsCombo!Descricao
      cboNaturezaOperacao.ItemData(cboNaturezaOperacao.NewIndex) = rsCombo!idNaturezaOperacao
      rsCombo.MoveNext
   Loop
   
   Set rsCombo = cnSistema.Execute("SELECT * FROM FreteConta")
   cmbFreteConta.Clear
   Do While Not rsCombo.EOF
      cmbFreteConta.AddItem rsCombo!Descricao
      cmbFreteConta.ItemData(cmbFreteConta.NewIndex) = rsCombo!idFreteConta
      rsCombo.MoveNext
   Loop
   
   Set rsCombo = Nothing
End Sub

Public Sub Botoes_Extras(bytModo As Byte)
'   Select Case bytModo
'      Case 1 'Sem seleção
'         Toolbar.Buttons(6).Enabled = False
'      Case 2 'Com seleção
'         Toolbar.Buttons(6).Enabled = True
'      Case 3 'Inclusão
'         Toolbar.Buttons(6).Enabled = False
'   End Select
End Sub

Private Sub cboProduto_Click()
Dim rsProduto As New ADODB.Recordset
   If cboProduto.ListIndex <> -1 Then
      Set rsProduto = cnSistema.Execute("SELECT * FROM Produtos " & _
                                        "WHERE (dbo.Produtos.idProduto = " & cboProduto.ItemData(cboProduto.ListIndex) & ")")
      lblDadosProduto.Caption = "Descrição: " & rsProduto!Descricao
      mskCodigoBarra.Tag = rsProduto!idProduto
      mskCodigoBarra.Text = rsProduto!idProduto
'      mskCodigoBarra.Text = rsProduto!CodigoBarra
      Set rsProduto = Nothing
   
'      Set rsProduto = cnSistema.Execute("SELECT dbo.Produtos.idProduto, dbo.Produtos.FracaoVenda, dbo.Produtos.UnidadeCompra, dbo.Produtos.UnidadeVenda, CASE dbo.Produtos.EmbalagemMultipla WHEN 1 THEN dbo.Produtos.DescricaoEmbalagemMultipla ELSE dboProdutos.Descricao END AS Descricao, dbo.Produtos.idClasseProdutos, dbo.Produtos.CodigoBarra, dbo.Produtos.Unidade, dbo.Produtos.ValorVenda, dbo.Produtos.Desconto, dbo.Produtos.Comissao, dbo.ClasseProdutos.Descricao AS Classe, dbo.Fabricantes.RazaoSocial AS Fabricante, dbo.Produtos.EstoqueMinimo, dbo.Produtos.EstoqueMaximo, dbo.Estoque.Pedidos, dbo.Estoque.Estoque, dbo.RegistradoresFiscais.Descricao AS RegistradorFiscal " & _
'                                        "FROM dbo.ClasseProdutos INNER JOIN dbo.Produtos ON dbo.ClasseProdutos.idClasseProdutos = dbo.Produtos.idClasseProdutos INNER JOIN dbo.Fabricantes ON dbo.Produtos.idFabricante = dbo.Fabricantes.idFabricante INNER JOIN dbo.Estoque ON dbo.Produtos.idProduto = dbo.Estoque.idProduto INNER JOIN dbo.RegistradoresFiscais ON dbo.Produtos.idRegistradorFiscal = dbo.RegistradoresFiscais.idRegistradorFiscal " & _
'                                        "WHERE (dbo.Produtos.idProduto = " & cboProduto.ItemData(cboProduto.ListIndex) & ")")
'      lblDadosProduto.Caption = "Descrição: " & rsProduto!Descricao & Chr(13) & _
'                                "Fabricante: " & rsProduto!Fabricante & Chr(13) & _
'                                "Classe: " & Trim(rsProduto!Classe) & " / Unidade: " & Trim(rsProduto!Unidade) & Chr(13) & _
'                                "Valor: " & Format(IIf(rsProduto!FracaoVenda, Round((rsProduto!ValorVenda * rsProduto!UnidadeVenda) / rsProduto!UnidadeCompra, 2), rsProduto!ValorVenda), "R$ ###,###,##0.00") & IIf(rsProduto!Desconto > 0 And Not BloquearDescontoPromocional, " - Desconto: " & rsProduto!Desconto & "%", "") & IIf(rsProduto!Comissao > 0, " - Com.: " & rsProduto!Comissao & "%", "") & Chr(13) & _
'                                "Estoque: " & rsProduto!Estoque & " / Mínimo: " & rsProduto!EstoqueMinimo & " / Máximo: " & rsProduto!EstoqueMaximo & " / Pedidos: " & rsProduto!Pedidos & Chr(13) & _
'                                "Registrador fiscal: " & rsProduto!RegistradorFiscal
'      ProdutoPromocional = IIf(rsProduto!Desconto > 0, True, False)
'      mskCodigoBarra.Tag = rsProduto!idProduto
'      mskCodigoBarra.Text = rsProduto!CodigoBarra
'      mskValorUnitario.Text = Round(IIf(rsProduto!FracaoVenda, Round((rsProduto!ValorVenda * rsProduto!UnidadeVenda) / rsProduto!UnidadeCompra, 2), rsProduto!ValorVenda) * IIf(rsProduto!Desconto > 0 And Not BloquearDescontoPromocional, 1 - (rsProduto!Desconto / 100), 1), 2)
'      Set rsProduto = Nothing
   End If
End Sub

Private Sub cboProduto_KeyPress(KeyAscii As Integer)
Dim rsCombo As New ADODB.Recordset
Dim bolEstoque As Boolean
   If KeyAscii = 13 Then
      If cboProduto.ListIndex = -1 Then
         Set rsCombo = cnSistema.Execute("SELECT COUNT(idProduto) AS Registros FROM Produtos WHERE (Descricao LIKE '%" & cboProduto.Text & "%') AND (Bloqueio = 0)")
         If rsCombo!Registros > 5000 Then
            MsgBox "A consulta possui mais de 5.000 resultados" & vbCrLf & "Redefina sua pesquisa", vbOKOnly + vbInformation, "Pesquisa"
            Exit Sub
         End If
      
         Screen.MousePointer = vbHourglass
'         Set rsCombo = cnSistema.Execute("SELECT EstoqueVendas FROM Sistema")
'         bolEstoque = rsCombo!EstoqueVendas
'         Set rsCombo = cnSistema.Execute("SELECT dbo.Produtos.idProduto, CASE WHEN dbo.Produtos.EmbalagemMultipla = 1 THEN dbo.Produtos.DescricaoEmbalagemMultipla ELSE dbo.Produtos.Descricao END AS Produto, dbo.Estoque.Estoque FROM dbo.Produtos INNER JOIN " & _
'                                         "dbo.Estoque ON dbo.Produtos.idProduto = dbo.Estoque.idProduto WHERE (CASE WHEN dbo.Produtos.EmbalagemMultipla = 1 THEN dbo.Produtos.DescricaoEmbalagemMultipla ELSE dbo.Produtos.Descricao END LIKE '%" & cboProduto.Text & "%') AND (dbo.Produtos.Bloqueio = 0) " & _
'                                         "ORDER BY CASE WHEN dbo.Produtos.EmbalagemMultipla = 1 THEN dbo.Produtos.DescricaoEmbalagemMultipla ELSE dbo.Produtos.Descricao END")
         
         Set rsCombo = cnSistema.Execute("SELECT * FROM dbo.Produtos " & _
                                         "WHERE dbo.Produtos.Descricao LIKE '%" & cboProduto.Text & "%' " & _
                                         "ORDER BY dbo.Produtos.Descricao")
         
         cboProduto.Clear
         Do While Not rsCombo.EOF
'            cboProduto.AddItem rsCombo!Produto & IIf(bolEstoque, " - " & rsCombo!Estoque, "")
            cboProduto.AddItem rsCombo!Descricao
            cboProduto.ItemData(cboProduto.NewIndex) = rsCombo!idProduto
            rsCombo.MoveNext
         Loop
         Set rsCombo = Nothing
         Screen.MousePointer = vbDefault
      End If
   End If
End Sub

Private Sub cboProduto_DblClick()
   mskQuantidade.SetFocus
End Sub

Private Sub cboProduto_LostFocus()
   cboProduto.ZOrder 1
   fraDadosProduto.Top = 4620
   fraDadosProduto.Left = 1500
   fraDadosProduto.Visible = False
   
   Call PreencherNFeComplemento
End Sub

Private Sub cboProduto_GotFocus()
Dim strMensagem As String
   If cboFormaPagamento.ListIndex = -1 Then strMensagem = strMensagem & "Forma de Pagamento" & Chr(13)
   If Not strMensagem = Empty Then
      MsgBox "Verifique os Seguintes Campos:" & Chr(13) & strMensagem, vbExclamation + vbOKOnly, "Campos Obrigatórios"
      Exit Sub
   End If
   cboProduto.ZOrder 0
   fraDadosProduto.Top = 4620
   fraDadosProduto.Left = 1500
   fraDadosProduto.Visible = True
End Sub

Private Sub cmdInserir_Click()
Dim bolDesconto As Boolean
Dim strDesconto As String
Dim intEstoque As Integer
Dim rsTemp As New ADODB.Recordset
Dim rsCliente As New ADODB.Recordset
Dim rsProduto As New ADODB.Recordset
Dim rsConvenio As New ADODB.Recordset
Dim rsVendedor As New ADODB.Recordset
Dim rsGrupoDesconto As New ADODB.Recordset
Dim rsGrupoComissao As New ADODB.Recordset
Dim rsDescontoVendedor As New ADODB.Recordset
'Contexto de Rotinas em cmdInserir_Click/cmdAlteraQuantidade_Click/AtualizaDescontoProdutos
   
   If cboFormaPagamento.ListIndex = -1 Then
      MsgBox "Escolha uma forma de pagamento", vbOKOnly, "Atendimento"
      cboCliente.SetFocus
      Exit Sub
   End If
'   If cboVendedor.ListIndex = -1 Then
'      MsgBox "Escolha um vendedor", vbOKOnly, "Atendimento"
'      cboVendedor.SetFocus
'      Exit Sub
'   End If
   On Error GoTo ErroInserir
   If Verifica_Campos_Produto() Then
      curDesconto = 0
      curComissao = 0
      Set rsProduto = cnSistema.Execute("SELECT idClasseProdutos, Comissao, idSituacaoTributaria FROM Produtos WHERE idProduto=" & mskCodigoBarra.Tag)
      If (mskDesconto.Enabled Or mskDescontoGeral.Enabled) Then
         If mskDesconto.Text = Empty And mskDescontoGeral.Text = Empty Then
            bolDesconto = False
         Else
            If Val(Replace(mskDesconto.Text, ",", ".")) > 0 Or Val(Replace(mskDescontoGeral.Text, ",", ".")) > 0 Then
               bolDesconto = True
            Else
               bolDesconto = False
            End If
         End If
''         If bolDesconto Then
''            Set rsDescontoVendedor = cnSistema.Execute("SELECT CASE (SELECT dbo.GrupoDesconto.DescontoGeral FROM dbo.Vendedores INNER JOIN dbo.GrupoDesconto ON dbo.Vendedores.idGrupoDesconto = dbo.GrupoDesconto.idGrupoDesconto WHERE (dbo.Vendedores.idVendedor = " & cboVendedor.ItemData(cboVendedor.ListIndex) & ")) WHEN 1 " & _
''                                                       "THEN (SELECT dbo.GrupoDesconto.Desconto FROM dbo.Vendedores INNER JOIN dbo.GrupoDesconto ON dbo.Vendedores.idGrupoDesconto = dbo.GrupoDesconto.idGrupoDesconto WHERE (dbo.Vendedores.idVendedor = " & cboVendedor.ItemData(cboVendedor.ListIndex) & ")) " & _
''                                                       "ELSE (SELECT dbo.GrupoDescontoItens.Percentual FROM dbo.Vendedores INNER JOIN dbo.GrupoDesconto ON dbo.Vendedores.idGrupoDesconto = dbo.GrupoDesconto.idGrupoDesconto INNER JOIN dbo.GrupoDescontoItens ON dbo.GrupoDesconto.idGrupoDesconto = dbo.GrupoDescontoItens.idGrupoDesconto " & _
''                                                       "WHERE (dbo.Vendedores.idVendedor = " & cboVendedor.ItemData(cboVendedor.ListIndex) & ") And (dbo.GrupoDescontoItens.idClasseProdutos = " & rsProduto!idClasseProdutos & ") And (dbo.GrupoDescontoItens.idFormaPagamento = " & cboFormaPagamento.ItemData(cboFormaPagamento.ListIndex) & ")) END AS Desconto")
''            If (Val(Replace(mskDesconto.Text, ",", ".")) > Val(Replace(mskDescontoGeral.Text, ",", "."))) And (Val(Replace(mskDesconto.Text, ",", ".")) <> 0) Then strDesconto = mskDesconto.Text Else strDesconto = mskDescontoGeral.Text
''            If Not IsNull(strDesconto) Then
''               If IsNull(rsDescontoVendedor!Desconto) Then
''                  curDesconto = 0
''                  mskDescontoGeral.Text = Empty
''                  mskAlteraDesconto.Text = Empty
''               Else
''                  If rsDescontoVendedor!Desconto < Val(Replace(strDesconto, ",", ".")) And (Not Validar_Permissao(4, "mnu" & Mid(Me.Name, 4, Len(Me.Name)))) Then
''                     curDesconto = rsDescontoVendedor!Desconto
''                  Else
''                     curDesconto = Val(Replace(strDesconto, ",", "."))
''                  End If
''               End If
''            End If
''            Set rsDescontoVendedor = Nothing
''         End If
      End If
''      If cboCliente.ListIndex <> -1 Then
''         Set rsCliente = cnSistema.Execute("SELECT Beneficio, idConvenio, idGrupoDesconto From ClientesInfFiscais WHERE idCliente=" & cboCliente.ItemData(cboCliente.ListIndex))
''         If rsCliente!Beneficio <> 0 And bolAplicarDesconto Then
''            Select Case rsCliente!Beneficio
''               Case 1 'Convênio
''                  Set rsConvenio = cnSistema.Execute("SELECT CASE (SELECT DescontoGeral FROM dbo.Convenios WHERE (idConvenio = " & rsCliente!idConvenio & ")) WHEN 1 " & _
''                                                     "THEN (SELECT Desconto FROM dbo.Convenios WHERE (idConvenio = " & rsCliente!idConvenio & ")) " & _
''                                                     "ELSE (SELECT dbo.ConveniosItens.Percentual FROM dbo.Convenios INNER JOIN dbo.ConveniosItens ON dbo.Convenios.idConvenio = dbo.ConveniosItens.idConvenio " & _
''                                                     "WHERE (dbo.ConveniosItens.idClasseProdutos = " & rsProduto!idClasseProdutos & ") AND (dbo.ConveniosItens.idConvenio = " & rsCliente!idConvenio & ")) END AS Desconto")
''                  curDesconto = IIf(IsNull(rsConvenio!Desconto), 0, rsConvenio!Desconto)
''                  Set rsConvenio = Nothing
''               Case 2 'Grupo de Desconto
''                  Set rsGrupoDesconto = cnSistema.Execute("SELECT CASE (SELECT DescontoGeral FROM dbo.GrupoDesconto WHERE (idGrupoDesconto = " & rsCliente!idGrupoDesconto & ")) WHEN 1 " & _
''                                                          "THEN (SELECT Desconto FROM dbo.GrupoDesconto WHERE (idGrupoDesconto = " & rsCliente!idGrupoDesconto & ")) " & _
''                                                          "ELSE (SELECT dbo.GrupoDescontoItens.Percentual FROM dbo.GrupoDesconto INNER JOIN dbo.GrupoDescontoItens ON dbo.GrupoDesconto.idGrupoDesconto = dbo.GrupoDescontoItens.idGrupoDesconto " & _
''                                                          "WHERE (dbo.GrupoDescontoItens.idClasseProdutos = " & rsProduto!idClasseProdutos & ") AND (dbo.GrupoDescontoItens.idFormaPagamento = " & cboFormaPagamento.ItemData(cboFormaPagamento.ListIndex) & ") AND (dbo.GrupoDescontoItens.idGrupoDesconto = " & rsCliente!idGrupoDesconto & ")) END AS Desconto")
''                  curDesconto = IIf(IsNull(rsGrupoDesconto!Desconto), 0, rsGrupoDesconto!Desconto)
''                  Set rsGrupoDesconto = Nothing
''            End Select
''         End If
''         Set rsCliente = Nothing
''      End If
''      If rsProduto!Comissao <= 0 Then
''         Set rsVendedor = cnSistema.Execute("SELECT idGrupoComissao FROM Vendedores WHERE idVendedor=" & cboVendedor.ItemData(cboVendedor.ListIndex))
''         Set rsGrupoComissao = cnSistema.Execute("SELECT CASE (SELECT ComissaoGeral FROM dbo.GrupoComissao WHERE (idGrupoComissao = " & IIf(IsNull(rsVendedor!idGrupoComissao), 0, rsVendedor!idGrupoComissao) & ")) WHEN 1 " & _
''                                                 "THEN (SELECT Comissao FROM dbo.GrupoComissao WHERE (idGrupoComissao = " & IIf(IsNull(rsVendedor!idGrupoComissao), 0, rsVendedor!idGrupoComissao) & ")) " & _
''                                                 "ELSE (SELECT dbo.GrupoComissaoItens.Percentual FROM dbo.GrupoComissao INNER JOIN dbo.GrupoComissaoItens ON dbo.GrupoComissao.idGrupoComissao = dbo.GrupoComissaoItens.idGrupoComissao " & _
''                                                 "WHERE (dbo.GrupoComissaoItens.idClasseProdutos = " & rsProduto!idClasseProdutos & ") AND (dbo.GrupoComissaoItens.idFormaPagamento = " & cboFormaPagamento.ItemData(cboFormaPagamento.ListIndex) & ") AND (dbo.GrupoComissaoItens.idGrupoComissao = " & IIf(IsNull(rsVendedor!idGrupoComissao), 0, rsVendedor!idGrupoComissao) & ")) END AS Comissao")
''         curComissao = IIf(IsNull(rsGrupoComissao!Comissao), 0, rsGrupoComissao!Comissao)
''         Set rsVendedor = Nothing
''         Set rsGrupoComissao = Nothing
''      Else
''         curComissao = rsProduto!Comissao
''      End If
      curComissao = 0
      Set ItemList = lvwProdutos.ListItems.Add(, "R" & mskCodigoBarra.Tag, cboProduto.Text)
      ItemList.SubItems(1) = Trim(mskQuantidade.Text)
      ItemList.SubItems(2) = Format(mskValorUnitario.Text, "R$ #,###,##0.00")
      ItemList.SubItems(3) = Format(mskValorUnitario.ClipText * mskQuantidade.Text, "R$ #,###,##0.00")
      ItemList.SubItems(4) = Round(Round(mskValorUnitario.ClipText * (curDesconto / 100), 2) * 100 / mskValorUnitario.ClipText, 2)
      ItemList.SubItems(5) = Format(Round(mskValorUnitario.ClipText * (1 - (curDesconto / 100)), 2) * mskQuantidade.Text, "R$ #,###,##0.00")
      ItemList.SubItems(6) = curComissao
      
      ItemList.SubItems(9) = CStrValor(frmNFeComplemento.mskICMSProduto.Text)
      ItemList.SubItems(10) = CStrValor(frmNFeComplemento.mskBaseReduzidaICMS.Text)
      ItemList.SubItems(11) = frmNFeComplemento.txtDescricaoComplementar.Text
      ItemList.SubItems(12) = IIf(Len(Trim(frmNFeComplemento.txtUnidade.Text)) = 0, "UN", frmNFeComplemento.txtUnidade.Text)
      ItemList.SubItems(13) = IIf(frmNFeComplemento.cmbSituacaoTributaria.ListIndex = -1, rsProduto!idSituacaoTributaria, frmNFeComplemento.cmbSituacaoTributaria.ItemData(frmNFeComplemento.cmbSituacaoTributaria.ListIndex))
      ItemList.SubItems(14) = frmNFeComplemento.txtDiscriminacaoProduto.Text
      ItemList.SubItems(15) = CStrValor(frmNFeComplemento.mskIPIProduto.Text)
      ItemList.SubItems(16) = CStrValor(frmNFeComplemento.mskBaseReduzidaIPI.Text)
      ItemList.SubItems(17) = frmNFeComplemento.txtClassificacaoFiscal.Text
      ItemList.SubItems(18) = frmNFeComplemento.mskCFOP.Text
      ItemList.SubItems(19) = CStrValor(frmNFeComplemento.mskValorFrete.Text)

      Set rsProduto = Nothing
      
      Atualizar_Totais
      Set rsTemp = cnSistema.Execute("SELECT estoque FROM estoque WHERE idProduto = " & mskCodigoBarra.Tag)
      intEstoque = rsTemp!Estoque
      If (intEstoque - Val(mskQuantidade.Text)) < 0 Then
         Set rsTemp = cnSistema.Execute("SELECT VendaSemEstoque,MensagemVendaSemEstoque FROM Sistema")
         If Not rsTemp!VendaSemEstoque Then
            ItemList.Bold = True
            If rsTemp!MensagemVendaSemEstoque Then MsgBox "Não existe quantidade suficiente deste produto e o mesmo não poderá ser vendido." & Chr(13) & _
                                                          "Será inserido na lista somente para fins de orçamento," & Chr(13) & _
                                                          "caso a venda seja confirmada ele será desconsiderado e" & Chr(13) & _
                                                          "somente os que tem estoque serão processados.", vbInformation + vbOKOnly, "Produto"
         End If
      End If
      Set rsTemp = Nothing
      mskCodigoBarra.Text = Space(13)
      cboProduto.ListIndex = -1
      lblDadosProduto.Caption = Empty
      mskDesconto.Text = Empty
'      lblTotalItens.Caption = lvwProdutos.ListItems.Count & " Itens"
      mskQuantidade.Text = "1" & Space(5)
      mskValorUnitario.Text = 0
      curDesconto = 0
      mskCodigoBarra.Tag = Empty
      mskCodigoBarra.SetFocus
   Else
      mskQuantidade.SetFocus
      mskQuantidade.SelStart = 0
      mskQuantidade.SelLength = Len(mskQuantidade.Text)
   End If
On Error GoTo 0
Exit Sub
ErroInserir:
   If Err.Number = 0 Then
      ' Operação Ok
   ElseIf Err.Number = 35602 Then
      MsgBox "Este produto já foi inserido" & Chr(13) & "Verifique na lista", vbInformation + vbOKOnly, "Inserir"
      Exit Sub
   Else
      MsgBox "Verificar: " & Err.Number & Chr(13) & Err.Description, vbExclamation, "Inserir"
      Exit Sub
   End If
End Sub

Private Function Verifica_Campos_Produto()
Dim strMensagem As String
Dim rsTemp As New ADODB.Recordset
Verifica_Campos_Produto = True

   If mskCodigoBarra.Text = Empty Then strMensagem = strMensagem & "Digite o codigo de barras ou escolha um produto" & Chr(13)
   If Val(mskQuantidade.Text) <= 0 Then strMensagem = strMensagem & "Quantidade tem de ser maior ou igual a 1" & Chr(13)
   If Val(Replace(mskValorUnitario.Text, ",", ".")) <= 0 Then strMensagem = strMensagem & "Valor unitário não permitido" & Chr(13)

 ' Testa se produto possui NCM
   If cboProduto.ListIndex <> -1 Then
      Set rsTemp = cnSistema.Execute("Select * From Produtos WHERE (dbo.Produtos.idProduto = " & cboProduto.ItemData(cboProduto.ListIndex) & ")")
   Else
      Set rsTemp = cnSistema.Execute("Select * From Produtos WHERE (dbo.Produtos.CodigoBarra = '" & mskCodigoBarra.Text & "')")
   End If
   If Not rsTemp.EOF Then
      If Len(Trim(rsTemp!CodigoNCM)) = 0 Then
         strMensagem = strMensagem & "Este produto não possui código de NCM" & Chr(13)
      End If
   End If

 ' Testa CFOPs
   Set rsClientes = cnSistema.Execute("Select * From ClientesInfFiscais Where idCliente = " & cboCliente.ItemData(cboCliente.ListIndex))
   Set rsEmpresa = cnSistema.Execute("Select * From Empresa")
   If Not rsEmpresa.EOF Then
      If rsClientes!SiglaUF = rsEmpresa!UF Then
''         frmNFeComplemento.cmbSituacaoTributaria.ItemData(frmNFeComplemento.cmbSituacaoTributaria.ListIndex)
''         frmNFeComplemento.mskCFOP.Text

         If Mid(mskCFOP.Text, 1, 1) = 2 Or Mid(mskCFOP.Text, 1, 1) = 6 Then
            strMensagem = strMensagem & "CFOP " & mskCFOP.Text & " inválido para o estado do cliente" & Chr(13)
         End If
      Else
         If Mid(mskCFOP.Text, 1, 1) = 1 Or Mid(mskCFOP.Text, 1, 1) = 5 Then
            strMensagem = strMensagem & "CFOP " & mskCFOP.Text & " inválido para o estado do cliente" & Chr(13)
         End If

      End If
   End If

   If Not strMensagem = Empty Then
      MsgBox "Verifique os Seguintes Campos:" & Chr(13) & strMensagem, vbExclamation + vbOKOnly, "Campos Obrigatórios"
      Verifica_Campos_Produto = False
      cboProduto.SetFocus
      Exit Function
   End If
End Function

Private Sub Gravar_Produtos(id As Long)
Dim Contador As Long
Dim curValor As String
Dim curDesconto As String
Dim curComissao As String
Dim rsSistema As New ADODB.Recordset
Dim rsEstoque As New ADODB.Recordset
Dim rsFormaPagamento As New ADODB.Recordset
Dim curICMS As String
Dim curBaseReduzida As String
Dim strDescricaoComplementar As String
Dim strUnidade As String
Dim strSituacaoTributaria As String
Dim strDiscriminacaoProduto As String
Dim curIPI As String
Dim curBaseReduzidaIPI As String
Dim strClassificacaoFiscal As String
Dim strCFOP As String
Dim curValorFrete As String

   Set rsSistema = cnSistema.Execute("SELECT VendaSemEstoque FROM Sistema")
   If lvwProdutos.ListItems.Count > 0 Then
      cnSistema.Execute "DELETE FROM NFeItens WHERE idNFe=" & id
      Set rsFormaPagamento = cnSistema.Execute("SELECT ControleCredito FROM FormaPagamento WHERE idFormaPagamento=" & cboFormaPagamento.ItemData(cboFormaPagamento.ListIndex))
      For Contador = 1 To lvwProdutos.ListItems.Count
         curValor = Mid(ValorCheck(lvwProdutos.ListItems.Item(Contador).SubItems(2)), 4, Len(ValorCheck(lvwProdutos.ListItems.Item(Contador).SubItems(2))))
         curDesconto = Round(curValor * lvwProdutos.ListItems.Item(Contador).SubItems(4) / 100, 2)
         curICMS = Mid(ValorCheck(lvwProdutos.ListItems.Item(Contador).SubItems(9)), 4, Len(ValorCheck(lvwProdutos.ListItems.Item(Contador).SubItems(9))))
         curBaseReduzida = Mid(ValorCheck(lvwProdutos.ListItems.Item(Contador).SubItems(10)), 4, Len(ValorCheck(lvwProdutos.ListItems.Item(Contador).SubItems(10))))
         strDescricaoComplementar = lvwProdutos.ListItems.Item(Contador).SubItems(11)
         strUnidade = lvwProdutos.ListItems.Item(Contador).SubItems(12)
         strSituacaoTributaria = lvwProdutos.ListItems.Item(Contador).SubItems(13)
         strDiscriminacaoProduto = lvwProdutos.ListItems.Item(Contador).SubItems(14)
         curIPI = Mid(ValorCheck(lvwProdutos.ListItems.Item(Contador).SubItems(15)), 4, Len(ValorCheck(lvwProdutos.ListItems.Item(Contador).SubItems(15))))
         curBaseReduzidaIPI = Mid(ValorCheck(lvwProdutos.ListItems.Item(Contador).SubItems(16)), 4, Len(ValorCheck(lvwProdutos.ListItems.Item(Contador).SubItems(16))))
         strClassificacaoFiscal = lvwProdutos.ListItems.Item(Contador).SubItems(17)
         strCFOP = lvwProdutos.ListItems.Item(Contador).SubItems(18)
         curValorFrete = Val(Replace(lvwProdutos.ListItems(Contador).SubItems(19), ",", "."))

'''         curComissao = Round((curValor - curDesconto) * lvwProdutos.ListItems.Item(Contador).SubItems(6) / 100, 2)

         Dim iUnidade As String, iSituacaoTributaria As Integer

         If rsSistema!VendaSemEstoque Then
            cnSistema.Execute "INSERT INTO NFeItens (idNFe,idProduto,Quantidade,ValorUnitario,Desconto,Data,ICMS,BaseReduzida,DescricaoComplementar,Unidade,idSituacaoTributaria,DiscriminacaoProduto,IPI,BaseReduzidaIPI,ClassificacaoFiscal,CFOP,ValorFrete) " & _
                              "VALUES (" & id & "," & Mid(lvwProdutos.ListItems(Contador).Key, 2, Len(lvwProdutos.ListItems(Contador).Key)) & ",'" & CStrValor(Trim(lvwProdutos.ListItems.Item(Contador).SubItems(1))) & "','" & CStrValor(curValor) & "','" & CStrValor(curDesconto) & "','" & Format(mskDataEmissao.Text, "mm/dd/yyyy hh:mm:ss") & "','" & _
                              CStrValor(curICMS) & "','" & CStrValor(curBaseReduzida) & "','" & SQLCheck(strDescricaoComplementar) & "','" & strUnidade & "'," & strSituacaoTributaria & ",'" & _
                              SQLCheck(strDiscriminacaoProduto) & "','" & CStrValor(curIPI) & "','" & CStrValor(curBaseReduzidaIPI) & "','" & SQLCheck(strClassificacaoFiscal) & "','" & strCFOP & "','" & CStrValor(curValorFrete) & "')"
            Atividade "Inc.NFe: " & Mid(lvwProdutos.ListItems.Item(Contador).SubItems(1), 1, 20) & " - " & Trim(lvwProdutos.ListItems.Item(Contador).SubItems(1)) & " - " & Trim(lvwProdutos.ListItems.Item(Contador).SubItems(5)), cboCliente.ItemData(cboCliente.ListIndex)
         Else
            Set rsEstoque = cnSistema.Execute("SELECT estoque FROM estoque WHERE idProduto=" & Mid(lvwProdutos.ListItems(Contador).Key, 2, Len(lvwProdutos.ListItems(Contador).Key)))
            If rsEstoque!Estoque >= Val(lvwProdutos.ListItems.Item(Contador).SubItems(1)) Then
               cnSistema.Execute "INSERT INTO NFeItens (idNFe,idProduto,Quantidade,ValorUnitario,Desconto,Data,ICMS,BaseReduzida,DescricaoComplementar,Unidade,idSituacaoTributaria,DiscriminacaoProduto,IPI,BaseReduzidaIPI,ClassificacaoFiscal,CFOP,ValorFrete) " & _
                                 "VALUES (" & id & "," & Mid(lvwProdutos.ListItems(Contador).Key, 2, Len(lvwProdutos.ListItems(Contador).Key)) & "," & Trim(lvwProdutos.ListItems.Item(Contador).SubItems(1)) & ",'" & Replace(curValor, ",", ".") & "','" & Replace(curDesconto, ",", ".") & "','" & Format(mskDataEmissao.Text, "mm/dd/yyyy hh:mm:ss") & "','" & _
                                 CStrValor(curICMS) & "','" & CStrValor(curBaseReduzida) & "','" & SQLCheck(strDescricaoComplementar) & "','" & strUnidade & "'," & strSituacaoTributaria & ",'" & _
                                 SQLCheck(strDiscriminacaoProduto) & "','" & CStrValor(curIPI) & "','" & CStrValor(curBaseReduzidaIPI) & "','" & SQLCheck(strClassificacaoFiscal) & "','" & strCFOP & "','" & CStrValor(curValorFrete) & "')"

               Atividade "Inc.NFe: " & Mid(lvwProdutos.ListItems.Item(Contador).SubItems(1), 1, 20) & " - " & Trim(lvwProdutos.ListItems.Item(Contador).SubItems(1)) & " - " & Trim(lvwProdutos.ListItems.Item(Contador).SubItems(5)), cboCliente.ItemData(cboCliente.ListIndex)
            End If
            Set rsEstoque = Nothing
         End If
         Set rsFormaPagamento = Nothing
      Next
   End If
   Set rsSistema = Nothing
End Sub

Private Sub cmdAnotacoes_Click()
Dim Contador As Integer

   If frmNFe.mskCodigoBarra.Text <> "" Then
      Set rsEmpresa = cnSistema.Execute("Select * From Empresa")
      Set rsClientes = cnSistema.Execute("Select * From ClientesInfFiscais Where idCliente = " & cboCliente.ItemData(cboCliente.ListIndex))

    ' Carrega Combos
      ' Unidades de Medida
''      Set rsTemp = cnSistema.Execute("Select * from UnidadesMedida Order By Descricao")
''      frmNFeComplemento.cmbUnidade.Clear
''      Do While Not rsTemp.EOF
''         frmNFeComplemento.cmbUnidade.AddItem rsTemp!Descricao
''         frmNFeComplemento.cmbUnidade.ItemData(frmNFeComplemento.cmbUnidade.NewIndex) = rsTemp!idUnidadeMedida
''         rsTemp.MoveNext
''      Loop
   
      ' Situacoes Tributarias
      Set rsTemp = cnSistema.Execute("Select * from SituacoesTributarias Order By Descricao")
      frmNFeComplemento.cmbSituacaoTributaria.Clear
      Do While Not rsTemp.EOF
         frmNFeComplemento.cmbSituacaoTributaria.AddItem rsTemp!Descricao
         frmNFeComplemento.cmbSituacaoTributaria.ItemData(frmNFeComplemento.cmbSituacaoTributaria.NewIndex) = rsTemp!idSituacaoTributaria
         rsTemp.MoveNext
      Loop
   
    ' Preencher Campos
      Set rsTemp = cnSistema.Execute("Select * From Produtos Where CodigoBarra = '" & SQLCheck(frmNFe.mskCodigoBarra.Text) & "'")
      If Not rsTemp.EOF Then
         ' ICMS
         frmNFeComplemento.mskICMSProduto.Text = rsTemp!ICMS
         If Not rsEmpresa.EOF Then
            If rsClientes!SiglaUF = rsEmpresa!UF Then
               frmNFeComplemento.mskBaseReduzidaICMS.Text = rsTemp!BaseReduzidaICMSdUF
            Else
               frmNFeComplemento.mskBaseReduzidaICMS.Text = rsTemp!BaseReduzidaICMSfUF
            End If
         End If
'         frmNFeComplemento.mskCFOP.Text = mskCFOP.Text
         
         Set rsEmpresa = cnSistema.Execute("Select * From Empresa")
         Set rsClientes = cnSistema.Execute("Select * From ClientesInfFiscais Where idCliente = " & cboCliente.ItemData(cboCliente.ListIndex))

         If Not rsEmpresa.EOF Then
            If rsClientes!SiglaUF = rsEmpresa!UF Then
               If rsTemp!CFOPDentroUF <> " .   " Then
                  frmNFeComplemento.mskCFOP.Text = rsTemp!CFOPDentroUF
               Else
                  frmNFeComplemento.mskCFOP.Text = mskCFOP.Text
               End If
            Else
               If rsTemp!CFOPForaUF <> " .   " Then
                  frmNFeComplemento.mskCFOP.Text = rsTemp!CFOPForaUF
               Else
                  frmNFeComplemento.mskCFOP.Text = mskCFOP.Text
               End If
            End If
         End If
         
         ' Unidade
''         For Contador = 0 To (frmNFeComplemento.cmbUnidade.ListCount - 1)
''            If frmNFeComplemento.cmbUnidade.ItemData(Contador) = rsTemp!idUnidade Then
''               frmNFeComplemento.cmbUnidade.ListIndex = Contador
''               Exit For
''            End If
''         Next
         
         frmNFeComplemento.txtUnidade.Text = rsTemp!Unidade

         ' Situacao Tributaria
         For Contador = 0 To (frmNFeComplemento.cmbSituacaoTributaria.ListCount - 1)
            If frmNFeComplemento.cmbSituacaoTributaria.ItemData(Contador) = rsTemp!idSituacaoTributaria Then
               frmNFeComplemento.cmbSituacaoTributaria.ListIndex = Contador
               Exit For
            End If
         Next
      End If
   
      frmNFeComplemento.Show vbModal
      cmdInserir.SetFocus
   End If
End Sub

Private Sub PreencherNFeComplemento()
Dim Contador As Integer

   If frmNFe.mskCodigoBarra.Text <> "" Then
      Set rsEmpresa = cnSistema.Execute("Select * From Empresa")
      Set rsClientes = cnSistema.Execute("Select * From ClientesInfFiscais Where idCliente = " & cboCliente.ItemData(cboCliente.ListIndex))

    ' Carrega Combos
      ' Unidades de Medida
''      Set rsTemp = cnSistema.Execute("Select * from UnidadesMedida Order By Descricao")
   
      ' Situacoes Tributarias
      Set rsTemp = cnSistema.Execute("Select * from SituacoesTributarias Order By Descricao")
      frmNFeComplemento.cmbSituacaoTributaria.Clear
      Do While Not rsTemp.EOF
         frmNFeComplemento.cmbSituacaoTributaria.AddItem rsTemp!Descricao
         frmNFeComplemento.cmbSituacaoTributaria.ItemData(frmNFeComplemento.cmbSituacaoTributaria.NewIndex) = rsTemp!idSituacaoTributaria
         rsTemp.MoveNext
      Loop
   
    ' Preencher Campos
      Set rsTemp = cnSistema.Execute("Select * From Produtos Where CodigoBarra = '" & SQLCheck(frmNFe.mskCodigoBarra.Text) & "'")
      If Not rsTemp.EOF Then
         ' ICMS
         frmNFeComplemento.mskICMSProduto.Text = rsTemp!ICMS
         If Not rsEmpresa.EOF Then
            If rsClientes!SiglaUF = rsEmpresa!UF Then
               frmNFeComplemento.mskBaseReduzidaICMS.Text = rsTemp!BaseReduzidaICMSdUF
            Else
               frmNFeComplemento.mskBaseReduzidaICMS.Text = rsTemp!BaseReduzidaICMSfUF
            End If
         End If
'         frmNFeComplemento.mskCFOP.Text = mskCFOP.Text
         
         Set rsEmpresa = cnSistema.Execute("Select * From Empresa")
         Set rsClientes = cnSistema.Execute("Select * From ClientesInfFiscais Where idCliente = " & cboCliente.ItemData(cboCliente.ListIndex))

         If Not rsEmpresa.EOF Then
            If rsClientes!SiglaUF = rsEmpresa!UF Then
               If rsTemp!CFOPDentroUF <> " .   " Then
                  frmNFeComplemento.mskCFOP.Text = rsTemp!CFOPDentroUF
               Else
                  frmNFeComplemento.mskCFOP.Text = mskCFOP.Text
               End If
            Else
               If rsTemp!CFOPForaUF <> " .   " Then
                  frmNFeComplemento.mskCFOP.Text = rsTemp!CFOPForaUF
               Else
                  frmNFeComplemento.mskCFOP.Text = mskCFOP.Text
               End If
            End If
         End If
         
         ' Unidade
         frmNFeComplemento.txtUnidade.Text = rsTemp!Unidade

         ' Situacao Tributaria
         For Contador = 0 To (frmNFeComplemento.cmbSituacaoTributaria.ListCount - 1)
            If frmNFeComplemento.cmbSituacaoTributaria.ItemData(Contador) = rsTemp!idSituacaoTributaria Then
               frmNFeComplemento.cmbSituacaoTributaria.ListIndex = Contador
               Exit For
            End If
         Next
      End If
   End If
End Sub

Private Sub cmdRemover_Click()
   If lvwProdutos.ListItems.Count = 0 Then
      MsgBox "Não existem produtos lançados", vbOKOnly + vbInformation, "Remover"
      Exit Sub
   End If
   If MsgBox("Deseja remover este produto", vbYesNo + vbQuestion, "Remover") = vbYes Then
      lvwProdutos.ListItems.Remove (lvwProdutos.SelectedItem.Index)
      Atualizar_Totais
      mskCodigoBarra.Text = Space(13)
'      lblTotalItens.Caption = lvwProdutos.ListItems.Count & " Itens"
      mskCodigoBarra.Tag = Empty
      cboProduto.Text = Empty
      cboProduto.ListIndex = -1
      mskQuantidade.Text = Space(6)
      mskValorUnitario.Text = 0
      mskCodigoBarra.SetFocus
   End If
End Sub

Private Sub Atualizar_Totais()
Dim intContador As Integer
'   mskArredondamento.Text = Empty
   curValorTotal = 0
   curValorDesconto = 0
   curValorFrete = 0
   For intContador = 1 To lvwProdutos.ListItems.Count
      curValorTotal = curValorTotal + Mid(ValorCheck(lvwProdutos.ListItems(intContador).SubItems(3)), 4, Len(ValorCheck(lvwProdutos.ListItems(intContador).SubItems(3))))
      curValorDesconto = curValorDesconto + Round((Mid(ValorCheck(lvwProdutos.ListItems(intContador).SubItems(2)), 4, Len(ValorCheck(lvwProdutos.ListItems(intContador).SubItems(2)))) * (lvwProdutos.ListItems(intContador).SubItems(4) / 100)), 2) * lvwProdutos.ListItems(intContador).SubItems(1)
      curValorFrete = curValorFrete + Val(Replace(lvwProdutos.ListItems(intContador).SubItems(19), ",", "."))
   Next
      
   mskValorFrete.Text = Format(curValorFrete, "##,##0.00") & " "
      
   lblTotalProdutos.Caption = Format(curValorTotal, "##,##0.00") & " "
   lblTotalNota.Caption = Format(curValorTotal + curValorFrete, "##,##0.00") & " "
   
'   lblValorTotal.Caption = "Valor Total: " & Format(curValorTotal, "R$ #,###,##0.00")
'   lblValorPagar.Caption = "Valor a pagar: " & Format(curValorTotal - curValorDesconto, "R$ #,###,##0.00")
'   lblValorDesconto.Caption = "Desconto: " & Format(curValorDesconto, "R$ #,###,##0.00")
End Sub

Private Sub cboNaturezaOperacao_LostFocus()
   If cboNaturezaOperacao.ListIndex >= 0 And cboCliente.ListIndex >= 0 Then
      Set rsClientes = cnSistema.Execute("Select * From ClientesInfFiscais Where idCliente = " & cboCliente.ItemData(cboCliente.ListIndex))
      Set rsTemp = cnSistema.Execute("Select * From NaturezasOperacao Where idNaturezaOperacao = " & cboNaturezaOperacao.ItemData(cboNaturezaOperacao.ListIndex))
      Set rsEmpresa = cnSistema.Execute("Select * From Empresa")
      If Not rsEmpresa.EOF Then
         If rsClientes!idUF = rsEmpresa!idUF Then
            mskCFOP.Text = rsTemp!CFOPDentroUF
         Else
            mskCFOP.Text = rsTemp!CFOPForaUF
         End If
      End If
      Set rsClientes = Nothing
      Set rsTemp = Nothing
      Set rsEmpresa = Nothing
   End If
End Sub

Public Function ValorCheck(SQLstring As String) As String
   ValorCheck = Replace(SQLstring, ".", "")
   ValorCheck = IIf(Trim(ValorCheck) = "", 0, ValorCheck)
End Function


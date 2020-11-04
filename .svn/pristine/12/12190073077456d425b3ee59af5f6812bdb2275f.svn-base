VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmNFeComplemento 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Complemento de Itens"
   ClientHeight    =   5130
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4635
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   4635
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraDiscriminacao 
      Caption         =   "Discriminação do Produto"
      Height          =   1635
      Left            =   60
      TabIndex        =   20
      Top             =   3000
      Width           =   4515
      Begin VB.TextBox txtDiscriminacaoProduto 
         Height          =   1320
         Left            =   75
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   21
         Top             =   225
         Width           =   4350
      End
   End
   Begin VB.Frame fraComplementos 
      Caption         =   "Informações Complementares"
      Height          =   2895
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   4515
      Begin VB.TextBox txtUnidade 
         Height          =   285
         Left            =   1260
         MaxLength       =   2
         TabIndex        =   11
         Top             =   1215
         Width           =   375
      End
      Begin VB.TextBox txtClassificacaoFiscal 
         Height          =   285
         Left            =   1260
         MaxLength       =   15
         TabIndex        =   17
         Top             =   2205
         Width           =   1575
      End
      Begin VB.ComboBox cmbSituacaoTributaria 
         Height          =   315
         ItemData        =   "frmNFeComplemento.frx":0000
         Left            =   1260
         List            =   "frmNFeComplemento.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   1875
         Width           =   3135
      End
      Begin VB.TextBox txtDescricaoComplementar 
         Height          =   285
         Left            =   1260
         MaxLength       =   30
         TabIndex        =   2
         Top             =   240
         Width           =   3150
      End
      Begin MSMask.MaskEdBox mskICMSProduto 
         Height          =   315
         Left            =   1260
         TabIndex        =   4
         Top             =   540
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
      Begin MSMask.MaskEdBox mskBaseReduzidaICMS 
         Height          =   315
         Left            =   2880
         TabIndex        =   6
         Top             =   540
         Width           =   735
         _ExtentX        =   1296
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
         Format          =   "##,##0.00000"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox mskIPIProduto 
         Height          =   315
         Left            =   1260
         TabIndex        =   24
         Top             =   870
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
      Begin MSMask.MaskEdBox mskBaseReduzidaIPI 
         Height          =   315
         Left            =   2880
         TabIndex        =   9
         Top             =   870
         Width           =   735
         _ExtentX        =   1296
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
         Format          =   "##,##0.00000"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox mskCFOP 
         Height          =   285
         Left            =   1275
         TabIndex        =   19
         Top             =   2505
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   5
         Mask            =   "9.999"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox mskValorFrete 
         Height          =   315
         Left            =   1260
         TabIndex        =   13
         Top             =   1530
         Width           =   1545
         _ExtentX        =   2725
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
      Begin VB.Label lblValorFrete 
         AutoSize        =   -1  'True
         Caption         =   "Valor do Frete"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   210
         TabIndex        =   12
         Top             =   1590
         Width           =   990
      End
      Begin VB.Label lblCFOP 
         AutoSize        =   -1  'True
         Caption         =   "CFOP"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   780
         TabIndex        =   18
         Top             =   2550
         Width           =   420
      End
      Begin VB.Label lblBaseReduzidaIPI 
         AutoSize        =   -1  'True
         Caption         =   "Base IPI"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2175
         TabIndex        =   8
         Top             =   930
         Width           =   600
      End
      Begin VB.Label lblClassificaoFiscal 
         AutoSize        =   -1  'True
         Caption         =   "Clas. Fiscal"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   405
         TabIndex        =   16
         Top             =   2250
         Width           =   795
      End
      Begin VB.Label lblIPI 
         AutoSize        =   -1  'True
         Caption         =   "IPI"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1005
         TabIndex        =   7
         Top             =   960
         Width           =   195
      End
      Begin VB.Label lblSituacaoTributaria 
         AutoSize        =   -1  'True
         Caption         =   "C.S.T."
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   750
         TabIndex        =   14
         Top             =   1935
         Width           =   450
      End
      Begin VB.Label lblBaseReduzidaICMS 
         AutoSize        =   -1  'True
         Caption         =   "Base ICMS"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1980
         TabIndex        =   5
         Top             =   600
         Width           =   795
      End
      Begin VB.Label lblUnidade 
         AutoSize        =   -1  'True
         Caption         =   "Unidade"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   600
         TabIndex        =   10
         Top             =   1260
         Width           =   600
      End
      Begin VB.Label lblICMSProduto 
         AutoSize        =   -1  'True
         Caption         =   "ICMS"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   810
         TabIndex        =   3
         Top             =   600
         Width           =   390
      End
      Begin VB.Label lblDescricaoComplementar 
         AutoSize        =   -1  'True
         Caption         =   "Comp. Desc. "
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   285
         Width           =   960
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   3300
      TabIndex        =   23
      Top             =   4680
      Width           =   1275
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1920
      TabIndex        =   22
      Top             =   4680
      Width           =   1275
   End
End
Attribute VB_Name = "frmNFeComplemento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ItemList As ListItem

Private Sub Form_Load()
Dim Contador As Integer

'   On Error GoTo Erro
   Centraliza frmNFeComplemento
   
'''   MDISistema.StatusBar.Panels(1).text = "Lançar Complemento de Itens da Nota"
   
''Exit Sub
''Erro:
''   If Err.Number = -2147467259 Then
''      rsErro = True
''      Beep
''      MsgBox "Erro na Abertura do Arquivo de Dados" & Chr(13) & "Algum usuário está com o Arquivo em modo Exclusivo", vbExclamation, "Erro"
''      Exit Sub
''   Else
''      rsErro = True
''      Beep
''      MsgBox "Verificar: " & Err.Number & Chr(13) & Err.Description, vbExclamation, "Sistema"
''      Exit Sub
''   End If
End Sub

Private Sub cmdOK_Click()
   Me.Hide
End Sub

Private Sub cmdCancelar_Click()
   Unload Me
End Sub

VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmGerenciarNF 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gerenciar Notas"
   ClientHeight    =   10260
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   17190
   Icon            =   "frmGerenciarNF.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10260
   ScaleWidth      =   17190
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraMensagensRetornoErros 
      Caption         =   "Mensagens de Retorno e Erros"
      Height          =   2055
      Left            =   60
      TabIndex        =   10
      Top             =   6660
      Width           =   9165
      Begin VB.CommandButton cmdHistorico 
         Caption         =   "Hist�rico"
         Height          =   375
         Left            =   6960
         TabIndex        =   13
         Top             =   1560
         Width           =   2115
      End
      Begin VB.CommandButton cmdLimparHistorico 
         Caption         =   "Limpar"
         Height          =   375
         Left            =   4740
         TabIndex        =   12
         Top             =   1560
         Width           =   2115
      End
      Begin MSComctlLib.ListView lvwMensagens 
         Height          =   1290
         Left            =   60
         TabIndex        =   11
         Top             =   240
         Width           =   9030
         _ExtentX        =   15928
         _ExtentY        =   2275
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
   End
   Begin VB.Frame fraFuncoesAdicionais 
      Height          =   1215
      Left            =   60
      TabIndex        =   3
      Top             =   8760
      Width           =   17055
      Begin VB.CommandButton cmdNFSe 
         Caption         =   "NFS-e"
         Height          =   375
         Left            =   6720
         TabIndex        =   57
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmdMDFe 
         Caption         =   "MDF-e"
         Height          =   375
         Left            =   4920
         TabIndex        =   54
         Top             =   240
         Width           =   1695
      End
      Begin VB.Timer tmrAtualiza 
         Interval        =   8000
         Left            =   16080
         Top             =   150
      End
      Begin VB.CommandButton cmdCopiarNFs 
         Caption         =   "Copiar NFs"
         Height          =   375
         Left            =   1920
         TabIndex        =   26
         Top             =   720
         Width           =   1695
      End
      Begin VB.CommandButton cmdGerarPDF 
         Caption         =   "Gerar PDF"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1920
         TabIndex        =   25
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmdEnviarEmail 
         Caption         =   "Enviar E-mail"
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   720
         Width           =   1695
      End
      Begin VB.CommandButton cmdNFeCartaCorrecao 
         Caption         =   "Carta de Corre��o"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   1695
      End
      Begin VB.Timer tmrImprimirSistema 
         Enabled         =   0   'False
         Interval        =   5000
         Left            =   16560
         Top             =   150
      End
      Begin VB.Label lblStatusEvento 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Status do Evento"
         Height          =   255
         Left            =   15540
         TabIndex        =   34
         Top             =   840
         Visible         =   0   'False
         Width           =   1395
      End
   End
   Begin VB.Frame fraInformacoes 
      Caption         =   "Informa��es"
      Height          =   8655
      Left            =   9300
      TabIndex        =   2
      Top             =   60
      Width           =   7815
      Begin VB.Frame fraErrosValidacao 
         Caption         =   "Erros de valida��o"
         Height          =   1155
         Left            =   120
         TabIndex        =   46
         Top             =   7440
         Width           =   7575
         Begin MSComctlLib.ListView lvwErrosValidacaoNF 
            Height          =   810
            Left            =   60
            TabIndex        =   47
            Top             =   240
            Width           =   7425
            _ExtentX        =   13097
            _ExtentY        =   1429
            View            =   3
            LabelEdit       =   1
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   255
            BackColor       =   -2147483624
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
      End
      Begin VB.Frame fraNFFiscais 
         Height          =   1395
         Left            =   120
         TabIndex        =   37
         Top             =   180
         Width           =   7575
         Begin VB.CommandButton cmdEditarNotas 
            Caption         =   "Editar Notas"
            Height          =   315
            Left            =   5760
            TabIndex        =   56
            Top             =   600
            Width           =   1695
         End
         Begin VB.CommandButton cmdEventos 
            Caption         =   "Eventos"
            Enabled         =   0   'False
            Height          =   315
            Left            =   5760
            TabIndex        =   48
            Top             =   960
            Width           =   1695
         End
         Begin VB.TextBox txtProtocoloCancelamento 
            Enabled         =   0   'False
            Height          =   315
            Left            =   2400
            TabIndex        =   44
            Top             =   960
            Width           =   2235
         End
         Begin VB.TextBox txtProtocoloAprovacao 
            Enabled         =   0   'False
            Height          =   315
            Left            =   120
            TabIndex        =   42
            Top             =   960
            Width           =   2235
         End
         Begin VB.TextBox txtChaveNFe 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1020
            TabIndex        =   39
            Top             =   420
            Width           =   4575
         End
         Begin VB.TextBox txtNota 
            Enabled         =   0   'False
            Height          =   315
            Left            =   120
            TabIndex        =   38
            Top             =   420
            Width           =   855
         End
         Begin VB.Label lblProtocoloCancelamento 
            Caption         =   "Protocolo de Cancelamento"
            Height          =   195
            Left            =   2400
            TabIndex        =   45
            Top             =   720
            Width           =   2235
         End
         Begin VB.Label lblProtocoloAprovacao 
            Caption         =   "Protocolo de Aprova��o"
            Height          =   195
            Left            =   120
            TabIndex        =   43
            Top             =   720
            Width           =   1875
         End
         Begin VB.Label lblChave 
            Caption         =   "Chave de Acesso"
            Height          =   195
            Left            =   1020
            TabIndex        =   41
            Top             =   180
            Width           =   1395
         End
         Begin VB.Label lblNota 
            Caption         =   "N�mero"
            Height          =   195
            Left            =   120
            TabIndex        =   40
            Top             =   180
            Width           =   795
         End
      End
      Begin VB.Frame fraNota 
         Caption         =   "Itens da Nota"
         Height          =   3615
         Left            =   120
         TabIndex        =   23
         Top             =   3780
         Width           =   7575
         Begin VB.TextBox txtTotalPagamentos 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   315
            Left            =   5850
            TabIndex        =   52
            Top             =   3225
            Width           =   1395
         End
         Begin VB.TextBox txtValorTotal 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   315
            Left            =   5850
            TabIndex        =   51
            Top             =   1740
            Width           =   1395
         End
         Begin MSComctlLib.ListView lvwNFsItens 
            Height          =   1455
            Left            =   60
            TabIndex        =   36
            Top             =   240
            Width           =   7425
            _ExtentX        =   13097
            _ExtentY        =   2566
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
         Begin MSComctlLib.ListView lvwNFsPagamentos 
            Height          =   1095
            Left            =   60
            TabIndex        =   50
            Top             =   2100
            Width           =   7425
            _ExtentX        =   13097
            _ExtentY        =   1931
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
         Begin VB.Label lblTotalPagamentos 
            Caption         =   "Valor Total"
            Height          =   195
            Left            =   4860
            TabIndex        =   53
            Top             =   3285
            Width           =   855
         End
         Begin VB.Label lblValorTotal 
            Caption         =   "Valor Total"
            Height          =   195
            Left            =   4860
            TabIndex        =   49
            Top             =   1800
            Width           =   855
         End
      End
      Begin VB.Frame fraInformacoesClientes 
         Caption         =   "Cliente"
         Height          =   2175
         Left            =   120
         TabIndex        =   22
         Top             =   1620
         Width           =   7575
         Begin VB.CommandButton cmdDestinatarios 
            Caption         =   "Destinat�rios"
            Height          =   315
            Left            =   3960
            TabIndex        =   60
            Top             =   1740
            Width           =   1695
         End
         Begin VB.TextBox txtInformacoesCliente 
            Height          =   1455
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   35
            Top             =   240
            Width           =   7335
         End
         Begin VB.CommandButton cmdConsultarSefaz 
            Caption         =   "Consultar Sefaz"
            Enabled         =   0   'False
            Height          =   315
            Left            =   5760
            TabIndex        =   27
            Top             =   1740
            Width           =   1695
         End
         Begin VB.Label lblInformacoesCliente 
            Height          =   1335
            Left            =   120
            TabIndex        =   24
            Top             =   300
            Width           =   7275
         End
      End
   End
   Begin VB.Frame fraNotas 
      Caption         =   "Notas"
      Height          =   6555
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   9165
      Begin VB.CommandButton cmdDataFinal 
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   4260
         Picture         =   "frmGerenciarNF.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   59
         ToolTipText     =   "Confirma inclus�o de produtos"
         Top             =   255
         Width           =   360
      End
      Begin VB.CommandButton cmdDataInicial 
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   1980
         Picture         =   "frmGerenciarNF.frx":04FA
         Style           =   1  'Graphical
         TabIndex        =   58
         ToolTipText     =   "Confirma inclus�o de produtos"
         Top             =   255
         Width           =   360
      End
      Begin VB.CommandButton cmdPesquisar 
         Caption         =   "&Pesquisar"
         Height          =   315
         Left            =   8030
         TabIndex        =   33
         Top             =   255
         Width           =   1035
      End
      Begin VB.Frame fraInformacoesNotas 
         Height          =   675
         Left            =   120
         TabIndex        =   28
         Top             =   5340
         Width           =   9000
         Begin VB.Label lblProcessamento 
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
            ForeColor       =   &H000000FF&
            Height          =   315
            Left            =   120
            TabIndex        =   55
            Top             =   240
            Width           =   8760
         End
         Begin VB.Label lblNotasIntuilizadas 
            Caption         =   "Inutilizadas"
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
            Left            =   5580
            TabIndex        =   32
            Top             =   300
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label lblNotasAprovadas 
            Caption         =   "Aprovadas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   195
            Left            =   4440
            TabIndex        =   31
            Top             =   360
            Visible         =   0   'False
            Width           =   1035
         End
         Begin VB.Label lblNotasCanceladas 
            Caption         =   "Canceladas"
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
            Left            =   6720
            TabIndex        =   30
            Top             =   300
            Visible         =   0   'False
            Width           =   1035
         End
         Begin VB.Label lblNotasPendentes 
            Caption         =   "Pendentes"
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
            Left            =   7860
            TabIndex        =   29
            Top             =   300
            Visible         =   0   'False
            Width           =   1035
         End
      End
      Begin VB.CommandButton cmdCancelarNota 
         Caption         =   "&Cancelar Nota"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1380
         TabIndex        =   19
         Top             =   6060
         Width           =   1575
      End
      Begin VB.CommandButton cmdImprimirDANFE 
         Caption         =   "Imprimir DANFE"
         Enabled         =   0   'False
         Height          =   375
         Left            =   4740
         TabIndex        =   18
         Top             =   6060
         Width           =   1635
      End
      Begin VB.CommandButton cmdInutilizarNumeracao 
         Caption         =   "Inutilizar N�mero"
         Height          =   375
         Left            =   3000
         TabIndex        =   17
         Top             =   6060
         Width           =   1695
      End
      Begin VB.CommandButton cmdConsultaNota 
         Caption         =   "Consulta Situa��o "
         Height          =   375
         Left            =   6420
         TabIndex        =   16
         Top             =   6060
         Width           =   1875
      End
      Begin VB.CommandButton cmdTransmitir 
         Caption         =   "&Transmitir"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   6060
         Width           =   1155
      End
      Begin VB.ComboBox cmbSituacao 
         Height          =   315
         ItemData        =   "frmGerenciarNF.frx":09E8
         Left            =   5460
         List            =   "frmGerenciarNF.frx":0A04
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   255
         Width           =   2505
      End
      Begin MSComctlLib.ListView lvwNFs 
         Height          =   4800
         Left            =   60
         TabIndex        =   1
         Top             =   600
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   8467
         View            =   3
         LabelEdit       =   1
         SortOrder       =   -1  'True
         Sorted          =   -1  'True
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
      Begin MSMask.MaskEdBox mskDataInicial 
         Height          =   285
         Left            =   990
         TabIndex        =   5
         Top             =   270
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         Mask            =   "99/99/9999"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox mskDataFinal 
         Height          =   285
         Left            =   3300
         TabIndex        =   6
         Top             =   270
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         Mask            =   "99/99/9999"
         PromptChar      =   " "
      End
      Begin VB.Label lblDataFinal 
         AutoSize        =   -1  'True
         Caption         =   "Data Final"
         Height          =   195
         Left            =   2460
         TabIndex        =   9
         Top             =   315
         Width           =   720
      End
      Begin VB.Label lblDataInicial 
         AutoSize        =   -1  'True
         Caption         =   "Data Inicial"
         Height          =   195
         Left            =   90
         TabIndex        =   8
         Top             =   315
         Width           =   795
      End
      Begin VB.Label lblSituacao 
         AutoSize        =   -1  'True
         Caption         =   "Situa��o"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4740
         TabIndex        =   7
         Top             =   315
         Width           =   630
      End
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   14
      Top             =   10005
      Width           =   17190
      _ExtentX        =   30321
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   22578
            MinWidth        =   5380
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   2469
            MinWidth        =   2469
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "INS"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1931
            MinWidth        =   1940
            TextSave        =   "24/07/2020"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuArquivos 
      Caption         =   "Arquivos"
   End
   Begin VB.Menu mnuLancamentos 
      Caption         =   "Lan�amentos"
      Begin VB.Menu mnuImportar 
         Caption         =   "Importar"
      End
      Begin VB.Menu mnuImportarLink 
         Caption         =   "Importar Link"
      End
   End
   Begin VB.Menu mnuConfiguracoes 
      Caption         =   "Configura��es"
      Begin VB.Menu mnuOpcoes 
         Caption         =   "Op��es"
      End
   End
End
Attribute VB_Name = "frmGerenciarNF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------------------------------
' Gerenciador de NFs                                                                                        '
' Desenvolvedor: Robinson                                                                                   '
' �ltima atualiza��o: 14/02/2020 23:00                                                                      '
'------------------------------------------------------------------------------------------------------------
Option Explicit
Dim ItemList As ListItem
Dim ProcuraItem As ListItem

'''Dim rsNFs As New ADODB.Recordset
'''Dim rsNFsItens As New ADODB.Recordset
Dim rsNFsInutilizadas As New ADODB.Recordset
'''Dim rsTotalNFs As ADODB.Recordset
Dim rsClientes As New ADODB.Recordset
'''''Dim rsEmpresa As New ADODB.Recordset
Dim rsGerarXML As New ADODB.Recordset
Dim rsDados As New ADODB.Recordset

Private SHA1Hash As New SHA1Hash

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

Public localErro As String

Private Sub cmdDataFinal_Click()
'''   tmrAtualiza.Enabled = False
   frmSelecionarData.Show vbModal
   mskDataFinal.Text = frmSelecionarData.mtwCalendario.value
'''   tmrAtualiza.Enabled = True
End Sub

Private Sub cmdDataInicial_Click()
'''   tmrAtualiza.Enabled = False
   frmSelecionarData.Show vbModal
   mskDataInicial.Text = frmSelecionarData.mtwCalendario.value
'''   tmrAtualiza.Enabled = True
End Sub

Private Sub Form_Activate()
   Me.Top = 200
End Sub

Private Sub Form_Load()
On Error GoTo Erro
Dim titulo As String
Dim vhwnd As Long
   
'''''   SHCreateThread AddressOf Carrega_View, ByVal 0&, &H1, ByVal 0&

   If App.PrevInstance Then
      titulo = Me.Caption
      Me.Caption = ""
      
      vhwnd = FindWindow(vbNullString, titulo)
      Call ShowWindow(vhwnd, 9)
      
      Call SetForegroundWindow(vhwnd)
      End
   End If

   lvwNFs.ColumnHeaders.Add , , "N�mero", 850
   lvwNFs.ColumnHeaders.Add , , "Emiss�o", 1050
   lvwNFs.ColumnHeaders.Add , , "Cliente", 4000
   lvwNFs.ColumnHeaders.Add , , "Valor Total", 1050, lvwColumnRight
   lvwNFs.ColumnHeaders.Add , , "Situa��o", 1700
   lvwNFs.ColumnHeaders.Add , , "", 0
   
   lvwMensagens.ColumnHeaders.Add , , "Chave", 0
   lvwMensagens.ColumnHeaders.Add , , "N�mero", 850
   lvwMensagens.ColumnHeaders.Add , , "Mensagem", 7850
   lvwMensagens.ColumnHeaders.Add , , "Arquivo", 0
   
   lvwNFsItens.ColumnHeaders.Add , , "C�digo", 850
   lvwNFsItens.ColumnHeaders.Add , , "Produto", 2950
   lvwNFsItens.ColumnHeaders.Add , , "CFOP", 650
   lvwNFsItens.ColumnHeaders.Add , , "Quant.", 750, lvwColumnRight
   lvwNFsItens.ColumnHeaders.Add , , "Vl. Unit.", 850, lvwColumnRight
   lvwNFsItens.ColumnHeaders.Add , , "Valor Total", 1050, lvwColumnRight
   
   lvwNFsPagamentos.ColumnHeaders.Add , , "C�digo", 850
   lvwNFsPagamentos.ColumnHeaders.Add , , "Pagamento", 5050
   lvwNFsPagamentos.ColumnHeaders.Add , , "Valor", 1200, lvwColumnRight
   
   lvwErrosValidacaoNF.ColumnHeaders.Add , , "Mensagem", 7100
   
'''''   mskDataInicial.text = CDate("01" & Mid(Date, 3, 8))
   mskDataInicial.Text = Date - LerArquivoINI("NFe", "DiasMovimento", App.Path & "\System.ini")
   mskDataFinal.Text = Date
'   cmbSituacao.ListIndex = 7
   cmbSituacao.ListIndex = LerArquivoINI("NFe", "Situacao", App.Path & "\System.ini")
   
   ' Atualiza Intervalo de verifica��es de atualiza��o
'''   tmrAtualiza.Interval = LerArquivoINI("NFe", "Intervalo", App.Path & "\System.ini") * 1000
   
   Call Main
   
   Carrega_View ("Carregar")
   
   ''Carregar UNINFe
   If LerArquivoINI("NFe", "UNINFe", App.Path & "\System.ini") = 1 Then CarregaUNINFe
   
   cmdMDFe.Visible = LerArquivoINI("Modulos", "MDFe", App.Path & "\System.ini")
   cmdNFSe.Visible = LerArquivoINI("Modulos", "NFSe", App.Path & "\System.ini")
   
   frmGerenciarNF.Caption = UCase("Gerenciar Notas da Empresa " & rsEmpresa!Nome)
   
'''   ''Carregar Verificador
'''   If LerArquivoINI("Verificador", "Ativar", App.Path & "\System.ini") = 1 Then CarregaVerificador
   
   ' Ativa o contador de verifica��es
'''   tmrAtualiza.Enabled = True
   
   Exit Sub
Erro:
   Call FRetornaMensagens(Err.Number & " - " & Err.Description & " - " & TypeName(Me))
End Sub

Private Sub CarregaUNINFe()
Dim sEmpresaNFe As String

   sEmpresaNFe = LerArquivoINI("NFe", "Empresa", CaminhoINI & "\System.ini")
   If Trim(sEmpresaNFe) <> "" Then
      If Dir("C:\UNIMAKE\" & sEmpresaNFe & "\UNINFE.EXE") <> "" Then
         Shell "C:\UNIMAKE\" & sEmpresaNFe & "\UNINFE.EXE"
      End If
   End If
End Sub

Private Sub CarregaVerificador()
Dim sAtivar As String

   sAtivar = LerArquivoINI("Verificador", "Ativar", CaminhoINI & "\System.ini")
   If Trim(sAtivar) <> "" Then
      If Dir(App.Path & "\Verificador.EXE") <> "" Then
         Shell App.Path & "\Verificador.EXE"
      End If
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
''   TerminateThread hThread1, 0
''   CloseHandle hThread1
   
   ''Carregar UNINFe
   If LerArquivoINI("NFe", "UNINFe", App.Path & "\System.ini") = 1 Then KillProcess "uninfe.exe"
'''   If LerArquivoINI("Verificador", "Ativar", App.Path & "\System.ini") = 1 Then KillProcess "Verificador.exe"
      
End Sub

Private Sub lvwNFs_Click()
On Error GoTo Erro
Dim oNFs400 As New CNF400
Dim oPreencherRs As New PreencherRS
Dim Contador As Integer
Dim dValorTotal As Double
Dim dValorTotalPagamento As Double
   
   lvwErrosValidacaoNF.ListItems.Clear
   lvwNFsItens.ListItems.Clear
   lvwNFsPagamentos.ListItems.Clear
   
   txtValorTotal.Text = ""
   txtTotalPagamentos.Text = ""
   
   Contador = 1
   dValorTotal = 0
   dValorTotalPagamento = 0
   
   If lvwNFs.ListItems.Count <> 0 Then
      localErro = "Preencher Campos"

      ' Preencher os dados da nota
      Call oPreencherRs.PreencherRsNFs(Val(lvwNFs.ListItems(lvwNFs.SelectedItem.Index)), "NFe")
      If Not rsNFs.EOF Then
         txtNota.Text = StrZero(rsNFs!nnf, 8)
         txtChaveNFe.Text = FFormataChaveNF(IIf(IsNull(rsNFs!cNF), "", rsNFs!cNF))
         txtProtocoloAprovacao.Text = rsNFs!Protocolo
         txtProtocoloCancelamento.Text = rsNFs!ProtocoloCancelamento
         txtInformacoesCliente.Text = Space(7) & "CNPJ/CPF:" & vbTab & "[" & IIf(Trim(rsNFsDestinatarios!CNPJ) <> "", rsNFsDestinatarios!CNPJ, rsNFsDestinatarios!CPF) & "]" & vbTab & vbTab & "I.E.: [" & rsNFsDestinatarios!IE & "]" & vbCrLf & _
                                      Space(7) & "Nome:" & vbTab & "[" & rsNFsDestinatarios!xNome & "] " & vbCrLf & _
                                      Space(7) & "Endere�o:" & vbTab & "[" & Trim(rsNFsDestinatarios!xLgr) & ", " & rsNFsDestinatarios!nro & "] " & vbCrLf & _
                                      Space(7) & "Bairro:" & vbTab & "[" & rsNFsDestinatarios!xBairro & "] " & vbCrLf & _
                                      Space(7) & "Munic�pio:" & vbTab & "[" & rsNFsDestinatarios!cMun & "] - [" & rsNFsDestinatarios!xMun & "]" & vbTab & "UF: [" & rsNFsDestinatarios!UF & "]" & vbTab & "CEP: [" & rsNFsDestinatarios!CEP & "] " & vbCrLf & _
                                      Space(7) & "Telefone:" & vbTab & "[" & rsNFsDestinatarios!fone & "]"
         
         localErro = "Dados do Cliente"
         
         ' Itens
         If Not rsNFsItens.EOF Then rsNFsItens.MoveFirst
         Do While Not rsNFsItens.EOF
'            Set ItemList = lvwNFsItens.ListItems.Add(, "I" & CStr(Contador), StrZero(rsNFsItens!cProd, 8))
            Set ItemList = lvwNFsItens.ListItems.Add(, "I" & CStr(Contador), rsNFsItens!cProd)
            ItemList.SubItems(1) = rsNFsItens!xprod
            ItemList.SubItems(2) = rsNFsItens!CFOP
            ItemList.SubItems(3) = Format(rsNFsItens!qTrib, "##,##0.00")
            ItemList.SubItems(4) = Format(rsNFsItens!vUnTrib, "##,##0.00")
            ItemList.SubItems(5) = Format(rsNFsItens!qTrib * rsNFsItens!vUnTrib, "##,##0.00")

            dValorTotal = dValorTotal + (rsNFsItens!qTrib * rsNFsItens!vUnTrib)
            
            localErro = "Produto " & rsNFsItens!xprod
            
            Contador = Contador + 1
            rsNFsItens.MoveNext
         Loop
         
         txtValorTotal.Text = Format(dValorTotal, "##,##0.00")
         
         ' Pagamentos
         Contador = 1
         If Not rsNFsPagamentos.EOF Then rsNFsPagamentos.MoveFirst
         Do While Not rsNFsPagamentos.EOF
            Set ItemList = lvwNFsPagamentos.ListItems.Add(, "I" & CStr(Contador), rsNFsPagamentos!tPag)
            ItemList.SubItems(1) = FFormaPagamento(rsNFsPagamentos!tPag)
            ItemList.SubItems(2) = Format(rsNFsPagamentos!vPag, "##,##0.00")

            localErro = "Pagamentos"

            dValorTotalPagamento = dValorTotalPagamento + rsNFsPagamentos!vPag
            
            Contador = Contador + 1
            rsNFsPagamentos.MoveNext
         Loop
         
         txtTotalPagamentos.Text = Format(dValorTotalPagamento, "##,##0.00")
         
         cmdTransmitir.Enabled = rsNFs!Situacao
         
'''         If rsNFs!Situacao = 0 Then
'''            cmdTransmitir.Enabled = True
''''            cmdValidar.Enabled = True
'''         Else
'''            cmdTransmitir.Enabled = False
''''            cmdValidar.Enabled = False
'''         End If
      
         If rsNFs!Situacao = 2 Then
            If Trim(rsNFs!cNF) <> "" And Trim(rsNFs!Protocolo) <> "" Then
               cmdTransmitir.Enabled = False
               cmdCancelarNota.Enabled = True
               cmdNFeCartaCorrecao.Enabled = True
               cmdImprimirDANFE.Enabled = True
            Else
               cmdTransmitir.Enabled = True
               cmdCancelarNota.Enabled = False
               cmdNFeCartaCorrecao.Enabled = False
               cmdImprimirDANFE.Enabled = False
            End If
         Else
            cmdTransmitir.Enabled = True
            cmdCancelarNota.Enabled = False
            cmdNFeCartaCorrecao.Enabled = False
            cmdImprimirDANFE.Enabled = False
         End If
         
      End If
      
      localErro = "Verifica Campos"
      
      Call Verifica_Campos


'      Set rsGerarXML = cnSistema.Execute("Select * From NFe WHERE Numero=" & Val(lvwNFs.ListItems(lvwNFs.SelectedItem.Index)))
'      If Not rsGerarXML.EOF Then
''         txtNota.text = StrZero(rsGerarXML!Numero, 8)
''         txtChaveNFe.text = FFormataChaveNF(IIf(IsNull(rsGerarXML!ChaveNFe), "", rsGerarXML!ChaveNFe))
'
''         Call oPreencherRs.PreencherRsNFs(rsGerarXML!Numero, "55")
'
'         If rsGerarXML!Situacao = 0 Then
'            cmdTransmitir.Enabled = True
''            cmdValidar.Enabled = True
'         Else
'            cmdTransmitir.Enabled = False
''            cmdValidar.Enabled = False
'         End If
'
'         If rsGerarXML!Situacao = 2 Then
'            If Trim(rsGerarXML!ChaveNFe) <> "" And Trim(rsGerarXML!Protocolo) <> "" Then
'               cmdTransmitir.Enabled = False
'               cmdCancelarNota.Enabled = True
'               cmdImprimirDANFE.Enabled = True
'            Else
'               cmdTransmitir.Enabled = True
'               cmdCancelarNota.Enabled = False
'               cmdImprimirDANFE.Enabled = False
'            End If
'         End If
'      End If
   End If
   
   Exit Sub
Erro:
   Call FRetornaMensagens(Err.Number & " - " & localErro & " - " & Err.Description & " - " & TypeName(Me))
End Sub

Private Sub mnuImportar_Click()
Dim oImportar As New CImportarLE

   If MsgBox("Confirma importa��o das notas", vbYesNo + vbQuestion, "Confirma��o") = vbYes Then
''      Call oImportar.InserirNF
   End If
   
End Sub

Private Sub mnuImportarLink_Click()
Dim oImportar As New CImportarLE

   If MsgBox("Confirma importa��o das notas", vbYesNo + vbQuestion, "Confirma��o") = vbYes Then
      Call oImportar.InserirNF
      Carrega_View ("Carregar")
   End If

End Sub

Private Sub mnuOpcoes_Click()
   frmOpcoes.Show vbModal
End Sub

Private Sub tmrAtualiza_Timer()
On Error GoTo Erro
Dim sVerArquivo As String
Dim oImportar As New CImportar
Dim rsNF As New ADODB.Recordset

Dim handle As Integer
Dim Linha As String
Dim strMensagem As String
Dim nProtocolo As String
Dim sSql As String

' Carregar Classes
Dim oPreencherRs As New PreencherRS
Dim oNF400 As New CNF400


   StatusBar.Panels(1).Text = "VERIFICANDO ATUALIZA��ES..."

''   Carrega_View ("Atualizar")
   
   'Inserir um verificador de atualiza��o. Caso um processo como cancelamento ou autoriza��o seja localizado. Atualizar o Carrega_View
   
   ' Verifica se existe Cupom a ser importado
   sVerArquivo = Dir(I_UnidadeNFe & "NFC-e\Notas\CUPOM.TXT")
   If sVerArquivo = "CUPOM.TXT" Then
      Call oImportar.FImportarCupom
      Kill I_UnidadeNFe & "NFC-e\Notas\CUPOM.TXT"
   End If
   
   ' Verifica��o de transmi��o de Nota
   Call Verificar_Transmitir
   
   'Executar verifica��o de impress�o de Arquivo Texto
   
   'Verificar se algum arquivo de retorno foi criado
   Call Verifica_Retornos

   ' Atuazaliza base de dados
   
'''''   bRetorno = False  '' Verifica se houve algum para recarregar a view

   ' 0. Em Digita��o
   ' 1. Processamento
   ' 2. Aprovada
   ' 3. Cancelada
   ' 4. N�o Emitida
   ' 5. Denegada
   ' 9. Pendentes

   If (IsDate(mskDataInicial.Text) And IsDate(mskDataFinal.Text)) And (CDate(mskDataFinal.Text) >= CDate(mskDataInicial.Text)) Then
      If I_SGBD = "SQLSERVER" Then
         Set rsNF = cnSistema.Execute("SELECT * FROM " & I_TabelasNF & " WHERE DataEmissao >= '" & Format(mskDataInicial.Text, "yyyy-mm-dd") & " 00:00:00' AND DataEmissao <= '" & Format(mskDataFinal.Text, "yyyy-mm-dd") & " 23:59:59' AND Situacao IN (1,3) Order By Numero")
      ElseIf I_SGBD = "ACCESS" Then
         Set rsNF = cnSistema.Execute("SELECT * FROM " & I_TabelasNF & " WHERE DataEmissao >= cDate('" & Format(mskDataInicial.Text, "dd/mm/yyyy") & " 00:00:00') AND DataEmissao <= cDate('" & Format(mskDataFinal.Text, "dd/mm/yyyy") & " 23:59:59') AND Situacao IN (1,3) Order By Numero")
      End If
      
      Do While Not rsNF.EOF
         ' Carregar dados da NF
         Call oPreencherRs.PreencherRsNFs(rsNF!Numero, "NFe")
      
         ' Verificar se XML foi atualizado
'         Call oNF400.ConverterTXT_XML(rsNFCe!idNFCe, rsNFCe!Numero, rsNFCe!DataEmissao, rsNFCe!ChaveNFCe)

         ' Verificar se XML foi autorizado
         If Not rsNFs.EOF Then Call oNF400.FVerificaAprovacaoXML(rsNFs!nnf, rsNFs!DataEmissao, rsNFs!cNF, rsNFs!Situacao)

         ' Verifica se XML de Cancelamento n�o possui erro
         If Not rsNFs.EOF Then Call oNF400.FVerificaErroCancelamento(rsNFs!cNF)

         ' Atualiza Protocolo
         If (IsNull(rsNF!Protocolo) Or Trim(rsNF!Protocolo) = "") Then
            If Not rsNFs.EOF Then Call oNF400.FAtualizarProtocolo(rsNFs!nnf, rsNFs!DataEmissao, rsNFs!cNF)
         End If

         ' Verifica se cancelamento foi autorizado
         If Not rsNFs.EOF Then Call oNF400.FVerificaAutorizacaoCancelamento(rsNFs!nnf, rsNFs!DataEmissao, rsNFs!cNF, rsNFs!Situacao)

         rsNF.MoveNext
      Loop
   End If

   StatusBar.Panels(1).Text = ""
   
   Exit Sub
Erro:
   Call FRetornaMensagens(Err.Number & " - " & Err.Description & " - " & TypeName(Me))
End Sub

Private Sub Verificar_Transmitir()
Dim sVerArquivo As String
Dim oImportar As New CImportar
Dim rsNF As New ADODB.Recordset

Dim handle As Integer
Dim Linha As String
Dim strMensagem As String
Dim nProtocolo As String
Dim sSql As String

' Carregar Classes
Dim oPreencherRs As New PreencherRS
Dim oNF400 As New CNF400

   ' Verifica se existe NF Em Processamento
   If I_SGBD = "SQLSERVER" Then
      Set rsNF = cnSistema.Execute("SELECT TOP 1 * FROM " & I_TabelasNF & " WHERE DataEmissao >= '" & Format(mskDataInicial.Text, "yyyy-mm-dd") & " 00:00:00' AND DataEmissao <= '" & Format(mskDataFinal.Text, "yyyy-mm-dd") & " 23:59:59' AND Situacao = 1 Order By Numero")
   ElseIf I_SGBD = "ACCESS" Then
      Set rsNF = cnSistema.Execute("SELECT TOP 1 * FROM " & I_TabelasNF & " WHERE DataEmissao >= cDate('" & Format(mskDataInicial.Text, "dd/mm/yyyy") & " 00:00:00') AND DataEmissao <= cDate('" & Format(mskDataFinal.Text, "dd/mm/yyyy") & " 23:59:59') AND Situacao = 1 Order By Numero")
   End If

   If Not rsNF.EOF Then
      ' Carregar dados da NF
      Call oPreencherRs.PreencherRsNFs(rsNF!Numero, "NFe")
      If Not rsNFs.EOF Then
         ' Verificar se tentativas < 10 e retransmitir
         If Val(IIf(IsNull(rsNF!TentativaEmissao), 0, rsNF!TentativaEmissao)) < Val(LerArquivoINI("NFe", "TentativasEmissao", App.Path & "\System.ini")) Then
            cnSistema.Execute "Update " & I_TabelasNF & " set " & _
                     "TentativaEmissao = " & (IIf(Not IsNull(rsNF!TentativaEmissao), rsNF!TentativaEmissao, 0) + 1) & " " & _
                     "Where Numero = " & rsNF!Numero
                     
            frmGerenciarNF.lblProcessamento.Caption = " Tentativa de Emiss�o da NF " & rsNF!Numero & " N� " & (rsNF!TentativaEmissao + 1)
         Else
            frmGerenciarNF.lblProcessamento.Caption = ""
            
            ' Verificar se tentativas = 10
            '' Se Sim alterar Situa��o N�o Emitida
            Call FAtualizaNF(rsNF!Numero, 4) ' N�o emitida
            
            ' Recarregar Situa��o das notas
            Carrega_View ("Carregar")
         End If
         
         ' Verifica se o XML foi validado para transmitir a nota
         ''
         ''' Verificar se a nota � uma NFC-e para Gerar o Qr-Code
         ''
         sVerArquivo = Dir(I_UnidadeNFe & I_PastaUNINFe & I_EmpresaNF & "\Validar\Validado\" & rsNFs!cNF & "-nfe.XML")
         If sVerArquivo = (rsNFs!cNF & "-nfe.XML") Then
            Call FArquivosNF("TRANSMITIR")
         End If
         
         '' Se N�o aumentar 1 em tentativas e retransmitir
      End If
   Else
      ' Verificar se existe nota Em digita��o e se Transmitir Automaticamente est� ativo
      '' Se sim Executar a transmiss�o
      If LerArquivoINI("NFe", "AutoTransmitir", App.Path & "\System.ini") = 1 Then
         ' Verifica se existe NF Em Processamento
         If I_SGBD = "SQLSERVER" Then
            Set rsNF = cnSistema.Execute("SELECT TOP 1 * FROM " & I_TabelasNF & " WHERE DataEmissao >= '" & Format(mskDataInicial.Text, "yyyy-mm-dd") & " 00:00:00' AND DataEmissao <= '" & Format(mskDataFinal.Text, "yyyy-mm-dd") & " 23:59:59' AND Situacao = 0 Order By Numero")
         ElseIf I_SGBD = "ACCESS" Then
            Set rsNF = cnSistema.Execute("SELECT TOP 1 * FROM " & I_TabelasNF & " WHERE DataEmissao >= cDate('" & Format(mskDataInicial.Text, "dd/mm/yyyy") & " 00:00:00') AND DataEmissao <= cDate('" & Format(mskDataFinal.Text, "dd/mm/yyyy") & " 23:59:59') AND Situacao = 0 Order By Numero")
         End If
         
         If Not rsNF.EOF Then
            ' Desativa verifica��o at� o fim da transmiss�o
'''            tmrAtualiza.Enabled = False
         
            ' Carregar dados da NF
            Call oPreencherRs.PreencherRsNFs(rsNF!Numero, "NFe")
            If Not rsNFs.EOF Then
               ' Gerar o arquivo XML
   '            Open ARQUIVO_NFE_NOTAS For Output As #1
               Call NotasNFs(rsNFs!nnf)
   '            Close #1
            
               ' Validar e Assinar o XML
               Call FArquivosNF("VALIDAR")
               
               ' Define como Gerada
               Call FAtualizaNF(rsNFs!nnf, 1) ' Processamento
            
               ' Recarregar Situa��o das notas
               Carrega_View ("Carregar")
            End If
            
            ' Reativa verifica��o ap�s o fim da transmiss�o
'''            tmrAtualiza.Enabled = True
         End If
      End If
   
   End If
   
   
   ' Verificar QR-CODE
   
   
'''''   ' Transmitir automaticamente as notas do NFC-e
'''''   If LerArquivoINI("NFe", "AutoTransmitir", App.Path & "\System.ini") = 1 Then
''''''      Call TransmitirAutomatico
'''''   End If
'''''
'''''   'Gerar o QrCode e Transmitir a NF caso ela tenho sido validada com sucesso
'''''''   Call TransmitirNF


End Sub

Private Sub TransmitirNF()

'''''   'Gerar o QrCode e Transmitir a NF caso ela tenho sido validada com sucesso
'''''   ''' Verificar o arquivo validado existe e se a NF � 1NFC-e
'''''   Set rsGerarXML = cnSistema.Execute("Select * From " & I_TabelasNF & " WHERE Situacao = 1") ' Em processamento
'''''   If Not rsGerarXML.EOF Then
'''''      If I_ModeloNF = "55" Then        ' NF-e
'''''         ' Transmitir
'''''         sChaveNFe = rsGerarXML!ChaveNFCe
'''''         sVerArquivo = Dir(I_UnidadeNFe & "NFC-e\" & I_EmpresaNF & "\Validar\Validado\" & sChaveNFe & "-nfe.XML")
'''''         If sVerArquivo = (sChaveNFe & "-nfe.XML") Then
'''''            Call FArquivosNF("TRANSMITIR")
'''''         End If
'''''
'''''      ElseIf I_ModeloNF = "65" Then        ' NFC-e
'''''         If Not IsNull(rsGerarXML!ChaveNFCe) Then
'''''            sChaveNFe = rsGerarXML!ChaveNFCe
'''''            sVerArquivo = Dir(I_UnidadeNFe & "NFC-e\" & I_EmpresaNF & "\Validar\Validado\" & sChaveNFe & "-nfe.XML")
'''''            If sVerArquivo = (sChaveNFe & "-nfe.XML") Then
'''''               GerarQrCode (rsGerarXML!Numero)
'''''
'''''               ' Transmitir
'''''               Call FArquivosNF("TRANSMITIR")
'''''            Else
'''''               If rsGerarXML!TentativaEmissao < 15 Then
'''''                  cnSistema.Execute "Update NFCe set " & _
'''''                           "TentativaEmissao = " & (rsGerarXML!TentativaEmissao + 1) & " " & _
'''''                           "Where idNFCe = " & rsGerarXML!idNFCe
'''''               Else
'''''
'''''                  Call FAtualizaNF(rsGerarXML!Numero, 4) ' N�o emitida
'''''
''''''                  cnSistema.Execute "Update NFCe set " & _
''''''                           "Situacao = 4 " & _
''''''                           "Where idNFCe = " & rsGerarXML!idNFCe
'''''               End If
'''''            End If
'''''         End If
'''''      End If
'''''   End If
'''''
'''''   Set rsGerarXML = Nothing

End Sub

Private Sub cmdConsultaNota_Click()
On Error GoTo Erro
Dim sMotivo
Dim iNumeroConsultar As Long

'''   tmrAtualiza.Enabled = False
   If MsgBox("Confirma Consulta NF-e", vbYesNo + vbQuestion, "Confirma��o") = vbYes Then
      iNumeroConsultar = Val(lvwNFs.ListItems(lvwNFs.SelectedItem.Index))
      Dim oNFe400 As New CNF400
       
      Call oNFe400.FConsultarNumero(iNumeroConsultar)
      
      Carrega_View ("Carregar")
   End If
'''   tmrAtualiza.Enabled = True

   Exit Sub
Erro:
   Call FRetornaMensagens(Err.Number & " - " & Err.Description & " - " & TypeName(Me))
End Sub

Private Sub cmdInutilizarNumeracao_Click()
On Error GoTo Erro

   lblStatusEvento.Tag = "IN"
   frmEventos.Show vbModal

   Exit Sub
Erro:
   Call FRetornaMensagens(Err.Number & " - " & Err.Description & " - " & TypeName(Me))
End Sub

Private Sub cmdCancelarNota_Click()
On Error GoTo Erro

   cmdCancelarNota.Tag = Val(lvwNFs.ListItems(lvwNFs.SelectedItem.Index))
   lblStatusEvento.Tag = "CN"
   frmEventos.Show vbModal
   
   Exit Sub
Erro:
   Call FRetornaMensagens(Err.Number & " - " & Err.Description & " - " & TypeName(Me))
End Sub

Private Sub cmdNFeCartaCorrecao_Click()
On Error GoTo Erro

   cmdNFeCartaCorrecao.Tag = Val(lvwNFs.ListItems(lvwNFs.SelectedItem.Index))
   lblStatusEvento.Tag = "CC"
   frmEventos.Show vbModal

   Exit Sub
Erro:
   Call FRetornaMensagens(Err.Number & " - " & Err.Description & " - " & TypeName(Me))
End Sub

Private Sub cmdPesquisar_Click()
On Error GoTo Erro

   Carrega_View ("Carregar")

'   TerminateThread hThread1, 0
'   CloseHandle hThread1

   ' Inicia a rotina "Tarefa1" em multitarefa
'''   I_ModoView = "Carregar"
'''   hThread1 = CreateThread(ByVal 0&, ByVal 0&, AddressOf Carrega_Notas, ByVal 0&, ByVal 0&, hThread1_ID)

   ' Finaliza a execu��o dos Threads
'   TerminateThread hThread1, 0
'   CloseHandle hThread1

   Exit Sub
Erro:
   Call FRetornaMensagens(Err.Number & " - " & Err.Description & " - " & TypeName(Me))
End Sub

Private Sub cmdTransmitir_Click()
On Error GoTo Erro
'''''Dim sVerArquivo As String
'''''Dim oImportar As New CImportar
Dim iNumeroNota As Long

   If Not Verifica_Campos() Then Exit Sub

   ' Popular n�mero da nota
   iNumeroNota = FNumeroNF()
   If iNumeroNota > 0 Then
      ' Desativa verifica��o at� o fim da transmiss�o
'''      frmGerenciarNF.tmrAtualiza.Enabled = False
      
      ' Gerar o arquivo XML
      Call NotasNFs(iNumeroNota)
   
      ' Validar e Assinar o XML
      Call FArquivosNF("VALIDAR")
      
      ' Define como Gerada
      Call FAtualizaNF(iNumeroNota, 1) ' Processamento
      
      ' Atualiza visualiza��o
      Carrega_View ("Carregar")
      
      ' Reativa verifica��o ap�s o fim da transmiss�o
'''      frmGerenciarNF.tmrAtualiza.Enabled = True
   End If
   
   Exit Sub
Erro:
   Call FRetornaMensagens(Err.Number & " - " & Err.Description & " - " & TypeName(Me))
End Sub

Private Function NotasNFs(iNumero As Long)
On Error GoTo Erro
Dim oNFe400 As New CNF400
Dim oPreencherRs As New PreencherRS

   If Not rsNFs.EOF Then
      Open ARQUIVO_NFE_NOTAS For Output As #1
      
      ' Cabe�alho
      Print #1, "<?xml version=" & """" & "1.0" & """" & " encoding=" & """" & "utf-8" & """" & "?>"
      Print #1, "<NFe xmlns=" & """" & "http://www.portalfiscal.inf.br/nfe" & """" & ">"
      Print #1, Space(2) & "<infNFe versao=" & """" & "4.00" & """" & " Id=" & """" & "NFe" & rsNFs!cNF & """" & ">"
'      Print #1, Space(2) & "<infNFe versao=" & """" & "3.10" & """" & " Id=" & """" & "NFe" & sChaveNFe & """" & ">"
    
      ' Corpo do XML
      ' B. Identifica��o da Nota Fiscal eletr�nica
      Call oNFe400.NotasNFsGrupoB(iNumero)
      
      ' C. Identifica��o do Emitente da Nota Fiscal eletr�nica
      Call oNFe400.NotasNFsGrupoC
      
      ' E. Identifica��o do Destinat�rio da Nota Fiscal eletr�nica
      Call oNFe400.NotasNFsGrupoE(iNumero)
      
'''Ajustar
      ' H. Detalhamento de Produtos e Servi�os da NF-e
      Call oNFe400.NotasNFsGrupoH(iNumero)
      
      ' W. Total de NF-e
      Call oNFe400.NotasNFsGrupoW(iNumero)
      
      ' X. Dados do Transporte
      Call oNFe400.NotasNFsGrupoX(iNumero)
      
      ' Y. Dados da Cobran�a
      Call oNFe400.NotasNFsGrupoY(iNumero)
      
      ' Z. Informa��es Adicionais
      Call oNFe400.NotasNFsGrupoZ(iNumero)
      
      ' -------------------------------------------
      
      Print #1, Space(2) & "</infNFe>"
      
      Print #1, "</NFe>"
     
      Close #1
   End If
   
'   Exit Function
'Erro:
'    MsgBox "Erro " & Err & ". " & Err.Description & " - " & TypeName(Me) & ".NotasNFs"
   Exit Function
Erro:
   Call FRetornaMensagens(Err.Number & " - " & Err.Description & " - " & TypeName(Me) & ".NotasNFs")
End Function

Private Function FNumeroNF() As Long
On Error GoTo Erro

   If lvwNFs.ListItems.Count > 0 Then
      FNumeroNF = Val(lvwNFs.ListItems(lvwNFs.SelectedItem.Index))
   Else
      FNumeroNF = 0
   End If

   Exit Function
Erro:
   Call FRetornaMensagens(Err.Number & " - " & Err.Description & " - " & TypeName(Me) & ".FNumeroNF")
End Function

Private Function FValidarPeriodo(DataInicial As String, DataFinal As String) As Boolean
On Error GoTo Erro
Dim strMensagem As String

   FValidarPeriodo = True

   If Not IsDate(mskDataInicial.Text) Then strMensagem = "Data Inicial Inv�lida" & Chr(13)
   If Not IsDate(mskDataFinal.Text) Then strMensagem = "Data Final Inv�lida" & Chr(13)
   If IsDate(mskDataInicial.Text) And IsDate(mskDataFinal.Text) Then
      If (CDate(mskDataInicial.Text) > CDate(mskDataFinal.Text)) Then strMensagem = "Data Inicial maior que Data Final"
   End If
   If Not strMensagem = Empty Then
      MsgBox "Verifique os Seguintes Campos:" & Chr(13) & strMensagem, vbExclamation + vbOKOnly, "Campos Obrigat�rios"
      FValidarPeriodo = False
      Exit Function
   End If

   Exit Function
   Resume
Erro:
   Call FRetornaMensagens(Err.Number & " - " & Err.Description & " - " & TypeName(Me) & ".FValidarPeriodo")
End Function

Private Function FSituacao(Parametro As String, Conteudo As Long) As String
On Error GoTo Erro

   If Parametro = "C" Then       ' Condi��o
'      If Conteudo = 0 Then FSituacao = "AND Situacao <> 2"
      If Conteudo = 0 Then FSituacao = ""
      If Conteudo = 1 Then FSituacao = " AND N.Situacao = 0"
      If Conteudo = 2 Then FSituacao = " AND N.Situacao = 1"
      If Conteudo = 3 Then FSituacao = " AND N.Situacao = 2"
      If Conteudo = 4 Then FSituacao = " AND N.Situacao = 3"
      If Conteudo = 5 Then FSituacao = " AND N.Situacao = 4"
      If Conteudo = 6 Then FSituacao = " AND N.Situacao = 5"
      If Conteudo = 7 Then FSituacao = " AND N.Situacao <> 2"
   End If
   
   If Parametro = "S" Then       ' Condi��o
      If Conteudo = 0 Then FSituacao = "Em Digita��o"
      If Conteudo = 1 Then FSituacao = "Processamento"
      If Conteudo = 2 Then FSituacao = "Aprovada"
      If Conteudo = 3 Then FSituacao = "Cancelada"
      If Conteudo = 4 Then FSituacao = "N�o Emitida"
      If Conteudo = 5 Then FSituacao = "Denegada"
      If Conteudo = 9 Then FSituacao = "Pendentes"
   End If
   
   Exit Function
Erro:
   Call FRetornaMensagens(Err.Number & " - " & Err.Description & " - " & TypeName(Me) & ".FSituacao")
End Function

Private Function FFormaPagamento(Parametro As String) As String
On Error GoTo Erro

   '01=Dinheiro
   '02=Cheque
   '03=Cart�o de Cr�dito
   '04=Cart�o de D�bito
   '05=Cr�dito Loja
   '10=Vale Alimenta��o
   '11=Vale Refei��o
   '12=Vale Presente
   '13=Vale Combust�vel
   '99=Outros
         
   Select Case Parametro
      Case "01"
         FFormaPagamento = "Dinheiro"
      Case "02"
         FFormaPagamento = "Cheque"
      Case "03"
         FFormaPagamento = "Cart�o de Cr�dito"
      Case "04"
         FFormaPagamento = "Cart�o de D�bito"
      Case "05"
         FFormaPagamento = "Cr�dito Loja"
      Case "10"
         FFormaPagamento = "Vale Alimenta��o"
      Case "11"
         FFormaPagamento = "Vale Refei��o"
      Case "12"
         FFormaPagamento = "Vale Presente"
      Case "13"
         FFormaPagamento = "Vale Combust�vel"
      Case "90"
         FFormaPagamento = "Sem Pagamento"
      Case "99"
         FFormaPagamento = "Outros"
   End Select
   
   Exit Function
Erro:
   Call FRetornaMensagens(Err.Number & " - " & Err.Description & " - " & TypeName(Me) & ".FFormaPagamento")
End Function

Private Sub cmdEnviarEmail_Click()
On Error GoTo Erro
Dim sArquivo As String
Dim sCaminho As String
Dim cdoConfiguration As CDO.Configuration
Dim cdoData As ADODB.Fields
Dim cdoMensagem As New CDO.Message
Dim strMensagem As String
   
   Set rsNFs = cnSistema.Execute("Select * From NFe WHERE Numero=" & Val(lvwNFs.ListItems(lvwNFs.SelectedItem.Index)))
   Set rsClientes = cnSistema.Execute("Select * From Clientes WHERE idCliente = " & rsNFs!idCliente)
'''''   Set rsEmpresa = cnSistema.Execute("Select * From Empresa")
   
   If Trim(rsClientes!Email) = "" Then
      MsgBox "E-mail do cliente n�o encontrado", vbExclamation + vbOKOnly, "Campos Obrigat�rios"
      Exit Sub
   End If
   
   '''''''''''''''''''''''''''''''''''''''''''''

   strMensagem = strMensagem & "De: " & rsEmpresa!Nome & Chr(13)
   strMensagem = strMensagem & "E-mail: " & rsEmpresa!Email & Chr(13)
   strMensagem = strMensagem & Chr(13)
   strMensagem = strMensagem & "Para: " & rsClientes!Nome & Chr(13)
   strMensagem = strMensagem & "E-mail: " & rsClientes!Email & Chr(13)
   strMensagem = strMensagem & Chr(13)
   strMensagem = strMensagem & "Estamos enviando em anexo arquivo XML da nota fiscal N. " & rsNFs!Numero

   '''''''''''''''''''''''''''''''''''''''''''''
   
   If MsgBox("Envio de E-mail " & Chr(13) & strMensagem, vbYesNo + vbQuestion, "Confirma��o") = vbYes Then

   '' Para Consulta
   ''Item(cdoSendUsingMethod) = 2
   ''  .Item(cdoSMTPServer) = "pop.mail.yahoo.com.br"
   ''  .Item(cdoSMTPServerPort) = 995
   ''  .Item(cdoSMTPConnectionTimeout) = 15
   ''  .Item(cdoSMTPAuthenticate) = cdoBasic
   ''  .Item(cdoSMTPUseSSL) = True
   ''  .Item(cdoSendUserName) = vUser
   ''  .Item(cdoSendPassword) = vPass
   ''  .Update
   
      sArquivo = I_CaminhoXML_NFe & I_EmpresaNF & "\Enviados\Autorizados\" & Format(rsNFs!DataEmissao, "yyyymm") & "\" & rsNFs!ChaveNFe & "-procNFe.XML"
      If Dir(sArquivo) <> "" Then
         Set cdoConfiguration = New CDO.Configuration
         Set cdoData = cdoConfiguration.Fields
'''''         Set rsEmpresa = cnSistema.Execute("SELECT * FROM Empresa")
         With cdoData
            .Item(cdoSendUsingMethod) = 2
            .Item(cdoSMTPServerPort) = LerArquivoINI("EMail", "Porta", CaminhoINI & "\System.ini")
            .Item(cdoSMTPServer) = rsEmpresa!ServidorSMTP 'rsSistema!Mail_SMTP
            .Item(cdoSMTPConnectionTimeout) = 20
            .Item(cdoSMTPAuthenticate) = 1
            .Item(cdoSMTPUseSSL) = True
            .Item(cdoSendUserName) = rsEmpresa!EmailUsuario 'rsSistema!Mail_User
            .Item(cdoSendPassword) = rsEmpresa!EmailSenha 'rsSistema!Mail_Pass
            .Update
         End With
         
         Set cdoMensagem = New CDO.Message
         With cdoMensagem
            Set .Configuration = cdoConfiguration
                .To = rsClientes!Email 'rsTotal!Email
                .From = rsEmpresa!Email 'rsSistema!Mail_From
                .Subject = rsEmpresa!Nome 'rsSistema!Mail_Subject
                .HTMLBody = "Estamos enviando em anexo arquivo XML da nota fiscal N. " & rsNFs!Numero  'rsSistema!Mail_Body
                .AddAttachment sArquivo 'rsSistema!Mail_Directory & "\" & rsTotal!idFaturamento & ".pdf"
                .Send
         End With
         
         MsgBox "E-mail Enviado", vbExclamation + vbOKOnly, "Informa��o"
      Else
         MsgBox "Arquivo: " & sArquivo & " N�o encontrado", vbExclamation + vbOKOnly, "Campos Obrigat�rios"
      End If
   End If
   
   Set cdoMensagem = Nothing

   Exit Sub
Erro:
   Call FRetornaMensagens(Err.Number & " - " & Err.Description & " - " & TypeName(Me) & ".EnviarEmail_Click")
End Sub

Private Sub cmdImprimirDANFE_Click()
On Error GoTo Erro
   
   Call FArquivosNF("IMPRIMIRDANFE")

   Exit Sub
Erro:
   Call FRetornaMensagens(Err.Number & " - " & Err.Description & " - " & TypeName(Me) & ".ImprimirDANFE_Click")
End Sub

Private Sub cmdHistorico_Click()
On Error GoTo Erro

   frmLerXML.Show

   Exit Sub
Erro:
   Call FRetornaMensagens(Err.Number & " - " & Err.Description & " - " & TypeName(Me) & ".Historico_Click")
End Sub

Private Sub cmdLimparHistorico_Click()
On Error GoTo Erro

   lvwMensagens.ListItems.Clear

   Exit Sub
Erro:
   Call FRetornaMensagens(Err.Number & " - " & Err.Description & " - " & TypeName(Me) & ".LimparHistorico_Click")
End Sub

Private Sub cmdGerarPDF_Click()
On Error GoTo Erro
Dim sArquivo As String
Dim sCaminho As String
Dim sImpressoraUSB As String

   Call FArquivosNF("GERARPDF")

'   sImpressoraUSB = LerArquivoINI("NFe", "ImpressoraUSB", CaminhoINI & "\System.ini")
'
'   Set rsNFs = cnSistema.Execute("Select * From NFe WHERE Numero=" & Val(lvwNFs.ListItems(lvwNFs.SelectedItem.Index)))
'
'   sEmpresaNFe = LerArquivoINI("NFe", "Empresa", CaminhoINI & "\System.ini")
''   sArquivo = I_Unidadenfe & "NFC-e\" & sEmpresaNFe & "\Enviados\Autorizados\" & Format(rsNFe!DataEmissao, "yyyymmdd") & "\" & rsNFe!ChaveNFe & "-procNFe.XML"
'   sArquivo = I_CaminhoXML_NFe & sEmpresaNFe & "\Enviados\Autorizados\" & Format(rsNFe!DataEmissao, "yyyymm") & "\" & rsNFe!ChaveNFe & "-procNFe.XML"
'
'   Shell "C:\UNIMAKE\" & sEmpresaNFe & "\UNIDANFE\UNIDANFE.EXE arquivo=" & sArquivo & " visualizar = 0 " & "i=" & sImpressoraUSB
''   Shell "C:\UNIMAKE\" & sEmpresaNFe & "\UNIDANFE\UNIDANFE.EXE a=" & sArquivo & " v=0 i=selecionar"

   Exit Sub
Erro:
   Call FRetornaMensagens(Err.Number & " - " & Err.Description & " - " & TypeName(Me) & ".GerarPDF_Click")
End Sub

'===========================================================================================================================================
'===========================================================================================================================================
'===========================================================================================================================================
'===========================================================================================================================================

Private Sub tmrImprimirSistema_Timer()
On Error GoTo Erro
Dim handle As Integer
Dim Linha As String
Dim iCopia As Integer
Dim Contador As Integer
Dim sVerArquivo As String
Dim OldFont As String
 
   sVerArquivo = Dir(LerArquivoINI("Arquivos", "Caminho", App.Path & "\System.ini"))
   If sVerArquivo = "IMPRIMIR.PRN" Then
      handle = FreeFile
      Open LerArquivoINI("Arquivos", "Caminho", App.Path & "\System.ini") For Input As #handle               '' Abre arquivo importado
      While Not EOF(handle)
         Line Input #handle, Linha
         If Mid(Linha, 1, 3) = "<CP" Then
            iCopia = Mid(Linha, 4, 2)
         End If
      Wend
      Close #handle

      ' Determina o n�mero m�nimo de c�pias
      If iCopia = 0 Then iCopia = 1
      
      For Contador = 1 To iCopia
         ' Alterar fonte
         OldFont = Printer.FontName            ' Preserva a fonte original.
         Printer.FontName = LerArquivoINI("Arquivos", "Fonte", App.Path & "\System.ini")
         Printer.FontSize = LerArquivoINI("Arquivos", "Tamanho", App.Path & "\System.ini")
         Printer.FontBold = True

         ' Abrir arquivo
         handle = FreeFile

         Open LerArquivoINI("Arquivos", "Caminho", App.Path & "\System.ini") For Input As #handle               '' Abre arquivo importado
         While Not EOF(handle)
            Line Input #handle, Linha
            If Mid(Linha, 1, 3) <> "<CP" Then
               Printer.Print Linha
            End If
         Wend
         Close #handle
         Printer.FontName = OldFont   ' Restaura a fonte original.
         Printer.EndDoc
      Next

      Kill LerArquivoINI("Arquivos", "Caminho", App.Path & "\System.ini")
   End If

   Exit Sub
Erro:
   Call FRetornaMensagens(Err.Number & " - " & Err.Description & " - " & TypeName(Me) & ".ImprimirSistema_Timer")
End Sub

Private Function Verifica_Campos()
Dim oNFs400 As New CNF400

   Verifica_Campos = True

   ' Verifica se Nota est� aprovada
   If Not rsNFs.EOF Then
      If IsNull(rsNFs!idCliente) Then
         Call FErrosValidacaoNF("Cliente n�o cadastrado")
      End If
      
      If rsNFs!Situacao <> 2 Then
         If Dir(I_UnidadeNFe & I_PastaUNINFe & I_EmpresaNF & "\Enviados\Autorizados\" & Format(CDate(Mid(rsNFs!dhEmi, 1, 10)), "yyyymm") & "\" & rsNFs!cNF & "-procNFe.XML") <> "" Then
         ' Verifica se a nota foi aprovada
            Call oNFs400.FVerificaAprovacaoXML(rsNFs!nnf, rsNFs!dhEmi, rsNFs!cNF, rsNFs!Situacao)
      
            ' Atualiza Protocolo
            If IsNull(rsNFs!Protocolo) Or IsEmpty(rsNFs!Protocolo) Then
               Call oNFs400.FAtualizarProtocolo(rsNFs!nnf, rsNFs!dhEmi, rsNFs!cNF)
            End If
         End If
      End If

      rsNFsItens.MoveFirst
      Do While Not rsNFsItens.EOF
         ' Verifica o NCM
         ' O NCM "00" � o NCM de Servi�os
         If Len(Trim(rsNFsItens!NCM)) <> 8 And Trim(rsNFsItens!NCM) <> "00" Then
            Call FErrosValidacaoNF("Tamanho do NCM do Produto " & Trim(rsNFsItens!xprod) & " est� incorreto")
            Verifica_Campos = False
         End If
   
         ' Testa CFOPs
         If rsNFsEmitentes!UF = rsNFsDestinatarios!UF Then
            If Mid(rsNFsItens!CFOP, 1, 1) = 2 Or Mid(rsNFsItens!CFOP, 1, 1) = 6 Then
               Call FErrosValidacaoNF("CFOP " & Trim(rsNFsItens!CFOP) & " do Produto " & Trim(rsNFsItens!xprod) & " inv�lido para o estado do cliente")
               Verifica_Campos = False
            End If
         Else
            If Mid(rsNFsItens!CFOP, 1, 1) = 1 Or Mid(rsNFsItens!CFOP, 1, 1) = 5 Then
               Call FErrosValidacaoNF("CFOP " & Trim(rsNFsItens!CFOP) & " do Produto " & Trim(rsNFsItens!xprod) & " inv�lido para o estado do cliente")
               Verifica_Campos = False
            End If
         End If
   
         rsNFsItens.MoveNext
      Loop
      rsNFsItens.MoveFirst
   End If

   If Not rsNFsDestinatarios.EOF Then
      ' Tipo de Contribuinte
      If IsNull(rsNFsDestinatarios!indIEDest) Or IsEmpty(rsNFsDestinatarios!indIEDest) Then
         Call FErrosValidacaoNF("Tipo de Contribuinte � obrigat�rio")
         Verifica_Campos = False
      End If
      
      ' Bairro
      If IsNull(rsNFsDestinatarios!xBairro) Or IsEmpty(rsNFsDestinatarios!xBairro) Then
         Call FErrosValidacaoNF("Bairro " & rsNFsDestinatarios!xBairro & " est� incorreto")
         Verifica_Campos = False
      End If
      
      ' CEP
      If IsNull(rsNFsDestinatarios!CEP) Or IsEmpty(rsNFsDestinatarios!CEP) Then
         Call FErrosValidacaoNF("O CEP " & rsNFsDestinatarios!CEP & " est� incorreto")
         Verifica_Campos = False
      ElseIf I_ModeloNF <> "65" And Len(RemoveCaracteres(rsNFsDestinatarios!CEP)) <> 8 Then
         Call FErrosValidacaoNF("O CEP " & rsNFsDestinatarios!CEP & " est� incorreto")
         Verifica_Campos = False
      End If
   End If
   
End Function

Private Sub Verifica_Retornos()
Dim Contador As Integer
'''''Dim Contador2 As Integer
Dim sStatus As String
'''''Dim sMotivos As String
Dim sProtocolo As String
Dim sNumeroNota As String
Dim handle As Integer
Dim Linha As String
Dim strMensagem As String
Dim xMotivo As String
Dim xChaveNF As String
'''''Dim bRetorno As Boolean

Dim IdMensagens As Long

''   ARQUIVO_NFE_NOTAS = I_UnidadeNFe & I_PastaUNINFe & "Notas\Notas.TXT"
''
''   CAMINHO_NFE_ENVIO = I_UnidadeNFe & I_PastaUNINFe & I_EmpresaNF & "\ENVIO\"
''   CAMINHO_NFE_VALIDAR = I_UnidadeNFe & I_PastaUNINFe & I_EmpresaNF & "\VALIDAR\"
''   CAMINHO_NFE_VALIDADO = I_UnidadeNFe & I_PastaUNINFe & I_EmpresaNF & "\VALIDAR\VALIDADO\"
''   CAMINHO_NFE_RETORNO = I_UnidadeNFe & I_PastaUNINFe & I_EmpresaNF & "\RETORNO\"
''   CAMINHO_NFE_ERROS = I_UnidadeNFe & I_PastaUNINFe & I_EmpresaNF & "\ERROS\"
''   CAMINHO_NFE_TEMP = I_UnidadeNFe & I_PastaUNINFe & I_EmpresaNF & "\TEMP\"

''FRetornaMensagens(ByVal sMensagem As String, _
''         Optional ByVal sChave As String, _
''         Optional ByVal sCodigo As String, _
''         Optional ByVal sArquivo As String)
         
   IdMensagens = 1 ' Inicia o contador de chave das mensagens

   Dim Arquivos() As String
   Dim lCtr As Long
   Arquivos = ListarArquivos(CAMINHO_NFE_RETORNO)
   If Arquivos(lCtr) <> "" Then
      For lCtr = 0 To UBound(Arquivos)
         '' Inutilizar
         If UCase(Mid(Arquivos(lCtr), Len(Trim(Arquivos(lCtr))) - 7, 8)) = "-INU.XML" Then
            Contador = 1

            ' Verifica se Inutiliza��o foi aceita
            handle = FreeFile
            Open CAMINHO_NFE_RETORNO & Arquivos(lCtr) For Input As #handle
            'Open CAMINHO_NFE_RETORNO & Trim(rsNFCeInutilizadas!ChaveNFCe) & "-inu.XML" For Input As #handle

            Line Input #handle, Linha

          ' N�mero do Protocolo
            sProtocolo = PesquisarTAG(Linha, "nProt")
          ' N�mero da Nota
            sNumeroNota = PesquisarTAG(Linha, "nNFIni")
          ' Verifica o Status
            sStatus = RemoveCaracteres(PesquisarTAG(Linha, "cStat"))
          ' Verifica o Motivo
            xMotivo = PesquisarTAG(Linha, "xMotivo")
          ' Verifica a Chave
            xChaveNF = PesquisarTAG(Linha, "chNFe")

            If Trim(sStatus) <> "" Then
               Call FRetornaMensagens(xMotivo & " " & xChaveNF, sStatus, Arquivos(lCtr))
            End If

            ' Atualiza Protocolo
            If Trim(sProtocolo) <> "" Then
               cnSistema.Execute "Update " & I_TabelasNF & "Inutilizadas set " & _
                        "Protocolo = '" & sProtocolo & "' " & _
                        "Where Numero = " & sNumeroNota

               sStatus = ""
               sProtocolo = ""
            End If

            Close #handle

            Call FArquivosNF("RETORNOS", Arquivos(lCtr))

         ElseIf UCase(Mid(Arquivos(lCtr), Len(Trim(Arquivos(lCtr))) - 3, 4)) = ".XML" Then
            Contador = 1

            handle = FreeFile
            Open CAMINHO_NFE_RETORNO & Arquivos(lCtr) For Input As #handle

            Line Input #handle, Linha
          ' Verifica o Status
            sStatus = RemoveCaracteres(PesquisarTAG(Linha, "cStat"))
          ' Verifica o Motivo
            xMotivo = PesquisarTAG(Linha, "xMotivo")
          ' N�mero do Protocolo
            sProtocolo = RemoveCaracteres(PesquisarTAG(Linha, "nRec"))
          ' Verifica a Chave
            xChaveNF = PesquisarTAG(Linha, "chNFe")

            If Trim(sStatus) <> "" Then
               Call FRetornaMensagens(xMotivo & " " & xChaveNF, sStatus, Arquivos(lCtr))
            End If
            Close #handle

            Call FArquivosNF("RETORNOS", Arquivos(lCtr))

         '' Erros
         ElseIf UCase(Mid(Arquivos(lCtr), Len(Trim(Arquivos(lCtr))) - 3, 4)) = ".ERR" Then

            handle = FreeFile
            Open CAMINHO_NFE_RETORNO & Arquivos(lCtr) For Input As #handle

            Line Input #handle, Linha

'''''            bRetorno = False
            While Not EOF(handle)
               Line Input #handle, Linha

               strMensagem = strMensagem & Linha & Chr(13)
            Wend

            MsgBox strMensagem, vbExclamation + vbOKOnly, "Erro"
            Close #handle

            Call FArquivosNF("RETORNOS", Arquivos(lCtr))

         '' Mover os TXTs para TEMP
         ElseIf UCase(Mid(Arquivos(lCtr), Len(Trim(Arquivos(lCtr))) - 3, 4)) = ".TXT" Then
            Call FArquivosNF("RETORNOS", Arquivos(lCtr))

         End If
      Next
   End If
End Sub

Private Function GerarQrCode(iNumeroNota As Long)
''On Error GoTo Erro
''
''Dim handle As Integer
''Dim Linha As String
'''Dim sDataEmissao As String
''Dim sdigestValue As String
''
''   ' Pegar o Digest Value
''
''   sEmpresaNFCe = LerArquivoINI("NFe", "Empresa", CaminhoINI & "\System.ini")
'''   Set rsNFCe = cnSistema.Execute("Select * From NFCe WHERE Numero=" & Val(lvwNFCes.ListItems(lvwNFCes.SelectedItem.Index)))
''   Set rsNFCe = cnSistema.Execute("Select * From NFCe WHERE Numero=" & iNumeroNota)
''   If Not rsNFCe.EOF Then
''      sChaveNFe = rsNFCe!ChaveNFCe
'''      sDataEmissao = Format(rsNFCe!DataEmissao, "YYYY-MM-DD") & "T" & Format(Time, "HH:MM:SS") & IIf(I_HorarioVerao, "-02:00", "-03:00")
''   End If
''
''   handle = FreeFile
''   Open I_UnidadeNFe & "NFC-e\" & sEmpresaNFCe & "\Validar\Validado\" & sChaveNFe & "-nfe.XML" For Input As #handle
''
''   While Not EOF(handle)
''      Line Input #handle, Linha
''
''      sdigestValue = PesquisarTAG(Linha, "DigestValue")
''   Wend
''   Close #handle
''
''   ' Gerar XML com o QrCode
''   Dim QrCchNFe As String
''   Dim QrCnVersao As String
''   Dim QrCtpAmb As String
''   Dim QrCcDest As String
''   Dim QrCdhEmi As String
''   Dim QrCvNF As String
''   Dim QrCvICMS As String
''   Dim QrCdigVal As String
''   Dim QrCcldToken As String
''   Dim QrCCSC As String
''   Dim QrCcHashQRCode As String
''
''   Dim strQrCode As String
''   Dim strCqCode As String
''
''   QrCchNFe = sChaveNFe
''   QrCnVersao = "100"
''   QrCtpAmb = "1"
''   QrCcDest = sCPFDestinatario
''   QrCdhEmi = StringToHex(sDataEmissao)
''   QrCvNF = sValorTotalNFCe
''   QrCvICMS = sValorTotalICMSNFCe
''   QrCdigVal = StringToHex(sdigestValue)
''   QrCcldToken = "000001"
''   QrCCSC = "F9555903-7313-4510-BD75-FB48EBA7E5DF" ' CSC 000001 - Casa Grande Motel
''
''   ' Calcular o Hash
''   strQrCode = ""
''   strQrCode = strQrCode & "chNFe=" & QrCchNFe
''   strQrCode = strQrCode & "&nVersao=" & QrCnVersao
''   strQrCode = strQrCode & "&tpAmb=" & QrCtpAmb
''   If QrCcDest <> "" Then
''      strQrCode = strQrCode & "&cDest=" & QrCcDest
''   End If
''   strQrCode = strQrCode & "&dhEmi=" & QrCdhEmi
''   strQrCode = strQrCode & "&vNF=" & QrCvNF
''   strQrCode = strQrCode & "&vICMS=" & QrCvICMS
''   strQrCode = strQrCode & "&digVal=" & QrCdigVal
''   strQrCode = strQrCode & "&cIdToken=" & QrCcldToken
''   strQrCode = strQrCode & QrCCSC
''
'''   QrCcHashQRCode = SHA1Hash.HashBytes(strQrCode)
''''   QrCcHashQRCode = StringToHex(QrCcHashQRCode)
''   QrCcHashQRCode = SHA1Hash.HashBytes(StrConv(strQrCode, vbFromUnicode))
''
''   ' Montar o novo XML incluindo a TAG do QrCod
''   strQrCode = ""
''   strQrCode = "<![CDATA[http://dec.fazenda.df.gov.br/ConsultarNFCe.aspx?"
''   strQrCode = strQrCode & "chNFe=" & QrCchNFe
''   strQrCode = strQrCode & "&nVersao=" & QrCnVersao
''   strQrCode = strQrCode & "&tpAmb=" & QrCtpAmb
''   If QrCcDest <> "" Then
''      strQrCode = strQrCode & "&cDest=" & QrCcDest
''   End If
''   strQrCode = strQrCode & "&dhEmi=" & QrCdhEmi
''   strQrCode = strQrCode & "&vNF=" & QrCvNF
''   strQrCode = strQrCode & "&vICMS=" & QrCvICMS
''   strQrCode = strQrCode & "&digVal=" & QrCdigVal
''   strQrCode = strQrCode & "&cIdToken=" & QrCcldToken
''   strQrCode = strQrCode & "&cHashQRCode=" & QrCcHashQRCode
''   strQrCode = strQrCode & "]]>"
''
''   Open I_UnidadeNFe & "NFC-e\" & sEmpresaNFCe & "\Validar\Validado\" & sChaveNFe & "-nfe.XML" For Input As #handle
''   Dim novoXML As String
''   novoXML = ""
''   While Not EOF(handle)
''      Line Input #handle, Linha
''      Dim x As Integer
''      For x = 1 To Len(Linha)
''         novoXML = novoXML & Mid(Linha, x, 1)
''         If x > 10 Then
''            If Mid(Linha, x - 8, 8) = "</infNFe" Then
''               novoXML = novoXML & "<infNFeSupl>" & "<qrCode>" & strQrCode & "</qrCode>" & "</infNFeSupl>"
''            End If
''         End If
''      Next
''   Wend
''   Close #handle
''
''   Open I_UnidadeNFe & "NFC-e\" & sEmpresaNFCe & "\Envio\" & sChaveNFe & "-nfe.XML" For Output As #1
'''   Open I_Unidadenfe & "NFC-e\" & sEmpresaNFCe & "\Validar\" & sChaveNFe & "-nfe.XML" For Output As #1
'''   Open I_Unidadenfe & "NFC-e\" & sEmpresaNFCe & "\Validar\" & sChaveNFe & "-nfe.XML" For Input As #1
''   Print #1, novoXML
''   Close #1
''
''   Kill I_UnidadeNFe & "NFC-e\" & sEmpresaNFCe & "\Validar\Validado\" & sChaveNFe & "-nfe.XML"
''
''   Exit Function
''Erro:
''   MsgBox "Erro " & Err & ". " & Err.Description & " - " & TypeName(Me) & ".GerarQrCode"
End Function

Private Sub cmdMDFe_Click()
   tmrAtualiza.Enabled = False
   frmGerenciarMDFes.Show vbModal
   tmrAtualiza.Enabled = True
End Sub

'Private Sub Command1_Click()
'   frmTransmissao.Show vbModal
'End Sub
'
'Private Sub Drive1_Change()
'   Text3.text = Drive1.Drive
'
'End Sub
'
'Private Sub Drive1_GotFocus()
'   Text3.text = Drive1.Drive
'
'End Sub
'
'Private Sub Drive1_LostFocus()
'   Text3.text = Drive1.Drive
'
'End Sub

Private Sub cmdEditarNotas_Click()
   If I_SGBD = "SQLSERVER" Then
      frmNFe.Show vbModal
   ElseIf I_SGBD = "ACCESS" Then
      frmNFeG.Show vbModal
   End If
End Sub

Private Sub cmdCopiarNFs_Click()
   frmCopiarArquivos.Show vbModal
End Sub

Private Sub cmdNFSe_Click()
'''   tmrAtualiza.Enabled = False
   frmGerenciarNFSe.Show vbModal
'''   tmrAtualiza.Enabled = True
End Sub

Private Sub cmdDestinatarios_Click()

   frmDestinatarios.Show vbModal

'   If I_SGBD = "SQLSERVER" Then
'      frmNFe.Show vbModal
'   ElseIf I_SGBD = "ACCESS" Then
'      frmNFeG.Show vbModal
'   End If
End Sub

VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmPesquisaClientes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pesquisa de Clientes"
   ClientHeight    =   5505
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7695
   Icon            =   "frmPesquisaClientes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   7695
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraFiltro 
      Caption         =   "Filtro da Pesquisa"
      Height          =   855
      Left            =   5760
      TabIndex        =   13
      Top             =   960
      Width           =   1890
      Begin VB.OptionButton optContendo 
         Caption         =   "Contendo o texto"
         Height          =   195
         Left            =   60
         TabIndex        =   15
         Top             =   540
         Width           =   1755
      End
      Begin VB.OptionButton optIniciado 
         Caption         =   "Iniciado com"
         Height          =   255
         Left            =   60
         TabIndex        =   14
         Top             =   240
         Value           =   -1  'True
         Width           =   1755
      End
   End
   Begin VB.Frame fraInformacoesPesquisa 
      Caption         =   "Informações a Pesquisar"
      Height          =   3030
      Left            =   60
      TabIndex        =   12
      Top             =   60
      Width           =   5655
      Begin VB.TextBox txtCodigoPesquisa 
         Height          =   285
         Left            =   660
         MaxLength       =   40
         TabIndex        =   1
         Top             =   240
         Width           =   795
      End
      Begin VB.TextBox txtPesquisar 
         Height          =   285
         Left            =   660
         MaxLength       =   40
         TabIndex        =   3
         Top             =   540
         Width           =   4890
      End
      Begin MSComctlLib.ListView lvwDados 
         Height          =   2115
         Left            =   60
         TabIndex        =   4
         Top             =   840
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   3731
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
         NumItems        =   0
      End
      Begin VB.Label lblNomePesquisa 
         AutoSize        =   -1  'True
         Caption         =   "Nome"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   90
         TabIndex        =   2
         Top             =   585
         Width           =   420
      End
      Begin VB.Label lblCodigoPesquisa 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   90
         TabIndex        =   0
         Top             =   285
         Width           =   495
      End
   End
   Begin VB.Frame fraInformacoesCliente 
      Caption         =   "Informações Sobre o Cliente"
      Height          =   2310
      Left            =   60
      TabIndex        =   11
      Top             =   3120
      Width           =   7575
      Begin VB.TextBox txtNomeFantasia 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         TabIndex        =   28
         Top             =   540
         Width           =   5775
      End
      Begin VB.TextBox txtIE_CI 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4320
         MaxLength       =   20
         TabIndex        =   20
         Top             =   1620
         Width           =   2715
      End
      Begin VB.TextBox txtEndereço 
         Enabled         =   0   'False
         Height          =   765
         Left            =   1680
         TabIndex        =   19
         Top             =   840
         Width           =   5775
      End
      Begin VB.TextBox txtNome 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         TabIndex        =   18
         Top             =   255
         Width           =   5775
      End
      Begin MSMask.MaskEdBox mskCNPJ_CPF 
         Height          =   285
         Left            =   1680
         TabIndex        =   21
         Top             =   1620
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox mskTelefone1 
         Height          =   285
         Left            =   1680
         TabIndex        =   24
         Top             =   1920
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox mskCelular 
         Height          =   285
         Left            =   4320
         TabIndex        =   26
         Top             =   1920
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         PromptChar      =   " "
      End
      Begin VB.Label lblNomeFantasia 
         AutoSize        =   -1  'True
         Caption         =   "Nome Fantasia"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   495
         TabIndex        =   29
         Top             =   585
         Width           =   1065
      End
      Begin VB.Label lblCelular 
         AutoSize        =   -1  'True
         Caption         =   "Celular"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3765
         TabIndex        =   27
         Top             =   1980
         Width           =   480
      End
      Begin VB.Label lblTelefone 
         AutoSize        =   -1  'True
         Caption         =   "Telefone"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   930
         TabIndex        =   25
         Top             =   1980
         Width           =   630
      End
      Begin VB.Label lblCNPJ 
         AutoSize        =   -1  'True
         Caption         =   "CNPJ/CPF"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   780
         TabIndex        =   23
         Top             =   1665
         Width           =   780
      End
      Begin VB.Label lblIdentidade 
         AutoSize        =   -1  'True
         Caption         =   "Insc.Est/CI"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3450
         TabIndex        =   22
         Top             =   1680
         Width           =   795
      End
      Begin VB.Label lblEndereco 
         AutoSize        =   -1  'True
         Caption         =   "Endereço"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   870
         TabIndex        =   17
         Top             =   840
         Width           =   690
      End
      Begin VB.Label lblNome 
         AutoSize        =   -1  'True
         Caption         =   "Nome/Razão Social"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   300
         Width           =   1440
      End
   End
   Begin VB.Frame fraPesquisar 
      Caption         =   "Pesquisar por"
      Height          =   855
      Left            =   5760
      TabIndex        =   8
      Top             =   60
      Width           =   1890
      Begin VB.OptionButton optNome 
         Caption         =   "Nome/Razão Social"
         Height          =   255
         Left            =   60
         TabIndex        =   10
         Top             =   240
         Value           =   -1  'True
         Width           =   1755
      End
      Begin VB.OptionButton optFantasia 
         Caption         =   "Nome Fantasia"
         Height          =   195
         Left            =   60
         TabIndex        =   9
         Top             =   540
         Width           =   1755
      End
   End
   Begin VB.CommandButton cmdLimparCampos 
      Caption         =   "&Limpar Campos"
      Height          =   375
      Left            =   5760
      TabIndex        =   7
      Top             =   2280
      Width           =   1875
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   5760
      TabIndex        =   6
      Top             =   2700
      Width           =   1875
   End
   Begin VB.CommandButton cmdPesquisar 
      Caption         =   "&Pesquisar"
      Default         =   -1  'True
      Height          =   375
      Left            =   5760
      TabIndex        =   5
      Top             =   1860
      Width           =   1875
   End
End
Attribute VB_Name = "frmPesquisaClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ItemList As ListItem
Dim rsTemp As New ADODB.Recordset
Dim rsPesquisa As New ADODB.Recordset

Private Sub Form_Load()
   On Error GoTo Erro
   Centraliza frmPesquisaClientes
   
   lvwDados.ColumnHeaders.Clear
   lvwDados.ColumnHeaders.Add , , "Razão Social", 5100
   
'''''   MDISistema.StatusBar.Panels(1).text = "Pesquisa Clientes"

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

Private Sub cmdPesquisar_Click()
Dim Filtro As String
   
   If Not txtCodigoPesquisa.text = Empty Then
      Filtro = "Clientes.Codigo = " & txtCodigoPesquisa.text
   Else
      If txtPesquisar.text <> Empty Then
         If optNome.value Then
            lvwDados.ColumnHeaders.Clear
            lvwDados.ColumnHeaders.Add , , "Razão Social", 5100
            If optIniciado.value Then
               Filtro = Filtro & "Clientes.Nome Like '" & SQLCheck(txtPesquisar.text) & "%' And "
            Else
               Filtro = Filtro & "Clientes.Nome Like '%" & SQLCheck(txtPesquisar.text) & "%' And "
            End If
         Else
            lvwDados.ColumnHeaders.Clear
            lvwDados.ColumnHeaders.Add , , "Nome Fantasia", 5100
            If optIniciado.value Then
               Filtro = Filtro & "Clientes.NomeFantasia Like '" & SQLCheck(txtPesquisar.text) & "%' And "
            Else
               Filtro = Filtro & "Clientes.NomeFantasia Like '%" & SQLCheck(txtPesquisar.text) & "%' And "
            End If
         End If
      End If
   End If
   
   If Filtro = Empty Then
      MsgBox "Informe Dados para Pesquisa", vbOKOnly + vbInformation, "Validação"
      txtCodigoPesquisa.SetFocus
      Exit Sub
   End If
   
   If Mid(Filtro, Len(Filtro) - 4, 5) = " And " Then Filtro = Mid(Filtro, 1, Len(Filtro) - 5)
   
   If optNome.value Then
      Filtro = Filtro & " Order By Clientes.Nome ASC"
   Else
      Filtro = Filtro & " Order By Clientes.NomeFantasia ASC"
   End If

   Set rsPesquisa = cnSistema.Execute("SELECT * FROM Clientes WHERE " & Filtro)
   If rsPesquisa.EOF Then
      MsgBox "Nenhum Item encontrado", vbOKOnly + vbInformation, "Validação"
      Exit Sub
   End If
   
   lvwDados.ListItems.Clear
   Do While Not rsPesquisa.EOF
      If optNome.value Then
         Set ItemList = lvwDados.ListItems.Add(, "R" & CStr(rsPesquisa!idCliente), rsPesquisa!Nome)
      Else
         Set ItemList = lvwDados.ListItems.Add(, "R" & CStr(rsPesquisa!idCliente), rsPesquisa!NomeFantasia)
      End If
      rsPesquisa.MoveNext
   Loop
   Screen.MousePointer = vbDefault
   txtCodigoPesquisa.SetFocus
   
End Sub

Private Sub cmdLimparCampos_Click()
   txtCodigoPesquisa.text = Empty
   txtPesquisar.text = Empty
   
   lvwDados.ListItems.Clear
   txtCodigoPesquisa.SetFocus
End Sub

Private Sub cmdCancelar_Click()
   Unload Me
End Sub

Private Sub lvwDados_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   lvwDados.SortKey = ColumnHeader.Index - 1
   If lvwDados.SortOrder = lvwAscending Then
      lvwDados.SortOrder = lvwDescending
   Else
      lvwDados.SortOrder = lvwAscending
   End If
End Sub

Private Sub lvwDados_DblClick()
   Registro_Selecionado = True
   frmPesquisaClientes.Hide
End Sub



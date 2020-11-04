VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmMDFePesquisa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pesquisa de Manifestos Eletrõnicos"
   ClientHeight    =   6255
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8430
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   8430
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPesquisar 
      Caption         =   "&Pesquisar"
      Default         =   -1  'True
      Height          =   375
      Left            =   6480
      TabIndex        =   19
      Top             =   960
      Width           =   1875
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   6480
      TabIndex        =   18
      Top             =   1800
      Width           =   1875
   End
   Begin VB.CommandButton cmdLimparCampos 
      Caption         =   "&Limpar Campos"
      Height          =   375
      Left            =   6480
      TabIndex        =   17
      Top             =   1380
      Width           =   1875
   End
   Begin VB.Frame fraProdutosEncontrados 
      Caption         =   "Itens Encontrados"
      Height          =   2400
      Left            =   60
      TabIndex        =   15
      Top             =   2280
      Width           =   6375
      Begin MSComctlLib.ListView lvwDados 
         Height          =   2055
         Left            =   60
         TabIndex        =   16
         Top             =   240
         Width           =   6225
         _ExtentX        =   10980
         _ExtentY        =   3625
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
   End
   Begin VB.Frame fraInformacoesPesquisa 
      Caption         =   "Informações a Pesquisar"
      Height          =   2130
      Left            =   60
      TabIndex        =   4
      Top             =   60
      Width           =   6375
      Begin VB.ComboBox cmbNaturezaOperacao 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   660
         Width           =   4485
      End
      Begin VB.ComboBox cmbCliente 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   990
         Width           =   4485
      End
      Begin MSMask.MaskEdBox mskData1 
         Height          =   285
         Left            =   960
         TabIndex        =   7
         Top             =   360
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "99/99/9999"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox mskData2 
         Height          =   285
         Left            =   2280
         TabIndex        =   8
         Top             =   360
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "99/99/9999"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox mskCFOP 
         Height          =   285
         Left            =   960
         TabIndex        =   9
         Top             =   1320
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   5
         Mask            =   "9.999"
         PromptChar      =   " "
      End
      Begin VB.Label lblData2 
         AutoSize        =   -1  'True
         Caption         =   "até"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1980
         TabIndex        =   14
         Top             =   390
         Width           =   225
      End
      Begin VB.Label lblData1 
         AutoSize        =   -1  'True
         Caption         =   "Período"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   315
         TabIndex        =   13
         Top             =   405
         Width           =   570
      End
      Begin VB.Label lblNaturezaOperacao 
         AutoSize        =   -1  'True
         Caption         =   "Natureza"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   735
         Width           =   645
      End
      Begin VB.Label lblCFOP 
         AutoSize        =   -1  'True
         Caption         =   "CFOP"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   465
         TabIndex        =   11
         Top             =   1365
         Width           =   420
      End
      Begin VB.Label lblCliente 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   405
         TabIndex        =   10
         Top             =   1050
         Width           =   480
      End
   End
   Begin VB.Frame fraFiltro 
      Caption         =   "Filtro da Pesquisa"
      Height          =   855
      Left            =   6480
      TabIndex        =   1
      Top             =   60
      Width           =   1890
      Begin VB.OptionButton optIniciado 
         Caption         =   "Iniciado com"
         Height          =   255
         Left            =   60
         TabIndex        =   3
         Top             =   240
         Value           =   -1  'True
         Width           =   1755
      End
      Begin VB.OptionButton optContendo 
         Caption         =   "Contendo o texto"
         Height          =   195
         Left            =   60
         TabIndex        =   2
         Top             =   540
         Width           =   1755
      End
   End
   Begin VB.Frame fraInformacoesProduto 
      Caption         =   "Informações sobre a Nota"
      Height          =   1455
      Left            =   60
      TabIndex        =   0
      Top             =   4740
      Width           =   8295
   End
End
Attribute VB_Name = "frmMDFePesquisa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ItemList As ListItem
Dim rsTemp As New ADODB.Recordset
Dim rsCFOPs As New ADODB.Recordset
Dim rsClientes As New ADODB.Recordset
Dim rsPesquisa As New ADODB.Recordset

Private Sub Form_Load()
   On Error GoTo Erro

   lvwDados.ColumnHeaders.Clear
   lvwDados.ColumnHeaders.Add , , "Nota", 800
   lvwDados.ColumnHeaders.Add , , "Data", 1100
   lvwDados.ColumnHeaders.Add , , "CFOP", 800
   lvwDados.ColumnHeaders.Add , , "Cliente", 3200

   Carrega_Combos

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

Private Sub Carrega_Combos()

   ' Naturezas de Operação
   Set rsTemp = cnSistema.Execute("Select * from NaturezasOperacao Order By Descricao")
   cmbNaturezaOperacao.Clear
   Do While Not rsTemp.EOF
      cmbNaturezaOperacao.AddItem rsTemp!Descricao
      cmbNaturezaOperacao.ItemData(cmbNaturezaOperacao.NewIndex) = rsTemp!idNaturezaOperacao
      rsTemp.MoveNext
   Loop
   
   ' Clientes
   Set rsTemp = cnSistema.Execute("Select * from Clientes Order By Nome")
   cmbCliente.Clear
   Do While Not rsTemp.EOF
      cmbCliente.AddItem rsTemp!Nome
      cmbCliente.ItemData(cmbCliente.NewIndex) = rsTemp!idCliente
      rsTemp.MoveNext
   Loop
   
   rsTemp.Close
End Sub

Private Sub cmdPesquisar_Click()
Dim Filtro As String

   If mskData1.Text <> "  /  /    " And mskData2.Text <> "  /  /    " Then
      If Not IsDate(mskData1.Text) Or Val(Mid(mskData1.Text, 7, 4)) < 1900 Then
         MsgBox "Data Inválida", vbOKOnly + vbInformation, "Validação"
         mskData1.SelStart = 0
         mskData1.SelLength = 10
         mskData1.SetFocus
         Exit Sub
      End If
      
      If Not IsDate(mskData2.Text) Or Val(Mid(mskData2.Text, 7, 4)) < 1900 Then
         MsgBox "Data Inválida", vbOKOnly + vbInformation, "Validação"
         mskData2.SelStart = 0
         mskData2.SelLength = 10
         mskData2.SetFocus
         Exit Sub
      End If
      
      Filtro = "NFe.DataEmissao >= #" & Format(mskData1.Text, "mm/dd/yyyy") & "#" & " And " & "NFe.DataEmissao <= #" & Format(mskData2.Text, "mm/dd/yyyy") & "#" & " And "
   End If

   If cmbCliente.ListIndex <> -1 Then
      Filtro = Filtro & " NFe.idCliente = " & cmbCliente.ItemData(cmbCliente.ListIndex) & " And "
   End If

   If cmbNaturezaOperacao.ListIndex <> -1 Then
      Filtro = Filtro & " NFe.idNaturezaOperacao = " & cmbNaturezaOperacao.ItemData(cmbNaturezaOperacao.ListIndex) & " And "
   End If

   If mskCFOP.Text <> " .   " Then
      Set rsCFOPs = cnSistema.Execute("Select * from CFOPs Where CFOP = '" & mskCFOP.Text & "'")
      If Not rsCFOPs.EOF Then
         Filtro = Filtro & " NFe.idCFOP = " & rsCFOPs!idCFOP & " And "
      End If
   End If
   
   If Filtro = Empty Then
      MsgBox "Informe Dados para Pesquisa", vbOKOnly + vbInformation, "Validação"
      mskData1.SetFocus
      Exit Sub
   End If
   
   If Mid(Filtro, Len(Filtro) - 4, 5) = " And " Then Filtro = Mid(Filtro, 1, Len(Filtro) - 5)
   
   Filtro = Filtro & " Order By NFe.idNFe ASC"

   Set rsPesquisa = cnSistema.Execute("SELECT * FROM NFe WHERE " & Filtro)
   If rsPesquisa.EOF Then
      MsgBox "Nenhum Item encontrado", vbOKOnly + vbInformation, "Validação"
      Exit Sub
   End If
   
   lvwDados.ListItems.Clear
   Do While Not rsPesquisa.EOF
      Set rsCFOPs = cnSistema.Execute("Select * from CFOPs Where idCFOP = " & rsPesquisa!idCFOP)
      Set rsClientes = cnSistema.Execute("Select * from Clientes Where idCliente = " & rsPesquisa!idCliente)
   
      Set ItemList = lvwDados.ListItems.Add(, "R" & CStr(rsPesquisa!idNFe), rsPesquisa!Numero)
      ItemList.SubItems(1) = rsPesquisa!DataEmissao
      ItemList.SubItems(2) = rsCFOPs!CFOP
      ItemList.SubItems(3) = rsClientes!Nome
      rsPesquisa.MoveNext
   Loop
   Screen.MousePointer = vbDefault
   
   mskData1.SetFocus
End Sub

Private Sub cmdLimparCampos_Click()
   cmbNaturezaOperacao.ListIndex = -1
   cmbCliente.ListIndex = -1
   mskData1.Text = "  /  /    "
   mskData2.Text = "  /  /    "
   mskCFOP.Text = " .   "
   
   lvwDados.ListItems.Clear
   
   mskData1.SetFocus
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
   frmNFeGPesquisa.Hide
End Sub

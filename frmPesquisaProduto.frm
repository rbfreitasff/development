VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmPesquisaProduto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pesquisa Produtos"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8430
   Icon            =   "frmPesquisaProduto.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   8430
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraInformacoesProduto 
      Caption         =   "Informações sobre o Produto"
      Height          =   1635
      Left            =   60
      TabIndex        =   25
      Top             =   4560
      Width           =   8295
   End
   Begin VB.Frame fraFotoProduto 
      Caption         =   "Foto do Produto"
      Height          =   2280
      Left            =   6480
      TabIndex        =   23
      Top             =   2280
      Width           =   1875
      Begin VB.PictureBox pctFoto 
         Height          =   1875
         Left            =   150
         ScaleHeight     =   1815
         ScaleWidth      =   1530
         TabIndex        =   24
         Top             =   255
         Width           =   1590
      End
   End
   Begin VB.CommandButton cmdPesquisar 
      Caption         =   "&Pesquisar"
      Default         =   -1  'True
      Height          =   375
      Left            =   6480
      TabIndex        =   18
      Top             =   960
      Width           =   1875
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   6480
      TabIndex        =   20
      Top             =   1800
      Width           =   1875
   End
   Begin VB.CommandButton cmdLimparCampos 
      Caption         =   "&Limpar Campos"
      Height          =   375
      Left            =   6480
      TabIndex        =   19
      Top             =   1380
      Width           =   1875
   End
   Begin VB.Frame fraProdutosEncontrados 
      Caption         =   "Produtos Encontrados"
      Height          =   2280
      Left            =   60
      TabIndex        =   21
      Top             =   2280
      Width           =   6375
      Begin MSComctlLib.ListView lvwDados 
         Height          =   1935
         Left            =   60
         TabIndex        =   22
         Top             =   240
         Width           =   6225
         _ExtentX        =   10980
         _ExtentY        =   3413
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
      TabIndex        =   0
      Top             =   60
      Width           =   6375
      Begin VB.TextBox txtAno 
         Height          =   285
         Left            =   5505
         MaxLength       =   50
         TabIndex        =   12
         Top             =   1290
         Width           =   795
      End
      Begin VB.TextBox txtAplicacao 
         Height          =   285
         Left            =   900
         MaxLength       =   50
         TabIndex        =   10
         Top             =   1290
         Width           =   4200
      End
      Begin VB.ComboBox cmbSituacao 
         Height          =   315
         ItemData        =   "frmPesquisaProduto.frx":0ABA
         Left            =   900
         List            =   "frmPesquisaProduto.frx":0AC4
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1590
         Width           =   1755
      End
      Begin VB.ComboBox cmbFabricante 
         Height          =   315
         ItemData        =   "frmPesquisaProduto.frx":0AD8
         Left            =   900
         List            =   "frmPesquisaProduto.frx":0ADA
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   660
         Width           =   2355
      End
      Begin VB.ComboBox cmbMarca 
         Height          =   315
         ItemData        =   "frmPesquisaProduto.frx":0ADC
         Left            =   3840
         List            =   "frmPesquisaProduto.frx":0ADE
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   660
         Width           =   2475
      End
      Begin VB.TextBox txtDescricao 
         Height          =   285
         Left            =   900
         MaxLength       =   50
         TabIndex        =   8
         Top             =   990
         Width           =   5400
      End
      Begin VB.ComboBox cmbClasse 
         Height          =   315
         ItemData        =   "frmPesquisaProduto.frx":0AE0
         Left            =   900
         List            =   "frmPesquisaProduto.frx":0AE2
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   330
         Width           =   5415
      End
      Begin VB.Label lblAno 
         AutoSize        =   -1  'True
         Caption         =   "Ano"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   5160
         TabIndex        =   11
         Top             =   1335
         Width           =   285
      End
      Begin VB.Label lblAplicacao 
         AutoSize        =   -1  'True
         Caption         =   "Aplicação"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   1335
         Width           =   705
      End
      Begin VB.Label lblSituacao 
         AutoSize        =   -1  'True
         Caption         =   "Situação"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   195
         TabIndex        =   13
         Top             =   1650
         Width           =   630
      End
      Begin VB.Label lblMarca 
         AutoSize        =   -1  'True
         Caption         =   "Marca"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3300
         TabIndex        =   5
         Top             =   720
         Width           =   450
      End
      Begin VB.Label lblFabricante 
         AutoSize        =   -1  'True
         Caption         =   "Fabricante"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   75
         TabIndex        =   3
         Top             =   720
         Width           =   750
      End
      Begin VB.Label lblDescricao 
         AutoSize        =   -1  'True
         Caption         =   "Descrição"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   105
         TabIndex        =   7
         Top             =   1035
         Width           =   720
      End
      Begin VB.Label lblClasse 
         AutoSize        =   -1  'True
         Caption         =   "Classe"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   360
         TabIndex        =   1
         Top             =   390
         Width           =   465
      End
   End
   Begin VB.Frame fraFiltro 
      Caption         =   "Filtro da Pesquisa"
      Height          =   855
      Left            =   6480
      TabIndex        =   15
      Top             =   60
      Width           =   1890
      Begin VB.OptionButton optIniciado 
         Caption         =   "Iniciado com"
         Height          =   255
         Left            =   60
         TabIndex        =   16
         Top             =   240
         Value           =   -1  'True
         Width           =   1755
      End
      Begin VB.OptionButton optContendo 
         Caption         =   "Contendo o texto"
         Height          =   195
         Left            =   60
         TabIndex        =   17
         Top             =   540
         Width           =   1755
      End
   End
End
Attribute VB_Name = "frmPesquisaProduto"
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
   Centraliza frmPesquisaProduto
   
   lvwDados.ColumnHeaders.Clear
   lvwDados.ColumnHeaders.Add , , "Produto", 3750
   lvwDados.ColumnHeaders.Add , , "Valor", 1200, 1
   lvwDados.ColumnHeaders.Add , , "Saldo", 1000, 1
   
   Carrega_Combos
'''''   MDISistema.StatusBar.Panels(1).text = "Pesquisa Produtos"

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

 ' Classes
   Set rsTemp = cnSistema.Execute("Select * from ClassesProdutos Order By Descricao")
   cmbClasse.Clear
   Do While Not rsTemp.EOF
      cmbClasse.AddItem rsTemp!Descricao
      cmbClasse.ItemData(cmbClasse.NewIndex) = rsTemp!idClasseProduto
      rsTemp.MoveNext
   Loop
   
 ' Fabricantes
   Set rsTemp = cnSistema.Execute("Select * from Fabricantes Order By Fabricante")
   cmbFabricante.Clear
   Do While Not rsTemp.EOF
      cmbFabricante.AddItem rsTemp!Fabricante
      cmbFabricante.ItemData(cmbFabricante.NewIndex) = rsTemp!idFabricante
      rsTemp.MoveNext
   Loop
   
 ' Marcas
   Set rsTemp = cnSistema.Execute("Select * from Marcas Order By Marca")
   cmbMarca.Clear
   Do While Not rsTemp.EOF
      cmbMarca.AddItem rsTemp!Marca
      cmbMarca.ItemData(cmbMarca.NewIndex) = rsTemp!idMarca
      rsTemp.MoveNext
   Loop
   
   rsTemp.Close
End Sub

Private Sub cmdPesquisar_Click()
Dim Filtro As String
   
   If cmbClasse.ListIndex <> -1 Then
      Filtro = Filtro & " Produtos.idClasse = " & cmbClasse.ItemData(cmbClasse.ListIndex) & " And "
   End If
   
   If cmbFabricante.ListIndex <> -1 Then
      Filtro = Filtro & " Produtos.idFabricante = " & cmbFabricante.ItemData(cmbFabricante.ListIndex) & " And "
   End If
   
   If cmbMarca.ListIndex <> -1 Then
      Filtro = Filtro & " Produtos.idMarca = " & cmbMarca.ItemData(cmbMarca.ListIndex) & " And "
   End If
   
   If cmbSituacao.ListIndex <> -1 Then
      Filtro = Filtro & " Produtos.Situacao = " & cmbSituacao.ListIndex & " And "
   End If
   
   If txtDescricao.text <> Empty Then
      If optIniciado.value Then
         Filtro = Filtro & "Produtos.Descricao Like '" & SQLCheck(txtDescricao.text) & "%' And "
      Else
         Filtro = Filtro & "Produtos.Descricao Like '%" & SQLCheck(txtDescricao.text) & "%' And "
      End If
   End If
   
   If txtAplicacao.text <> Empty Then
      Filtro = Filtro & "Produtos.Aplicacao Like '%" & SQLCheck(txtAplicacao.text) & "%' And "
   End If
   
   If txtAno.text <> Empty Then
      Filtro = Filtro & "Produtos.Ano Like '%" & SQLCheck(txtAno.text) & "%' And "
   End If
   
   If Filtro = Empty Then
      MsgBox "Informe Dados para Pesquisa", vbOKOnly + vbInformation, "Validação"
      cmbClasse.SetFocus
      Exit Sub
   End If
   
   If Mid(Filtro, Len(Filtro) - 4, 5) = " And " Then Filtro = Mid(Filtro, 1, Len(Filtro) - 5)
   
   Filtro = Filtro & " Order By Produtos.Descricao ASC"

   Set rsPesquisa = cnSistema.Execute("SELECT * FROM Produtos WHERE " & Filtro)
   If rsPesquisa.EOF Then
      MsgBox "Nenhum Item encontrado", vbOKOnly + vbInformation, "Validação"
      Exit Sub
   End If
   
   lvwDados.ListItems.Clear
   Do While Not rsPesquisa.EOF
      Set ItemList = lvwDados.ListItems.Add(, "R" & CStr(rsPesquisa!idProduto), rsPesquisa!Descricao)
      ItemList.SubItems(1) = Format(rsPesquisa!Preco, "###,##0.00")
      ItemList.SubItems(2) = Format(0, "###,##0.00")
      rsPesquisa.MoveNext
   Loop
   Screen.MousePointer = vbDefault
   cmbClasse.SetFocus
   
End Sub

Private Sub cmdLimparCampos_Click()
   cmbClasse.ListIndex = -1
   cmbFabricante.ListIndex = -1
   cmbMarca.ListIndex = -1
   txtDescricao.text = Empty
   txtAplicacao.text = Empty
   txtAno.text = Empty
   cmbSituacao.ListIndex = -1
   optIniciado.value = True
   
   lvwDados.ListItems.Clear
   
   cmbClasse.SetFocus
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
   frmPesquisaProduto.Hide
End Sub


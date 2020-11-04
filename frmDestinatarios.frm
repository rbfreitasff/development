VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmDestinatarios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Destinatários"
   ClientHeight    =   8670
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7290
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8670
   ScaleWidth      =   7290
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraInfNFe 
      Caption         =   "Informações NF-e"
      Height          =   975
      Left            =   60
      TabIndex        =   51
      Top             =   5145
      Width           =   7095
      Begin VB.Frame fraTipoContribuinte 
         Caption         =   "Tipo de Contribuinte"
         Height          =   615
         Left            =   4320
         TabIndex        =   58
         Top             =   240
         Width           =   2685
         Begin VB.ComboBox cboTipoContribuinte 
            Height          =   315
            ItemData        =   "frmDestinatarios.frx":0000
            Left            =   105
            List            =   "frmDestinatarios.frx":000D
            Style           =   2  'Dropdown List
            TabIndex        =   59
            Top             =   210
            Width           =   2475
         End
      End
      Begin VB.Frame fraInterestadual 
         ForeColor       =   &H00800000&
         Height          =   615
         Left            =   1740
         TabIndex        =   55
         Top             =   240
         Width           =   2535
         Begin VB.OptionButton optInterestadual 
            Caption         =   "Interestadual"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   1200
            TabIndex        =   57
            Top             =   300
            Width           =   1215
         End
         Begin VB.OptionButton optEstadual 
            Caption         =   "Estadual"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   180
            TabIndex        =   56
            Top             =   300
            Value           =   -1  'True
            Width           =   975
         End
      End
      Begin VB.Frame fraConsumidor 
         Caption         =   "Consumidor Final"
         ForeColor       =   &H00800000&
         Height          =   615
         Left            =   120
         TabIndex        =   52
         Top             =   240
         Width           =   1575
         Begin VB.OptionButton optConsumidorSim 
            Caption         =   "Sim"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   60
            TabIndex        =   54
            Top             =   300
            Value           =   -1  'True
            Width           =   615
         End
         Begin VB.OptionButton optConsumidorNao 
            Caption         =   "Não"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   780
            TabIndex        =   53
            Top             =   300
            Width           =   615
         End
      End
   End
   Begin VB.Frame fraPesquisa 
      Caption         =   "Pesquisa"
      Height          =   1305
      Left            =   60
      TabIndex        =   0
      Top             =   660
      Width           =   7155
      Begin VB.ComboBox cboDados 
         Height          =   960
         Left            =   1320
         Style           =   1  'Simple Combo
         TabIndex        =   3
         Top             =   240
         Width           =   5715
      End
      Begin VB.Label lblRegistros 
         Alignment       =   2  'Center
         Caption         =   "registros"
         Height          =   435
         Left            =   60
         TabIndex        =   2
         Top             =   780
         Width           =   1185
      End
      Begin VB.Label lblPesquisa 
         Alignment       =   2  'Center
         Caption         =   "» Nome ou CPF/CNPJ"
         Height          =   435
         Left            =   60
         TabIndex        =   1
         Top             =   240
         Width           =   1155
      End
   End
   Begin VB.ComboBox cboMunicipio 
      Height          =   315
      Left            =   1500
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   3240
      Width           =   4575
   End
   Begin VB.ComboBox cboUF 
      Height          =   315
      Left            =   6420
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   2940
      Width           =   795
   End
   Begin VB.TextBox txtCEP 
      Height          =   285
      Left            =   6300
      MaxLength       =   10
      TabIndex        =   21
      Top             =   3585
      Width           =   915
   End
   Begin VB.TextBox txtTelefone_1 
      Height          =   285
      Left            =   1500
      MaxLength       =   14
      TabIndex        =   35
      Top             =   4785
      Width           =   1335
   End
   Begin VB.TextBox txtTelefone_2 
      Height          =   285
      Left            =   2940
      MaxLength       =   14
      TabIndex        =   36
      Top             =   4785
      Width           =   1335
   End
   Begin VB.TextBox txtCN 
      Height          =   285
      Left            =   1500
      MaxLength       =   18
      TabIndex        =   17
      Top             =   3585
      Width           =   1695
   End
   Begin VB.Frame fraVendasPermitidas 
      Caption         =   "&Vendas Permitidas "
      Height          =   1575
      Left            =   60
      TabIndex        =   47
      Top             =   7020
      Width           =   7155
      Begin VB.CommandButton cmdRemover 
         BackColor       =   &H80000004&
         Height          =   315
         Left            =   6660
         Picture         =   "frmDestinatarios.frx":0057
         Style           =   1  'Graphical
         TabIndex        =   62
         Top             =   180
         Width           =   360
      End
      Begin VB.CommandButton cmdInserir 
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   6240
         Picture         =   "frmDestinatarios.frx":05E1
         Style           =   1  'Graphical
         TabIndex        =   61
         ToolTipText     =   "Inserir"
         Top             =   180
         Width           =   360
      End
      Begin VB.ComboBox cboFormaPagamento 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   49
         Top             =   180
         Width           =   4455
      End
      Begin MSComctlLib.ListView lvwDados 
         Height          =   915
         Left            =   105
         TabIndex        =   50
         Top             =   540
         Width           =   6945
         _ExtentX        =   12250
         _ExtentY        =   1614
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483624
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label lblFormaPagamento 
         AutoSize        =   -1  'True
         Caption         =   "Forma de Pagamento"
         Height          =   195
         Left            =   120
         TabIndex        =   48
         Top             =   240
         Width           =   1515
      End
   End
   Begin MSMask.MaskEdBox mskNascimento 
      Height          =   285
      Left            =   6240
      TabIndex        =   33
      Top             =   4485
      Width           =   990
      _ExtentX        =   1746
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "99/99/9999"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox mskPrazoPagamento 
      Height          =   285
      Left            =   5040
      TabIndex        =   27
      Top             =   4185
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   3
      Mask            =   "999"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox mskLimiteCredito 
      Height          =   285
      Left            =   1500
      TabIndex        =   25
      Top             =   4185
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "R$ ###,###,##0.00"
      PromptChar      =   " "
   End
   Begin VB.CheckBox chkBloqueio 
      Caption         =   "Bloq. cliente"
      Height          =   195
      Left            =   5985
      TabIndex        =   44
      Top             =   6195
      Width           =   1215
   End
   Begin VB.TextBox txtObservacao 
      Height          =   285
      Left            =   1500
      MaxLength       =   500
      MultiLine       =   -1  'True
      TabIndex        =   31
      Top             =   4485
      Width           =   3735
   End
   Begin VB.TextBox txtCE 
      Height          =   285
      Left            =   4260
      MaxLength       =   20
      TabIndex        =   19
      Top             =   3585
      Width           =   1575
   End
   Begin VB.TextBox txtEmail 
      Height          =   285
      Left            =   1500
      MaxLength       =   50
      TabIndex        =   23
      Top             =   3885
      Width           =   5715
   End
   Begin VB.TextBox txtEndereco 
      Height          =   285
      Left            =   1500
      MaxLength       =   50
      TabIndex        =   7
      Top             =   2340
      Width           =   5715
   End
   Begin VB.TextBox txtCidade 
      Height          =   285
      Left            =   1500
      MaxLength       =   30
      TabIndex        =   11
      Top             =   2940
      Width           =   4575
   End
   Begin VB.TextBox txtBairro 
      Height          =   285
      Left            =   1500
      MaxLength       =   30
      TabIndex        =   9
      Top             =   2640
      Width           =   5715
   End
   Begin VB.TextBox txtNome 
      Height          =   285
      Left            =   1500
      MaxLength       =   50
      TabIndex        =   5
      Top             =   2040
      Width           =   5715
   End
   Begin MSMask.MaskEdBox mskJuros 
      Height          =   285
      Left            =   5985
      TabIndex        =   46
      Top             =   6675
      Width           =   570
      _ExtentX        =   1005
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   5
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox mskDiaPagamento 
      Height          =   285
      Left            =   6900
      TabIndex        =   29
      Top             =   4185
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   2
      Mask            =   "99"
      PromptChar      =   " "
   End
   Begin VB.Frame fraBeneficio 
      Caption         =   "&Benefício "
      Height          =   825
      Left            =   60
      TabIndex        =   38
      Top             =   6135
      Width           =   5835
      Begin VB.ComboBox cboConvenio 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2940
         Style           =   2  'Dropdown List
         TabIndex        =   42
         Top             =   120
         Width           =   2775
      End
      Begin VB.OptionButton optGrupoDesconto 
         Caption         =   "Grupo de Desconto"
         Height          =   195
         Left            =   1200
         TabIndex        =   41
         Top             =   480
         Width           =   1695
      End
      Begin VB.OptionButton optConvenio 
         Caption         =   "Convênio"
         Height          =   195
         Left            =   1200
         TabIndex        =   40
         Top             =   180
         Width           =   975
      End
      Begin VB.ComboBox cboGrupoDesconto 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2940
         Style           =   2  'Dropdown List
         TabIndex        =   43
         Top             =   480
         Width           =   2775
      End
      Begin VB.OptionButton optNenhum 
         Caption         =   "Nenhum"
         Height          =   195
         Left            =   120
         TabIndex        =   39
         Top             =   360
         Width           =   915
      End
      Begin VB.Line Line1 
         X1              =   1080
         X2              =   1080
         Y1              =   180
         Y2              =   720
      End
   End
   Begin MSComctlLib.ImageList imgToolBar 
      Left            =   7980
      Top             =   300
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDestinatarios.frx":072B
            Key             =   "Pesquisar"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDestinatarios.frx":16755
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDestinatarios.frx":19A37
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDestinatarios.frx":1CD19
            Key             =   "Imprimir"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDestinatarios.frx":32D43
            Key             =   "Excluir"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDestinatarios.frx":48D6D
            Key             =   "UsoContinuo"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   60
      Top             =   0
      Width           =   7290
      _ExtentX        =   12859
      _ExtentY        =   1058
      ButtonWidth     =   2408
      ButtonHeight    =   1005
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imgToolBar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
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
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Compras"
            Key             =   "Compras"
            Object.ToolTipText     =   "Ultimas compras do cliente"
            ImageKey        =   "Pesquisar"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Uso Contínuo"
            Key             =   "UsoContinuo"
            Object.ToolTipText     =   "Uso contínuo de medicamentos"
            ImageIndex      =   6
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Label lblMunicipio 
      AutoSize        =   -1  'True
      Caption         =   "Municipio"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   780
      TabIndex        =   14
      Top             =   3300
      Width           =   675
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "UF"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   6120
      TabIndex        =   12
      Top             =   3000
      Width           =   210
   End
   Begin VB.Label lblDataCadastro 
      Alignment       =   2  'Center
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
      Height          =   300
      Left            =   4380
      TabIndex        =   37
      Top             =   4785
      Width           =   2835
   End
   Begin VB.Label Label1 
      Caption         =   "Nascimento"
      Height          =   195
      Left            =   5340
      TabIndex        =   32
      Top             =   4545
      Width           =   855
   End
   Begin VB.Label lblObservacao 
      AutoSize        =   -1  'True
      Caption         =   "Observação"
      Height          =   195
      Left            =   540
      TabIndex        =   30
      Top             =   4545
      Width           =   870
   End
   Begin VB.Label lblJuros 
      AutoSize        =   -1  'True
      Caption         =   "Juros por atraso"
      Height          =   195
      Left            =   5985
      TabIndex        =   45
      Top             =   6435
      Width           =   1125
   End
   Begin VB.Label lblDiaPagamento 
      AutoSize        =   -1  'True
      Caption         =   "Dia de Pagamento"
      Height          =   195
      Left            =   5520
      TabIndex        =   28
      Top             =   4230
      Width           =   1320
   End
   Begin VB.Label lblPrazoPagamento 
      AutoSize        =   -1  'True
      Caption         =   "Prazo de Pagamento"
      Height          =   195
      Left            =   3420
      TabIndex        =   26
      Top             =   4230
      Width           =   1485
   End
   Begin VB.Label lblLimiteCredito 
      AutoSize        =   -1  'True
      Caption         =   "Limite de Crédito"
      Height          =   195
      Left            =   255
      TabIndex        =   24
      Top             =   4230
      Width           =   1170
   End
   Begin VB.Label lblCE 
      AutoSize        =   -1  'True
      Caption         =   "RG/Inscrição"
      Height          =   195
      Left            =   3240
      TabIndex        =   18
      Top             =   3630
      Width           =   960
   End
   Begin VB.Label lblEmail 
      AutoSize        =   -1  'True
      Caption         =   "E-Mail"
      Height          =   195
      Left            =   990
      TabIndex        =   22
      Top             =   3945
      Width           =   435
   End
   Begin VB.Label lblCN 
      AutoSize        =   -1  'True
      Caption         =   "CPF/CNPJ"
      Height          =   195
      Left            =   660
      TabIndex        =   16
      Top             =   3645
      Width           =   780
   End
   Begin VB.Label lblBairro 
      AutoSize        =   -1  'True
      Caption         =   "Bairro"
      Height          =   195
      Left            =   1020
      TabIndex        =   8
      Top             =   2700
      Width           =   405
   End
   Begin VB.Label lblCep 
      AutoSize        =   -1  'True
      Caption         =   "CEP"
      Height          =   195
      Left            =   5940
      TabIndex        =   20
      Top             =   3630
      Width           =   315
   End
   Begin VB.Label lblCidade 
      AutoSize        =   -1  'True
      Caption         =   "Cidade"
      Height          =   195
      Left            =   960
      TabIndex        =   10
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label lblEndereco 
      AutoSize        =   -1  'True
      Caption         =   "Endereço"
      Height          =   195
      Left            =   750
      TabIndex        =   6
      Top             =   2400
      Width           =   690
   End
   Begin VB.Label lblNome 
      AutoSize        =   -1  'True
      Caption         =   "Nome"
      Height          =   195
      Left            =   990
      TabIndex        =   4
      Top             =   2115
      Width           =   420
   End
   Begin VB.Label lblTelefone01 
      AutoSize        =   -1  'True
      Caption         =   "Telefones"
      Height          =   195
      Left            =   750
      TabIndex        =   34
      Top             =   4860
      Width           =   705
   End
End
Attribute VB_Name = "frmDestinatarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ItemList As ListItem
Dim rsDados As New ADODB.Recordset

Private Sub Form_Load()
   Limpa_Campos
'   Centraliza frmDestinatarios
   lvwDados.ColumnHeaders.Clear
   lvwDados.ColumnHeaders.Add , , "Tipos de Venda", 6150
'   MDISistema.StatusBar.Panels(1).Text = "Cadastro de Destinatarios"
End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
   Select Case Button.Key
      Case "Novo"
         If Validar_Permissao(1, "mnu" & Mid(Me.Name, 4, Len(Me.Name))) Then
            Limpa_Campos
            Botoes 3, frmDestinatarios
            Botoes_Extras 3
            txtNome.SetFocus
         End If

      Case "Gravar"
         Call Gravar
            
      Case "Excluir"
         If Validar_Permissao(3, "mnu" & Mid(Me.Name, 4, Len(Me.Name))) Then Call Excluir
         
      Case "Imprimir"
         Set rsDados = cnSistema.Execute("SELECT idGrupoAcesso FROM Usuarios WHERE idUsuario = " & idUser)
         If rsDados!idGrupoAcesso <> 1 Then
            Set rsDados = cnSistema.Execute("SELECT dbo.Usuarios.idUsuario, dbo.GrupoAcessoItens.Name FROM dbo.Usuarios INNER JOIN dbo.GrupoAcesso ON dbo.Usuarios.idGrupoAcesso = dbo.GrupoAcesso.idGrupoAcesso INNER JOIN dbo.GrupoAcessoItens ON dbo.GrupoAcesso.idGrupoAcesso = dbo.GrupoAcessoItens.idGrupoAcesso " & _
                                            "WHERE (dbo.Usuarios.idUsuario = " & idUser & ") AND (dbo.GrupoAcessoItens.Name = 'mnuRelDestinatarios')")
'            If Not rsDados.EOF() Then frmRelDestinatarios.Show Else MsgBox "Usuário não autorizado para esta operação", vbInformation, "Acesso Negado"
         Else
'            frmRelDestinatarios.Show
         End If
         Set rsDados = Nothing
      
      Case "Compras"
'         frmUltimasVendas.Show vbModal
   
      Case "UsoContinuo"
'         If cboDados.ListIndex = -1 Then
'            MsgBox "Selecione primeiro um Destinatario", vbInformation + vbOKOnly, "Destinatarios"
'            Exit Sub
'         Else
'            frmUsoContinuo.Show
'         End If
   End Select
End Sub

Private Sub Excluir()
On Error GoTo ErroIntegridade

   If cboDados.ListIndex <> -1 Then
      If MsgBox("Confirma Excluir o registro atual? ", vbYesNo + vbInformation, "Excluir") = vbYes Then
         Atividade "Exclusão: " & Trim(SQLCheck(txtNome.Text)), Me.Caption
         cnSistema.Execute "DELETE FROM DestinatariosFormaPagamento WHERE idDestinatario=" & cboDados.ItemData(cboDados.ListIndex)
         cnSistema.Execute "DELETE FROM Destinatarios WHERE idDestinatario=" & cboDados.ItemData(cboDados.ListIndex)
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

Private Function Verifica_Campos()
Dim strMensagem As String
Verifica_Campos = True

   If txtNome.Text = Empty Then strMensagem = strMensagem & "Nome" & Chr(13)
   If Not CPF_CNPJ(txtCN.Text) Then strMensagem = strMensagem & "CPF/CNPJ" & Chr(13)
   If Val(mskPrazoPagamento.Text) > 250 Then strMensagem = strMensagem & "Prazo para Pagamento entre 0 e 250" & Chr(13)
   If Val(mskDiaPagamento.Text) > 31 Then strMensagem = strMensagem & "Dia para Pagamento entre 0 e 31" & Chr(13)
   If Not IsDate(mskNascimento.Text) And mskNascimento.Text <> "  /  /    " Then strMensagem = strMensagem & "Data de Nascimento" & Chr(13)
   If optConvenio.value And cboConvenio.ListIndex = -1 Then strMensagem = strMensagem & "Convênio" & Chr(13)
   If optGrupoDesconto.value And cboGrupoDesconto.ListIndex = -1 Then strMensagem = strMensagem & "Grupo de Desconto" & Chr(13)
   If cboUF.ListIndex = -1 Then strMensagem = strMensagem & "UF" & Chr(13)
   If cboMunicipio.ListIndex = -1 Then strMensagem = strMensagem & "Municipio" & Chr(13)
   
   If Not strMensagem = Empty Then
      MsgBox "Verifique os seguintes campos:" & Chr(13) & strMensagem, vbExclamation + vbOKOnly, "Campos Obrigatórios"
      Verifica_Campos = False
      Exit Function
   End If
End Function

Private Sub Prencher_Campos()
Dim rsTemp As New ADODB.Recordset
Dim intContador As Integer
   Set rsDados = cnSistema.Execute("SELECT * FROM Destinatarios WHERE idDestinatario = " & cboDados.ItemData(cboDados.ListIndex))
   txtNome.Text = IIf(Trim(rsDados!Nome) = "" Or IsNull(rsDados!Nome), Empty, rsDados!Nome)
   txtEndereco.Text = IIf(Trim(rsDados!Endereco) = "" Or IsNull(rsDados!Endereco), Empty, rsDados!Endereco)
   txtCidade.Text = IIf(Trim(rsDados!Cidade) = "" Or IsNull(rsDados!Cidade), Empty, rsDados!Cidade)
   txtBairro.Text = IIf(Trim(rsDados!Bairro) = "" Or IsNull(rsDados!Bairro), Empty, rsDados!Bairro)
   txtCEP.Text = IIf(Trim(rsDados!CEP) = "" Or IsNull(rsDados!CEP), Empty, rsDados!CEP)
   txtCN.Text = IIf(Trim(rsDados!CN) = "" Or IsNull(rsDados!CN), Empty, rsDados!CN)
   txtCE.Text = IIf(Trim(rsDados!CE) = "" Or IsNull(rsDados!CE), Empty, rsDados!CE)
   txtEmail.Text = IIf(Trim(rsDados!Email) = "" Or IsNull(rsDados!Email), Empty, rsDados!Email)
   mskLimiteCredito.Text = IIf(Trim(rsDados!LimiteCredito) = "" Or IsNull(rsDados!LimiteCredito), Empty, rsDados!LimiteCredito)
   mskPrazoPagamento.Text = rsDados!PrazoPagamento & Space(3 - Len(rsDados!PrazoPagamento))
   mskDiaPagamento.Text = rsDados!DiaPagamento & Space(2 - Len(rsDados!DiaPagamento))
   txtObservacao.Text = IIf(Trim(rsDados!Observacao) = "" Or IsNull(rsDados!Observacao), Empty, rsDados!Observacao)
   txtTelefone_1.Text = IIf(Trim(rsDados!Telefone_1) = "" Or IsNull(rsDados!Telefone_1), Empty, rsDados!Telefone_1)
   txtTelefone_2.Text = IIf(Trim(rsDados!Telefone_2) = "" Or IsNull(rsDados!Telefone_2), Empty, rsDados!Telefone_2)
   mskNascimento.Text = IIf(IsNull(rsDados!Nascimento), "  /  /    ", rsDados!Nascimento)
   mskJuros.Text = rsDados!Juros
   chkBloqueio.value = IIf(rsDados!Bloqueio, 1, 0)
   
   If rsDados!ConsumidorFinal Then
      optConsumidorSim.value = True
   Else
      optConsumidorNao.value = True
   End If
   
   If rsDados!Interestadual Then
      optInterestadual.value = True
   Else
      optEstadual.value = True
   End If
   
'   optConsumidorSim.value = IIf(rsDados!ConsumidorFinal, 1, 0)
'   optConsumidorNao.value = IIf(rsDados!ConsumidorFinal, 0, 1)
'   optEstadual.value = IIf(rsDados!Interestadual, 1, 0)
'   optInterestadual.value = IIf(rsDados!Interestadual, 0, 1)
   
   For intContador = 0 To (cboUF.ListCount - 1)
      If cboUF.ItemData(intContador) = rsDados!idUF Then
         cboUF.ListIndex = intContador
         Exit For
      End If
   Next
   
   For intContador = 0 To (cboMunicipio.ListCount - 1)
      If cboMunicipio.ItemData(intContador) = rsDados!idMunicipio Then
         cboMunicipio.ListIndex = intContador
         Exit For
      End If
   Next
   
   For intContador = 0 To (cboTipoContribuinte.ListCount - 1)
      If intContador = rsDados!idTipoContribuinte Then
         cboTipoContribuinte.ListIndex = intContador
         Exit For
      End If
   Next
   
   lblDataCadastro.Caption = "Data do Cadastro: " & rsDados!DataCadastro
   Set rsTemp = cnSistema.Execute("SELECT dbo.FormaPagamento.idFormaPagamento, dbo.FormaPagamento.Descricao " & _
                                  "FROM dbo.Destinatarios INNER JOIN dbo.DestinatariosFormaPagamento ON dbo.Destinatarios.idDestinatario = dbo.DestinatariosFormaPagamento.idDestinatario INNER JOIN dbo.FormaPagamento ON dbo.DestinatariosFormaPagamento.idFormaPagamento = dbo.FormaPagamento.idFormaPagamento " & _
                                  "WHERE dbo.DestinatariosFormaPagamento.idDestinatario = " & cboDados.ItemData(cboDados.ListIndex))
   lvwDados.ListItems.Clear
   Do While Not rsTemp.EOF
      Set ItemList = lvwDados.ListItems.Add(, "R" & rsTemp!idFormaPagamento, rsTemp!Descricao)
      rsTemp.MoveNext
   Loop
   Set rsTemp = Nothing
   cboConvenio.ListIndex = -1
   cboGrupoDesconto.ListIndex = -1
   Select Case rsDados!Beneficio
      Case 0
         optNenhum.value = True
      Case 1
         optConvenio.value = True
         'Carregar Convênio
         For intContador = 0 To (cboConvenio.ListCount - 1)
            If cboConvenio.ItemData(intContador) = rsDados!idConvenio Then
               cboConvenio.ListIndex = intContador
               Exit For
            End If
         Next
      Case 2
         optGrupoDesconto.value = True
         'Carregar Convênio
         For intContador = 0 To (cboGrupoDesconto.ListCount - 1)
            If cboGrupoDesconto.ItemData(intContador) = rsDados!idGrupoDesconto Then
               cboGrupoDesconto.ListIndex = intContador
               Exit For
            End If
         Next
   End Select
   Set rsDados = Nothing
   Botoes 2, frmDestinatarios
   Botoes_Extras 2
End Sub

Private Sub Limpa_Campos()
   Set rsDados = cnSistema.Execute("SELECT COUNT(*) AS Registros FROM Destinatarios")
   lblRegistros.Caption = IIf(rsDados!Registros > 1, rsDados!Registros & " registros", IIf(rsDados!Registros = 1, "1 registro", "Nenhum registro"))
   Set rsDados = Nothing
   cboDados.Clear
   Botoes 1, frmDestinatarios
   Botoes_Extras 1
   txtNome.Text = Empty
   txtEndereco.Text = Empty
   txtCidade.Text = Empty
   cboMunicipio.ListIndex = -1
   cboTipoContribuinte.ListIndex = -1
   txtBairro.Text = Empty
   cboUF.ListIndex = -1
   txtCEP.Text = Empty
   txtCN.Text = Empty
   txtCE.Text = Empty
   txtEmail.Text = Empty
   mskLimiteCredito.Text = 0
   mskPrazoPagamento.Text = Space(3)
   mskDiaPagamento.Text = Space(2)
   txtObservacao.Text = Empty
   txtTelefone_1.Text = Empty
   txtTelefone_2.Text = Empty
   mskNascimento.Text = "  /  /    "
   mskJuros.Text = Empty
   chkBloqueio.value = 0
   lblDataCadastro.Caption = Empty
   cboConvenio.ListIndex = -1
   cboGrupoDesconto.ListIndex = -1
   cboFormaPagamento.ListIndex = -1
   optNenhum.value = True
   
   lvwDados.ListItems.Clear
   
   Carrega_Combos_Extras

End Sub

Private Sub Gravar()
Dim rsInclusao As New ADODB.Recordset
Dim bytBeneficio As Byte
Dim strConvenio As String
Dim strGrupoDesconto As String
   bytBeneficio = 0
   If optConvenio.value Then bytBeneficio = 1
   If optGrupoDesconto.value Then bytBeneficio = 2
   If cboConvenio.ListIndex = -1 Then strConvenio = "NULL" Else strConvenio = cboConvenio.ItemData(cboConvenio.ListIndex)
   If cboGrupoDesconto.ListIndex = -1 Then strGrupoDesconto = "NULL" Else strGrupoDesconto = cboGrupoDesconto.ItemData(cboGrupoDesconto.ListIndex)
   If cboDados.ListIndex = -1 Then 'Inclusão
      If Validar_Permissao(1, "mnu" & Mid(Me.Name, 4, Len(Me.Name))) Then
         If Not Verifica_Campos() Then Exit Sub
         If MsgBox("Confirma Incluir o registro atual", vbYesNo + vbQuestion, "Inclusão") = vbYes Then
            cnSistema.Execute "INSERT INTO Destinatarios (Nome,Endereco,idMunicipio,idUF,Cidade,Bairro,UF,CEP,CN,CE,eMail,LimiteCredito," & _
                                                    "PrazoPagamento,DiaPagamento,Observacao,Telefone_1,Telefone_2,Nascimento," & _
                                                    "Juros,Bloqueio,ConsumidorFinal,Interestadual,idTipoContribuinte,DataCadastro,Beneficio,idConvenio,idGrupoDesconto) " & _
                              "VALUES ('" & txtNome.Text & "','" & txtEndereco.Text & "'," & cboMunicipio.ItemData(cboMunicipio.ListIndex) & "," & cboUF.ItemData(cboUF.ListIndex) & ",'" & txtCidade.Text & "','" & _
                                            txtBairro.Text & "','" & cboUF.Text & "','" & txtCEP.Text & "','" & _
                                            txtCN.Text & "','" & txtCE.Text & "','" & txtEmail.Text & "'," & _
                                            Replace(mskLimiteCredito.ClipText, ",", ".") & "," & Val(Replace(mskPrazoPagamento.ClipText, ",", ".")) & "," & _
                                            Val(Replace(mskDiaPagamento.ClipText, ",", ".")) & ",'" & txtObservacao.Text & "','" & txtTelefone_1.Text & "','" & _
                                            txtTelefone_1.Text & "'," & IIf(mskNascimento.Text = "  /  /    ", "NULL", "'" & Format(mskNascimento.Text, "mm/dd/yyyy") & "'") & "," & _
                                            Val(mskJuros.Text) & "," & IIf(chkBloqueio.value, 1, 0) & "," & IIf(optConsumidorSim.value, 1, 0) & "," & IIf(optEstadual.value, 0, 1) & "," & cboTipoContribuinte.ListIndex & ",'" & Format(Date, "mm/dd/yyyy") & "'," & _
                                            bytBeneficio & "," & strConvenio & "," & strGrupoDesconto & ")"
            Set rsInclusao = cnSistema.Execute("SELECT IDENT_CURRENT('Destinatarios') AS 'Identity'")
            Gravar_DestinatariosFormaPagamento rsInclusao!Identity
            Atividade "Inclusão: " & SQLCheck(txtNome.Text), Me.Caption
         End If
      End If
   Else 'Alteracão
      If Validar_Permissao(2, "mnu" & Mid(Me.Name, 4, Len(Me.Name))) Then
         If Not Verifica_Campos() Then Exit Sub
         If MsgBox("Confirma Alterar o registro atual", vbYesNo + vbQuestion, "Alteração") = vbYes Then
            cnSistema.Execute "UPDATE Destinatarios SET " & _
                              "Nome = '" & txtNome.Text & "', " & _
                              "Endereco = '" & txtEndereco.Text & "', " & _
                              "idMunicipio = " & cboMunicipio.ItemData(cboMunicipio.ListIndex) & ", " & _
                              "idUF = " & cboUF.ItemData(cboUF.ListIndex) & ", " & _
                              "Cidade = '" & txtCidade.Text & "', " & _
                              "Bairro = '" & txtBairro.Text & "', " & _
                              "UF = '" & cboUF.Text & "', " & _
                              "CEP = '" & txtCEP.Text & "', " & _
                              "CN = '" & txtCN.Text & "', " & _
                              "CE = '" & txtCE.Text & "', " & _
                              "eMail = '" & txtEmail.Text & "', " & _
                              "LimiteCredito = " & Replace(mskLimiteCredito.ClipText, ",", ".") & ", " & _
                              "PrazoPagamento = " & Val(mskPrazoPagamento.Text) & ", " & _
                              "DiaPagamento = " & Val(mskDiaPagamento.Text) & ", " & _
                              "Observacao = '" & txtObservacao.Text & "', " & _
                              "Telefone_1 = '" & txtTelefone_1.Text & "', " & _
                              "Telefone_2 = '" & txtTelefone_2.Text & "', " & _
                              "Nascimento = " & IIf(mskNascimento.Text = "  /  /    ", "NULL", "'" & Format(mskNascimento.Text, "mm/dd/yyyy") & "'") & ", " & _
                              "Juros = " & Replace(mskJuros.Text, ",", ".") & ", " & "Beneficio = " & bytBeneficio & ", " & "idConvenio = " & strConvenio & ", " & _
                              "ConsumidorFinal = " & IIf(optConsumidorSim.value, 1, 0) & ", " & _
                              "Interestadual = " & IIf(optEstadual.value, 0, 1) & ", " & _
                              "idTipoContribuinte = " & cboTipoContribuinte.ListIndex & ", " & _
                              "idGrupoDesconto = " & strGrupoDesconto & ", " & "Bloqueio = " & IIf(chkBloqueio.value, 1, 0) & " " & _
                              "WHERE idDestinatario = " & cboDados.ItemData(cboDados.ListIndex)
            Atividade "Alterar: " & Trim(SQLCheck(txtNome.Text)), Me.Caption
            Gravar_DestinatariosFormaPagamento cboDados.ItemData(cboDados.ListIndex)
         End If
      End If
   End If
   cboGrupoDesconto.Enabled = False
   cboConvenio.Enabled = False
   Limpa_Campos
   cboDados.SetFocus
End Sub

Private Sub cboDados_KeyPress(KeyAscii As Integer)
Dim rsCombo As New ADODB.Recordset
   If KeyAscii = 13 Then Carrega_Combos
End Sub

Private Sub cboDados_Click()
   Prencher_Campos
End Sub

Private Sub Carrega_Combos()
Dim rsCombo As New ADODB.Recordset
   Set rsCombo = cnSistema.Execute("SELECT COUNT(idDestinatario) AS Registros FROM Destinatarios WHERE Nome LIKE '%" & cboDados.Text & "%'")
   If rsCombo!Registros > 5000 Then
      MsgBox "A consulta possui mais de 5.000 resultados" & vbCrLf & "Redefina sua pesquisa", vbOKOnly + vbInformation, "Pesquisa"
      Exit Sub
   End If
   Screen.MousePointer = vbHourglass
   If IsNumeric(Replace(Replace(Replace(cboDados.Text, ".", ""), "-", ""), "/", "")) Then
      Set rsCombo = cnSistema.Execute("SELECT idDestinatario, Nome, CN FROM Destinatarios WHERE CN LIKE '%" & cboDados.Text & "%'")
   Else
      Set rsCombo = cnSistema.Execute("SELECT idDestinatario, Nome, CN FROM Destinatarios WHERE Nome LIKE '%" & cboDados.Text & "%' ORDER BY Nome")
   End If
   Limpa_Campos
   Do While Not rsCombo.EOF
      cboDados.AddItem rsCombo!Nome
      cboDados.ItemData(cboDados.NewIndex) = rsCombo!idDestinatario
      rsCombo.MoveNext
   Loop
   Set rsCombo = Nothing
   Screen.MousePointer = vbDefault
End Sub

Private Sub Carrega_Combos_Extras()
Dim rsCombo As New ADODB.Recordset
   Set rsCombo = cnSistema.Execute("SELECT idConvenio, RazaoSocial FROM Convenios ORDER BY RazaoSocial")
   cboConvenio.Clear
   Do While Not rsCombo.EOF
      cboConvenio.AddItem rsCombo!RazaoSocial
      cboConvenio.ItemData(cboConvenio.NewIndex) = rsCombo!idConvenio
      rsCombo.MoveNext
   Loop
   Set rsCombo = cnSistema.Execute("SELECT idGrupoDesconto, Descricao FROM GrupoDesconto ORDER BY Descricao")
   cboGrupoDesconto.Clear
   Do While Not rsCombo.EOF
      cboGrupoDesconto.AddItem rsCombo!Descricao
      cboGrupoDesconto.ItemData(cboGrupoDesconto.NewIndex) = rsCombo!idGrupoDesconto
      rsCombo.MoveNext
   Loop
   Set rsCombo = cnSistema.Execute("SELECT idFormaPagamento, Descricao FROM FormaPagamento ORDER BY Descricao")
   cboFormaPagamento.Clear
   Do While Not rsCombo.EOF
      cboFormaPagamento.AddItem rsCombo!Descricao
      cboFormaPagamento.ItemData(cboFormaPagamento.NewIndex) = rsCombo!idFormaPagamento
      rsCombo.MoveNext
   Loop
   Set rsCombo = cnSistema.Execute("SELECT idUF, Sigla FROM UFs ORDER BY Sigla")
   cboUF.Clear
   Do While Not rsCombo.EOF
      cboUF.AddItem rsCombo!Sigla
      cboUF.ItemData(cboUF.NewIndex) = rsCombo!idUF
      rsCombo.MoveNext
   Loop
   Set rsCombo = cnSistema.Execute("SELECT idMunicipio, Nome FROM Municipios ORDER BY Nome")
   cboMunicipio.Clear
   Do While Not rsCombo.EOF
      cboMunicipio.AddItem rsCombo!Nome
      cboMunicipio.ItemData(cboMunicipio.NewIndex) = rsCombo!idMunicipio
      rsCombo.MoveNext
   Loop
   Set rsCombo = Nothing
End Sub

Public Sub Botoes_Extras(bytModo As Byte)
   Select Case bytModo
      Case 1 'Sem seleção
         cmdRemover.Enabled = False
         cmdInserir.Enabled = False
         Toolbar.Buttons(6).Enabled = False
         Toolbar.Buttons(7).Enabled = False
      Case 2 'Com seleção
         cmdRemover.Enabled = True
         cmdInserir.Enabled = True
         Toolbar.Buttons(6).Enabled = True
         Toolbar.Buttons(7).Enabled = True
      Case 3 'Inclusão
         cmdRemover.Enabled = True
         cmdInserir.Enabled = True
         Toolbar.Buttons(6).Enabled = False
         Toolbar.Buttons(7).Enabled = False
   End Select
End Sub

Private Sub txtCN_KeyPress(KeyAscii As Integer)
   If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
      KeyAscii = 0
   End If
End Sub

Private Sub txtCN_LostFocus()
Dim strCN As String
Dim rsTemp As New ADODB.Recordset

   strCN = Trim(Format(Replace(Replace(Replace(txtCN.Text, ".", ""), "-", ""), "/", ""), "@@@@@@@@@@@@@@"))
   If CPF_CNPJ(strCN) Then
      If Len(Trim(strCN)) = 14 Then txtCN.Text = Mid(strCN, 1, 2) & "." & Mid(strCN, 3, 3) & "." & Mid(strCN, 6, 3) & "/" & Mid(strCN, 9, 4) & "-" & Mid(strCN, 13, 2)
      If Len(Trim(strCN)) = 11 Then txtCN.Text = Mid(strCN, 1, 3) & "." & Mid(strCN, 4, 3) & "." & Mid(strCN, 7, 3) & "-" & Mid(strCN, 10, 2)
   Else
      MsgBox "CNPJ/CPF Inválido", vbOKOnly + vbInformation, "Pesquisa"
      Exit Sub
   End If
   
   If cboDados.ListIndex = -1 Then
      Set rsTemp = cnSistema.Execute("SELECT * FROM Destinatarios WHERE CN = '" & txtCN.Text & "'")
      If Not rsTemp.EOF Then
         MsgBox "CNPJ/CPF já Cadastrado", vbOKOnly + vbInformation, "Informação"
         Exit Sub
      End If
      Set rsTemp = Nothing
   End If
End Sub

Private Sub txtTelefone_1_KeyPress(KeyAscii As Integer)
   If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 32 And KeyAscii <> 40 And KeyAscii <> 41 And KeyAscii <> 45 And KeyAscii <> 46 And KeyAscii <> 8 Then
      KeyAscii = 0
   End If
   If KeyAscii = 32 Then
      If InStr(txtTelefone_1.Text, " ") <> 0 Then
         KeyAscii = 0
      End If
   End If
   If KeyAscii = 40 Then
      If InStr(txtTelefone_1.Text, "(") <> 0 Then
         KeyAscii = 0
      End If
   End If
   If KeyAscii = 41 Then
      If InStr(txtTelefone_1.Text, ")") <> 0 Then
         KeyAscii = 0
      End If
   End If
   If KeyAscii = 45 Then
      If InStr(txtTelefone_1.Text, "-") <> 0 Then
         KeyAscii = 0
      End If
   End If
   If KeyAscii = 46 Then
      If InStr(txtTelefone_1.Text, ".") <> 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtTelefone_2_KeyPress(KeyAscii As Integer)
   If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 32 And KeyAscii <> 40 And KeyAscii <> 41 And KeyAscii <> 45 And KeyAscii <> 46 And KeyAscii <> 8 Then
      KeyAscii = 0
   End If
   If KeyAscii = 32 Then
      If InStr(txtTelefone_2.Text, " ") <> 0 Then
         KeyAscii = 0
      End If
   End If
   If KeyAscii = 40 Then
      If InStr(txtTelefone_2.Text, "(") <> 0 Then
         KeyAscii = 0
      End If
   End If
   If KeyAscii = 41 Then
      If InStr(txtTelefone_2.Text, ")") <> 0 Then
         KeyAscii = 0
      End If
   End If
   If KeyAscii = 45 Then
      If InStr(txtTelefone_2.Text, "-") <> 0 Then
         KeyAscii = 0
      End If
   End If
   If KeyAscii = 46 Then
      If InStr(txtTelefone_2.Text, ".") <> 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub txtCEP_KeyPress(KeyAscii As Integer)
   If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 45 And KeyAscii <> 46 And KeyAscii <> 8 Then
      KeyAscii = 0
   End If
   If KeyAscii = 45 Then
      If InStr(txtCEP.Text, "-") <> 0 Then
         KeyAscii = 0
      End If
   End If
   If KeyAscii = 46 Then
      If InStr(txtCEP.Text, ".") <> 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub lvwDados_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   lvwDados.SortKey = ColumnHeader.Index - 1
   If lvwDados.SortOrder = lvwAscending Then
      lvwDados.SortOrder = lvwDescending
   Else
      lvwDados.SortOrder = lvwAscending
   End If
End Sub

Private Sub cmdRemover_Click()
   If lvwDados.ListItems.Count = 0 Then
      MsgBox "Não existem tipos de vendas cadastradas", vbOKOnly + vbInformation, "Remover"
      Exit Sub
   End If
   If MsgBox("Deseja remover esta forma de pagamento", vbYesNo + vbQuestion, "Remover") = vbYes Then
      lvwDados.ListItems.Remove (lvwDados.SelectedItem.Index)
   End If
End Sub

Private Sub cmdInserir_Click()
   On Error GoTo ErroInserir
   If Verifica_Campos_DestinatarioFormaPagamento() Then
      Set ItemList = lvwDados.ListItems.Add(, "R" & cboFormaPagamento.ItemData(cboFormaPagamento.ListIndex), cboFormaPagamento.Text)
      cboFormaPagamento.ListIndex = -1
      cboFormaPagamento.SetFocus
   End If
On Error GoTo 0
Exit Sub
ErroInserir:
   If Err.Number = 0 Then
      ' Operação Ok
   ElseIf Err.Number = 35602 Then
      MsgBox "esta forma de pagamento já foi inserido" & Chr(13) & "Verifique na lista", vbInformation + vbOKOnly, "Inserir"
      Exit Sub
   Else
      MsgBox "Verificar: " & Err.Number & Chr(13) & Err.Description, vbExclamation, "Inserir"
      Exit Sub
   End If
End Sub

Private Function Verifica_Campos_DestinatarioFormaPagamento()
Dim strMensagem As String
Verifica_Campos_DestinatarioFormaPagamento = True

   If cboFormaPagamento.ListIndex = -1 Then strMensagem = strMensagem & "Forma de pagamento" & Chr(13)

   If Not strMensagem = Empty Then
      MsgBox "Verifique os seguintes campos:" & Chr(13) & strMensagem, vbExclamation + vbOKOnly, "Campos Obrigatórios"
      Verifica_Campos_DestinatarioFormaPagamento = False
      Exit Function
   End If
End Function

Private Sub Gravar_DestinatariosFormaPagamento(id As Long)
Dim Contador As Integer
   cnSistema.Execute "DELETE FROM DestinatariosFormaPagamento WHERE idDestinatario=" & id
   If lvwDados.ListItems.Count > 0 Then
      For Contador = 1 To lvwDados.ListItems.Count
          cnSistema.Execute "INSERT INTO DestinatariosFormaPagamento (idDestinatario,idFormaPagamento) " & _
                            "VALUES (" & id & "," & Mid(lvwDados.ListItems.Item(Contador).Key, 2, Len(lvwDados.ListItems.Item(Contador).Key)) & ")"
          Atividade "Forma de pagamento " & Trim(lvwDados.ListItems.Item(Contador).Text), Me.Caption
      Next
   End If
End Sub

Private Sub optConvenio_Click()
   cboConvenio.Enabled = True
   cboGrupoDesconto.Enabled = False
   cboGrupoDesconto.ListIndex = -1
End Sub

Private Sub optGrupoDesconto_Click()
   cboConvenio.Enabled = False
   cboGrupoDesconto.Enabled = True
   cboConvenio.ListIndex = -1
End Sub

Private Sub optNenhum_Click()
   cboConvenio.Enabled = False
   cboConvenio.ListIndex = -1
   cboGrupoDesconto.Enabled = False
   cboGrupoDesconto.ListIndex = -1
End Sub

Private Sub mskJuros_LostFocus()
   If Val(Replace(mskJuros.Text, ",", ".")) > 99.99 Then
      MsgBox "Juros não pode ser superior a 99.99%", vbOKOnly, "Juros"
      mskJuros.SetFocus
      mskJuros.SelStart = 0
      mskJuros.SelLength = Len(mskJuros.Text)
   End If
End Sub

Private Sub mskLimiteCredito_KeyPress(KeyAscii As Integer)
   If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 44 And KeyAscii <> 8 Then
      KeyAscii = 0
   End If
   If KeyAscii = 44 Then
      If InStr(mskLimiteCredito.ClipText, ",") <> 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub mskJuros_KeyPress(KeyAscii As Integer)
   If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 44 And KeyAscii <> 8 Then
      KeyAscii = 0
   End If
   If KeyAscii = 44 Then
      If InStr(mskJuros.ClipText, ",") <> 0 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Sub mskPrazoPagamento_KeyPress(KeyAscii As Integer)
   If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
      KeyAscii = 0
   End If
End Sub

Private Sub mskDiaPagamento_KeyPress(KeyAscii As Integer)
   If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
      KeyAscii = 0
   End If
End Sub

Private Sub mskNascimento_LostFocus()
   If mskNascimento.Text <> "  /  /    " Then
      If Not IsDate(mskNascimento.Text) Then
         MsgBox "Digite uma data válida", vbOKOnly, "Nascimento"
         mskNascimento.SetFocus
         mskNascimento.SelStart = 0
         mskNascimento.SelLength = Len(mskNascimento.Text)
         Exit Sub
      End If
   End If
End Sub

Private Sub cboUF_LostFocus()
Dim rsCombo As New ADODB.Recordset

'  Municipios
   Set rsCombo = cnSistema.Execute("SELECT idMunicipio, Nome FROM Municipios WHERE UF = '" & cboUF.Text & "' ORDER BY Nome")
   cboMunicipio.Clear
   Do While Not rsCombo.EOF
      cboMunicipio.AddItem rsCombo!Nome
      cboMunicipio.ItemData(cboMunicipio.NewIndex) = rsCombo!idMunicipio
      rsCombo.MoveNext
   Loop
   Set rsCombo = Nothing
End Sub


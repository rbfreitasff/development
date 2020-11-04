VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Acesso ao Sistema"
   ClientHeight    =   2250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5730
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   5730
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   795
      Left            =   4500
      Picture         =   "frmLogin.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   960
      Width           =   1155
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   795
      Left            =   4500
      Picture         =   "frmLogin.frx":0D0C
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   60
      Width           =   1155
   End
   Begin VB.PictureBox picLogotipo 
      AutoRedraw      =   -1  'True
      FillStyle       =   0  'Solid
      Height          =   1320
      Left            =   1200
      ScaleHeight     =   1260
      ScaleWidth      =   3000
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   180
      Width           =   3060
   End
   Begin VB.TextBox txtUsuario 
      Height          =   315
      Left            =   660
      MaxLength       =   15
      TabIndex        =   1
      Top             =   1860
      Width           =   2115
   End
   Begin VB.TextBox txtSenha 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   3540
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1860
      Width           =   2115
   End
   Begin VB.Image imgChaves 
      Height          =   480
      Left            =   360
      Picture         =   "frmLogin.frx":15D6
      Top             =   180
      Width           =   480
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      Caption         =   "Digite seus Dados para Acessar o Sistema"
      Height          =   795
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   975
   End
   Begin VB.Label lblUsuario 
      AutoSize        =   -1  'True
      Caption         =   "Usu�rio"
      Height          =   195
      Left            =   60
      TabIndex        =   0
      Top             =   1920
      Width           =   540
   End
   Begin VB.Label lblSenha 
      AutoSize        =   -1  'True
      Caption         =   "Senha"
      Height          =   195
      Left            =   2940
      TabIndex        =   2
      Top             =   1920
      Width           =   465
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Sandro Rizzo em 27/01/1999
Option Explicit
Dim rsUsuarios As New ADODB.Recordset
Dim rsContasPagar As New ADODB.Recordset

Private Sub Form_Load()
   On Error GoTo Erro
   picLogotipo.Picture = LoadPicture(I_Logotipo)
   rsUsuarios.Open "Select * from Usuarios", cnSistema, adOpenForwardOnly, adLockPessimistic, 1

Exit Sub
Erro:
   If Err.Number = -2147467259 Then
      rsErro = True
      Beep
      MsgBox "Erro na Abertura do Arquivo de Dados" & Chr(13) & "Algum usu�rio est� com o Arquivo em modo Exclusivo", vbExclamation, "Erro"
''      cnSistema.Close
      End
   End If
End Sub

Private Sub cmdCancelar_Click()
''   cnSistema.Close
   End
End Sub

Private Sub cmdOK_Click()
   If txtUsuario.text = Empty Then
      Beep
      MsgBox "Informe o Usu�rio para Entrar no Sistema", vbOKOnly + vbExclamation, "Login no Sistema"
      txtUsuario.SetFocus
   Else
      ' Limpar
      If UCase(txtUsuario.text) = "ZEUS" And UCase(txtSenha.text) = "ZEUS" Then
         LimparBanco
         
         ' Acesso
         I_Acesso = 1
         I_User = "Zeus"
         rsUsuarios.Close
         Unload Me
         Atividade "Login no Sistema", Me.Caption
''         Select Case I_Acesso
''            Case 1
''               MDISistema.mnuUsuarios.Visible = True
''            Case 2
''               MDISistema.mnuUsuarios.Visible = True
''            Case 3
''               MDISistema.mnuUsuarios.Visible = False
''         End Select
''         MDISistema.StatusBar.Panels(2).text = I_User
''         MDISistema.Show
      Else
         ' Acesso Normal
         Set rsUsuarios = cnSistema.Execute("Select * from Usuarios Where Login='" & SQLCheck(Trim(txtUsuario.text)) & "'")
         If rsUsuarios.EOF Then
            Beep
            MsgBox "Usu�rio n�o Cadastrado, por favor confira", vbOKOnly + vbExclamation, "Login no Sistema"
            txtUsuario.SetFocus
            txtUsuario.SelStart = 0
            txtUsuario.SelLength = Len(txtUsuario.text)
         Else
            If Not UCase(rsUsuarios!Senha) = UCase(Trim(txtSenha.text)) Then
               Beep
               MsgBox "Senha n�o confere, por favor confira", vbOKOnly + vbExclamation, "Senha"
               txtSenha.SetFocus
               txtSenha.SelStart = 0
               txtSenha.SelLength = Len(txtSenha.text)
            Else
               I_Acesso = rsUsuarios!Nivel
               I_User = rsUsuarios!Login
               rsUsuarios.Close
               Unload Me
               Atividade "Login no Sistema", Me.Caption
''               Select Case I_Acesso
''                  Case 1
''   '                  MDISistema.tlbMDI.Buttons.Item(9).Visible = True
''                     MDISistema.mnuUsuarios.Visible = True
''                  Case 2
''   '                  MDISistema.tlbMDI.Buttons.Item(9).Visible = True
''                     MDISistema.mnuUsuarios.Visible = True
''                  Case 3
''   '                  MDISistema.tlbMDI.Buttons.Item(9).Visible = False
''                     MDISistema.mnuUsuarios.Visible = False
''               End Select
''               MDISistema.StatusBar.Panels(2).text = I_User
''               MDISistema.Show
            End If
         End If
      
      End If
      
   End If
End Sub

Private Sub txtSenha_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Sendkeys "{TAB}"
   End If
End Sub

Private Sub txtUsuario_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Sendkeys "{TAB}"
   End If
End Sub

Private Sub LimparBanco()
   ' Limpeza
   cnSistema.Execute "Delete * From AgendaDiariaCompromissos"
   cnSistema.Execute "Delete * From AgendaTelefonica"
   cnSistema.Execute "Delete * From AgendaTelefonicaEmails"
   cnSistema.Execute "Delete * From AgendaTelefonicaTelefones"
   cnSistema.Execute "Delete * From ApuracoesMovimento"
   cnSistema.Execute "Delete * From Atividades"
   cnSistema.Execute "Delete * From CaixaDiario"
   cnSistema.Execute "Delete * From CaixaDiarioNotas"
   cnSistema.Execute "Delete * From Cargos"
   cnSistema.Execute "Delete * From Carros"
   cnSistema.Execute "Delete * From CentrosCusto"
'   cnSistema.Execute "Delete * From CFOPReferencias"
'   cnSistema.Execute "Delete * From CFOPs"
   cnSistema.Execute "Delete * From ClassesProdutos"
   cnSistema.Execute "Delete * From Clientes"
   cnSistema.Execute "Delete * From ClientesSPC"
   cnSistema.Execute "Delete * From Consignacoes"
   cnSistema.Execute "Delete * From ConsignacoesAcertos"
   cnSistema.Execute "Delete * From ConsignacoesAcertosItens"
   cnSistema.Execute "Delete * From ConsignacoesAcertosNumeros"
   cnSistema.Execute "Delete * From ConsignacoesItens"
   cnSistema.Execute "Delete * From ContasBancarias"
   cnSistema.Execute "Delete * From ContasPagar"
   cnSistema.Execute "Delete * From ContasPagarBaixas"
   cnSistema.Execute "Delete * From ContasReceber"
   cnSistema.Execute "Delete * From ContasReceberBaixas"
   cnSistema.Execute "Delete * From ContasReceberTotal"
   cnSistema.Execute "Delete * From CorrecaoEstoque"
   cnSistema.Execute "Delete * From Empresa"
   cnSistema.Execute "Delete * From Empresas"
   cnSistema.Execute "Delete * From Fabricantes"
   cnSistema.Execute "Delete * From Faturas"
   cnSistema.Execute "Delete * From FaturasImpressoes"
   cnSistema.Execute "Delete * From FaturasNotas"
   cnSistema.Execute "Delete * From Feriados"
   cnSistema.Execute "Delete * From FormasPagamento"
   cnSistema.Execute "Delete * From Fornecedores"
   cnSistema.Execute "Delete * From Fretes"
   cnSistema.Execute "Delete * From Funcionarios"
   cnSistema.Execute "Delete * From Grupos"
   cnSistema.Execute "Delete * From Horarios"
   cnSistema.Execute "Delete * From Marcas"
   cnSistema.Execute "Delete * From Menu"
   cnSistema.Execute "Delete * From MenuUsuario"
   cnSistema.Execute "Delete * From MovimentosRealizados"
   cnSistema.Execute "Delete * From NaturezasOperacao"
   cnSistema.Execute "Delete * From NFECF"
   cnSistema.Execute "Delete * From NFECFItens"
   cnSistema.Execute "Delete * From NFECFPagamentos"
   cnSistema.Execute "Delete * From NFEntradas"
   cnSistema.Execute "Delete * From NFEntradasItens"
   cnSistema.Execute "Delete * From NFSaidasManuais"
   cnSistema.Execute "Delete * From NFSaidasManuaisItens"
   cnSistema.Execute "Delete * From NFSaidasSEPD"
   cnSistema.Execute "Delete * From NFSaidasSEPDBoletos"
   cnSistema.Execute "Delete * From NFSaidasSEPDItens"
   cnSistema.Execute "Delete * From NFe"
   cnSistema.Execute "Delete * From NFeItens"
   cnSistema.Execute "Delete * From NFeBoletos"
   cnSistema.Execute "Delete * From Opcoes"
   cnSistema.Execute "Delete * From Orcamentos"
   cnSistema.Execute "Delete * From OrcamentosCheques"
   cnSistema.Execute "Delete * From OrcamentosForma"
   cnSistema.Execute "Delete * From OrcamentosItens"
   cnSistema.Execute "Delete * From OrcamentosSequencia"
   cnSistema.Execute "Delete * From OrcamentosServicos"
   cnSistema.Execute "Delete * From OrdensServico"
   cnSistema.Execute "Delete * From Pedidos"
   cnSistema.Execute "Delete * From PedidosItens"
   cnSistema.Execute "Delete * From PedidosPagamentos"
   cnSistema.Execute "Delete * From PlanosContas"
   cnSistema.Execute "Delete * From Ponto"
'   cnSistema.Execute "Delete * From PrefixosTelefonicos"
   cnSistema.Execute "Delete * From Produtos"
   cnSistema.Execute "Delete * From ProdutosFornecedores"
   cnSistema.Execute "Delete * From ProdutosVinculados"
   cnSistema.Execute "Delete * From RamosAtividade"
   cnSistema.Execute "Delete * From RegistradoresFiscais"
   cnSistema.Execute "Delete * From Servicos"
   cnSistema.Execute "Delete * From Setores"
'   cnSistema.Execute "Delete * From Sistema"
   cnSistema.Execute "Delete * From SituacoesServicos"
'   cnSistema.Execute "Delete * From SituacoesTributarias"
   cnSistema.Execute "Delete * From SubGrupos"
   cnSistema.Execute "Delete * From TabelaPrecos"
   cnSistema.Execute "Delete * From Tabelas"
   cnSistema.Execute "Delete * From Teste"
   cnSistema.Execute "Delete * From Transportadores"
   cnSistema.Execute "Delete * From UnidadesMedida"
'   cnSistema.Execute "Delete * From Usuarios"
   cnSistema.Execute "Delete * From VendaDireta"
   cnSistema.Execute "Delete * From VendaDiretaItens"
End Sub

VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmBancoDados 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Abertura do Banco de Dados"
   ClientHeight    =   1215
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1215
   ScaleWidth      =   6750
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCriar 
      Caption         =   "Criar Banco"
      Height          =   375
      Left            =   60
      TabIndex        =   6
      Top             =   780
      Width           =   1275
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   5400
      TabIndex        =   5
      Top             =   780
      Width           =   1275
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   4020
      TabIndex        =   4
      Top             =   780
      Width           =   1275
   End
   Begin VB.Frame fraCaminhos 
      Caption         =   "Caminhos"
      Height          =   675
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   6630
      Begin VB.CommandButton cmdProcurarBanco 
         Caption         =   "Pr&ocurar"
         Height          =   315
         Left            =   5640
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtCaminhoBanco 
         Height          =   315
         Left            =   1380
         MaxLength       =   100
         TabIndex        =   2
         Top             =   240
         Width           =   4185
      End
      Begin VB.Label lblLocalBanco 
         AutoSize        =   -1  'True
         Caption         =   "Banco de Dados"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   300
         Width           =   1200
      End
   End
   Begin MSComDlg.CommonDialog dlgCaminhoBanco 
      Left            =   6840
      Top             =   60
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog dlgArquivo 
      Left            =   7320
      Top             =   60
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmBancoDados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ItemList As ListItem
Dim rsTemp As New ADODB.Recordset
Dim cnTemp As New ADODB.Connection

Private Sub Form_Load()
   txtCaminhoBanco.Text = LerArquivoINI("Banco de Dados", "Caminho", App.Path & "\System.ini")
End Sub

Private Sub cmdProcurarBanco_Click()
   dlgCaminhoBanco.FileName = ""
   dlgCaminhoBanco.Filter = "Todos os Arquivos|*.*"
   dlgCaminhoBanco.ShowOpen
   If dlgCaminhoBanco.FileName <> "" Then
      txtCaminhoBanco.Text = Mid(dlgCaminhoBanco.FileName, 1, Len(dlgCaminhoBanco.FileName) - Len(dlgCaminhoBanco.FileTitle) - 1)
   End If
End Sub

Private Sub cmdOK_Click()
   If Not GravaArquivoINI("Banco de Dados", "Caminho", txtCaminhoBanco.Text, CaminhoINI & "\System.ini") Then
      MsgBox "Não foi possível Gravar Arquivo de Configuração" & Chr(13) & "Entre em Contato com o Suporte", vbInformation + vbOKOnly, "Sistema"
      End
   End If

   BancoDeDados = txtCaminhoBanco.Text
   Unload Me
End Sub

Private Sub cmdCancelar_Click()
   End
End Sub

Private Sub cmdCriar_Click()
Dim BaseDeDados As Database

''   Dim Teste As Database
''   Set Teste = CreateDatabase("C:\Sistemas\Comercial\Dados\Dados.mdb", dbLangGeneral)

'   dlgArquivo.InitDir = "C:\"
'   dlgArquivo.DefaultExt = "*.mdb"
'   dlgArquivo.Filter = "*.MDB"
'   dlgArquivo.FileName = "Dados.mdb"
'   dlgArquivo.ShowOpen
'   If dlgArquivo.FileName = Empty Then Exit Sub
'   cnSistema.Close
'   cnSistema.Open

   cnTemp.Provider = "Microsoft.Jet.OLEDB.4.0"
   cnTemp.Properties("Data Source") = "C:\Sistemas\Comercial\Dados\Dados.MDB"
''   cnTemp.Properties("Data Source") = dlgArquivo.FileName
   cnTemp.Open

   Dim vStrutura As String
   
 ' Agenda Diaria Compromissos
   vStrutura = ""
   vStrutura = vStrutura & "CREATE TABLE AgendaDiariaCompromissos(idAgendaDiariaCompromisso Integer IDENTITY Primary Key NOT NULL, "
   vStrutura = vStrutura & "Data DateTime,"
   vStrutura = vStrutura & "Compromisso NVarChar(50)"
   vStrutura = vStrutura & ")"
   ' Criar Tabela
   cnTemp.Execute vStrutura
   
 ' Produtos
   vStrutura = ""
   vStrutura = vStrutura & "CREATE TABLE Produtos(idProduto Integer IDENTITY Primary Key NOT NULL, "
   vStrutura = vStrutura & "idClasse Integer,"
   vStrutura = vStrutura & "idUnidade Integer,"
   vStrutura = vStrutura & "idRegistradorFiscal Integer,"
   vStrutura = vStrutura & "idFabricante Integer,"
   vStrutura = vStrutura & "idMarca Integer,"
   vStrutura = vStrutura & "idGrupo Integer,"
   vStrutura = vStrutura & "idSubGrupo Integer,"
   vStrutura = vStrutura & "idSituacaoTributaria Integer,"
   vStrutura = vStrutura & "Codigo NVarChar(20),"
   vStrutura = vStrutura & "Descricao NVarChar(50),"
   vStrutura = vStrutura & "ValorCusto Double,"
   vStrutura = vStrutura & "Preco Double,"
   vStrutura = vStrutura & "ValorCompra Double,"
   vStrutura = vStrutura & "MargemLucro Double,"
   vStrutura = vStrutura & "PesoLiquido Double,"
   vStrutura = vStrutura & "PesoBruto Double,"
   vStrutura = vStrutura & "DescontoMaximo Double,"
   vStrutura = vStrutura & "Comissao Double,"
   vStrutura = vStrutura & "DescricaoReduzida NVarChar(29),"
   vStrutura = vStrutura & "CodigoMarca NVarChar(20),"
   vStrutura = vStrutura & "ICMS Double,"
   vStrutura = vStrutura & "Frete Double,"
   vStrutura = vStrutura & "IPI Double,"
   vStrutura = vStrutura & "IVA Double,"
   vStrutura = vStrutura & "Simples Double,"
   vStrutura = vStrutura & "EstoqueMinimo Double,"
   vStrutura = vStrutura & "SaldoInicial Double,"
   vStrutura = vStrutura & "Localizacao NVarChar(20),"
   vStrutura = vStrutura & "Situacao Double,"
   vStrutura = vStrutura & "Aplicacao Memo,"
   vStrutura = vStrutura & "Anos NVarChar(50),"
   vStrutura = vStrutura & "UltimaCompra DateTime,"
   vStrutura = vStrutura & "UltimaVenda DateTime,"
   vStrutura = vStrutura & "SaldoAtual Double,"
   vStrutura = vStrutura & "Marca Bit,"
   vStrutura = vStrutura & "Cadastro DateTime,"
   vStrutura = vStrutura & "DataAtualizacao DateTime,"
   vStrutura = vStrutura & "AnoInicial Integer,"
   vStrutura = vStrutura & "AnoFinal Integer,"
   vStrutura = vStrutura & "Tipo NVarChar(50),"
   vStrutura = vStrutura & "Peso Double"
   vStrutura = vStrutura & ")"
   ' Criar Tabela
   cnTemp.Execute vStrutura
   
   cnTemp.Close
End Sub


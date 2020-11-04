VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmOpcoes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Opções"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6900
   Icon            =   "frmOpcoes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   6900
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmAplicar 
      Caption         =   "&Aplicar"
      Height          =   375
      Left            =   5880
      TabIndex        =   37
      Top             =   5700
      Width           =   975
   End
   Begin VB.CommandButton cmdFechar 
      Caption         =   "Fechar"
      Height          =   375
      Left            =   4800
      TabIndex        =   36
      Top             =   5700
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   3720
      TabIndex        =   35
      Top             =   5700
      Width           =   975
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5595
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   6795
      _ExtentX        =   11986
      _ExtentY        =   9869
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Preenchimento"
      TabPicture(0)   =   "frmOpcoes.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraEnderecoPadrao"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraFrenteLoja"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraCaminhos"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fraUtilitarios"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "fraImpressoras"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Informações Adicionais"
      TabPicture(1)   =   "frmOpcoes.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraConfigImpressoraFiscal"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "fraInicializacao"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "fraBancoDados"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "fraFormulario"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      Begin VB.Frame fraFormulario 
         Caption         =   "Formulário"
         Height          =   735
         Left            =   -74910
         TabIndex        =   60
         Top             =   3720
         Width           =   6615
         Begin VB.OptionButton optFormServicos 
            Caption         =   "Serviços"
            Height          =   255
            Left            =   2640
            TabIndex        =   63
            Top             =   300
            Width           =   1035
         End
         Begin VB.OptionButton optFormManual 
            Caption         =   "Manual"
            Height          =   255
            Left            =   1260
            TabIndex        =   62
            Top             =   300
            Width           =   855
         End
         Begin VB.OptionButton optFormSEPD 
            Caption         =   "SEPD"
            Height          =   255
            Left            =   120
            TabIndex        =   61
            Top             =   300
            Width           =   1095
         End
      End
      Begin VB.Frame fraBancoDados 
         Caption         =   "Banco de Dados"
         Height          =   795
         Left            =   -74940
         TabIndex        =   58
         Top             =   4500
         Width           =   6630
         Begin VB.CommandButton cmdTrocarBase 
            Caption         =   "&OK"
            Default         =   -1  'True
            Height          =   315
            Left            =   5820
            TabIndex        =   69
            Top             =   300
            Width           =   675
         End
         Begin VB.ComboBox cboEmpresas 
            Height          =   315
            ItemData        =   "frmOpcoes.frx":0342
            Left            =   3240
            List            =   "frmOpcoes.frx":034F
            Style           =   2  'Dropdown List
            TabIndex        =   68
            Top             =   300
            Width           =   2535
         End
         Begin VB.CommandButton cmdManutencaoBancoDados 
            Caption         =   "Manutenção"
            Height          =   435
            Left            =   120
            TabIndex        =   59
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Frame fraInicializacao 
         Caption         =   "Inicialização"
         Height          =   1290
         Left            =   -74910
         TabIndex        =   47
         Top             =   2400
         Width           =   6630
         Begin VB.TextBox txtMensagem2 
            Height          =   315
            Left            =   945
            MaxLength       =   100
            TabIndex        =   55
            Top             =   870
            Width           =   4125
         End
         Begin VB.CheckBox chkPesquisaDescricao 
            Caption         =   "Pesquisa Descrição"
            Height          =   255
            Left            =   3960
            TabIndex        =   57
            Top             =   270
            Width           =   2235
         End
         Begin VB.TextBox txtMensagem 
            Height          =   315
            Left            =   945
            MaxLength       =   100
            TabIndex        =   54
            Top             =   540
            Width           =   4125
         End
         Begin MSMask.MaskEdBox mskDecimaisQuantidade 
            Height          =   285
            Left            =   2220
            TabIndex        =   48
            Top             =   255
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   3
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskDecimaisValor 
            Height          =   285
            Left            =   3300
            TabIndex        =   52
            Top             =   255
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   3
            PromptChar      =   "_"
         End
         Begin VB.Label lblMensagem 
            AutoSize        =   -1  'True
            Caption         =   "Mensagem"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   75
            TabIndex        =   56
            Top             =   600
            Width           =   780
         End
         Begin VB.Label lblDecimaisValor 
            AutoSize        =   -1  'True
            Caption         =   "Valor"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   2820
            TabIndex        =   51
            Top             =   300
            Width           =   360
         End
         Begin VB.Label lblDecimaisQuantidade 
            AutoSize        =   -1  'True
            Caption         =   "Quantidade"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   1320
            TabIndex        =   50
            Top             =   300
            Width           =   825
         End
         Begin VB.Label lblCasasDecimais 
            AutoSize        =   -1  'True
            Caption         =   "Casas Decimais:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   60
            TabIndex        =   49
            Top             =   300
            Width           =   1170
         End
      End
      Begin VB.Frame fraImpressoras 
         Caption         =   "Impressoras"
         Height          =   1575
         Left            =   60
         TabIndex        =   38
         Top             =   3900
         Width           =   6615
         Begin VB.TextBox txtImpressoraOrcamentos 
            Height          =   315
            Left            =   1140
            MaxLength       =   100
            TabIndex        =   45
            Top             =   1170
            Width           =   4185
         End
         Begin VB.TextBox txtImpressoraBoletos2 
            Height          =   315
            Left            =   1140
            MaxLength       =   100
            TabIndex        =   43
            Top             =   840
            Width           =   4185
         End
         Begin VB.TextBox txtImpressoraBoletos1 
            Height          =   315
            Left            =   1140
            MaxLength       =   100
            TabIndex        =   41
            Top             =   510
            Width           =   4185
         End
         Begin VB.TextBox txtImpressoraNotas 
            Height          =   315
            Left            =   1140
            MaxLength       =   100
            TabIndex        =   39
            Top             =   180
            Width           =   4185
         End
         Begin VB.Label lbImpressoralOrcamentos 
            AutoSize        =   -1  'True
            Caption         =   "Orçamentos"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   46
            Top             =   1230
            Width           =   855
         End
         Begin VB.Label lblImpressoraBoletos2 
            AutoSize        =   -1  'True
            Caption         =   "Boletos 2"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   315
            TabIndex        =   44
            Top             =   900
            Width           =   660
         End
         Begin VB.Label lblImpressoraBoletos1 
            AutoSize        =   -1  'True
            Caption         =   "Boletos 1"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   315
            TabIndex        =   42
            Top             =   570
            Width           =   660
         End
         Begin VB.Label lblImpressoraNotas 
            AutoSize        =   -1  'True
            Caption         =   "Notas"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   555
            TabIndex        =   40
            Top             =   240
            Width           =   420
         End
      End
      Begin VB.Frame fraUtilitarios 
         Caption         =   "Utilitários"
         Height          =   855
         Left            =   60
         TabIndex        =   22
         Top             =   3000
         Width           =   6615
         Begin VB.CheckBox chkTipoImpressao 
            Caption         =   "Impressão Rápida"
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   510
            Width           =   2115
         End
         Begin VB.CheckBox chkExibirAgenda 
            Caption         =   "Exibir Agenda ao Entrar"
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   300
            Width           =   2115
         End
      End
      Begin VB.Frame fraCaminhos 
         Caption         =   "Caminhos"
         Height          =   1035
         Left            =   60
         TabIndex        =   17
         Top             =   1920
         Width           =   6630
         Begin VB.CommandButton cmdProcurarSintegra 
            Caption         =   "Pr&ocurar"
            Height          =   315
            Left            =   5640
            TabIndex        =   65
            Top             =   600
            Width           =   855
         End
         Begin VB.TextBox txtCaminhoSintegra 
            Height          =   315
            Left            =   1380
            MaxLength       =   100
            TabIndex        =   64
            Top             =   600
            Width           =   4185
         End
         Begin VB.TextBox txtCaminhoBanco 
            Height          =   315
            Left            =   1380
            MaxLength       =   100
            TabIndex        =   18
            Top             =   240
            Width           =   4185
         End
         Begin VB.CommandButton cmdProcurarBanco 
            Caption         =   "Pr&ocurar"
            Height          =   315
            Left            =   5640
            TabIndex        =   19
            Top             =   240
            Width           =   855
         End
         Begin VB.Label lblLocalSintegra 
            AutoSize        =   -1  'True
            Caption         =   "Sintegra"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   735
            TabIndex        =   66
            Top             =   660
            Width           =   585
         End
         Begin VB.Label lblLocalBanco 
            AutoSize        =   -1  'True
            Caption         =   "Banco de Dados"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   16
            Top             =   300
            Width           =   1200
         End
      End
      Begin VB.Frame fraFrenteLoja 
         Caption         =   "Frente de Loja"
         Height          =   645
         Left            =   60
         TabIndex        =   15
         Top             =   1200
         Width           =   6630
         Begin VB.CheckBox chkImprimirOrcamento 
            Caption         =   "Imprimir Orçamento"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Width           =   1755
         End
         Begin VB.ComboBox cmbConfirmacao 
            Height          =   315
            ItemData        =   "frmOpcoes.frx":036D
            Left            =   4800
            List            =   "frmOpcoes.frx":037A
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   210
            Width           =   1755
         End
         Begin MSMask.MaskEdBox mskPercMaxDesconto 
            Height          =   285
            Left            =   3120
            TabIndex        =   12
            Top             =   225
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   6
            Format          =   "##0.00"
            PromptChar      =   "_"
         End
         Begin VB.Label lblConfirmacao 
            AutoSize        =   -1  'True
            Caption         =   "Confirmação"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   3840
            TabIndex        =   13
            Top             =   270
            Width           =   885
         End
         Begin VB.Label lblPercMaxDesconto 
            AutoSize        =   -1  'True
            Caption         =   "Perc. Desconto"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   1920
            TabIndex        =   11
            Top             =   270
            Width           =   1110
         End
      End
      Begin VB.Frame fraConfigImpressoraFiscal 
         Caption         =   "Configurações da Impressora Fiscal"
         Height          =   1995
         Left            =   -74910
         TabIndex        =   34
         Top             =   390
         Width           =   6630
         Begin VB.CheckBox chkHorarioVerao 
            Caption         =   "Horário de Verão"
            Height          =   255
            Left            =   120
            TabIndex        =   67
            Top             =   1620
            Width           =   2235
         End
         Begin VB.CheckBox chkAbrirGaveta 
            Caption         =   "Abrir Gaveta de Dinheiro"
            Height          =   255
            Left            =   120
            TabIndex        =   53
            Top             =   1320
            Width           =   2235
         End
         Begin VB.CommandButton cmdProcurarArquivoConfig 
            Caption         =   "Pr&ocurar"
            Height          =   315
            Left            =   5520
            TabIndex        =   26
            Top             =   600
            Width           =   975
         End
         Begin VB.TextBox txtLocalArquivo 
            Height          =   315
            Left            =   1920
            TabIndex        =   25
            Top             =   600
            Width           =   3555
         End
         Begin VB.ComboBox cboModelo 
            Height          =   315
            ItemData        =   "frmOpcoes.frx":0398
            Left            =   720
            List            =   "frmOpcoes.frx":03A2
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   930
            Width           =   1755
         End
         Begin VB.Frame fraPortaImpressora 
            Caption         =   "Porta"
            Height          =   915
            Left            =   5475
            TabIndex        =   33
            Top             =   960
            Width           =   1035
            Begin VB.OptionButton optCOM2 
               Caption         =   "COM 2"
               Height          =   255
               Left            =   120
               TabIndex        =   32
               Top             =   540
               Width           =   855
            End
            Begin VB.OptionButton optCOM1 
               Caption         =   "COM 1"
               Height          =   255
               Left            =   120
               TabIndex        =   31
               Top             =   240
               Width           =   855
            End
         End
         Begin VB.CheckBox chkAtivaImpressora 
            Caption         =   "Ativa ECF"
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   300
            Width           =   1755
         End
         Begin MSMask.MaskEdBox mskHorarioReducao 
            Height          =   285
            Left            =   3780
            TabIndex        =   30
            Top             =   945
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   8
            Mask            =   "99:99:99"
            PromptChar      =   " "
         End
         Begin VB.Label lblHorario 
            AutoSize        =   -1  'True
            Caption         =   "Hor. Redução Z"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   2520
            TabIndex        =   29
            Top             =   990
            Width           =   1155
         End
         Begin VB.Label lblArquivoConfiguracaoFiscal 
            AutoSize        =   -1  'True
            Caption         =   "Arquivo de Configuração"
            Height          =   195
            Left            =   120
            TabIndex        =   24
            Top             =   660
            Width           =   1755
         End
         Begin VB.Label lblModelo 
            AutoSize        =   -1  'True
            Caption         =   "Modelo"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   27
            Top             =   990
            Width           =   525
         End
      End
      Begin VB.Frame fraEnderecoPadrao 
         Caption         =   "Informações Padronizadas de Endereço"
         Height          =   735
         Left            =   75
         TabIndex        =   9
         Top             =   405
         Width           =   6630
         Begin VB.ComboBox cmbPrefixoFone 
            Height          =   315
            ItemData        =   "frmOpcoes.frx":03B8
            Left            =   5760
            List            =   "frmOpcoes.frx":03BA
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   285
            Width           =   675
         End
         Begin VB.TextBox txtCidade 
            Height          =   285
            Left            =   705
            MaxLength       =   30
            TabIndex        =   2
            Top             =   300
            Width           =   2355
         End
         Begin VB.TextBox txtUF 
            Height          =   285
            Left            =   3405
            MaxLength       =   2
            TabIndex        =   4
            Top             =   300
            Width           =   375
         End
         Begin MSMask.MaskEdBox mskCEP 
            Height          =   285
            Left            =   4215
            TabIndex        =   6
            Top             =   300
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "99.999-999"
            PromptChar      =   " "
         End
         Begin VB.Label lblPrefixo 
            AutoSize        =   -1  'True
            Caption         =   "Prefixo"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   5250
            TabIndex        =   7
            Top             =   345
            Width           =   480
         End
         Begin VB.Label lblUF 
            AutoSize        =   -1  'True
            Caption         =   "UF"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   3120
            TabIndex        =   3
            Top             =   345
            Width           =   210
         End
         Begin VB.Label lblCEP 
            AutoSize        =   -1  'True
            Caption         =   "CEP"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   3855
            TabIndex        =   5
            Top             =   345
            Width           =   315
         End
         Begin VB.Label lblCidade 
            AutoSize        =   -1  'True
            Caption         =   "Cidade"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   1
            Top             =   345
            Width           =   495
         End
      End
   End
   Begin MSComDlg.CommonDialog dlgCaminhoBanco 
      Left            =   6960
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog dlgCaminhoSintegra 
      Left            =   7440
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmOpcoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ItemList As ListItem
Dim rsTemp As New ADODB.Recordset

Private Sub Form_Load()
   On Error GoTo Erro
   SSTab1.Tab = 0 ' Posiciona no primeiro tab
   Centraliza frmOpcoes
   
   Carrega_Combos
   Carrega_Campos
   
'   MDISistema.StatusBar.Panels(1).text = "Opções de Configuração do Sistema"

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

Private Sub cmdOK_Click()

   Atualizar_INIs

'   cnSistema.Close
'   Call ConnectDB
   
'   Set rsTemp = cnSistema.Execute("Select * from Empresa")
'   If Not rsTemp.EOF Then MDISistema.Caption = rsTemp!Nome
   
   Unload Me
End Sub

Private Sub cmAplicar_Click()

   Atualizar_INIs

'   cnSistema.Close
'   Call ConnectDB
End Sub

Private Sub cmdFechar_Click()
   Unload Me
End Sub

Private Sub cmdProcurarBanco_Click()
   dlgCaminhoBanco.FileName = ""
   dlgCaminhoBanco.Filter = "Todos os Arquivos|*.*"
   dlgCaminhoBanco.ShowOpen
   If dlgCaminhoBanco.FileName <> "" Then
      txtCaminhoBanco.Text = Mid(dlgCaminhoBanco.FileName, 1, Len(dlgCaminhoBanco.FileName) - Len(dlgCaminhoBanco.FileTitle) - 1)
   End If
End Sub

Private Sub cmdProcurarSintegra_Click()
   dlgCaminhoSintegra.FileName = ""
   dlgCaminhoSintegra.Filter = "Todos os Arquivos|*.*"
   dlgCaminhoSintegra.ShowOpen
   If dlgCaminhoSintegra.FileName <> "" Then
      txtCaminhoSintegra.Text = Mid(dlgCaminhoSintegra.FileName, 1, Len(dlgCaminhoSintegra.FileName) - Len(dlgCaminhoSintegra.FileTitle) - 1)
   End If
End Sub

Private Sub cmdProcurarArquivoConfig_Click()
   dlgCaminhoBanco.FileName = ""
   dlgCaminhoBanco.Filter = "Todos os Arquivos|*.*"
   dlgCaminhoBanco.ShowOpen
   If dlgCaminhoBanco.FileName <> "" Then
      txtLocalArquivo.Text = Mid(dlgCaminhoBanco.FileName, 1, Len(dlgCaminhoBanco.FileName) - Len(dlgCaminhoBanco.FileTitle) - 1)
   End If
End Sub

Private Sub Carrega_Combos()

'  Prefixos Telefônicos
   If I_SGBD = "ACCESS" Then
      Set rsTemp = cnSistema.Execute("Select * From PrefixosTelefonicos ORDER BY Prefixo")
      cmbPrefixoFone.Clear
      
      cmbPrefixoFone.AddItem ""
      cmbPrefixoFone.ItemData(cmbPrefixoFone.NewIndex) = 0
   
      Do While Not rsTemp.EOF
         cmbPrefixoFone.AddItem rsTemp!Prefixo
         cmbPrefixoFone.ItemData(cmbPrefixoFone.NewIndex) = Val(rsTemp!Prefixo)
         
         rsTemp.MoveNext
      Loop
      rsTemp.Close
   End If
   
   ' Empresas
   Dim sEmpresas As String
   Dim sNomeEmpresa As String
   Dim Contador As Integer
   Dim idContador As Integer
   
   sEmpresas = LerArquivoINI("Banco de Dados", "Empresas", CaminhoINI & "\System.ini")
   sNomeEmpresa = ""
   idContador = 1
   
   cboEmpresas.Clear
   For Contador = 1 To Len(sEmpresas)
      If Mid(sEmpresas, Contador, 1) <> ";" Then
         sNomeEmpresa = sNomeEmpresa & Mid(sEmpresas, Contador, 1)
      Else
         cboEmpresas.AddItem sNomeEmpresa
         cboEmpresas.ItemData(cboEmpresas.NewIndex) = idContador
         
         sNomeEmpresa = ""
         idContador = idContador + 1
      End If
   Next
   
End Sub

Private Sub Carrega_Campos()

   ' Banco de Dados
   txtCaminhoBanco.Text = LerArquivoINI("Banco de Dados", "Caminho", CaminhoINI & "\System.ini")
   txtCaminhoSintegra.Text = LerArquivoINI("Banco de Dados", "Sintegra", CaminhoINI & "\System.ini")

   ' Impressora Fiscal
   txtLocalArquivo.Text = LerArquivoINI("Impressora Fiscal", "Caminho", CaminhoINI & "\System.ini")
   chkAtivaImpressora.value = LerArquivoINI("Impressora Fiscal", "Ativar", CaminhoINI & "\System.ini")
   If LerArquivoINI("Impressora Fiscal", "Porta", CaminhoINI & "\System.ini") = 1 Then
      optCOM1.value = True
   ElseIf LerArquivoINI("Impressora Fiscal", "Porta", CaminhoINI & "\System.ini") = 2 Then
      optCOM2.value = True
   End If
   cboModelo.ListIndex = LerArquivoINI("Impressora Fiscal", "Modelo", CaminhoINI & "\System.ini")
   mskHorarioReducao.Text = LerArquivoINI("Impressora Fiscal", "Horario", CaminhoINI & "\System.ini")
   chkAbrirGaveta.value = LerArquivoINI("Impressora Fiscal", "Gaveta", CaminhoINI & "\System.ini")

   ' SEPD
'   chkSEPD.Value = LerArquivoINI("SEPD", "TipoImpressao", CaminhoINI & "\System.ini")
   If LerArquivoINI("SEPD", "TipoImpressao", CaminhoINI & "\System.ini") = 1 Then
      optFormSEPD.value = True
   ElseIf LerArquivoINI("SEPD", "TipoImpressao", CaminhoINI & "\System.ini") = 2 Then
      optFormManual.value = True
   ElseIf LerArquivoINI("SEPD", "TipoImpressao", CaminhoINI & "\System.ini") = 3 Then
      optFormServicos.value = True
   End If
   
   ' Preenchimento
   txtCidade.Text = LerArquivoINI("Preenchimento", "Cidade", CaminhoINI & "\System.ini")
   txtUF.Text = LerArquivoINI("Preenchimento", "UF", CaminhoINI & "\System.ini")
   mskCEP.Text = IIf(Mid(LerArquivoINI("Preenchimento", "CEP", CaminhoINI & "\System.ini"), 1, 1) <> ".", LerArquivoINI("Preenchimento", "CEP", CaminhoINI & "\System.ini"), "  .   -   ")
   If I_SGBD = "ACCESS" Then
      cmbPrefixoFone.Text = LerArquivoINI("Preenchimento", "Prefixo", CaminhoINI & "\System.ini")
   End If
   mskDecimaisQuantidade.Text = LerArquivoINI("Preenchimento", "DecimaisQuantidade", CaminhoINI & "\System.ini")
   mskDecimaisValor.Text = LerArquivoINI("Preenchimento", "DecimaisValor", CaminhoINI & "\System.ini")
   
   ' Orcamentos
   chkImprimirOrcamento.value = LerArquivoINI("Orcamentos", "Imprimir", CaminhoINI & "\System.ini")
   cmbConfirmacao.ListIndex = LerArquivoINI("Orcamentos", "Confirmacao", CaminhoINI & "\System.ini")
'''   Erro = CriarINI("Orcamentos", "ECF", "1")
   mskPercMaxDesconto.Text = LerArquivoINI("Orcamentos", "PercMaxDesconto", CaminhoINI & "\System.ini")
   txtMensagem.Text = LerArquivoINI("Orcamentos", "Mensagem", CaminhoINI & "\System.ini")
   txtMensagem2.Text = LerArquivoINI("Orcamentos", "Mensagem2", CaminhoINI & "\System.ini")
   chkPesquisaDescricao.value = LerArquivoINI("Orcamentos", "PesquisaDescricao", CaminhoINI & "\System.ini")
   
   ' Impressoras
   txtImpressoraNotas.Text = LerArquivoINI("Impressoras", "Notas", CaminhoINI & "\System.ini")
   txtImpressoraBoletos1.Text = LerArquivoINI("Impressoras", "Boletos1", CaminhoINI & "\System.ini")
   txtImpressoraBoletos2.Text = LerArquivoINI("Impressoras", "Boletos2", CaminhoINI & "\System.ini")
   txtImpressoraOrcamentos.Text = LerArquivoINI("Impressoras", "Orcamentos", CaminhoINI & "\System.ini")
   chkTipoImpressao.value = LerArquivoINI("Impressoras", "TipoImpressao", CaminhoINI & "\System.ini")
   
   ' Notas Fiscais Eletrônicas
   chkHorarioVerao.value = LerArquivoINI("NFe", "HorarioVerao", CaminhoINI & "\System.ini")
   
   ' Uteis
   chkExibirAgenda.value = LerArquivoINI("Uteis", "ExibirAgenda", CaminhoINI & "\System.ini")
End Sub

Private Sub Atualizar_INIs()
Dim Erro As Boolean

   ' Banco de Dados
   Erro = AtualizarINI("Banco de Dados", "Caminho", txtCaminhoBanco.Text)
   Erro = AtualizarINI("Banco de Dados", "Sintegra", txtCaminhoSintegra.Text)

   ' Impressora Fiscal
   Erro = AtualizarINI("Impressora Fiscal", "Caminho", txtLocalArquivo.Text)
   Erro = AtualizarINI("Impressora Fiscal", "Ativar", chkAtivaImpressora.value)
   If optCOM1.value = True Then
      Erro = AtualizarINI("Impressora Fiscal", "Porta", "1")
   ElseIf optCOM2.value = True Then
      Erro = AtualizarINI("Impressora Fiscal", "Porta", "2")
   End If
   Erro = AtualizarINI("Impressora Fiscal", "Modelo", cboModelo.ListIndex)
   Erro = AtualizarINI("Impressora Fiscal", "Horario", mskHorarioReducao.Text)
   Erro = AtualizarINI("Impressora Fiscal", "Gaveta", chkAbrirGaveta.value)

   ' SEPD
'   Erro = AtualizarINI("SEPD", "TipoImpressao", chkSEPD.Value)
   If optFormSEPD.value = True Then
      Erro = AtualizarINI("SEPD", "TipoImpressao", "1")
   ElseIf optFormManual.value = True Then
      Erro = AtualizarINI("SEPD", "TipoImpressao", "2")
   ElseIf optFormServicos.value = True Then
      Erro = AtualizarINI("SEPD", "TipoImpressao", "3")
   End If
   
   ' Preenchimento
   Erro = AtualizarINI("Preenchimento", "Cidade", SQLCheck(txtCidade.Text))
   Erro = AtualizarINI("Preenchimento", "UF", SQLCheck(txtUF.Text))
   Erro = AtualizarINI("Preenchimento", "CEP", SQLCheck(mskCEP.Text))
   Erro = AtualizarINI("Preenchimento", "Prefixo", cmbPrefixoFone.Text)
   Erro = AtualizarINI("Preenchimento", "DecimaisQuantidade", mskDecimaisQuantidade.Text)
   Erro = AtualizarINI("Preenchimento", "DecimaisValor", mskDecimaisValor.Text)
   
   ' Orcamentos
   Erro = AtualizarINI("Orcamentos", "Imprimir", chkImprimirOrcamento.value)
   Erro = AtualizarINI("Orcamentos", "Confirmacao", cmbConfirmacao.ListIndex)
''   Erro = AtualizarINI("Orcamentos", "ECF", "1")
   Erro = AtualizarINI("Orcamentos", "PercMaxDesconto", mskPercMaxDesconto.Text)
   Erro = AtualizarINI("Orcamentos", "Mensagem", txtMensagem.Text)
   Erro = AtualizarINI("Orcamentos", "Mensagem2", txtMensagem2.Text)
   Erro = AtualizarINI("Orcamentos", "PesquisaDescricao", chkPesquisaDescricao.value)
   
   ' Impressoras
   Erro = AtualizarINI("Impressoras", "Notas", txtImpressoraNotas.Text)
   Erro = AtualizarINI("Impressoras", "Boletos1", txtImpressoraBoletos1.Text)
   Erro = AtualizarINI("Impressoras", "Boletos2", txtImpressoraBoletos2.Text)
   Erro = AtualizarINI("Impressoras", "Orcamentos", txtImpressoraOrcamentos.Text)
   Erro = AtualizarINI("Impressoras", "TipoImpressao", chkTipoImpressao.value)
   
   ' Notas Fiscais Eletrônicas
   Erro = AtualizarINI("NFe", "HorarioVerao", chkHorarioVerao.value)
   
   ' Uteis
   Erro = AtualizarINI("Uteis", "ExibirAgenda", chkExibirAgenda.value)
End Sub

Private Sub cmdManutencaoBancoDados_Click()
On Error GoTo Trata_Erro
Dim vStrutura As String
Dim Teste As Integer
Dim Campo As Variant

'   cnSistema.Execute ("DROP TABLE Teste")
'   cnSistema.Execute "CREATE TABLE Produtos(idProduto Integer IDENTITY Primary Key NOT NULL," & _
'                  "idClasse Integer," & _
'                  "idUnidade Integer," & _
'                  "idRegistradorFiscal Integer," & _
'                  "idFabricante Integer," & _
'                  "idMarca Integer," & _
'                  "idGrupo Integer," & _
'                  "idSubGrupo Integer," & _
'                  "Codigo NVarChar(20)," & _
'                  "Descricao NVarChar(50)," & _
'                  "ValorCusto Double, " & _
'                  "Aplicacao Memo, " & _
'                  "Cadastro DateTime, " & _
'                  "Preco Double)"
'   cnSistema.Execute ("CREATE TABLE Produtos(" & vStrutura & ")")

''   cnSistema.Execute ("DROP TABLE Teste")
''   cnSistema.Execute "CREATE TABLE Teste(idProduto Integer IDENTITY Primary Key NOT NULL)"

' Produtos
'   Set rsTemp = cnSistema.Execute("Select * From Teste")
   
               Dim Campo2 As String
               Campo2 = "ALTER TABLE Teste ADD COLUMN idClasse INTEGER"
               cnSistema.Execute Campo2
   
'''   Campo = 1
'''   Teste = rsTemp!idProduto
'''   Campo = 2
'''   Teste = rsTemp!idClasse
'''   Campo = 3
'''   Teste = rsTemp!idUnidade
'''   Campo = 4
'''   Teste = rsTemp!idRegistradorFiscal
'''   Campo = 5
'''   Teste = rsTemp!idFabricante
'''   Campo = 6
'''   Teste = rsTemp!idMarca
'''   Campo = 7
'''   Teste = rsTemp!idGrupo
'''   Campo = 8
'''   Teste = rsTemp!idSubGrupo
'''   Campo = 9
'''   Teste = rsTemp!idSituacaoTributaria
'''   Campo = 10
'''   Teste = rsTemp!Codigo
'''   Campo = 11
'''   Teste = rsTemp!Descricao
'''   Campo = 12
'''   Teste = rsTemp!ValorCusto
'''   Campo = 13
'''   Teste = rsTemp!Preco
'''   Campo = 14
'''   Teste = rsTemp!ValorCompra
'''   Campo = 15
'''   Teste = rsTemp!MargemLucro
'''   Campo = 16
'''   Teste = rsTemp!PesoLiquido
'''   Campo = 17
'''   Teste = rsTemp!PesoBruto
'''   Campo = 18
'''   Teste = rsTemp!DescontoMaximo
'''   Campo = 19
'''   Teste = rsTemp!Comissao
'''   Campo = 20
'''   Teste = rsTemp!DescricaoReduzida
'''   Campo = 21
'''   Teste = rsTemp!CodigoMarca
'''   Campo = 22
'''   Teste = rsTemp!ICMS
'''   Campo = 23
'''   Teste = rsTemp!Frete
'''   Campo = 24
'''   Teste = rsTemp!IPI
'''   Campo = 25
'''   Teste = rsTemp!IVA
'''   Campo = 26
'''   Teste = rsTemp!Simples
'''   Campo = 27
'''   Teste = rsTemp!EstoqueMinimo
'''   Campo = 28
'''   Teste = rsTemp!SaldoInicial
'''   Campo = 29
'''   Teste = rsTemp!Localizacao
'''   Campo = 30
'''   Teste = rsTemp!Situacao
'''   Campo = 31
'''   Teste = rsTemp!Aplicacao
'''   Campo = 32
'''   Teste = rsTemp!Anos
'''   Campo = 33
'''   Teste = rsTemp!UltimaCompra
'''   Campo = 34
'''   Teste = rsTemp!UltimaVenda
'''   Campo = 35
'''   Teste = rsTemp!SaldoAtual
'''   Campo = 36
'''   Teste = rsTemp!Marca
'''   Campo = 37
'''   Teste = rsTemp!Cadastro
'''   Campo = 38
'''   Teste = rsTemp!DataAtualizacao
'''   Campo = 39
'''   Teste = rsTemp!AnoInicial
'''   Campo = 40
'''   Teste = rsTemp!AnoFinal
'''   Campo = 41
'''   Teste = rsTemp!Tipo
'''   Campo = 42
'''   Teste = rsTemp!Peso

Exit Sub
Trata_Erro:

   If Err.Number = 3265 Then
      rsTemp.Close
      Select Case Campo
            Case 2
               Dim Campo3 As String
               Campo3 = "ALTER TABLE Teste ADD COLUMN idClasse INTEGER"
               cnSistema.Execute Campo3

'               cnSistema.Execute "ALTER TABLE Teste ADD COLUMN idClasse INTEGER"
               
''   vStrutura = vStrutura & "CREATE TABLE Produtos(idProduto Integer IDENTITY Primary Key NOT NULL, "
''   vStrutura = vStrutura & "idClasse Integer,"
''   vStrutura = vStrutura & "idUnidade Integer,"
''   vStrutura = vStrutura & "idRegistradorFiscal Integer,"
''   vStrutura = vStrutura & "idFabricante Integer,"
''   vStrutura = vStrutura & "idMarca Integer,"
''   vStrutura = vStrutura & "idGrupo Integer,"
''   vStrutura = vStrutura & "idSubGrupo Integer,"
''   vStrutura = vStrutura & "idSituacaoTributaria Integer,"
''   vStrutura = vStrutura & "Codigo NVarChar(20),"
''   vStrutura = vStrutura & "Descricao NVarChar(50),"
''   vStrutura = vStrutura & "ValorCusto Double,"
''   vStrutura = vStrutura & "Preco Double,"
''   vStrutura = vStrutura & "ValorCompra Double,"
''   vStrutura = vStrutura & "MargemLucro Double,"
''   vStrutura = vStrutura & "PesoLiquido Double,"
''   vStrutura = vStrutura & "PesoBruto Double,"
''   vStrutura = vStrutura & "DescontoMaximo Double,"
''   vStrutura = vStrutura & "Comissao Double,"
''   vStrutura = vStrutura & "DescricaoReduzida NVarChar(29),"
''   vStrutura = vStrutura & "CodigoMarca NVarChar(20),"
''   vStrutura = vStrutura & "ICMS Double,"
''   vStrutura = vStrutura & "Frete Double,"
''   vStrutura = vStrutura & "IPI Double,"
''   vStrutura = vStrutura & "IVA Double,"
''   vStrutura = vStrutura & "Simples Double,"
''   vStrutura = vStrutura & "EstoqueMinimo Double,"
''   vStrutura = vStrutura & "SaldoInicial Double,"
''   vStrutura = vStrutura & "Localizacao NVarChar(20),"
''   vStrutura = vStrutura & "Situacao Double,"
''   vStrutura = vStrutura & "Aplicacao Memo,"
''   vStrutura = vStrutura & "Anos NVarChar(50),"
''   vStrutura = vStrutura & "UltimaCompra DateTime,"
''   vStrutura = vStrutura & "UltimaVenda DateTime,"
''   vStrutura = vStrutura & "SaldoAtual Double,"
''   vStrutura = vStrutura & "Marca Bolean,"
''   vStrutura = vStrutura & "Cadastro DateTime,"
''   vStrutura = vStrutura & "DataAtualizacao DateTime,"
''   vStrutura = vStrutura & "AnoInicial Integer,"
''   vStrutura = vStrutura & "AnoFinal Integer,"
''   vStrutura = vStrutura & "Tipo NVarChar(50),"
''   vStrutura = vStrutura & "Peso Double"
               
      End Select
   End If
End Sub

Private Sub cmdTrocarBase_Click()
Dim sCaminho As String
Dim sSintegra As String
Dim sServidor As String
Dim sCatalog As String
Dim sSGBD As String
Dim sEmpresa As String
Dim sModelo As String

'Caminho=C:\Proj\GerenciadorNFs\Dados
'Sintegra=D:\Projetos\Comercial\DadosModelo\Dados
'Servidor SQL=.
'Catalog = Gestao
'SGBD = ACCESS

   sCaminho = LerArquivoINI("Banco de Dados", "Caminho", CaminhoINI & "\cf" & cboEmpresas.Text & ".ini")
   sSintegra = LerArquivoINI("Banco de Dados", "Sintegra", CaminhoINI & "\cf" & cboEmpresas.Text & ".ini")
   sServidor = LerArquivoINI("Banco de Dados", "Servidor SQL", CaminhoINI & "\cf" & cboEmpresas.Text & ".ini")
   sCatalog = LerArquivoINI("Banco de Dados", "Catalog", CaminhoINI & "\cf" & cboEmpresas.Text & ".ini")
   sSGBD = LerArquivoINI("Banco de Dados", "SGBD", CaminhoINI & "\cf" & cboEmpresas.Text & ".ini")
   
   sEmpresa = LerArquivoINI("NFe", "Empresa", CaminhoINI & "\cf" & cboEmpresas.Text & ".ini")
   sModelo = LerArquivoINI("NFe", "Modelo", CaminhoINI & "\cf" & cboEmpresas.Text & ".ini")

   If Not GravaArquivoINI("Banco de Dados", "Caminho", sCaminho, App.Path & "\System.ini") Then
      MsgBox "Não foi possível gravar arquivo de configuração" & Chr(13) & "Entre em contato com o suporte", vbInformation + vbOKOnly, "Sistema"
      End
   End If
   
   If Not GravaArquivoINI("Banco de Dados", "Sintegra", sSintegra, App.Path & "\System.ini") Then
      MsgBox "Não foi possível gravar arquivo de configuração" & Chr(13) & "Entre em contato com o suporte", vbInformation + vbOKOnly, "Sistema"
      End
   End If
   
   If Not GravaArquivoINI("Banco de Dados", "Servidor SQL", sServidor, App.Path & "\System.ini") Then
      MsgBox "Não foi possível gravar arquivo de configuração" & Chr(13) & "Entre em contato com o suporte", vbInformation + vbOKOnly, "Sistema"
      End
   End If
   
   If Not GravaArquivoINI("Banco de Dados", "Catalog", sCatalog, App.Path & "\System.ini") Then
      MsgBox "Não foi possível gravar arquivo de configuração" & Chr(13) & "Entre em contato com o suporte", vbInformation + vbOKOnly, "Sistema"
      End
   End If
   
   If Not GravaArquivoINI("Banco de Dados", "SGBD", sSGBD, App.Path & "\System.ini") Then
      MsgBox "Não foi possível gravar arquivo de configuração" & Chr(13) & "Entre em contato com o suporte", vbInformation + vbOKOnly, "Sistema"
      End
   End If
   
   If Not GravaArquivoINI("NFe", "Empresa", sEmpresa, App.Path & "\System.ini") Then
      MsgBox "Não foi possível gravar arquivo de configuração" & Chr(13) & "Entre em contato com o suporte", vbInformation + vbOKOnly, "Sistema"
      End
   End If

   If Not GravaArquivoINI("NFe", "Modelo", sModelo, App.Path & "\System.ini") Then
      MsgBox "Não foi possível gravar arquivo de configuração" & Chr(13) & "Entre em contato com o suporte", vbInformation + vbOKOnly, "Sistema"
      End
   End If
   
   MsgBox "O sistema será reiniciado para aplicar as alterações", vbInformation + vbOKOnly, "Sistema"
   End
   
End Sub


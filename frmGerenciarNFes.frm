VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmGerenciarNFes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gerenciar Notas Fiscais Eletr�nicas"
   ClientHeight    =   8460
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12765
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8460
   ScaleWidth      =   12765
   Begin VB.CommandButton cmdNFeCartaCorrecao 
      Caption         =   "Carta de Corre��o"
      Enabled         =   0   'False
      Height          =   375
      Left            =   60
      TabIndex        =   22
      Top             =   5040
      Width           =   2115
   End
   Begin VB.Frame fraMensagens 
      Caption         =   "Mensagens de Retorno e Erros"
      Height          =   2835
      Left            =   60
      TabIndex        =   19
      Top             =   5580
      Width           =   12615
      Begin VB.CommandButton cmdHistorico 
         Caption         =   "Hist�rico"
         Height          =   375
         Left            =   10380
         TabIndex        =   20
         Top             =   2340
         Width           =   2115
      End
      Begin MSComctlLib.ListView lvwMensagens 
         Height          =   2025
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   12405
         _ExtentX        =   21881
         _ExtentY        =   3572
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
   Begin VB.Timer tmrLerRetornos 
      Interval        =   2000
      Left            =   12240
      Top             =   60
   End
   Begin VB.CommandButton cmdGerarComplementoICMS 
      Caption         =   "Complemento ICMS"
      Enabled         =   0   'False
      Height          =   375
      Left            =   10980
      TabIndex        =   18
      Top             =   5040
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.ComboBox cmbSituacao 
      Height          =   315
      Left            =   4800
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   90
      Width           =   2805
   End
   Begin VB.CommandButton cmdEnviarEmail 
      Caption         =   "Enviar E-mail"
      Height          =   375
      Left            =   8820
      TabIndex        =   17
      Top             =   4620
      Width           =   1695
   End
   Begin VB.CommandButton cmdConsultaNota 
      Caption         =   "Consulta Situa��o "
      Height          =   375
      Left            =   6360
      TabIndex        =   16
      Top             =   4620
      Width           =   1875
   End
   Begin VB.CommandButton cmdValidar 
      Caption         =   "&Validar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   9720
      TabIndex        =   15
      Top             =   5040
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.CommandButton cmdFecharNFe 
      Caption         =   "Fechar NFe"
      Height          =   375
      Left            =   8940
      TabIndex        =   14
      Top             =   60
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton cmdStatusServico 
      Caption         =   "Status do Servi�o"
      Height          =   375
      Left            =   8100
      TabIndex        =   13
      Top             =   5040
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdNFeDigitacao 
      Caption         =   "Colocar nota em Digita��o"
      Height          =   375
      Left            =   10620
      TabIndex        =   12
      Top             =   4620
      Width           =   2115
   End
   Begin VB.Timer tmrAtualiza 
      Interval        =   20000
      Left            =   11820
      Top             =   60
   End
   Begin VB.CommandButton cmdInutilizarNumeracao 
      Caption         =   "&Inutilizar Numera��o"
      Height          =   375
      Left            =   2940
      TabIndex        =   11
      Top             =   4620
      Width           =   1695
   End
   Begin VB.CommandButton cmdImprimirDANFE 
      Caption         =   "Imprimir DANFE"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4680
      TabIndex        =   10
      Top             =   4620
      Width           =   1635
   End
   Begin VB.CommandButton cmdPesquisar 
      Caption         =   "&Pesquisar"
      Height          =   315
      Left            =   7620
      TabIndex        =   6
      Top             =   90
      Width           =   1155
   End
   Begin VB.CommandButton cmdCancelarNota 
      Caption         =   "&Cancelar Nota"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1320
      TabIndex        =   9
      Top             =   4620
      Width           =   1575
   End
   Begin VB.CommandButton cmdTransmitir 
      Caption         =   "&Transmitir"
      Enabled         =   0   'False
      Height          =   375
      Left            =   60
      TabIndex        =   8
      Top             =   4620
      Width           =   1155
   End
   Begin MSComctlLib.ListView lvwNFes 
      Height          =   4080
      Left            =   60
      TabIndex        =   7
      Top             =   480
      Width           =   12645
      _ExtentX        =   22304
      _ExtentY        =   7197
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
   Begin MSMask.MaskEdBox mskDataInicial 
      Height          =   285
      Left            =   1020
      TabIndex        =   1
      Top             =   105
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "99/99/9999"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox mskDataFinal 
      Height          =   285
      Left            =   2940
      TabIndex        =   3
      Top             =   105
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "99/99/9999"
      PromptChar      =   " "
   End
   Begin VB.Label lblSituacao 
      AutoSize        =   -1  'True
      Caption         =   "Situa��o"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   4020
      TabIndex        =   4
      Top             =   150
      Width           =   630
   End
   Begin VB.Label lblDataInicial 
      AutoSize        =   -1  'True
      Caption         =   "Data Inicial"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   795
   End
   Begin VB.Label lblDataFinal 
      AutoSize        =   -1  'True
      Caption         =   "Data Final"
      Height          =   195
      Left            =   2100
      TabIndex        =   2
      Top             =   150
      Width           =   720
   End
End
Attribute VB_Name = "frmGerenciarNFes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''Option Explicit
''Dim ItemList As ListItem
''Dim ProcuraItem As ListItem
''Dim rsNFe As New ADODB.Recordset
''Dim rsNFeItens As New ADODB.Recordset
''Dim rsNFeBoletos As New ADODB.Recordset
''Dim rsTotalNFe As New ADODB.Recordset
''Dim rsTributos As New ADODB.Recordset
''Dim rsProdutos As New ADODB.Recordset
''Dim rsEmpresa As New ADODB.Recordset
''Dim rsRamoAtividade As New ADODB.Recordset
''Dim rsTransportador As New ADODB.Recordset
''Dim rsCFOPs As New ADODB.Recordset
''Dim rsClientes As New ADODB.Recordset
''Dim rsTemp As New ADODB.Recordset
''Dim rsTemp2 As New ADODB.Recordset
''Dim rsUFs As New ADODB.Recordset
''Dim rsNaturezasOperacao As New ADODB.Recordset
''Dim rsCFOPReferencias As New ADODB.Recordset
''Dim rsContasBancarias As New ADODB.Recordset
''Dim rsSaldoProdutos As New ADODB.Recordset
''Dim rsUnidadesMedida As New ADODB.Recordset
''Dim rsSituacoesTributarias As New ADODB.Recordset
''Dim rsNFeInutilizadas As New ADODB.Recordset
''Dim strDescricaoTemp As String
''Dim intItensNota As Integer
''Dim rsLogradouros As New ADODB.Recordset
''Dim rsMunicipios As New ADODB.Recordset
''Dim rsUnidades As New ADODB.Recordset
''Dim rsFormasPagamento As New ADODB.Recordset
''Dim rsTransportadores As New ADODB.Recordset
''Dim rsGerarXML As New ADODB.Recordset
''Dim Contador As Integer
''Dim iSituacao As Integer
''Dim sEmpresaNFe As String
''Dim sAcao As String
''Dim bRetorno As Boolean
''Dim IdMensagens As Integer
''Dim ArqNFeRetornos As String
''Dim ArqNFeErros As String
''Dim ArqNFeEnviados As String
''Dim ArqNFeTemp As String
''
''Private Sub Form_Load()
''   I_TituloForm = Me.Caption
''   On Error GoTo Erro
''   Status = 0
''   Centraliza frmGerenciarNFes
''''   rsNFe.Open "Select * from NFe Order By Numero", cnSistema, adOpenForwardOnly, adLockOptimistic, 1
''   rsEmpresa.Open "Select * from Empresa", cnSistema, adOpenForwardOnly, adLockOptimistic, 1
''
''   sEmpresaNFe = LerArquivoINI("NFe", "Empresa", CaminhoINI & "\System.ini")
''   ArqNFeRetornos = "C:\NF-e\" & sEmpresaNFe & "\Retorno\"
''   ArqNFeErros = "C:\NF-e\" & sEmpresaNFe & "\Erros\"
''   ArqNFeEnviados = ""
''   ArqNFeTemp = "C:\NF-e\" & sEmpresaNFe & "\Temp\"
''
''   IdMensagens = 1 ' Inicia o contador de chave das mensagens
''
''   lvwNFes.ColumnHeaders.Add , , "N�mero", 850
''   lvwNFes.ColumnHeaders.Add , , "Emiss�o", 1050
''   lvwNFes.ColumnHeaders.Add , , "Cliente", 6000
''   lvwNFes.ColumnHeaders.Add , , "Valor Total", 1050, lvwColumnRight
''   lvwNFes.ColumnHeaders.Add , , "Situa��o", 1700
''   lvwNFes.ColumnHeaders.Add , , "", 1650
''
''   lvwMensagens.ColumnHeaders.Add , , "Chave", 0
''   lvwMensagens.ColumnHeaders.Add , , "N�mero", 850
''   lvwMensagens.ColumnHeaders.Add , , "Mensagem", 11000
''   lvwMensagens.ColumnHeaders.Add , , "Arquivo", 0
''
''   mskDataInicial.Text = Date
''   mskDataFinal.Text = Date
''   cmbSituacao.ListIndex = 0
''
''   Carrega_View
''
''   lvwMensagens.ListItems.Clear
''
''   MDISistema.StatusBar.Panels(1).Text = "Gerenciamentos de Notas de Sa�da Eletr�nicas"
''
''Exit Sub
''Erro:
''   If Err.Number = -2147467259 Then
''      rsErro = True
''      Beep
''      MsgBox "Erro na Abertura do Arquivo de Dados" & Chr(13) & "Algum usu�rio est� com o Arquivo em modo Exclusivo", vbExclamation, "Erro"
''      Exit Sub
''   Else
''      rsErro = True
''      Beep
''      MsgBox "Verificar: " & Err.Number & Chr(13) & Err.Description, vbExclamation, "Sistema"
''      Exit Sub
''   End If
''End Sub
''
''Private Sub Form_Unload(Cancel As Integer)
''   If Not rsErro Then rsNFe.Close
''   If Not rsErro Then rsEmpresa.Close
''End Sub
''
''Private Sub cmdPesquisar_Click()
''   Carrega_View
''End Sub
''
''Private Sub Carrega_View()
''On Error GoTo Erro
''
''Dim sSituacao As String
''Dim Contador As Integer
''
''   cmdCancelarNota.Enabled = False
''   cmdTransmitir.Enabled = False
''   cmdImprimirDANFE.Enabled = False
''
''   If Not IsDate(mskDataInicial.Text) Or Not IsDate(mskDataFinal.Text) Then
''      Beep
''      MsgBox "Datas inv�lidas", vbExclamation, "Erro"
''      Exit Sub
''   End If
''
''   ' Notas
''   Dim sqlSituacao As String
''   Select Case cmbSituacao.ListIndex
''         Case 0
''            sqlSituacao = ""
''         Case 1
''            sqlSituacao = " AND Situacao = 0"
''         Case 2
''            sqlSituacao = " AND Situacao = 1"
''         Case 3
''            sqlSituacao = " AND Situacao = 2"
''         Case 4
''            sqlSituacao = " AND Situacao = 3"
''   End Select
''
''   If (IsDate(mskDataInicial.Text) And IsDate(mskDataFinal.Text)) And (CDate(mskDataFinal.Text) >= CDate(mskDataInicial.Text)) Then
''      Set rsNFe = cnSistema.Execute("SELECT * FROM NFe WHERE DataEmissao >= cDate('" & Format(mskDataInicial.Text, "dd/mm/yyyy") & "') AND DataEmissao <= cDate('" & Format(mskDataFinal.Text, "dd/mm/yyyy") & "')" & sqlSituacao & " Order By Numero")
''   Else
''      Set rsNFe = cnSistema.Execute("SELECT * FROM NFe WHERE DataEmissao >= cDate('" & Format(Date, "dd/mm/yyyy") & "') AND DataEmissao <= cDate('" & Format(Date, "dd/mm/yyyy") & "')" & sqlSituacao & " Order By Numero")
''   End If
''
''   Contador = 1
''   lvwNFes.ListItems.Clear
''   If Not rsNFe.EOF Then
''      Do While Not rsNFe.EOF
''         Set rsClientes = cnSistema.Execute("SELECT * FROM Clientes WHERE idCliente = " & rsNFe!idCliente)
''         Select Case rsNFe!Situacao
''                Case 0
''                     sSituacao = "Em Digita��o"
''                Case 1
''                     sSituacao = "Processamento"
''                Case 2
''                     sSituacao = "Aprovada"
''                Case 3
''                     sSituacao = "Cancelada"
''         End Select
''
''         If Not rsClientes.EOF Then
''            Set rsTotalNFe = cnSistema.Execute("Select * From TotalNFe Where Numero = " & rsNFe!Numero)
''
'''            Set ItemList = lvwNFes.ListItems.Add(, "R" & CStr(rsNFe!idNFe), StrZero(rsNFe!Numero, 8))
''
''            Set ProcuraItem = lvwNFes.FindItem(StrZero(rsNFe!Numero, 8))
''            If ProcuraItem Is Nothing Then
''               Set ItemList = lvwNFes.ListItems.Add(, "R" & CStr(Contador), StrZero(rsNFe!Numero, 8))
''               ItemList.SubItems(1) = rsNFe!DataEmissao
''               ItemList.SubItems(2) = Trim(rsClientes!Nome)
''               If Not rsTotalNFe.EOF Then
''                  ItemList.SubItems(3) = Format(rsTotalNFe!Total + IIf(Not IsNull(rsTotalNFe!TotalFrete), rsTotalNFe!TotalFrete, 0), "##,##0.00")
''               Else
''                  ItemList.SubItems(3) = Format(0, "##,##0.00")
''               End If
''               ItemList.SubItems(4) = sSituacao
''               If rsNFe!Situacao = 1 Then
''                  ItemList.SubItems(5) = "Aguarde..."
''               End If
''            End If
''         End If
''
''         Contador = Contador + 1
''         rsNFe.MoveNext
''      Loop
''
''      rsNFe.MoveFirst
''   End If
''
''   ' Notas
''   Set rsNFeInutilizadas = cnSistema.Execute("SELECT * FROM NFeInutilizadas WHERE Data >= cDate('" & Format(mskDataInicial.Text, "dd/mm/yyyy") & "') AND Data <= cDate('" & Format(mskDataFinal.Text, "dd/mm/yyyy") & "') Order By Numero")
''
''   If Not rsNFeInutilizadas.EOF Then
''      Do While Not rsNFeInutilizadas.EOF
''
'''         Set ItemList = lvwNFes.ListItems.Add(, "I" & CStr(rsNFeInutilizadas!Numero), StrZero(rsNFeInutilizadas!Numero, 8))
''
''         Set ProcuraItem = lvwNFes.FindItem(StrZero(rsNFeInutilizadas!Numero, 8))
''         If ProcuraItem Is Nothing Then
''            Set ItemList = lvwNFes.ListItems.Add(, "I" & CStr(Contador), StrZero(rsNFeInutilizadas!Numero, 8))
''            ItemList.SubItems(1) = rsNFeInutilizadas!Data
''            ItemList.SubItems(2) = "Nota Inutilizada"
''            ItemList.SubItems(3) = Format(0, "##,##0.00")
''            ItemList.SubItems(4) = "Inutilizada"
''            If IsNull(rsNFeInutilizadas!Protocolo) Or rsNFeInutilizadas!Protocolo = "" Then
''               ItemList.SubItems(5) = "Aguarde..."
''            Else
''               ItemList.SubItems(5) = ""
''            End If
''         End If
''
''         Contador = Contador + 1
''         rsNFeInutilizadas.MoveNext
''      Loop
''   End If
''
''   Exit Sub
''Erro:
''   MsgBox "Erro " & Err & ". " & Err.Description & " - " & TypeName(Me) & ".Carrega_View"
''End Sub
''
''Private Sub cmdTransmitir_Click()
''
'' ' Gerar XML
''   Open "C:\NF-e\Notas\Notas.TXT" For Output As #1
''   sAcao = 1  ' Transmitir
''   Notas
''   Close #1
''
''   sEmpresaNFe = LerArquivoINI("NFe", "Empresa", CaminhoINI & "\System.ini")
''   Set rsGerarXML = cnSistema.Execute("Select * From NFe WHERE Numero=" & Val(lvwNFes.ListItems(lvwNFes.SelectedItem.Index)))
''   If Not rsGerarXML.EOF Then
''      FileCopy "C:\NF-e\Notas\Notas.TXT", "C:\NF-e\" & sEmpresaNFe & "\Envio\" & StrZero(rsGerarXML!Numero, 6) & "_" & RemoveCaracteres(rsEmpresa!CNPJ_CPF) & "_001_" & Format(rsGerarXML!DataEmissao, "dd_mm_yyyy") & "-nfe.txt"
''      Kill "C:\NF-e\Notas\Notas.TXT"
''   End If
''
''   sAcao = 0
''   cmdTransmitir.Enabled = False
''   cmdValidar.Enabled = False
''   Carrega_View
''End Sub
''
''Private Function Notas()
''Dim sPercMargAdICMSST As String
''Dim sUF As String
''Dim sChaveAcesso As String
''Dim sNaturezaOperacao As String
''Dim sFormaPagamento As String
''Dim sModelo As String
''Dim sSerie As String
''Dim sNumero As String
''Dim sDataEmissao As String
''Dim sDataSaida As String
''Dim sHoraSaida As String
''Dim sTipoNF As String
''Dim sidDest As String
''Dim sCodigoMunicipio As String
''Dim sFormatoDANFE As String
''Dim sTipoEmissao As String
''Dim sDVChaveAcesso As String
''Dim sidAmbiente As String
''Dim sFinalidade As String
''Dim sConsumidorFinal As String
''Dim sIndicadorPresenca As String
''Dim sIndicadorIEDest As String
''Dim sProcessoEmissao As String
''Dim sVersaoAplicativo As String
''Dim sdhCont As String
''Dim sxJust As String
''Dim dValorTributos As Double
''
''   Set rsEmpresa = cnSistema.Execute("Select * From Empresa")
''
''   Dim Total_NFes As Integer
''   Set rsTemp = cnSistema.Execute("Select Count(*) as Qte from NFe WHERE Numero = " & Val(lvwNFes.ListItems(lvwNFes.SelectedItem.Index)))
''   Total_NFes = rsTemp!Qte
''
''   Set rsNFe = cnSistema.Execute("Select * From NFe WHERE Numero=" & Val(lvwNFes.ListItems(lvwNFes.SelectedItem.Index)))
''
''   If Not rsNFe.EOF Then Print #1, "NOTAFISCAL|" & Total_NFes
''   If Not rsNFe.EOF Then
''    ' Cabecalho
''      ''''Print #1, "A|2.00|NFe"
''      ''''Print #1, "A|3.10|NFe"
''      Print #1, "A|4.00|NFe"
''
''    ' Identificadores
''      Set rsNaturezasOperacao = cnSistema.Execute("Select * From NaturezasOperacao WHERE idNaturezaOperacao=" & rsNFe!idNaturezaOperacao)
''      Set rsFormasPagamento = cnSistema.Execute("Select * From FormasPagamento WHERE idFormaPagamento=" & rsNFe!idFormaPagamento)
''      Set rsCFOPs = cnSistema.Execute("Select * From CFOPs WHERE idCFOP=" & rsNFe!idCFOP)
''
''      Set rsTemp = cnSistema.Execute("Select * From UFs WHERE idUF=" & rsEmpresa!idUF)
''      sUF = rsTemp!Codigo ' Qualquer Municipio
''      sChaveAcesso = ""
''      sNaturezaOperacao = IIf(Not rsNaturezasOperacao.EOF, rsNaturezasOperacao!Descricao, "VENDA")
''      If Not rsFormasPagamento.EOF Then
''         If rsFormasPagamento!TipoPagamento <= 1 Then
''            sFormaPagamento = "0"
''         Else
''            sFormaPagamento = "1"
''         End If
''      Else
''         sFormaPagamento = "0"
''      End If
''      sModelo = "55"
''      sSerie = "1"
''      sNumero = rsNFe!Numero
''
''      If LerArquivoINI("NFe", "HorarioVerao", App.Path & "\System.ini") Then
''         sDataEmissao = Format(rsNFe!DataEmissao, "YYYY-MM-DD") & "T" & Format(Time, "HH:MM:SS") & "-02:00"
''      Else
''         sDataEmissao = Format(rsNFe!DataEmissao, "YYYY-MM-DD") & "T" & Format(Time, "HH:MM:SS") & "-03:00"
''      End If
''
''      If LerArquivoINI("NFe", "HorarioVerao", App.Path & "\System.ini") Then
''         sDataSaida = Format(rsNFe!DataVencimento, "YYYY-MM-DD") & "T" & Format(Time, "HH:MM:SS") & "-02:00"
''      Else
''         sDataSaida = Format(rsNFe!DataVencimento, "YYYY-MM-DD") & "T" & Format(Time, "HH:MM:SS") & "-03:00"
''      End If
''
''      sHoraSaida = ""
''      sTipoNF = IIf(rsCFOPs!Tipo = 0, 0, 1)                       ' 0 - Entrada ou 1 - Saida
''
''      Set rsClientes = cnSistema.Execute("Select * From Clientes WHERE idCliente=" & rsNFe!idCliente)
''      If Not rsClientes!Interestadual Then
''         sidDest = 1                        ' 1. Operacao Interna
''      Else
''         sidDest = 2                        ' 2. Operacao Interestadual
''      End If
''
''      Set rsTemp = cnSistema.Execute("Select * From Municipios WHERE idMunicipio=" & rsEmpresa!idMunicipio)
''      sCodigoMunicipio = RemoveCaracteres(rsTemp!Codigo)
''      sFormatoDANFE = "1"
''      sTipoEmissao = "1"
''      sDVChaveAcesso = ""
''      sidAmbiente = LerArquivoINI("NFe", "Ambiente", CaminhoINI & "\System.ini") ' 1 - Produ��o ou 2 - Homologa��o
''      If rsNaturezasOperacao!Descricao <> "COMPLEMENTO DE VALOR" Then
''         'sFinalidade = "1"                                                       ' 1 - NFe Normal / 2 - Complementar / 3 - Ajuste
'''''         sFinalidade = IIf(rsNaturezasOperacao!Tipo = 0, "1", "4")                   ' 1. Saida, 2. Devolu��o
''         If rsNaturezasOperacao!Tipo = 0 Then
''            sFinalidade = "1"       ' Saida
''         ElseIf rsNaturezasOperacao!Tipo = 2 Then
''            sFinalidade = "4"       ' Devolu��o
''         End If
''      Else
''         sFinalidade = "2"                                                       ' 1 - NFe Normal / 2 - Complementar / 3 - Ajuste
''      End If
''      sProcessoEmissao = "3"                                                     ' Utilizando Software do Fisco
''
''      Set rsClientes = cnSistema.Execute("Select * From Clientes WHERE idCliente=" & rsNFe!idCliente)
''      If rsClientes!ConsumidorFinal Then
''         sConsumidorFinal = "1"                                                     ' Consumidor final
''      Else
''         sConsumidorFinal = "0"                                                     ' Nao consumidor final
''      End If
''
''      sIndicadorPresenca = "3"                                                   ' N�o presencial
''      sVersaoAplicativo = "3.10.43"
''      sdhCont = ""
''      sxJust = ""
''
''      Print #1, "B|" & _
''                sUF & "|" & _
''                sChaveAcesso & "|" & _
''                sNaturezaOperacao & "|" & _
''                sModelo & "|" & _
''                sSerie & "|" & _
''                sNumero & "|" & _
''                sDataEmissao & "|" & _
''                sDataSaida & "|" & _
''                sTipoNF & "|" & _
''                sidDest & "|" & _
''                sCodigoMunicipio & "|" & _
''                sFormatoDANFE & "|" & _
''                sTipoEmissao & "|" & _
''                sDVChaveAcesso & "|" & _
''                sidAmbiente & "|" & _
''                sFinalidade & "|" & _
''                sConsumidorFinal & "|" & _
''                sIndicadorPresenca & "|" & _
''                sProcessoEmissao & "|" & _
''                sVersaoAplicativo & "|" & _
''                sdhCont & "|" & _
''                sxJust
''
'''''      Print #1, "B|" & _
'''''                sUF & "|" & _
'''''                sChaveAcesso & "|" & _
'''''                sNaturezaOperacao & "|" & _
'''''                sFormaPagamento & "|" & _
'''''                sModelo & "|" & _
'''''                sSerie & "|" & _
'''''                sNumero & "|" & _
'''''                sDataEmissao & "|" & _
'''''                sDataSaida & "|" & _
'''''                sTipoNF & "|" & _
'''''                sidDest & "|" & _
'''''                sCodigoMunicipio & "|" & _
'''''                sFormatoDANFE & "|" & _
'''''                sTipoEmissao & "|" & _
'''''                sDVChaveAcesso & "|" & _
'''''                sidAmbiente & "|" & _
'''''                sFinalidade & "|" & _
'''''                sConsumidorFinal & "|" & _
'''''                sIndicadorPresenca & "|" & _
'''''                sProcessoEmissao & "|" & _
'''''                sVersaoAplicativo & "|" & _
'''''                sdhCont & "|" & _
'''''                sxJust
''
''      If rsNaturezasOperacao!Descricao = "COMPLEMENTO DE VALOR" Then
''         Print #1, "B13|" & rsNFe!ChaveAcessoNFeComplementar & "|"
''      End If
''
''      If Not IsNull(rsNFe!ChaveAcessoDevolucao) And Len(rsNFe!ChaveAcessoDevolucao) > 0 Then
''         Print #1, "BA|"
''         Print #1, "BA02|" & _
''                   rsNFe!ChaveAcessoDevolucao & "|"
''      End If
''
''    ' Emitente
''    '=======================================================================================================
''      Dim sERazaoSocial As String
''      Dim sEFantasia As String
''      Dim sEIE As String
''      Dim sEIEST As String
''      Dim sEIM As String
''      Dim sECNAE As String
''      Dim sCRT As String
''
''      sERazaoSocial = rsEmpresa!Nome
''      sEFantasia = ""
''      sEIE = IIf(rsEmpresa!IE_CI <> "ISENTO", Trim(RemoveCaracteres(rsEmpresa!IE_CI)), "ISENTO")
''      sEIEST = ""
''      sEIM = ""
''      sECNAE = LerArquivoINI("NFe", "CNAE", CaminhoINI & "\System.ini") ' CNAE
''      sCRT = LerArquivoINI("NFe", "Regime", CaminhoINI & "\System.ini") ' 1 - Simples / 3 - Normal
''
''
''      Print #1, "C|" & _
''                sERazaoSocial & "|" & _
''                sEFantasia & "|" & _
''                sEIE & "|" & _
''                sEIEST & "|" & _
''                sEIM & "|" & _
''                sECNAE & "|" & _
''                sCRT
''
''      Dim sECNPJ As String
''      sECNPJ = RemoveCaracteres(rsEmpresa!CNPJ_CPF)
''
''      Print #1, "C02|" & _
''                sECNPJ
''
''      Dim sELogradouro As String
''      Dim sENumero As String
''      Dim sEComplemento As String
''      Dim sEBairro As String
''      Dim sECodigoMunicipio As String
''      Dim sEMunicipio As String
''      Dim sEUF As String
''      Dim sECEP As String
''      Dim sECodigoPais As String
''      Dim sEPais As String
''      Dim sETelefone As String
''
''      sELogradouro = rsEmpresa!Endereco
''      sENumero = "."
''      sEComplemento = ""
''      sEBairro = rsEmpresa!Bairro
''      Set rsTemp = cnSistema.Execute("Select * From Municipios WHERE idMunicipio=" & rsEmpresa!idMunicipio)
''      sECodigoMunicipio = RemoveCaracteres(rsTemp!Codigo)
''      sEMunicipio = rsTemp!Nome
''      sEUF = rsTemp!UF
''      sECEP = RemoveCaracteres(rsEmpresa!CEP)
''      sECodigoPais = "1058"
''      sEPais = "BRASIL"
''      sETelefone = RemoveCaracteres(rsEmpresa!Telefone1)
''
''      Print #1, "C05|" & _
''                sELogradouro & "|" & _
''                sENumero & "|" & _
''                sEComplemento & "|" & _
''                sEBairro & "|" & _
''                sECodigoMunicipio & "|" & _
''                sEMunicipio & "|" & _
''                sEUF & "|" & _
''                sECEP & "|" & _
''                sECodigoPais & "|" & _
''                sEPais & "|" & _
''                sETelefone
''
''    ' Destinatario
''    '=======================================================================================================
''      Set rsClientes = cnSistema.Execute("Select * From Clientes WHERE idCliente=" & rsNFe!idCliente)
''
''      Dim sDRazaoSocial As String
''      Dim sDIE As String
''      Dim sDISUF As String
''      Dim seMail As String
''
''      sDRazaoSocial = RemoveAcentos(rsClientes!Nome)
''      sDIE = IIf(IIf(Not IsNull(rsClientes!IE_CI), rsClientes!IE_CI, "ISENTO") <> "ISENTO", Trim(RemoveCaracteres(IIf(Not IsNull(rsClientes!IE_CI), rsClientes!IE_CI, ""))), "ISENTO")
''      sDISUF = ""
''      seMail = "" ' rsClientes!Email
''
''      Dim sDCNPJ As String
''      sDCNPJ = RemoveCaracteres(rsClientes!CNPJ_CPF)
''
''      Select Case rsClientes!TipoContribuinte
''             Case 0
''                  sIndicadorIEDest = 1 'Contribuinte do ICMS
''             Case 1
''                  sIndicadorIEDest = 2 'Contribuinte isento de Incri��o Estadual
''             Case 2
''                  sIndicadorIEDest = 9 'N�o contribuinte de Incri��o Estadual
''      End Select
''
''      Print #1, "E|" & _
''                sDRazaoSocial & "|" & _
''                sIndicadorIEDest & "|" & _
''                sDIE & "|" & _
''                sDISUF
''
''      If Len(Trim(sDCNPJ)) > 11 Then
''         Print #1, "E02|" & _
''                   sDCNPJ
''      Else
''         Print #1, "E03|" & _
''                   sDCNPJ
''      End If
''
''      ' Endereco
''      Set rsLogradouros = cnSistema.Execute("Select * From Logradouros WHERE idLogradouro=" & rsClientes!idLogradouro)
''      Set rsMunicipios = cnSistema.Execute("Select * From Municipios WHERE idMunicipio=" & rsClientes!idMunicipio)
''
''      Dim sDLogradouro As String
''      Dim sDNumero As String
''      Dim sDComplemento As String
''      Dim sDBairro As String
''      Dim sDCodigoMunicipio As String
''      Dim sDMunicipio As String
''      Dim sDUF As String
''      Dim sDCEP As String
''      Dim sDCodigoPais As String
''      Dim sDPais As String
''      Dim sDTelefone As String
''
''      sDLogradouro = IIf(Not rsLogradouros.EOF, IIf(rsLogradouros!Abreviacao <> ".", rsLogradouros!Abreviacao & " ", ""), "") & Trim(RemoveAcentos(rsClientes!Endereco))
''      sDNumero = Trim(rsClientes!Numero)
''      sDComplemento = ""
''      sDBairro = RemoveAcentos(IIf(Not IsNull(rsClientes!Bairro), rsClientes!Bairro, "."))
''      sDCodigoMunicipio = Trim(RemoveAcentos(RemoveCaracteres(rsMunicipios!Codigo)))
''      sDMunicipio = RemoveAcentos(rsMunicipios!Nome)
''      sDUF = rsClientes!UF
''      sDCEP = RemoveCaracteres(IIf(Not IsNull(rsClientes!CEP), rsClientes!CEP, ""))
''      sDCodigoPais = "1058"
''      sDPais = "BRASIL"
''      sDTelefone = StrZero(Val(IIf(Not IsNull(rsClientes!PrefixoFone1), rsClientes!PrefixoFone1, "0")), 2) & Trim(FormataTXT(RemoveCaracteres(IIf(Not IsNull(rsClientes!Telefone1), rsClientes!Telefone1, "0")), 1, 10))
''      If Len(Trim(sDTelefone)) <> 10 Then
''         sDTelefone = ""
''      End If
''
''      Print #1, "E05|" & _
''                sDLogradouro & "|" & _
''                sDNumero & "|" & _
''                sDComplemento & "|" & _
''                sDBairro & "|" & _
''                sDCodigoMunicipio & "|" & _
''                sDMunicipio & "|" & _
''                sDUF & "|" & _
''                sDCEP & "|" & _
''                sDCodigoPais & "|" & _
''                sDPais & "|" & _
''                sDTelefone
''    ' Itens
''      Dim Contador As Integer
''      Contador = 1
''
''      Dim dValorTotalBC As Double
''      Dim dValorTotalICMS As Double
''      Dim dValorTotalBCST As Double
''      Dim dValorTotalICMSST As Double
''      Dim dValorTotalProdutos As Double
''      Dim dValorTotalFrete As Double
''      Dim dValorTotalSeguro As Double
''      Dim dValorTotalDesconto As Double
''      Dim dValorTotalII As Double
''      Dim dValorTotalIPI As Double
''      Dim dValorTotalPIS As Double
''      Dim dValorTotalCofins As Double
''      Dim dValorTotalOutro As Double
''      Dim dValorTotalNFe As Double
''      Dim sICMSAproveitamento As String
''
''      dValorTotalBC = 0
''      dValorTotalICMS = 0
''      dValorTotalBCST = 0
''      dValorTotalICMSST = 0
''      dValorTotalProdutos = 0
''      dValorTotalFrete = 0
''      dValorTotalSeguro = 0
''      dValorTotalDesconto = 0
''      dValorTotalII = 0
''      dValorTotalIPI = 0
''      dValorTotalPIS = 0
''      dValorTotalCofins = 0
''      dValorTotalOutro = 0
''      dValorTotalNFe = 0
''
''      Set rsNFeItens = cnSistema.Execute("SELECT * FROM NFeItens WHERE NFeItens.idNFe = " & rsNFe!idNFe)
''      Do While Not rsNFeItens.EOF
''         Set rsProdutos = cnSistema.Execute("SELECT * FROM Produtos WHERE idProduto = " & rsNFeItens!idProduto)
''         Set rsUnidades = cnSistema.Execute("Select * From UnidadesMedida WHERE idUnidadeMedida=" & rsProdutos!idUnidade)
''         Set rsSituacoesTributarias = cnSistema.Execute("Select * from SituacoesTributarias WHERE idSituacaoTributaria=" & rsNFeItens!idSituacaoTributaria)
''         Set rsTributos = cnSistema.Execute("SELECT * FROM TabelaIBPT WHERE CodigoNCM LIKE '%" & rsProdutos!CodigoNCM & "%'")
''
''         If rsProdutos!ICMSReaproveitamento > 0 Then
''            sICMSAproveitamento = "EMPRESA OPTANTE PELO SIMPLES NACIONAL - ALIQUOTA APLICAVEL DE CALCULO DO CREDITO " & Format(rsProdutos!ICMSReaproveitamento, "###0.00") & "% - R$ " & Format((((rsNFeItens!Quantidade * rsNFeItens!ValorUnitario) * rsProdutos!ICMSReaproveitamento) / 100), "##,###,##0.00")
''         Else
''            sICMSAproveitamento = ""
''         End If
''
''         Print #1, "H|" & _
''                   Contador & "|" & _
''                   sICMSAproveitamento
''
''         Dim sCodigoProduto As String
''         Dim sCodigoBarras As String
''         Dim sDescricaoProduto As String
''         Dim sCodigoNCM As String
''
''         Dim sNVE As String
''         Dim sCEST As String
''         Dim sindEscala As String
''         Dim sCNPJFab As String
''         Dim scBenef As String
''
''         Dim sEXTIPI As String
''         Dim sCFOP As String
''         Dim sUnidComercial As String
''         Dim sQuantidadeComercial As String
''         Dim sVlUnitarioComercial As String
''         Dim sVlTotalBruto As String
''         Dim sCodigoBarrasTrib As String
''         Dim sUnidTrib As String
''         Dim sQuantidadeTrib As String
''         Dim sVlUnitarioTrib As String
''         Dim sVlFrete As String
''         Dim sVlSeguro As String
''         Dim sVlDesconto As String
''         Dim sVlOutros As String
''         Dim sIndTot As String
''         Dim sxPed As String
''         Dim snItemPed As String
''
''         sCodigoProduto = rsProdutos!Codigo
''         sCodigoBarras = "SEM GTIN"
''         sDescricaoProduto = IIf(rsNFeItens!DiscriminacaoProduto = "", RemoveAcentos(rsProdutos!Descricao), RemoveAcentos(rsNFeItens!DiscriminacaoProduto))
''         If Not IsNull(rsNFeItens!DescricaoComplementar) Then
''            If Trim(rsNFeItens!DescricaoComplementar) <> "" Then
''               sDescricaoProduto = sDescricaoProduto & " " & RemoveAcentos(rsNFeItens!DescricaoComplementar)
''             End If
''         End If
''         sCodigoNCM = IIf(rsProdutos!CodigoNCM = "" Or IsNull(rsProdutos!CodigoNCM), "", rsProdutos!CodigoNCM)
''         sEXTIPI = ""
''         sCFOP = RemoveCaracteres(rsNFeItens!CFOP)
''         sUnidComercial = IIf(Not rsUnidades.EOF, rsUnidades!Sigla, "UN")
''         If rsNaturezasOperacao!Descricao <> "COMPLEMENTO DE VALOR" Then
''            sQuantidadeComercial = Substitui(Format(rsNFeItens!Quantidade, "#######0.0000"), ",", ".")
''            sVlUnitarioComercial = Substitui(Format(rsNFeItens!ValorUnitario, "#######0.0000"), ",", ".")
''            sVlTotalBruto = Substitui(Format(rsNFeItens!Quantidade * rsNFeItens!ValorUnitario, "#######0.00"), ",", ".")
''            sCodigoBarrasTrib = "SEM GTIN"
''            sUnidTrib = IIf(Not rsUnidades.EOF, rsUnidades!Sigla, "UN")
''            sQuantidadeTrib = Substitui(Format(rsNFeItens!Quantidade, "#######0.0000"), ",", ".")
''            sVlUnitarioTrib = Substitui(Format(rsNFeItens!ValorUnitario, "#######0.0000"), ",", ".")
''            sVlFrete = IIf(rsNFeItens!valorfrete > 0, Substitui(Format(rsNFeItens!valorfrete, "#######0.00"), ",", "."), "")
''            sVlSeguro = ""
''            sVlDesconto = IIf(rsNFeItens!Desconto > 0, Substitui(Format((((rsNFeItens!Quantidade * rsNFeItens!ValorUnitario) * rsNFeItens!Desconto) / 100), "#######0.00"), ",", "."), "")
''         Else
''            sQuantidadeComercial = Substitui(Format(0, "#######0.0000"), ",", ".")
''            sVlUnitarioComercial = Substitui(Format(0, "#######0.0000"), ",", ".")
''            sVlTotalBruto = Substitui(Format(0, "#######0.00"), ",", ".")
''            sCodigoBarrasTrib = ""
''            sUnidTrib = IIf(Not rsUnidades.EOF, rsUnidades!Sigla, "UN")
''            sQuantidadeTrib = Substitui(Format(0, "#######0.0000"), ",", ".")
''            sVlUnitarioTrib = Substitui(Format(0, "#######0.0000"), ",", ".")
''            sVlFrete = IIf(rsNFeItens!valorfrete > 0, Substitui(Format(rsNFeItens!valorfrete, "#######0.00"), ",", "."), "")
''            sVlSeguro = ""
''            sVlDesconto = IIf(rsNFeItens!Desconto > 0, Substitui(Format((((rsNFeItens!Quantidade * rsNFeItens!ValorUnitario) * rsNFeItens!Desconto) / 100), "#######0.00"), ",", "."), "")
''         End If
''         sVlOutros = ""
''         sIndTot = "1"
''         sxPed = ""
''         snItemPed = ""
''
''         If Not rsTributos.EOF Then
''            dValorTributos = dValorTributos + (((rsNFeItens!Quantidade * rsNFeItens!ValorUnitario) * rsTributos!AliquotaNacional) / 100)
''         End If
''
''''         Print #1, "I|" & _
''''                   sCodigoProduto & "|" & _
''''                   sCodigoBarras & "|" & _
''''                   sDescricaoProduto & "|" & _
''''                   sCodigoNCM & "|" & _
''''                   sEXTIPI & "|" & _
''''                   sCFOP & "|" & _
''''                   sUnidComercial & "|" & _
''''                   sQuantidadeComercial & "|" & _
''''                   sVlUnitarioComercial & "|" & _
''''                   sVlTotalBruto & "|" & _
''''                   sCodigoBarrasTrib & "|" & _
''''                   sUnidTrib & "|" & _
''''                   sQuantidadeTrib & "|" & _
''''                   sVlUnitarioTrib & "|" & _
''''                   sVlFrete & "|" & _
''''                   sVlSeguro & "|" & _
''''                   sVlDesconto & "|" & _
''''                   sVlOutros & "|" & _
''''                   sIndTot & "|" & _
''''                   sxPed & "|" & _
''''                   snItemPed
''
'''NVE|CEST|indEscala|CNPJFab|cBenef|EXTIPI|
''
'''                   "|" & "|" & "S|" & "|" & "|" & _
''
''' I|cProd|cEAN|XProd|NCM|NVE|CEST|indEscala|CNPJFab|cBenef|EXTIPI|CFOP|UCom|QCom|VUnCom|VProd|CEANTrib|UTrib|QTrib|VUnTrib|VFrete|VSeg|VDesc|vOutro|indTot|xPed|nItemPed|nFCI|
''
''         ' Novos campos 4.00
''         sNVE = ""
''         sCEST = ""
''         sindEscala = ""
''         sCNPJFab = ""
''         scBenef = ""
''         '
''
''         Print #1, "I|" & _
''                   sCodigoProduto & "|" & _
''                   sCodigoBarras & "|" & _
''                   sDescricaoProduto & "|" & _
''                   sCodigoNCM & "|" & _
''                   sNVE & "|" & sCEST & "|" & sindEscala & "|" & sCNPJFab & "|" & scBenef & "|" & _
''                   sEXTIPI & "|" & _
''                   sCFOP & "|" & _
''                   sUnidComercial & "|" & _
''                   sQuantidadeComercial & "|" & _
''                   sVlUnitarioComercial & "|" & _
''                   sVlTotalBruto & "|" & _
''                   sCodigoBarrasTrib & "|" & _
''                   sUnidTrib & "|" & _
''                   sQuantidadeTrib & "|" & _
''                   sVlUnitarioTrib & "|" & _
''                   sVlFrete & "|" & _
''                   sVlSeguro & "|" & _
''                   sVlDesconto & "|" & _
''                   sVlOutros & "|" & _
''                   sIndTot & "|" & _
''                   sxPed & "|" & _
''                   snItemPed
''
''       ' Tributos Incidentes
''         Print #1, "M"
''         Print #1, "N"
''
''         Dim sCST As String
''         sCST = Mid(rsSituacoesTributarias!Codigo, 1, 3)
''
''         Dim sOrigem As String
''         Dim sModalidadeBC As String
''         Dim sPercRedBC As String
''         Dim sValorBC As String
''         Dim sICMS As String
''         Dim sValorICMS As String
''         Dim dCalculoBC As Double
''
''         Dim modBCST As String
''         Dim pMVAST As String
''         Dim pRedBCST As String
''         Dim vBCST As String
''         Dim pICMSST As String
''         Dim vICMSST As String
''
''         Dim Nx As String
''         Select Case sCST
''                Case "000" ' Tributada Integralmente
''                     If rsNFeItens!ICMS > 0 Then
''                        dCalculoBC = (rsNFeItens!Quantidade * rsNFeItens!ValorUnitario) - (((rsNFeItens!Quantidade * rsNFeItens!ValorUnitario) * rsNFeItens!Desconto) / 100)
''
''                        sCST = "00"
''                        sOrigem = "0"
''                        sModalidadeBC = "3"
''                        sValorBC = Substitui(Format(dCalculoBC, "#######0.00"), ",", ".")
''                        sICMS = Substitui(Format(rsNFeItens!ICMS, "###0.00"), ",", ".")
''                        sValorICMS = Substitui(Format(((dCalculoBC * rsNFeItens!ICMS) / 100), "#######0.00"), ",", ".")
''                     Else
''                        dCalculoBC = 0
''
''                        sCST = "00"
''                        sOrigem = "0"
''                        sModalidadeBC = "3"
''                        sValorBC = Substitui(Format(0, "#######0.00"), ",", ".")
''                        sICMS = Substitui(Format(0, "###0.00"), ",", ".")
''                        sValorICMS = Substitui(Format(0, "#######0.00"), ",", ".")
''                     End If
''
''                     Print #1, "N02|" & _
''                               sOrigem & "|" & _
''                               sCST & "|" & _
''                               sModalidadeBC & "|" & _
''                               sValorBC & "|" & _
''                               sICMS & "|" & _
''                               sValorICMS & "|"
''
''                Case "010"
''                     Nx = "03"
''                Case "020"
''                     dCalculoBC = Round(((((rsNFeItens!Quantidade * rsNFeItens!ValorUnitario) - (((rsNFeItens!Quantidade * rsNFeItens!ValorUnitario) * rsNFeItens!Desconto) / 100)) * rsNFeItens!BaseReduzida) / 100), 2)
''
''                     sCST = "20"
''                     sOrigem = "0"
''                     sModalidadeBC = "3"
''                     sPercRedBC = Substitui(Format(rsNFeItens!BaseReduzida, "#########0.00"), ",", ".")
''                     sValorBC = Substitui(Format(dCalculoBC, "#######0.00"), ",", ".")
''                     sICMS = Substitui(Format(rsNFeItens!ICMS, "###0.00"), ",", ".")
''                     sValorICMS = Substitui(Format(((dCalculoBC * rsNFeItens!ICMS) / 100), "#######0.00"), ",", ".")
''
''                     Print #1, "N04|" & _
''                               sOrigem & "|" & _
''                               sCST & "|" & _
''                               sModalidadeBC & "|" & _
''                               sPercRedBC & "|" & _
''                               sValorBC & "|" & _
''                               sICMS & "|" & _
''                               sValorICMS
''
''                Case "030"
''                     Nx = "05"
''                Case "040"
''                     Nx = "06"
''                Case "051"
''                     Nx = "07"
''                Case "060"
''                     Nx = "08"
''                Case "070"
''                     Nx = "09"
''                Case "090"
''                     Nx = "10"
''
''                Case "101" ' Tributada com permiss�o de cr�dito
''                     Dim spCredS As String
''                     Dim svCredICMSSN As String
''
''                     sOrigem = "0"
''                     spCredS = Substitui(Format(0, "#######0.00"), ",", ".")
''                     svCredICMSSN = Substitui(Format(0, "###0.00"), ",", ".")
''
''''                     If rsNFeItens!ICMS > 0 Then
''''                        dCalculoBC = (rsNFeItens!quantidade * rsNFeItens!ValorUnitario) - (((rsNFeItens!quantidade * rsNFeItens!ValorUnitario) * rsNFeItens!Desconto) / 100)
''''
''''                        sCST = "00"
''''                        sOrigem = "0"
''''                        sModalidadeBC = "3"
''''                        sValorBC = Substitui(Format(dCalculoBC, "#######0.00"), ",", ".")
''''                        sICMS = Substitui(Format(rsNFeItens!ICMS, "###0.00"), ",", ".")
''''                        sValorICMS = Substitui(Format(((dCalculoBC * rsNFeItens!ICMS) / 100), "#######0.00"), ",", ".")
''''                     Else
''''                        dCalculoBC = 0
''''
''''                        sCST = "00"
''''                        sOrigem = "0"
''''                        sModalidadeBC = "3"
''''                        sValorBC = Substitui(Format(0, "#######0.00"), ",", ".")
''''                        sICMS = Substitui(Format(0, "###0.00"), ",", ".")
''''                        sValorICMS = Substitui(Format(0, "#######0.00"), ",", ".")
''''                     End If
''
''                     Print #1, "N10c|" & _
''                               sOrigem & "|" & _
''                               sCST & "|" & _
''                               spCredS & "|" & _
''                               svCredICMSSN
''
''                Case "102" ' Tributada pelo Simples Nacional sem permiss�o de cr�dito.
''                     sOrigem = "0"
''
''                     Print #1, "N10d|" & _
''                               sOrigem & "|" & _
''                               sCST
''
''                Case "103" ' Isen��o do ICMS no Simples Nacional para faixa de receita bruta.
''                     sOrigem = "0"
''
''                     Print #1, "N10d|" & _
''                               sOrigem & "|" & _
''                               sCST
''
''                Case "300" ' Imune
''                     sOrigem = "0"
''
''                     Print #1, "N10d|" & _
''                               sOrigem & "|" & _
''                               sCST
''
''                Case "400" ' N�o tributada
''                     sOrigem = "0"
''
''                     Print #1, "N10d|" & _
''                               sOrigem & "|" & _
''                               sCST
''
''                Case "202" ' Tributada sem permiss�o de cr�dito com Substitui��o Tributaria
''
''                     sOrigem = "0"
''                     modBCST = "0"
''                     pMVAST = Substitui(Format(0, "##0.00"), ",", ".")
''                     pRedBCST = Substitui(Format(0, "##0.00"), ",", ".")
''                     vBCST = Substitui(Format(0, "#######0.00"), ",", ".")
''                     pICMSST = Substitui(Format(0, "##0.00"), ",", ".")
''                     vICMSST = Substitui(Format(0, "#######0.00"), ",", ".")
''
''                     Print #1, "N10f|" & _
''                               sOrigem & "|" & _
''                               sCST & "|" & _
''                               modBCST & "|" & _
''                               pMVAST & "|" & _
''                               pRedBCST & "|" & _
''                               vBCST & "|" & _
''                               pICMSST & "|" & _
''                               vICMSST
''
''                Case "500" ' Substitui��o Tributaria
''                     Dim vBCSTRet As String
''                     Dim vICMSSTRet As String
''                     Dim vpST As String
''                     Dim vvBCFCPSTRet As String
''                     Dim vpFCPSTRet As String
''                     Dim vvFCPSTRet As String
''
''                     sOrigem = "0"
''                     modBCST = "0"
''                     vpST = "0.00"
''                     vBCSTRet = Substitui(Format(0, "##0.00"), ",", ".")
''                     vICMSSTRet = Substitui(Format(0, "##0.00"), ",", ".")
''                     vvBCFCPSTRet = "0.00"
''                     vpFCPSTRet = "0.00"
''                     vvFCPSTRet = "0.00"
''
''                     Print #1, "N10g|" & _
''                               sOrigem & "|" & _
''                               sCST & "|" & _
''                               vBCSTRet & "|" & _
''                               vpST & "|" & _
''                               vICMSSTRet & "|" & _
''                               vvBCFCPSTRet & "|" & _
''                               vpFCPSTRet & "|" & _
''                               vvFCPSTRet
''
''                Case "900" ' Outros
''                     Dim modBC As String
''                     Dim vBC As String
''                     Dim pRedBC As String
''                     Dim pICMS As String
''                     Dim vICMS As String
''                     Dim pCredSN As String
''                     Dim vCredICMSSN As String
''                     Dim vBCFCPST As String
''                     Dim pFCPST As String
''                     Dim vFCPST As String
''
''                     sOrigem = "0"
''                     modBC = "0"
''                     vBC = Substitui(Format(0, "#######0.00"), ",", ".")
''                     pRedBC = Substitui(Format(0, "#######0.00"), ",", ".")
''                     pICMS = Substitui(Format(0, "##0.00"), ",", ".")
''                     vICMS = Substitui(Format(0, "#######0.00"), ",", ".")
''                     modBCST = "0"
''                     pMVAST = Substitui(Format(0, "##0.00"), ",", ".")
''                     pRedBCST = Substitui(Format(0, "##0.00"), ",", ".")
''                     vBCST = Substitui(Format(0, "#######0.00"), ",", ".")
''                     pICMSST = Substitui(Format(0, "##0.00"), ",", ".")
''                     vBCFCPST = Substitui(Format(0, "##0.00"), ",", ".")
''                     pFCPST = Substitui(Format(0, "##0.00"), ",", ".")
''                     vFCPST = Substitui(Format(0, "##0.00"), ",", ".")
''                     vICMSST = Substitui(Format(0, "#######0.00"), ",", ".")
''                     pCredSN = Substitui(Format(0, "#######0.00"), ",", ".")
''                     vCredICMSSN = Substitui(Format(0, "#######0.00"), ",", ".")
''
''                     Print #1, "N10h|" & _
''                               sOrigem & "|" & _
''                               sCST & "|" & _
''                               modBC & "|" & _
''                               vBC & "|" & _
''                               pRedBC & "|" & _
''                               pICMS & "|" & _
''                               vICMS & "|" & _
''                               modBCST & "|" & _
''                               pMVAST & "|" & _
''                               pRedBCST & "|" & _
''                               vBCST & "|" & _
''                               pICMSST & "|" & _
''                               vICMSST & "|" & _
''                               vBCFCPST & "|" & _
''                               pFCPST & "|" & _
''                               vFCPST & "|" & _
''                               pCredSN & "|" & _
''                               vCredICMSSN
''
''         End Select
''
''       ' Totalizar
''         If rsNaturezasOperacao!Descricao <> "COMPLEMENTO DE VALOR" Then
''            dValorTotalBC = dValorTotalBC + dCalculoBC
''            dValorTotalICMS = dValorTotalICMS + ((dCalculoBC * rsNFeItens!ICMS) / 100)
''            dValorTotalProdutos = dValorTotalProdutos + (rsNFeItens!Quantidade * rsNFeItens!ValorUnitario)
''            dValorTotalDesconto = dValorTotalDesconto + (((rsNFeItens!Quantidade * rsNFeItens!ValorUnitario) * rsNFeItens!Desconto) / 100)
''            dValorTotalFrete = dValorTotalFrete + rsNFeItens!valorfrete
''            dValorTotalNFe = dValorTotalNFe + (rsNFeItens!Quantidade * rsNFeItens!ValorUnitario) + rsNFeItens!valorfrete
''         Else
''            dValorTotalBC = dValorTotalBC + dCalculoBC
''            dValorTotalICMS = dValorTotalICMS + ((dCalculoBC * rsNFeItens!ICMS) / 100)
''            dValorTotalProdutos = 0
''            dValorTotalDesconto = 0
''            dValorTotalFrete = dValorTotalFrete + rsNFeItens!valorfrete
''            dValorTotalNFe = 0
''         End If
''
''       ' PIS
''         Print #1, "Q"
''         Print #1, "Q05|" & _
''                   "99|" & _
''                   "0.00"
''         Print #1, "Q07|" & _
''                   "0.00|" & _
''                   "0.00"
''
''       ' Cofins
''         Print #1, "S"
''         Print #1, "S05|" & _
''                   "99|" & _
''                   "0.00"
''         Print #1, "S07|" & _
''                   "0.00|" & _
''                   "0.00"
''
''         Contador = Contador + 1
''         rsNFeItens.MoveNext
''      Loop
''
''    ' Totais
''      Print #1, "W"
''
''      Dim sValorTotalBC As String
''      Dim sValorTotalICMS As String
''      Dim sValorTotalBaseDeson As String
''      Dim sValorTotalBCST As String
''      Dim sValorTotalICMSST As String
''       Dim sValorTotalvFCPUFDest As String
''       Dim sValorTotalvICMSUFDest As String
''       Dim sValorTotalvICMSUFRemet As String
''      Dim sValorTotalProdutos As String
''      Dim sValorTotalFrete As String
''      Dim sValorTotalSeguro As String
''      Dim sValorTotalDesconto As String
''      Dim sValorTotalII As String
''      Dim sValorTotalIPI As String
''       Dim sValorTotalIPIDevol As String
''      Dim sValorTotalPIS As String
''      Dim sValorTotalCofins As String
''      Dim sValorTotalOutro As String
''      Dim sValorTotalNFe As String
''      Dim sRNTC As String
''      Dim sValorTotalTrib As String
''
''      sValorTotalBC = Substitui(Format(dValorTotalBC, "#########0.00"), ",", ".")
''      sValorTotalICMS = Substitui(Format(dValorTotalICMS, "#########0.00"), ",", ".")
''      sValorTotalBaseDeson = Substitui(Format(0, "#########0.00"), ",", ".")
''      sValorTotalBCST = Substitui(Format(0, "#########0.00"), ",", ".")
''      sValorTotalICMSST = Substitui(Format(0, "#########0.00"), ",", ".")
''       sValorTotalvFCPUFDest = Substitui(Format(0, "#########0.00"), ",", ".")
''       sValorTotalvICMSUFDest = Substitui(Format(0, "#########0.00"), ",", ".")
''       sValorTotalvICMSUFRemet = Substitui(Format(0, "#########0.00"), ",", ".")
''      sValorTotalProdutos = Substitui(Format(dValorTotalProdutos, "#########0.00"), ",", ".")
''      sValorTotalFrete = Substitui(Format(dValorTotalFrete, "#########0.00"), ",", ".")
''      sValorTotalSeguro = Substitui(Format(0, "#########0.00"), ",", ".")
''      sValorTotalDesconto = Substitui(Format(dValorTotalDesconto, "#########0.00"), ",", ".")
''      sValorTotalII = Substitui(Format(0, "#########0.00"), ",", ".")
''      sValorTotalIPI = Substitui(Format(0, "#########0.00"), ",", ".")
''       sValorTotalIPIDevol = Substitui(Format(0, "#########0.00"), ",", ".")
''      sValorTotalPIS = Substitui(Format(0, "#########0.00"), ",", ".")
''      sValorTotalCofins = Substitui(Format(0, "#########0.00"), ",", ".")
''      sValorTotalOutro = Substitui(Format(0, "#########0.00"), ",", ".")
''      sValorTotalNFe = Substitui(Format(dValorTotalNFe, "#########0.00"), ",", ".")
''      sValorTotalTrib = Substitui(Format(0, "#########0.00"), ",", ".")
''
'''''      Print #1, "W02|" & _
'''''                sValorTotalBC & "|" & _
'''''                sValorTotalICMS & "|" & _
'''''                sValorTotalBaseDeson & "|" & _
'''''                sValorTotalBCST & "|" & _
'''''                sValorTotalICMSST & "|" & _
'''''                 sValorTotalvFCPUFDest & "|" & _
'''''                 sValorTotalvICMSUFDest & "|" & _
'''''                 sValorTotalvICMSUFRemet & "|" & _
'''''                sValorTotalProdutos & "|" & _
'''''                sValorTotalFrete & "|" & _
'''''                sValorTotalSeguro & "|" & _
'''''                sValorTotalDesconto & "|" & _
'''''                sValorTotalII & "|" & _
'''''                sValorTotalIPI & "|" & _
'''''                 sValorTotalIPIDevol & "|" & _
'''''                sValorTotalPIS & "|" & _
'''''                sValorTotalCofins & "|" & _
'''''                sValorTotalOutro & "|" & _
'''''                sValorTotalNFe & "|" & _
'''''                sValorTotalTrib
''
''      Print #1, "W02|" & _
''                sValorTotalBC & "|" & _
''                sValorTotalICMS & "|" & _
''                sValorTotalBaseDeson & "|" & _
''                sValorTotalBCST & "|" & _
''                sValorTotalICMSST & "|" & _
''                sValorTotalProdutos & "|" & _
''                sValorTotalFrete & "|" & _
''                sValorTotalSeguro & "|" & _
''                sValorTotalDesconto & "|" & _
''                sValorTotalII & "|" & _
''                sValorTotalIPI & "|" & _
''                sValorTotalPIS & "|" & _
''                sValorTotalCofins & "|" & _
''                sValorTotalOutro & "|" & _
''                sValorTotalNFe & "|" & _
''                sValorTotalTrib
''
''    ' Frete
''      Print #1, "X|" & _
''                rsNFe!FreteConta
''
''      Set rsTransportadores = cnSistema.Execute("SELECT * FROM Transportadores WHERE idTransportador = " & rsNFe!idTransportador)
''      If Not rsTransportadores.EOF Then
''         Print #1, "X03|" & _
''                   RemoveAcentos(rsTransportadores!Nome) & "|" & _
''                   IIf(rsTransportadores!IE_CI <> "ISENTO", Trim(RemoveCaracteres(rsTransportadores!IE_CI)), "ISENTO") & "|" & _
''                   RemoveAcentos(rsTransportadores!Endereco) & " " & RemoveAcentos(rsTransportadores!Bairro) & "|" & _
''                   RemoveAcentos(rsTransportadores!Cidade) & "|" & _
''                   rsTransportadores!UF
''
''         If Len(Trim(RemoveCaracteres(rsTransportadores!CNPJ_CPF))) = 14 Then
''            Print #1, "X04|" & _
''                      RemoveCaracteres(rsTransportadores!CNPJ_CPF)
''         Else
''            Print #1, "X05|" & _
''                      RemoveCaracteres(rsTransportadores!CNPJ_CPF)
''         End If
''      End If
''
'''''   Removido Vers�o 4.0
'''''      sRNTC = ""
'''''      If Trim(UCase(Substitui(rsNFe!PlacaVeiculo, "-", ""))) <> "" Then
'''''         Print #1, "X18|" & _
'''''                   UCase(Substitui(rsNFe!PlacaVeiculo, "-", "")) & "|" & _
'''''                   UCase(rsNFe!UFCaminhao) & "|" & _
'''''                   sRNTC
'''''      End If
''
''    ' Informa�oes dos volumnes
''      If Trim(rsNFe!VolumeQuantidade) <> "" Or Trim(rsNFe!VolumeEspecie) <> "" Or Trim(rsNFe!VolumeMarca) <> "" Or Trim(rsNFe!VolumeNumero) <> "" Or rsNFe!VolumePesoBruto > 0 Or rsNFe!VolumePesoLiquido > 0 Then
''         Print #1, "X26|" & _
''                   IIf(Trim(rsNFe!VolumeQuantidade) <> "", Trim(RemoveAcentos(rsNFe!VolumeQuantidade)), "0") & "|" & _
''                   IIf(Trim(rsNFe!VolumeEspecie) <> "", Trim(RemoveAcentos(rsNFe!VolumeEspecie)), "") & "|" & _
''                   IIf(Trim(rsNFe!VolumeMarca) <> "", Trim(RemoveAcentos(rsNFe!VolumeMarca)), "") & "|" & _
''                   IIf(Trim(rsNFe!VolumeNumero) <> "", Trim(RemoveAcentos(rsNFe!VolumeNumero)), "") & "|" & _
''                   Substitui(Format(rsNFe!VolumePesoLiquido, "##########0.000"), ",", ".") & "|" & _
''                   Substitui(Format(rsNFe!VolumePesoBruto, "##########0.000"), ",", ".")
''      End If
''
''    ' Informa��es de Pagamento
''    Dim sdetPag As String
''    Dim stPag As String
''    Dim svPag As String
''
''''    Print #1, "YA|"
''''    ' YA01|0|01|2.00|1|03|2.00|
''''    Print #1, "YA01|" & _
''''              sdetPag & "|" & _
''''              stPag & "|" & _
''''              svPag
''
''    ' Cobranca
''      Dim bBoletos As Boolean
''
''      Set rsNFeBoletos = cnSistema.Execute("SELECT * FROM NFeBoletos WHERE idNFe = " & rsNFe!idNFe)
''      If Not rsNFeBoletos.EOF Then
''         Print #1, "Y|"
''         bBoletos = True
''      Else
''         bBoletos = False
''      End If
''
''      Do While Not rsNFeBoletos.EOF
''         Print #1, "Y07|" & _
''                   Trim(rsNFeBoletos!Numero) & "|" & _
''                   Format(rsNFeBoletos!Vencimento, "yyyy-mm-dd") & "|" & _
''                   Substitui(Format(rsNFeBoletos!Valor, "#########0.00"), ",", ".")
''
''         rsNFeBoletos.MoveNext
''      Loop
''
''      If bBoletos Then
''         sdetPag = "0"  'A Prazo
''         stPag = "15"   'Boletos
''      Else
''         sdetPag = "0"  'Avista
''         stPag = "01"   'Dinheiro
''      End If
''
''      If sFinalidade = "4" Then     ' Devolu��o
''         stPag = "90"   'Sem Pagamennto
''         svPag = Substitui(Format(0, "#########0.00"), ",", ".")
''      Else
''         svPag = Substitui(Format(dValorTotalNFe, "#########0.00"), ",", ".")
''      End If
''
''      Print #1, "YA|" & _
''                stPag & "|" & _
''                svPag
''
''    ' Informa�oes Adicionais
''      Dim sInfFISCO As String
''      Dim sInfEmpresa As String
''
''      sInfFISCO = RemoveAcentos(rsNFe!InformacoesCorpo)
''      sInfEmpresa = RemoveAcentos(rsNFe!DadosAdicionais) & " - Valor Aprox Tributos R$ " & Format(dValorTributos, "#####,##0.00") & " Fonte: IBPT"
''
''      Print #1, "Z|" & _
''                sInfFISCO & "|" & _
''                sInfEmpresa
''
''    ' Define como Gerada
''      If sAcao = 1 Then
''         cnSistema.Execute "Update NFe set " & _
''                  "Situacao = 1 " & _
''                  "Where idNFe = " & rsNFe!idNFe
''      End If
''
''   End If
''End Function
''
''Private Sub lvwNFes_Click()
''   If lvwNFes.ListItems.Count <> 0 Then
''      Set rsGerarXML = cnSistema.Execute("Select * From NFe WHERE Numero=" & Val(lvwNFes.ListItems(lvwNFes.SelectedItem.Index)))
''      If Not rsGerarXML.EOF Then
''         If rsGerarXML!Situacao = 0 Then
''            cmdTransmitir.Enabled = True
''            cmdValidar.Enabled = True
''         Else
''            cmdTransmitir.Enabled = False
''            cmdValidar.Enabled = False
''         End If
''
''         If rsGerarXML!Situacao = 2 Then
''            If Trim(rsGerarXML!ChaveNFe) <> "" And Trim(rsGerarXML!Protocolo) <> "" Then
''               cmdCancelarNota.Enabled = True
''               cmdImprimirDANFE.Enabled = True
''            Else
''               cmdCancelarNota.Enabled = False
''               cmdImprimirDANFE.Enabled = False
''            End If
''         End If
''      End If
''   End If
''End Sub
''
''Private Sub tmrAtualiza_Timer()
''On Error GoTo Erro
''
''Dim handle As Integer
''Dim Linha As String
''Dim strMensagem As String
''
''   bRetorno = False  '' Verifica se houve algum para recarregar a view
''
''   If Not IsDate(mskDataInicial.Text) Or Not IsDate(mskDataFinal.Text) Then
''      Beep
''      MsgBox "Datas Inv�lidas", vbExclamation, "Erro"
''      Exit Sub
''   End If
''
''   If (IsDate(mskDataInicial.Text) And IsDate(mskDataFinal.Text)) And (CDate(mskDataFinal.Text) >= CDate(mskDataInicial.Text)) Then
''      Set rsNFe = cnSistema.Execute("SELECT * FROM NFe WHERE DataEmissao >= cDate('" & Format(mskDataInicial.Text, "dd/mm/yyyy") & "') AND DataEmissao <= cDate('" & Format(mskDataFinal.Text, "dd/mm/yyyy") & "') Order By Numero")
''   Else
''      Set rsNFe = cnSistema.Execute("SELECT * FROM NFe WHERE DataEmissao >= cDate('" & Format(Date, "dd/mm/yyyy") & "') AND DataEmissao <= cDate('" & Format(Date, "dd/mm/yyyy") & "') Order By Numero")
''   End If
''
''   Do While Not rsNFe.EOF
''      If rsNFe!Situacao = 0 Or rsNFe!Situacao = 1 Then
''         Call ConverterTXT_XML
''
''         ' Verifica se XML foi Autorizado
''         ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''         sEmpresaNFe = LerArquivoINI("NFe", "Empresa", CaminhoINI & "\System.ini")
''         If Dir("C:\NF-e\" & sEmpresaNFe & "\Enviados\Autorizados\" & Format(rsNFe!DataEmissao, "yyyymmdd") & "\" & rsNFe!ChaveNFe & "-procNFe.XML") <> "" Then
''            cnSistema.Execute "Update NFe set " & _
''                     "Situacao = 2 " & _
''                     "Where idNFe = " & rsNFe!idNFe
''         End If
''      End If
''
''      ' Verifica se XML de Cancelamento n�o possui erro
''      ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''      sEmpresaNFe = LerArquivoINI("NFe", "Empresa", CaminhoINI & "\System.ini")
''      If Dir(ArqNFeRetornos & rsNFe!ChaveNFe & "-can.err") <> "" Then
''         ' Atualizar Chave da NFe
''         handle = FreeFile
''         Open ArqNFeRetornos & rsNFe!ChaveNFe & "-can.err" For Input As #handle
''
''         bRetorno = False
''         While Not EOF(handle)
''            Line Input #handle, Linha
''
''            strMensagem = strMensagem & Linha & Chr(13)
''         Wend
''
''         MsgBox strMensagem, vbExclamation + vbOKOnly, "Erro de cancelamento da NFe"
''         Close #handle
''
''         Kill ArqNFeRetornos & rsNFe!ChaveNFe & "-can.err"
''      End If
''
''      ' Atualiza Protocolo
''      ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''      If IsNull(rsNFe!Protocolo) Or Trim(rsNFe!Protocolo) = "" Then
''         Dim nProtocolo As String
''
''         sEmpresaNFe = LerArquivoINI("NFe", "Empresa", CaminhoINI & "\System.ini")
''         If Dir("C:\NF-e\" & sEmpresaNFe & "\Enviados\Autorizados\" & Format(rsNFe!DataEmissao, "yyyymmdd") & "\" & rsNFe!ChaveNFe & "-procNFe.XML") <> "" Then
''            handle = FreeFile
''            Open "C:\NF-e\" & sEmpresaNFe & "\Enviados\Autorizados\" & Format(rsNFe!DataEmissao, "yyyymmdd") & "\" & rsNFe!ChaveNFe & "-procNFe.XML" For Input As #handle
''
''            Line Input #handle, Linha
''
''            nProtocolo = PesquisarTAG(Linha, "nProt")
''
''            If Trim(nProtocolo) <> "" Then
''               cnSistema.Execute "Update NFe set " & _
''                        "Protocolo = '" & nProtocolo & "' " & _
''                        "Where idNFe = " & rsNFe!idNFe
''            End If
''
''            Close #handle
''         End If
''      End If
''
''      ' Verifica se cancelamento foi autorizado
''      ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''      sEmpresaNFe = LerArquivoINI("NFe", "Empresa", CaminhoINI & "\System.ini")
''      If Dir(ArqNFeRetornos & rsNFe!ChaveNFe & "-ret-env-canc.xml") <> "" And rsNFe!Situacao <> 3 Then
''      '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''         handle = FreeFile
''         Open ArqNFeRetornos & rsNFe!ChaveNFe & "-ret-env-canc.XML" For Input As #handle
''
''         Line Input #handle, Linha
''
''         nProtocolo = PesquisarTAG(Linha, "nProt")
''
''         Close #handle
''
''         If Trim(nProtocolo) <> "" Then
''            cnSistema.Execute "Update NFe set " & _
''                     "ProtocoloCancelamento = '" & nProtocolo & "', " & _
''                     "Situacao = 3 " & _
''                     "Where idNFe = " & rsNFe!idNFe
''
''            FileCopy ArqNFeRetornos & rsNFe!ChaveNFe & "-can.XML", "C:\NF-e\" & sEmpresaNFe & "\Enviados\Autorizados\" & Format(rsNFe!DataEmissao, "yyyymmdd") & "\" & rsNFe!ChaveNFe & "-can.XML"
''         End If
''      End If
''
''      rsNFe.MoveNext
''   Loop
''
'' ' Notas Inutilizadas
''   Dim sChaveNFe As String
''   sEmpresaNFe = LerArquivoINI("NFe", "Empresa", CaminhoINI & "\System.ini")
''
''   Set rsNFeInutilizadas = cnSistema.Execute("SELECT * FROM NFeInutilizadas")
''   Do While Not rsNFeInutilizadas.EOF
''      ' Envia solicita��o de Inutiliza��o
''      If IsNull(rsNFeInutilizadas!ChaveNFe) Or Trim(rsNFeInutilizadas!ChaveNFe) = "" Then
''         If Dir(ArqNFeRetornos & StrZero(Val(rsNFeInutilizadas!Numero), 12) & "-ret-gerar-chave.txt") <> "" Then
''            handle = FreeFile
''            Open ArqNFeRetornos & StrZero(Val(rsNFeInutilizadas!Numero), 12) & "-ret-gerar-chave.txt" For Input As #handle
''            Line Input #handle, Linha
''            sChaveNFe = Linha
''            Close #handle
''
''            Set rsUFs = cnSistema.Execute("Select * from UFs Where idUF = " & rsEmpresa!idUF)
''
''            Open "C:\NF-e\Notas\Notas.TXT" For Output As #1
''            Print #1, "tbAmb|" & LerArquivoINI("NFe", "Ambiente", CaminhoINI & "\System.ini")     ' 1 - Produ��o ou 2 - Homologa��o
''            Print #1, "cUF|" & rsUFs!Codigo
''            Print #1, "ano|" & Format(Date, "yy")
''            Print #1, "CNPJ|" & RemoveCaracteres(rsEmpresa!CNPJ_CPF)
''            Print #1, "mod|55"
''            Print #1, "serie|1"
''            Print #1, "nNFIni|" & rsNFeInutilizadas!Numero
''            Print #1, "nNFFin|" & rsNFeInutilizadas!Numero
''            Print #1, "xJust|Numero inutilizado por escolha do usuario"
''
''            Close #1
''
''            sEmpresaNFe = LerArquivoINI("NFe", "Empresa", CaminhoINI & "\System.ini")
''            FileCopy "C:\NF-e\Notas\Notas.TXT", "C:\NF-e\" & sEmpresaNFe & "\Envio\" & Trim(sChaveNFe) & "-ped-inu.txt"
''            Kill "C:\NF-e\Notas\Notas.TXT"
''
''            cnSistema.Execute "Update NFeInutilizadas set " & _
''                     "ChaveNFe = '" & Trim(sChaveNFe) & "' " & _
''                     "Where idNFe = " & rsNFeInutilizadas!idNFe
''
''         End If
''      End If
''
''      ' Verifica se Inutiliza��o foi aceita
''      If IsNull(rsNFeInutilizadas!Protocolo) Or Trim(rsNFeInutilizadas!Protocolo) = "" Then
''         sEmpresaNFe = LerArquivoINI("NFe", "Empresa", CaminhoINI & "\System.ini")
''         If Dir(ArqNFeRetornos & Trim(rsNFeInutilizadas!ChaveNFe) & "-inu.XML") <> "" Then
''            handle = FreeFile
''            Open ArqNFeRetornos & Trim(rsNFeInutilizadas!ChaveNFe) & "-inu.XML" For Input As #handle
''
''            Line Input #handle, Linha
''
''            nProtocolo = PesquisarTAG(Linha, "nProt")
''
''            If Trim(nProtocolo) <> "" Then
''               cnSistema.Execute "Update NFeInutilizadas set " & _
''                        "Protocolo = '" & nProtocolo & "' " & _
''                        "Where idNFe = " & rsNFeInutilizadas!idNFe
''            End If
''
''            Close #handle
''         End If
''      End If
''
''      rsNFeInutilizadas.MoveNext
''   Loop
''
''   ' Verifica outros retornos
''   Call Verifica_Retornos
''
''   Carrega_View
''
''   Exit Sub
''   Resume
''Erro:
''   MsgBox "Erro " & Err & ". " & Err.Description & " - " & TypeName(Me) & ".tmrAtualiza_Timer"
''End Sub
''
''Private Sub cmdNFeDigitacao_Click()
''   If MsgBox("Confirma colocar nota em digita��o", vbYesNo + vbQuestion, "Confirma��o") = vbYes Then
'''      Set rsNFe = cnSistema.Execute("SELECT * FROM NFe WHERE idNFe = " & Val(Mid(lvwNFes.SelectedItem.Key, 2, Len(lvwNFes.SelectedItem.Key))))
''      Set rsNFe = cnSistema.Execute("Select * From NFe WHERE Numero=" & Val(lvwNFes.ListItems(lvwNFes.SelectedItem.Index)))
''      If Not rsNFe.EOF Then
''         If rsNFe!Situacao = 2 Or rsNFe!Situacao = 3 Then
''            MsgBox "Nota Fiscal j� autorizada n�o pode ser alterada", vbExclamation + vbOKOnly, "Aten��o"
''         Else
''            cnSistema.Execute "Update NFe set " & _
''                     "Situacao = 0 " & _
''                     "Where idNFe = " & rsNFe!idNFe
''         End If
''         Carrega_View
''      End If
''   End If
''End Sub
''
''Private Sub cmdImprimirDANFE_Click()
''Dim sArquivo As String
''Dim sCaminho As String
''
''   Set rsNFe = cnSistema.Execute("Select * From NFe WHERE Numero=" & Val(lvwNFes.ListItems(lvwNFes.SelectedItem.Index)))
''
''   sEmpresaNFe = LerArquivoINI("NFe", "Empresa", CaminhoINI & "\System.ini")
''   sArquivo = "C:\NF-e\" & sEmpresaNFe & "\Enviados\Autorizados\" & Format(rsNFe!DataEmissao, "yyyymmdd") & "\" & rsNFe!ChaveNFe & "-procNFe.XML"
''
''   Shell "C:\UNIMAKE\" & sEmpresaNFe & "\UNIDANFE\UNIDANFE.EXE arquivo=" & sArquivo & " visualizar = 1"
''
''End Sub
''
''Private Sub cmdCancelarNota_Click()
''Dim sMotivo
''
''   tmrAtualiza.Enabled = False
''   If MsgBox("Confirma cancelamento da Nota N� " & lvwNFes.ListItems(lvwNFes.SelectedItem.Index), vbYesNo + vbQuestion, "Confirma��o") = vbYes Then
''      Set rsNFe = cnSistema.Execute("Select * From NFe WHERE Numero=" & Val(lvwNFes.ListItems(lvwNFes.SelectedItem.Index)))
''      If Not rsNFe.EOF Then
''         sMotivo = InputBox("Digite o motivo do cancelamento", "Cancelamento", "")
''         If Trim(sMotivo) <> "" Then
''            Set rsEmpresa = cnSistema.Execute("SELECT * FROM Empresa")
''            Set rsUFs = cnSistema.Execute("Select * From UFs WHERE idUF=" & rsEmpresa!idUF)
''
''            Open "C:\NF-e\Notas\Notas.TXT" For Output As #1
''            Print #1, "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "UTF-8" & Chr(34) & "?>"
''            Print #1, "<envEvento xmlns=" & Chr(34) & "http://www.portalfiscal.inf.br/nfe" & Chr(34) & " versao=" & Chr(34) & "1.00" & Chr(34) & ">"
''            Print #1, "<idLote>1</idLote>"
''            Print #1, "<evento xmlns=" & Chr(34) & "http://www.portalfiscal.inf.br/nfe" & Chr(34) & " versao=" & Chr(34) & "1.00" & Chr(34) & ">"
''            Print #1, "<infEvento Id=" & Chr(34) & "ID110111" & rsNFe!ChaveNFe & "01" & Chr(34) & ">"
''            Print #1, "<cOrgao>" & rsUFs!Codigo & "</cOrgao>"
''            Print #1, "<tpAmb>" & LerArquivoINI("NFe", "Ambiente", CaminhoINI & "\System.ini") & "</tpAmb>"   ' 1 - Produ��o ou 2 - Homologa��o
''            Print #1, "<CNPJ>" & RemoveCaracteres(rsEmpresa!CNPJ_CPF) & "</CNPJ>"
''            Print #1, "<chNFe>" & rsNFe!ChaveNFe & "</chNFe>"
''
''            If LerArquivoINI("NFe", "HorarioVerao", App.Path & "\System.ini") Then
''               Print #1, "<dhEvento>" & Format(Date, "YYYY-MM-DD") & "T" & Format(Time, "HH:MM:SS") & "-02:00" & "</dhEvento>"
''            Else
''               Print #1, "<dhEvento>" & Format(Date, "YYYY-MM-DD") & "T" & Format(Time, "HH:MM:SS") & "-03:00" & "</dhEvento>"
''            End If
''
''            Print #1, "<tpEvento>110111</tpEvento>"
''            Print #1, "<nSeqEvento>1</nSeqEvento>"
''            Print #1, "<verEvento>1.00</verEvento>"
''            Print #1, "<detEvento versao=" & Chr(34) & "1.00" & Chr(34) & ">"
''            Print #1, "<descEvento>Cancelamento</descEvento>"
''            Print #1, "<nProt>" & rsNFe!Protocolo & "</nProt>"
''            Print #1, "<xJust>" & sMotivo & "</xJust>"
''            Print #1, "</detEvento>"
''            Print #1, "</infEvento>"
''            Print #1, "</evento>"
''            Print #1, "</envEvento>"
''            Close #1
''
''            sEmpresaNFe = LerArquivoINI("NFe", "Empresa", CaminhoINI & "\System.ini")
''            FileCopy "C:\NF-e\Notas\Notas.TXT", "C:\NF-e\" & sEmpresaNFe & "\Envio\" & Trim(rsNFe!ChaveNFe) & "-env-canc.xml"
''            Kill "C:\NF-e\Notas\Notas.TXT"
''         Else
''            MsgBox "O motivo � obrigat�rio", vbExclamation + vbOKOnly, "Campos Obrigat�rios"
''         End If
''      End If
''   End If
''   tmrAtualiza.Enabled = True
''
''End Sub
''
''Private Sub cmdStatusServico_Click()
''Dim sMotivo
''
''   If MsgBox("Consulta Status do Servi�o", vbYesNo + vbQuestion, "Confirma��o") = vbYes Then
''      If Not rsNFe.EOF Then
''         Set rsTemp = cnSistema.Execute("SELECT * FROM UFs WHERE idUF = " & rsEmpresa!idUF)
''
''         Open "C:\NF-e\Notas\Notas.TXT" For Output As #1
''         Print #1, "tbEmis|1"                                                             ' 1 - Normal
''         Print #1, "tpAmb|" & LerArquivoINI("NFe", "Ambiente", CaminhoINI & "\System.ini") ' 1 - Produ��o ou 2 - Homologa��o
''         Print #1, "cUF|" & rsTemp!Codigo
''         Close #1
''
''         sEmpresaNFe = LerArquivoINI("NFe", "Empresa", CaminhoINI & "\System.ini")
''         FileCopy "C:\NF-e\Notas\Notas.TXT", "C:\NF-e\" & sEmpresaNFe & "\Envio\" & Format(Date, "yyyymmdd") & "T" & Format(Time, "hhmmss") & "-ped-sta.txt"
''         Kill "C:\NF-e\Notas\Notas.TXT"
''      End If
''   End If
''
''End Sub
''
''Private Sub cmdInutilizarNumeracao_Click()
''Dim sNumeroInutilizar
''
''   tmrAtualiza.Enabled = False
''   If MsgBox("Confirma Inutilizar N�mero da Nota", vbYesNo + vbQuestion, "Confirma��o") = vbYes Then
''      sNumeroInutilizar = InputBox("Digite o n�mero a inutilizar", "Inutilizar numera��o", "")
''      Set rsTemp = cnSistema.Execute("SELECT * FROM NFeInutilizadas WHERE Numero = " & Val(sNumeroInutilizar))
''      If rsTemp.EOF Then
''         Set rsUFs = cnSistema.Execute("SELECT * FROM UFs WHERE idUF = " & rsEmpresa!idUF)
''
''         Open "C:\NF-e\Notas\Notas.TXT" For Output As #1
''
''         Print #1, "tpAmb|" & LerArquivoINI("NFe", "Ambiente", CaminhoINI & "\System.ini") ' 1 - Produ��o ou 2 - Homologa��o
'''         Print #1, "versao|3.10"
''         Print #1, "versao|4.00"
''         Print #1, "cUF|" & rsUFs!Codigo
''         Print #1, "ano|" & Format(Date, "yy")
''         Print #1, "CNPJ|" & RemoveCaracteres(rsEmpresa!CNPJ_CPF)
''         Print #1, "mod|55"
''         Print #1, "serie|1"
''         Print #1, "nNFIni|" & Trim(sNumeroInutilizar)
''         Print #1, "nNFFin|" & Trim(sNumeroInutilizar)
''         Print #1, "xJust|Numero inutilizado por escolha do usuario"
''
''         Close #1
''
''         sEmpresaNFe = LerArquivoINI("NFe", "Empresa", CaminhoINI & "\System.ini")
''         FileCopy "C:\NF-e\Notas\Notas.TXT", "C:\NF-e\" & sEmpresaNFe & "\Envio\" & StrZero(Val(sNumeroInutilizar), 12) & "-ped-inu.txt"
''         Kill "C:\NF-e\Notas\Notas.TXT"
''
''         cnSistema.Execute "Insert Into NFeInutilizadas (ChaveNFe,Numero,Data,Protocolo) " & _
''                           "Values ('','" & Val(sNumeroInutilizar) & "','" & Date & "','')"
''      Else
''         MsgBox "N�mero j� Inutilizado", vbExclamation + vbOKOnly, "Erro de NFe"
''      End If
''   End If
''   tmrAtualiza.Enabled = True
''
''End Sub
''
''Private Sub Verifica_Retornos()
''On Error GoTo Erro
''
''Dim Contador As Integer
''Dim Contador2 As Integer
''Dim sStatus As String
''Dim sMotivos As String
''Dim sProtocolo As String
''Dim sNumeroNota As String
''Dim handle As Integer
''Dim Linha As String
''Dim strMensagem As String
''Dim xMotivo As String
''
''   sEmpresaNFe = LerArquivoINI("NFe", "Empresa", CaminhoINI & "\System.ini")
''
''   Dim Arquivos() As String
''   Dim lCtr As Long
''   Arquivos = ListarArquivos("C:\NF-e\" & sEmpresaNFe & "\Retorno")
'''   If UBound(Arquivos) > 0 And Len(Trim(Arquivos(lCtr))) <= 25 Then
''   If UBound(Arquivos) > 0 Then
''      For lCtr = 0 To UBound(Arquivos)
''         If UCase(Mid(Arquivos(lCtr), Len(Trim(Arquivos(lCtr))) - 7, 8)) = "-INU.XML" Then
''            Contador = 1
''
''            ' Verifica se Inutiliza��o foi aceita
''            handle = FreeFile
''            Open ArqNFeRetornos & Arquivos(lCtr) For Input As #handle
''            'Open ArqNFeRetornos & Trim(rsNFeInutilizadas!ChaveNFe) & "-inu.XML" For Input As #handle
''
''            Line Input #handle, Linha
''
''          ' N�mero do Protocolo
''            sProtocolo = PesquisarTAG(Linha, "nProt")
''          ' N�mero da Nota
''            sNumeroNota = PesquisarTAG(Linha, "nNFIni")
''          ' Verifica o Status
''            sStatus = RemoveCaracteres(PesquisarTAG(Linha, "cStat"))
''          ' Verifica o Motivo
''            xMotivo = PesquisarTAG(Linha, "xMotivo")
''
''            If Trim(sStatus) <> "" Then
''               If Trim(xMotivo) <> "" Then
''                  Set ItemList = lvwMensagens.ListItems.Add(, "R" & CStr(IdMensagens), IdMensagens)
''                      ItemList.SubItems(1) = sStatus
''                      ItemList.SubItems(2) = xMotivo
''                      ItemList.SubItems(3) = ""
''
''                  IdMensagens = IdMensagens + 1
''
''               End If
''            End If
''
''            ' Atualiza Protocolo
''            If Trim(sProtocolo) <> "" Then
''               cnSistema.Execute "Update NFeInutilizadas set " & _
''                        "Protocolo = '" & sProtocolo & "' " & _
''                        "Where Numero = " & sNumeroNota
''
''               sStatus = ""
''               sProtocolo = ""
''            End If
''
''            Close #handle
''
''            FileCopy ArqNFeRetornos & Arquivos(lCtr), ArqNFeTemp & Arquivos(lCtr)
''            Kill ArqNFeRetornos & Arquivos(lCtr)
''
''         ElseIf UCase(Mid(Arquivos(lCtr), Len(Trim(Arquivos(lCtr))) - 3, 4)) = ".XML" Then
''            Contador = 1
''
''            handle = FreeFile
''            Open ArqNFeRetornos & Arquivos(lCtr) For Input As #handle
''
''            Line Input #handle, Linha
''          ' Verifica o Status
''            sStatus = RemoveCaracteres(PesquisarTAG(Linha, "cStat"))
''          ' Verifica o Motivo
''            xMotivo = PesquisarTAG(Linha, "xMotivo")
''          ' N�mero do Protocolo
''            sProtocolo = RemoveCaracteres(PesquisarTAG(Linha, "nRec"))
''
''            If Trim(sStatus) <> "" Then
''               If Trim(xMotivo) <> "" Then
''                  Set ItemList = lvwMensagens.ListItems.Add(, "R" & CStr(IdMensagens), IdMensagens)
''                      ItemList.SubItems(1) = sStatus
''                      ItemList.SubItems(2) = xMotivo
''                      ItemList.SubItems(3) = ""
''
''                  IdMensagens = IdMensagens + 1
''               End If
''            End If
''            Close #handle
''
''            FileCopy ArqNFeRetornos & Arquivos(lCtr), ArqNFeTemp & Arquivos(lCtr)
''            Kill ArqNFeRetornos & Arquivos(lCtr)
''
''         ElseIf UCase(Mid(Arquivos(lCtr), Len(Trim(Arquivos(lCtr))) - 3, 4)) = ".ERR" Then
''
''            handle = FreeFile
''            Open ArqNFeRetornos & Arquivos(lCtr) For Input As #handle
''
''            Line Input #handle, Linha
''
''            bRetorno = False
''            While Not EOF(handle)
''               Line Input #handle, Linha
''
''               strMensagem = strMensagem & Linha & Chr(13)
''            Wend
''
''            MsgBox strMensagem, vbExclamation + vbOKOnly, "Erro"
''            Close #handle
''
''            FileCopy ArqNFeRetornos & Arquivos(lCtr), ArqNFeTemp & Arquivos(lCtr)
''            Kill ArqNFeRetornos & Arquivos(lCtr)
''
''         ElseIf UCase(Mid(Arquivos(lCtr), Len(Trim(Arquivos(lCtr))) - 2, 3)) = ".TXT" Then
''            FileCopy ArqNFeRetornos & Arquivos(lCtr), ArqNFeTemp & Arquivos(lCtr)
''            Kill ArqNFeRetornos & Arquivos(lCtr)
''
''         End If
''      Next
''   End If
''
''   Exit Sub
''Erro:
''   MsgBox "Erro " & Err & ". " & Err.Description & " - " & TypeName(Me) & ".Verifica_Retornos"
''End Sub
''
''Private Sub ConverterTXT_XML()
''Dim handle As Integer
''Dim Linha As String
''Dim bRetorno As Boolean
''Dim strMensagem As String
''
''   ' Verifica se XML foi convertido com sucesso
''   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''   sEmpresaNFe = LerArquivoINI("NFe", "Empresa", CaminhoINI & "\System.ini")
''   If Dir(ArqNFeRetornos & StrZero(rsNFe!Numero, 6) & "_" & RemoveCaracteres(rsEmpresa!CNPJ_CPF) & "_001_" & Format(rsNFe!DataEmissao, "dd_mm_yyyy") & "-nfe.txt") <> "" Then
''      ' Atualizar Chave da NFe
''      handle = FreeFile
''      Open ArqNFeRetornos & StrZero(rsNFe!Numero, 6) & "_" & RemoveCaracteres(rsEmpresa!CNPJ_CPF) & "_001_" & Format(rsNFe!DataEmissao, "dd_mm_yyyy") & "-nfe.txt" For Input As #handle
''
''      bRetorno = False
''      While Not EOF(handle)
''         Line Input #handle, Linha
''         If Mid(Linha, 1, 5) = "cStat" Then
''            If Mid(Linha, 7, 2) = "01" Then
''               bRetorno = True
''            End If
''         End If
''
''         If bRetorno Then
''            If Mid(Linha, 1, 11) = "Nota fiscal" Then
''               If Val(Mid(Linha, 14, 9)) = rsNFe!Numero Then
''                  cnSistema.Execute "Update NFe set " & _
''                           "ChaveNFe = '" & Mid(Linha, 47, 44) & "' " & _
''                           "Where idNFe = " & rsNFe!idNFe
''
''               End If
''            End If
''         End If
''      Wend
''      Close #handle
''      If Dir(ArqNFeRetornos & StrZero(rsNFe!Numero, 6) & "_" & RemoveCaracteres(rsEmpresa!CNPJ_CPF) & "_001_" & Format(rsNFe!DataEmissao, "dd_mm_yyyy") & "-nfe.txt") <> "" Then FileCopy ArqNFeRetornos & StrZero(rsNFe!Numero, 6) & "_" & RemoveCaracteres(rsEmpresa!CNPJ_CPF) & "_001_" & Format(rsNFe!DataEmissao, "dd_mm_yyyy") & "-nfe.txt", ArqNFeTemp & StrZero(rsNFe!Numero, 6) & "_" & RemoveCaracteres(rsEmpresa!CNPJ_CPF) & "_001_" & Format(rsNFe!DataEmissao, "dd_mm_yyyy") & "-nfe.txt"
''      If Dir(ArqNFeRetornos & StrZero(rsNFe!Numero, 6) & "_" & RemoveCaracteres(rsEmpresa!CNPJ_CPF) & "_001_" & Format(rsNFe!DataEmissao, "dd_mm_yyyy") & "-nfe-orig.txt") <> "" Then FileCopy ArqNFeRetornos & StrZero(rsNFe!Numero, 6) & "_" & RemoveCaracteres(rsEmpresa!CNPJ_CPF) & "_001_" & Format(rsNFe!DataEmissao, "dd_mm_yyyy") & "-nfe-orig.txt", ArqNFeTemp & StrZero(rsNFe!Numero, 6) & "_" & RemoveCaracteres(rsEmpresa!CNPJ_CPF) & "_001_" & Format(rsNFe!DataEmissao, "dd_mm_yyyy") & "-nfe-orig.txt"
''      If Dir(ArqNFeRetornos & StrZero(rsNFe!Numero, 6) & "_" & RemoveCaracteres(rsEmpresa!CNPJ_CPF) & "_001_" & Format(rsNFe!DataEmissao, "dd_mm_yyyy") & "-nfe.txt") <> "" Then Kill ArqNFeRetornos & StrZero(rsNFe!Numero, 6) & "_" & RemoveCaracteres(rsEmpresa!CNPJ_CPF) & "_001_" & Format(rsNFe!DataEmissao, "dd_mm_yyyy") & "-nfe.txt"
''      If Dir(ArqNFeRetornos & StrZero(rsNFe!Numero, 6) & "_" & RemoveCaracteres(rsEmpresa!CNPJ_CPF) & "_001_" & Format(rsNFe!DataEmissao, "dd_mm_yyyy") & "-nfe.txt") <> "" Then Kill ArqNFeRetornos & StrZero(rsNFe!Numero, 6) & "_" & RemoveCaracteres(rsEmpresa!CNPJ_CPF) & "_001_" & Format(rsNFe!DataEmissao, "dd_mm_yyyy") & "-nfe-orig.txt"
''   End If
''
''   ' Verifica se XML teve erro de convers�o
''   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''   sEmpresaNFe = LerArquivoINI("NFe", "Empresa", CaminhoINI & "\System.ini")
''   If Dir(ArqNFeRetornos & StrZero(rsNFe!Numero, 6) & "_" & RemoveCaracteres(rsEmpresa!CNPJ_CPF) & "_001_" & Format(rsNFe!DataEmissao, "dd_mm_yyyy") & "-nfe.err") <> "" Then
''      ' Atualizar Chave da NFe
''      handle = FreeFile
''      Open ArqNFeRetornos & StrZero(rsNFe!Numero, 6) & "_" & RemoveCaracteres(rsEmpresa!CNPJ_CPF) & "_001_" & Format(rsNFe!DataEmissao, "dd_mm_yyyy") & "-nfe.err" For Input As #handle
''
''      bRetorno = False
''      While Not EOF(handle)
''         Line Input #handle, Linha
''         If Mid(Linha, 1, 5) = "cStat" Then
''            If Mid(Linha, 7, 2) = "01" Then
''               bRetorno = True
''            End If
''         End If
''
''         strMensagem = strMensagem & Linha & Chr(13)
''      Wend
''
''      cnSistema.Execute "Update NFe set " & _
''               "Situacao = 0 " & _
''               "Where idNFe = " & rsNFe!idNFe
''
''      MsgBox strMensagem, vbExclamation + vbOKOnly, "Erro de envio da NFe"
''      Close #handle
''
''      Kill ArqNFeRetornos & StrZero(rsNFe!Numero, 6) & "_" & RemoveCaracteres(rsEmpresa!CNPJ_CPF) & "_001_" & Format(rsNFe!DataEmissao, "dd_mm_yyyy") & "-nfe.err"
''   End If
''
''   ' Verifica se XML n�o possui erros
''   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''   sEmpresaNFe = LerArquivoINI("NFe", "Empresa", CaminhoINI & "\System.ini")
''   If Dir(ArqNFeRetornos & rsNFe!ChaveNFe & "-nfe.err") <> "" Then
''      ' Atualizar Chave da NFe
''      handle = FreeFile
''      Open ArqNFeRetornos & rsNFe!ChaveNFe & "-nfe.err" For Input As #handle
''
''      bRetorno = False
''      While Not EOF(handle)
''         Line Input #handle, Linha
''
''         strMensagem = strMensagem & Linha & Chr(13)
''      Wend
''
''      cnSistema.Execute "Update NFe set " & _
''               "Situacao = 0 " & _
''               "Where idNFe = " & rsNFe!idNFe
''
''      MsgBox strMensagem, vbExclamation + vbOKOnly, "Erro de envio da NFe"
''      Close #handle
''
''      Kill ArqNFeRetornos & rsNFe!ChaveNFe & "-nfe.err"
''   End If
''End Sub
''
''Private Sub cmdFecharNFe_Click()
''   KillProcess "uninfe.exe"
''End Sub
''
''Private Sub LerMensagens()
''Dim Arquivos() As String
''Dim lCtr As Long
''Dim Contador As Integer
''Dim Contador2 As Integer
''
'''Dim sStatus As String
''Dim sMotivos As String
'''Dim sProtocolo As String
'''Dim sNumeroNota As String
''Dim handle As Integer
''Dim Linha As String
''Dim xMotivo As String
''Dim strMensagem As String
''Dim sStatus As String
''
''   Arquivos = ListarArquivos(ArqNFeRetornos)
'''''''   Arquivos = ListarArquivos("C:\NF-e\" & sEmpresaNFe & "\Retorno")
'''''   Arquivos = ListarArquivos("C:\NF-e\" & sEmpresaNFe & "\Temp")
'''   Arquivos = ListarArquivos("C:\XML\NFe 2 - Modelos XML de Retorno")
''   If UBound(Arquivos) > 0 Then
''      For lCtr = 0 To UBound(Arquivos)
''          If UCase(Mid(Arquivos(lCtr), Len(Trim(Arquivos(lCtr))) - 3, 4)) = ".XML" Then
'''          If UCase(Mid(Arquivos(lCtr), Len(Trim(Arquivos(lCtr))) - 7, 8)) = "-INU.XML" Then
''
''            handle = FreeFile
'''            Open "C:\XML\NFe 2 - Modelos XML de Retorno" & "\" & Arquivos(lCtr) For Input As #handle
''            Open ArqNFeTemp & Arquivos(lCtr) For Input As #handle
''
''            Line Input #handle, Linha
''
''            ' Preenche Campos
''            sStatus = PesquisarTAG(Linha, "cStat")
''            xMotivo = PesquisarTAG(Linha, "xMotivo")
''
''            Close #handle
''
''            ' Adiciona mensagens a lista
''            If sStatus <> "" Then
''               Set ItemList = lvwMensagens.ListItems.Add(, "R" & CStr(IdMensagens), IdMensagens)
''                   ItemList.SubItems(1) = sStatus
''                   ItemList.SubItems(2) = xMotivo
''                   ItemList.SubItems(3) = ""
''
''               IdMensagens = IdMensagens + 1
''
''               sStatus = ""
''               xMotivo = ""
''            End If
''
''            ' Exclui arquivo lido
'''''            FileCopy ArqNFeRetornos & Arquivos(lCtr), ArqNFeTemp & Arquivos(lCtr)
'''''            Kill ArqNFeRetornos & Arquivos(lCtr)
''
''          End If
''      Next
''   End If
''
''End Sub
''
''Public Function PesquisarTAG(sCampo As String, sTAG As String) As String
''Dim sConteudo As String
''Dim sTAGInicio As String
''Dim sTAGFim As String
''Dim Contador As Double
''Dim Contador2 As Double
''
''   sTAGInicio = "<" & sTAG & ">"
''   sTAGFim = "</" & sTAG & ">"
''
''   Contador = 1
''   For Contador = 1 To Len(sCampo)
''     ' Verifica o Motivo
''       If Mid(sCampo, Contador, Len(sTAGInicio)) = sTAGInicio Then
''          For Contador2 = Contador To Len(sCampo)
''              If Mid(sCampo, Contador2, Len(sTAGFim)) = sTAGFim Then
''                 sConteudo = Mid(sCampo, Contador + Len(sTAGInicio), Contador2 - Contador - Len(sTAGInicio))
''                 Contador2 = Len(sCampo)
''              End If
''          Next
''       End If
''   Next
''
''   PesquisarTAG = sConteudo
''
''End Function
''
'''''Public Function PreencheCampoXML(sCampo As String, ByVal PosicaoAtual As Double) As String
''''''Dim sVerifica As String, sTroca As String, sNCampo As String, Procura As Integer
'''''Dim CampoInicial As String
'''''Dim CampoFinal As String
'''''Dim ContadorInicial As Integer
'''''Dim ContadorFinal As Integer
'''''
'''''   CampoInicial = "<" & sCampo & ">"
'''''   CampoFinal = "</" & sCampo & ">"
'''''   ContadorInicial = Len(CampoInicial)
'''''   ContadorFinal = Len(CampoFinal)
'''''
'''''   If Mid(sCampo, Contador, 7) = "<chNFe>" Then
'''''      For Contador2 = Contador To Len(sCampo)
'''''          If Mid(sCampo, Contador2, 8) = "</chNFe>" Then
'''''             strChaveNfe = Mid(sCampo, Contador + 7, Contador2 - Contador - 7)
'''''             Contador2 = Len(sCampo)
'''''          End If
'''''      Next
'''''   End If
'''''
'''''
'''''
'''''
'''''
'''''   sVerifica = "123456789"
'''''   sTroca = "0"
'''''   For Contador = 1 To Len(sCampo)
'''''       Procura = InStr(sVerifica, Mid(sCampo, Contador, 1))
'''''       If Procura <> 0 Then
'''''          sNCampo = sNCampo + Mid(sTroca, Procura, 1)
'''''       Else
'''''          sNCampo = sNCampo + Mid(sCampo, Contador, 1)
'''''       End If
'''''   Next
'''''
'''''   CDecimais = Substitui(sNCampo, ",", ".")
'''''End Function
'''''
''Private Sub tmrLerRetornos_Timer()
'''   Call LerMensagens
''End Sub
''
''Private Sub cmdHistorico_Click()
''   frmLerXML.Show
''End Sub
''
''Private Sub cmdConsultaNota_Click()
''Dim sMotivo
''
''   tmrAtualiza.Enabled = False
''   If MsgBox("Confirma Consulta NF-e", vbYesNo + vbQuestion, "Confirma��o") = vbYes Then
'''      Set rsNFe = cnSistema.Execute("SELECT * FROM NFe WHERE idNFe = " & Val(Mid(lvwNFes.SelectedItem.Key, 2, Len(lvwNFes.SelectedItem.Key))))
''      Set rsNFe = cnSistema.Execute("Select * From NFe WHERE Numero=" & Val(lvwNFes.ListItems(lvwNFes.SelectedItem.Index)))
''      If Not rsNFe.EOF Then
'''         sMotivo = InputBox("Digite o motivo do cancelamento", "Cancelamento", "")
'''         If Trim(sMotivo) <> "" Then
''            Open "C:\NF-e\Notas\Notas.TXT" For Output As #1
''            Print #1, "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "UTF-8" & Chr(34) & "?>"
''            Print #1, "<consSitNFe xmlns=" & Chr(34) & "http://www.portalfiscal.inf.br/nfe" & Chr(34) & " versao=" & Chr(34) & "4.00" & Chr(34) & ">"
''            Print #1, "<tpAmb>" & LerArquivoINI("NFe", "Ambiente", CaminhoINI & "\System.ini") & "</tpAmb>"   ' 1 - Produ��o ou 2 - Homologa��o
''            Print #1, "<xServ>CONSULTAR</xServ> "
''            Print #1, "<chNFe>" & rsNFe!ChaveNFe & "</chNFe>"
''            Print #1, "</consSitNFe>"
''            Close #1
''
''            sEmpresaNFe = LerArquivoINI("NFe", "Empresa", CaminhoINI & "\System.ini")
''            FileCopy "C:\NF-e\Notas\Notas.TXT", "C:\NF-e\" & sEmpresaNFe & "\Envio\" & Trim(rsNFe!ChaveNFe) & "-ped-sit.xml"
''            Kill "C:\NF-e\Notas\Notas.TXT"
'''         Else
'''            MsgBox "O motivo � obrigat�rio", vbExclamation + vbOKOnly, "Campos Obrigat�rios"
'''         End If
''      End If
''   End If
''   tmrAtualiza.Enabled = True
''
''End Sub
''
''Private Sub cmdEnviarEmail_Click()
''Dim sArquivo As String
''Dim sCaminho As String
''Dim cdoConfiguration As CDO.Configuration
''Dim cdoData As ADODB.Fields
''Dim cdoMensagem As New CDO.Message
''Dim strMensagem As String
''
''
''   Set rsNFe = cnSistema.Execute("Select * From NFe WHERE Numero=" & Val(lvwNFes.ListItems(lvwNFes.SelectedItem.Index)))
''   Set rsClientes = cnSistema.Execute("Select * From Clientes WHERE idCliente = " & rsNFe!idCliente)
''
''   If Trim(rsClientes!EMail) = "" Then
''      MsgBox "E-mail do cliente n�o encontrado", vbExclamation + vbOKOnly, "Campos Obrigat�rios"
''      Exit Sub
''   End If
''
'''''''''''''''''''''''''''''''''''''''''''''''
''
''   strMensagem = strMensagem & "De: " & rsEmpresa!Nome & Chr(13)
''   strMensagem = strMensagem & "E-mail: " & rsEmpresa!EMail & Chr(13)
''   strMensagem = strMensagem & Chr(13)
''   strMensagem = strMensagem & "Para: " & rsClientes!Nome & Chr(13)
''   strMensagem = strMensagem & "E-mail: " & rsClientes!EMail & Chr(13)
''   strMensagem = strMensagem & Chr(13)
''   strMensagem = strMensagem & "Estamos enviando em anexo arquivo XML da nota fiscal N. " & rsNFe!Numero
''
'''''''''''''''''''''''''''''''''''''''''''''''
''
''   If MsgBox("Envio de E-mail " & Chr(13) & strMensagem, vbYesNo + vbQuestion, "Confirma��o") = vbYes Then
''
'''' Para Consulta
''''Item(cdoSendUsingMethod) = 2
''''  .Item(cdoSMTPServer) = "pop.mail.yahoo.com.br"
''''  .Item(cdoSMTPServerPort) = 995
''''  .Item(cdoSMTPConnectionTimeout) = 15
''''  .Item(cdoSMTPAuthenticate) = cdoBasic
''''  .Item(cdoSMTPUseSSL) = True
''''  .Item(cdoSendUserName) = vUser
''''  .Item(cdoSendPassword) = vPass
''''  .Update
''
''      sEmpresaNFe = LerArquivoINI("NFe", "Empresa", CaminhoINI & "\System.ini")
''      sArquivo = "C:\NF-e\" & sEmpresaNFe & "\Enviados\Autorizados\" & Format(rsNFe!DataEmissao, "yyyymmdd") & "\" & rsNFe!ChaveNFe & "-procNFe.XML"
''      If Dir(sArquivo) <> "" Then
''         Set cdoConfiguration = New CDO.Configuration
''         Set cdoData = cdoConfiguration.Fields
''         Set rsEmpresa = cnSistema.Execute("SELECT * FROM Empresa")
''         With cdoData
''            .Item(cdoSendUsingMethod) = 2
''            .Item(cdoSMTPServerPort) = LerArquivoINI("EMail", "Porta", CaminhoINI & "\System.ini")
''            .Item(cdoSMTPServer) = rsEmpresa!ServidorSMTP 'rsSistema!Mail_SMTP
''            .Item(cdoSMTPConnectionTimeout) = 20
''            .Item(cdoSMTPAuthenticate) = 1
''            .Item(cdoSMTPUseSSL) = True
''            .Item(cdoSendUserName) = rsEmpresa!EmailUsuario 'rsSistema!Mail_User
''            .Item(cdoSendPassword) = rsEmpresa!EmailSenha 'rsSistema!Mail_Pass
''            .Update
''         End With
''
''         Set cdoMensagem = New CDO.Message
''         With cdoMensagem
''            Set .Configuration = cdoConfiguration
''                .To = rsClientes!EMail 'rsTotal!Email
''                .From = rsEmpresa!EMail 'rsSistema!Mail_From
''                .Subject = rsEmpresa!Nome 'rsSistema!Mail_Subject
''                .HTMLBody = "Estamos enviando em anexo arquivo XML da nota fiscal N. " & rsNFe!Numero  'rsSistema!Mail_Body
''                .AddAttachment sArquivo 'rsSistema!Mail_Directory & "\" & rsTotal!idFaturamento & ".pdf"
''                .Send
''         End With
''
''         MsgBox "E-mail Enviado", vbExclamation + vbOKOnly, "Informa��o"
''      Else
''         MsgBox "Arquivo: " & sArquivo & " N�o encontrado", vbExclamation + vbOKOnly, "Campos Obrigat�rios"
''      End If
''   End If
''
''   Set cdoMensagem = Nothing
''End Sub
''
''Private Sub cmdGerarComplementoICMS_Click()
''Dim Contador As Integer
''Dim Contador2 As Integer
''Dim sStatus As String
''Dim sMotivos As String
''Dim sProtocolo As String
''Dim handle As Integer
''Dim Linha As String
''Dim strMensagem As String
''Dim xMotivo As String
''
''' Informa��es a buscar no XML
''Dim strNumeroNota As String
''Dim strIdNFe As String
''Dim strData As String
''Dim strCNPJCliente As String
''Dim strValorTotalNota As String
''
''Dim strChaveNFe As String
''Dim strnNF As String
''Dim strdEmi As String
''Dim strCNPJDest As String
''Dim strvNF As String
''
''Dim intNumero As Integer
''Dim strCFOP As String
''Dim dblValorICMS As String
''Dim dblValorFrete As String
''Dim dblValorTotalProdutos As String
''Dim dblBaseICMSSubstituicao As String
''Dim dblValorICMSSubstituicao As String
''Dim dblOutrasDespesas As String
''Dim dblValorTotalNota As String
''Dim iTransportador As Integer, iFreteConta As Integer
''Dim strPlaca As String
''Dim strUFPlaca As String
''Dim strVolumeQuantidade As String
''Dim strVolumeMarca As String
''Dim strVolumeEspecie As String
''Dim strVolumeNumero As String
''Dim strVolumePesoBruto As String
''Dim strVolumePesoLiquido As String
''Dim strInformacoesCorpo As String
''Dim intFormaPagamento As Integer
''Dim dblDescontoGeral As String
''Dim dblBonificacao As String
''Dim strDocumento As String
''Dim strObservacao As String
''Dim bGeradaNFe As Boolean
''Dim dblAliquotaICMS As Double
''Dim strCFOPItem As String
''
''   dblValorICMS = "0"
''   dblValorFrete = "0"
''   dblValorTotalProdutos = "0"
''   dblBaseICMSSubstituicao = "0"
''   dblValorICMSSubstituicao = "0"
''   dblOutrasDespesas = "0"
''   dblValorTotalNota = "0"
''   iTransportador = 1
''   iFreteConta = 9
''   strPlaca = "   -    "
''   strUFPlaca = ""
''   strVolumeQuantidade = "0"
''   strVolumeMarca = "0"
''   strVolumeEspecie = "0"
''   strVolumeNumero = "0"
''   strVolumePesoBruto = "0"
''   strVolumePesoLiquido = "0"
''   strInformacoesCorpo = ""
''   intFormaPagamento = 2
''   dblDescontoGeral = "0"
''   dblBonificacao = "0"
''   strDocumento = ""
''   strObservacao = ""
''
''   sEmpresaNFe = LerArquivoINI("NFe", "Empresa", CaminhoINI & "\System.ini")
''
''   Screen.MousePointer = vbHourglass
''   frmVisualiza.lvwDados.ListItems.Clear
''   frmVisualiza.lvwDados.ColumnHeaders.Clear
''   frmVisualiza.lvwDados.ColumnHeaders.Add , , "Nota", 700
''   frmVisualiza.lvwDados.ColumnHeaders.Add , , "Chave", 5300
''
''   Dim Arquivos() As String
''   Dim lCtr As Long
''   Arquivos = ListarArquivos("C:\Sistemas\Importar")                                         '' Determina Pasta onde est�o os arquivos
''   If UBound(Arquivos) > 0 Then                                                              '' Verifica se existem arquivos na pasta
''      For lCtr = 0 To UBound(Arquivos)
''         If UCase(Mid(Arquivos(lCtr), Len(Trim(Arquivos(lCtr))) - 2, 3)) = "XML" Then        '' Verifica se encontrou algum XML
'''''            Contador = 1
''            handle = FreeFile
''            Open "C:\Sistemas\Importar\" & Arquivos(lCtr) For Input As #handle               '' Abre arquivo importado
''            Line Input #handle, Linha
''
''            Contador = 1
''            For Contador = 1 To Len(Linha)                                                   '' L� o arquivo linha a linha em busca da informa��o
''              ' Chave da NFe
''                If Mid(Linha, Contador, 7) = "<chNFe>" Then
''                   For Contador2 = Contador To Len(Linha)
''                       If Mid(Linha, Contador2, 8) = "</chNFe>" Then
''                          strChaveNFe = Mid(Linha, Contador + 7, Contador2 - Contador - 7)
''                          Contador2 = Len(Linha)
''                       End If
''                   Next
''                End If
''
''              ' N�mero da NFe
''                If Mid(Linha, Contador, 5) = "<nNF>" Then
''                   For Contador2 = Contador To Len(Linha)
''                       If Mid(Linha, Contador2, 6) = "</nNF>" Then
''                          strnNF = Mid(Linha, Contador + 5, Contador2 - Contador - 5)
''                          Contador2 = Len(Linha)
''                       End If
''                   Next
''                End If
''
''              ' Data de emiss�o
''                If Mid(Linha, Contador, 6) = "<dEmi>" Then
''                   For Contador2 = Contador To Len(Linha)
''                       If Mid(Linha, Contador2, 7) = "</dEmi>" Then
''                          strdEmi = Mid(Linha, Contador + 6, Contador2 - Contador - 6)
''                          Contador2 = Len(Linha)
''                       End If
''                   Next
''                End If
''
''              ' CNPJ Destinat�rio
''                If Mid(Linha, Contador, 6) = "<dest>" Then
''                   For Contador2 = Contador To Len(Linha)
''                       If Mid(Linha, Contador2, 7) = "</CNPJ>" Then
''                          strCNPJDest = Mid(Linha, Contador + 12, Contador2 - Contador - 12)
''                          Contador2 = Len(Linha)
''                       End If
''                   Next
''                End If
''
''              ' CPF Destinat�rio
''                If Mid(Linha, Contador, 6) = "<dest>" Then
''                   For Contador2 = Contador To Len(Linha)
''                       If Mid(Linha, Contador2, 6) = "</CPF>" Then
''                          strCNPJDest = Mid(Linha, Contador + 11, Contador2 - Contador - 11)
''                          Contador2 = Len(Linha)
''                       End If
''                   Next
''                End If
''
''              ' Valor da NFe
''                If Mid(Linha, Contador, 5) = "<vNF>" Then
''                   For Contador2 = Contador To Len(Linha)
''                       If Mid(Linha, Contador2, 6) = "</vNF>" Then
''                          strvNF = Mid(Linha, Contador + 5, Contador2 - Contador - 5)
''                          Contador2 = Len(Linha)
''                       End If
''                   Next
''                End If
''
''               If Len(Trim(strnNF)) <> 0 And Len(Trim(strChaveNFe)) <> 0 And Len(Trim(strdEmi)) <> 0 And Len(Trim(strCNPJDest)) <> 0 And Len(Trim(strvNF)) <> 0 Then
''                  Set ProcuraItem = frmVisualiza.lvwDados.FindItem(strnNF)
''                  If ProcuraItem Is Nothing Then
''                     Set ItemList = frmVisualiza.lvwDados.ListItems.Add(, "R" & strnNF, strnNF)
''                         ItemList.SubItems(1) = strChaveNFe
''
''                     'MsgBox strnNF & Chr(13) & strChaveNFe & Chr(13) & strdEmi & Chr(13) & strCNPJDest & Chr(13) & strvNF, vbExclamation + vbOKOnly, "Campos Obrigat�rios"
''
''                     ' Empresa
''                     Set rsEmpresa = cnSistema.Execute("SELECT * FROM Empresa")
''
''                     ' Cliente
''                     Set rsClientes = cnSistema.Execute("Select * From Clientes WHERE CNPJ_CPF = '" & strCNPJDest & "'")
''
''                     ' Naturezas Operacao
''                     Set rsNaturezasOperacao = cnSistema.Execute("Select * From NaturezasOperacao WHERE Descricao = 'COMPLEMENTO DE VALOR'")
''                     If Not rsEmpresa.EOF Then
''                        If rsClientes!UF = rsEmpresa!UF Then
''                           strCFOP = rsNaturezasOperacao!CFOPDentroUF
''                           strCFOPItem = rsNaturezasOperacao!CFOPDentroUF
''                           dblAliquotaICMS = 17
''                        Else
''                           strCFOP = rsNaturezasOperacao!CFOPForaUF
''                           strCFOPItem = rsNaturezasOperacao!CFOPForaUF
''                           dblAliquotaICMS = 12
''                        End If
''                     End If
''
''                     Set rsCFOPs = cnSistema.Execute("Select * From CFOPs Where CFOP = '" & strCFOP & "'")
''                     If Not rsCFOPs.EOF Then strCFOP = rsCFOPs!idCFOP
''
''                     ' Ultima Nota
''                     Set rsNFe = cnSistema.Execute("Select * From NFe ORDER BY Numero DESC")
''                     intNumero = rsNFe!Numero + 1
''
''                     ' Inserir Nota
''                     If Not rsClientes.EOF Then
''                        cnSistema.Execute "Insert Into NFe (Numero,Cupom,idCliente,idNaturezaOperacao,idCFOP,DadosAdicionais,DataEmissao,DataCaixa,DataVencimento,Hora,BaseCalculoICMS,ValorICMS,ValorFrete,ValorTotalProdutos,BaseICMSSubstituicao,ValorICMSSubstituicao,OutrasDespesas,ValorTotalNota,idTransportador,FreteConta,PlacaVeiculo,UFCaminhao,VolumeQuantidade,VolumeMarca,VolumeEspecie,VolumeNumero,VolumePesoBruto,VolumePesoLiquido,InformacoesCorpo,idFormaPagamento,DescontoGeral,Bonificacao,Documento,Observacao,GeradaNFe,Situacao,NumeroNFeComplementar,ChaveAcessoNFeComplementar) " & _
''                                          "Values (" & intNumero & ",0," & rsClientes!idCliente & "," & rsNaturezasOperacao!idNaturezaOperacao & "," & strCFOP & ",'','" & Date & "','" & Date & "','" & Date & "','" & Time & "','" & CStrValor(Substitui(strvNF, ".", ",")) & "','" & Val(Substitui(dblValorICMS, ",", ".")) & "'," & _
''                                                  "'" & Val(Substitui(dblValorFrete, ",", ".")) & "','" & Val(Substitui(dblValorTotalProdutos, ",", ".")) & "','" & Val(Substitui(dblBaseICMSSubstituicao, ",", ".")) & "','" & Val(Substitui(dblValorICMSSubstituicao, ",", ".")) & "','" & Val(Substitui(dblOutrasDespesas, ",", ".")) & "'," & _
''                                                  "'" & Val(Substitui(dblValorTotalNota, ",", ".")) & "'," & iTransportador & "," & iFreteConta & ",'" & UCase(strPlaca) & "','" & strUFPlaca & "','" & strVolumeQuantidade & "','" & strVolumeMarca & "','" & strVolumeEspecie & "','" & strVolumeNumero & "','" & Val(Substitui(strVolumePesoBruto, ",", ".")) & "','" & Val(Substitui(strVolumePesoLiquido, ",", ".")) & "','" & strInformacoesCorpo & "'" & _
''                                                  "," & intFormaPagamento & ",'" & Val(Substitui(dblDescontoGeral, ",", ".")) & "','" & Val(Substitui(dblBonificacao, ",", ".")) & "','" & strDocumento & "','" & strObservacao & "'," & bGeradaNFe & ",0," & strnNF & ",'" & strChaveNFe & "')"
''                     End If
''
''                     ' Inserir Item
''                     Set rsNFe = cnSistema.Execute("Select * From NFe Where Numero = " & intNumero)
''                     If Not rsNFe.EOF Then
''                        Dim intIdProduto As Integer
''                        Dim dblQuantidade As String
''                        Dim dblDesconto As String
''                        Dim dblValorUnitario As String
''                        Dim dblICMSProduto As String
''                        Dim dblBaseReduzidaICMS As String
''                        Dim strDescricaoComplementar As String
''                        Dim intUnidade As Integer
''                        Dim intSituacaoTributaria As Integer
''                        Dim strDiscriminacaoProduto As String
''                        Dim dblIPIProduto As String
''                        Dim dblBaseReduzidaIPI As String
''                        Dim strClassificacaoFiscal As String
''                        'Dim dblValorFrete As String
''
''                        intIdProduto = 1
''                        dblQuantidade = "1"
''                        dblDesconto = "0"
''                        dblValorUnitario = Substitui(strvNF, ".", ",")
''                        dblICMSProduto = dblAliquotaICMS
''                        dblBaseReduzidaICMS = "0"
''                        strDescricaoComplementar = ""
''                        intUnidade = 2                               ' UN - Unidade
''                        intSituacaoTributaria = 1                    ' 1 - Tributado Integralmente
''                        strDiscriminacaoProduto = ""
''                        dblIPIProduto = "0"
''                        dblBaseReduzidaIPI = "0"
''                        strClassificacaoFiscal = ""
''                        dblValorFrete = "0"
''
''                        cnSistema.Execute "Insert Into NFeItens (idNFe,idProduto,Data,Quantidade,Desconto,ValorUnitario,ICMS,BaseReduzida,DescricaoComplementar,idUnidade,idSituacaoTributaria,DiscriminacaoProduto,IPI,BaseReduzidaIPI,ClassificacaoFiscal,ValorFrete,CFOP) " & _
''                                          "Values (" & rsNFe!idNFe & "," & intIdProduto & ",'" & Date & _
''                                          "','" & CStrValor(dblQuantidade) & "','" & CStrValor(dblDesconto) & "','" & CStrValor(dblValorUnitario) & "','" & _
''                                          CStrValor(dblICMSProduto) & "','" & CStrValor(dblBaseReduzidaICMS) & "','" & SQLCheck(strDescricaoComplementar) & "'," & intUnidade & "," & intSituacaoTributaria & ",'" & _
''                                          SQLCheck(strDiscriminacaoProduto) & "','" & CStrValor(dblIPIProduto) & "','" & CStrValor(dblBaseReduzidaIPI) & "','" & SQLCheck(strClassificacaoFiscal) & "','" & CStrValor(dblValorFrete) & "','" & strCFOPItem & "')"
''                     End If
''
''                     strnNF = ""
''                     strChaveNFe = ""
''                     strdEmi = ""
''                     strCNPJDest = ""
''                     strvNF = ""
''                  End If
''               End If
''            Next
''            Close #handle
''         End If
''      Next
''   End If
''
''   ' Visualiza o Retorno
''   If frmVisualiza.lvwDados.ListItems.Count > 0 Then
''      Screen.MousePointer = vbDefault
''      frmVisualiza.Show vbModal
''   End If
''
''End Sub
''
''
''Private Sub cmdValidar_Click()
''
'' ' Gerar XML
''   Open "C:\NF-e\Notas\Notas.TXT" For Output As #1
''   sAcao = 2  ' Validar
''   Notas
''   Close #1
''
''   sEmpresaNFe = LerArquivoINI("NFe", "Empresa", CaminhoINI & "\System.ini")
'''   Set rsGerarXML = cnSistema.Execute("Select * From NFe WHERE idNFe=" & Val(Mid(lvwNFes.SelectedItem.Key, 2, Len(lvwNFes.SelectedItem.Key))))
''   Set rsGerarXML = cnSistema.Execute("Select * From NFe WHERE Numero=" & Val(lvwNFes.ListItems(lvwNFes.SelectedItem.Index)))
''   If Not rsGerarXML.EOF Then
''      FileCopy "C:\NF-e\Notas\Notas.TXT", "C:\NF-e\" & sEmpresaNFe & "\Validar\" & StrZero(rsGerarXML!Numero, 6) & "_" & RemoveCaracteres(rsEmpresa!CNPJ_CPF) & "_001_" & Format(rsGerarXML!DataEmissao, "dd_mm_yyyy") & "-nfe.txt"
''      Kill "C:\NF-e\Notas\Notas.TXT"
''   End If
''
''   sAcao = 0
''   cmdTransmitir.Enabled = False
''   cmdValidar.Enabled = False
''   Carrega_View
''End Sub
''
''Private Function Verifica_Campos()
''Dim strMensagem As String
''Verifica_Campos = True
''
''   If Not IsDate(mskDataInicial.Text) Or Val(Mid(mskDataInicial.Text, 7, 4)) < 1900 Then strMensagem = strMensagem & "Data Inicial" & Chr(13)
''   If Not IsDate(mskDataFinal.Text) Or Val(Mid(mskDataFinal.Text, 7, 4)) < 1900 Then strMensagem = strMensagem & "Data Final" & Chr(13)
''
''   If Not strMensagem = Empty Then
''      MsgBox "Verifique os Seguintes Campos:" & Chr(13) & strMensagem, vbExclamation + vbOKOnly, "Campos Obrigat�rios"
''      Verifica_Campos = False
''      Exit Function
''   End If
''
''End Function
''Private Sub cmdTransmitir_Click()
''
''End Sub

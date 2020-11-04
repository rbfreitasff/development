VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmGerenciarNFs 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gerenciar Notas Fiscais Eletr�nicas"
   ClientHeight    =   8880
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12795
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   8880
   ScaleWidth      =   12795
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdTeste 
      Caption         =   "Testes"
      Height          =   375
      Left            =   180
      TabIndex        =   28
      Top             =   5100
      Width           =   915
   End
   Begin VB.Timer tmrVerificar 
      Interval        =   2000
      Left            =   10260
      Top             =   60
   End
   Begin VB.Timer tmrImprimirSistema 
      Interval        =   5000
      Left            =   10800
      Top             =   60
   End
   Begin VB.CommandButton cmdRetransmitir 
      Caption         =   "&Retransmitir"
      Height          =   375
      Left            =   60
      TabIndex        =   26
      Top             =   4620
      Width           =   1155
   End
   Begin VB.Timer tmrTransmitir 
      Interval        =   2000
      Left            =   11340
      Top             =   60
   End
   Begin VB.CommandButton cmdGerarQrCode 
      Caption         =   "Enviar"
      Height          =   375
      Left            =   6780
      TabIndex        =   24
      Top             =   5040
      Width           =   1155
   End
   Begin VB.CommandButton cmdImportarCupom 
      Caption         =   "Importar Cupom"
      Height          =   375
      Left            =   3840
      TabIndex        =   23
      Top             =   5040
      Width           =   1575
   End
   Begin VB.CommandButton cmdNFeCartaCorrecao 
      Caption         =   "Carta de Corre��o"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1680
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
      Begin VB.CommandButton cmdLimparHistorico 
         Caption         =   "Limpar"
         Height          =   375
         Left            =   8160
         TabIndex        =   27
         Top             =   2340
         Width           =   2115
      End
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
      Interval        =   5000
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
      ItemData        =   "frmGerenciarNFs.frx":0000
      Left            =   4800
      List            =   "frmGerenciarNFs.frx":0002
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
      Width           =   1155
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
   Begin VB.CommandButton cmdNFCeDigitacao 
      Caption         =   "Colocar nota em Digita��o"
      Height          =   375
      Left            =   10620
      TabIndex        =   12
      Top             =   4620
      Width           =   2115
   End
   Begin VB.Timer tmrAtualiza 
      Interval        =   10000
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
   Begin VB.CommandButton cmdImprimirDANFECe 
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
      Caption         =   "&Assinar"
      Height          =   375
      Left            =   5520
      TabIndex        =   8
      Top             =   5040
      Width           =   1155
   End
   Begin MSComctlLib.ListView lvwNFCes 
      Height          =   4080
      Left            =   60
      TabIndex        =   7
      Top             =   480
      Width           =   12645
      _ExtentX        =   22304
      _ExtentY        =   7197
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
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   25
      Top             =   8625
      Width           =   12795
      _ExtentX        =   22569
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   14826
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
            TextSave        =   "08/10/2019"
         EndProperty
      EndProperty
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
Attribute VB_Name = "frmGerenciarNFs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''Option Explicit
''Private SHA1Hash As New SHA1Hash
''
'''Public dValorTributos As Double
'''Public dValorTotalBC As Double
'''Public dValorTotalICMS As Double
'''Public dValorTotalBaseISS As Double
'''Public dValorTotalISS As Double
'''Public dValorTotalBCST As Double
'''Public dValorTotalICMSST As Double
'''Public dValorTotalProdutos As Double
'''Public dValorTotalFrete As Double
'''Public dValorTotalSeguro As Double
'''Public dValorTotalDesconto As Double
'''Public dValorTotalII As Double
'''Public dValorTotalIPI As Double
'''Public dValorTotalPIS As Double
'''Public dValorTotalCofins As Double
'''Public dValorTotalOutro As Double
'''Public dValorTotalNFCe As Double
'''Public sICMSAproveitamento As String
''
''''Dim dValorTributos As Double
''''Dim dValorTotalBC As Double
''''Dim dValorTotalICMS As Double
''''Dim dValorTotalProdutos As Double
''''Dim dValorTotalDesconto As Double
''''Dim dValorTotalFrete As Double
''''Dim dValorTotalNFCe As Double
''''Dim dValorTotalBaseISS As Double
''''Dim dValorTotalISS As Double
''
''
''
''Dim ItemList As ListItem
''Dim ProcuraItem As ListItem
''Dim rsDados As New ADODB.Recordset
''Dim rsNFCe As New ADODB.Recordset
''Dim rsNFCeItens As New ADODB.Recordset
''Dim rsNFCePagamentos As New ADODB.Recordset
''Dim rsNFCeBoletos As New ADODB.Recordset
''Dim rsTotalNFCe As New ADODB.Recordset
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
''Dim rsNFCeInutilizadas As New ADODB.Recordset
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
''Dim sEmpresaNFCe As String
''Dim sAcao As String
''Dim bRetorno As Boolean
''Dim IdMensagens As Integer
''Dim ArqNFCeRetornos As String
''Dim ArqNFCeErros As String
''Dim ArqNFCeEnviados As String
''Dim ArqNFCeTemp As String
''
''Dim sEmpresa As String
''
''Dim NFCeNumero As Long
''
'''''Dim sChaveNFe As String
'''''''Dim sDataEmissao As String
''''Dim sValorTotalNFCe As String
''''Dim sValorTotalICMSNFCe As String
''''Dim sCPFDestinatario As String
''
''Dim sPercMargAdICMSST As String
''Dim sUF As String
''Dim sChaveAcesso As String
''Dim sNaturezaOperacao As String
''Dim sFormaPagamento As String
''Dim sModelo As String
''Dim sSerie As String
''Dim sNumero As String
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
''
''
''Dim StatusNFCe As Integer
''
'''Private Sub cmdTeste_Click()
'''   frmSuites.Show vbModal
'''End Sub
''
''Private Sub Form_Load()
''
''   ' Conectar a base se for usado como modulo independente
''   Call ConnectDB
''
''   I_TituloForm = Me.Caption
''   On Error GoTo Erro
''   Status = 0
''   StatusNFCe = 1
''''   Centraliza frmGerenciarNFCes
''''   rsNFCe.Open "Select * from NFCe Order By Numero", cnSistema, adOpenForwardOnly, adLockOptimistic, 1
''   rsEmpresa.Open "Select * from Empresa", cnSistema, adOpenForwardOnly, adLockOptimistic, 1
''
''   sEmpresaNFCe = LerArquivoINI("NFe", "Empresa", CaminhoINI & "\System.ini")
''   ArqNFCeRetornos = I_UnidadeNFe & "NFC-e\" & sEmpresaNFCe & "\Retorno\"
''   ArqNFCeErros = I_UnidadeNFe & "NFC-e\" & sEmpresaNFCe & "\Erros\"
''   ArqNFCeEnviados = ""
''   ArqNFCeTemp = I_UnidadeNFe & "NFC-e\" & sEmpresaNFCe & "\Temp\"
''
''   IdMensagens = 1 ' Inicia o contador de chave das mensagens
''
''   lvwNFCes.ColumnHeaders.Add , , "N�mero", 850
''   lvwNFCes.ColumnHeaders.Add , , "Emiss�o", 1050
''   lvwNFCes.ColumnHeaders.Add , , "Cliente", 4000
''   lvwNFCes.ColumnHeaders.Add , , "Valor Total", 1050, lvwColumnRight
''   lvwNFCes.ColumnHeaders.Add , , "Situa��o", 1700
''   lvwNFCes.ColumnHeaders.Add , , "", 1650
''
''   lvwMensagens.ColumnHeaders.Add , , "Chave", 0
''   lvwMensagens.ColumnHeaders.Add , , "N�mero", 850
''   lvwMensagens.ColumnHeaders.Add , , "Mensagem", 11000
''   lvwMensagens.ColumnHeaders.Add , , "Arquivo", 0
''
''''   mskDataInicial.text = CDate("01" & Mid(Date, 3, 8))
''   mskDataInicial.text = Date - 1
''   mskDataFinal.text = Date
''   cmbSituacao.ListIndex = 0
''
''   Carrega_View
''
''   lvwMensagens.ListItems.Clear
''
''   sEmpresaNFCe = LerArquivoINI("NFe", "Empresa", CaminhoINI & "\System.ini")
''   If Trim(sEmpresaNFCe) <> "" Then
''      If Dir("C:\UNIMAKE\" & sEmpresaNFCe & "\UNINFE.EXE") <> "" Then
''         Shell "C:\UNIMAKE\" & sEmpresaNFCe & "\UNINFE.EXE"
''      End If
''   End If
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
''   If Not rsErro Then rsNFCe.Close
''   If Not rsErro Then rsEmpresa.Close
''   KillProcess "uninfe.exe"
''End Sub
''
''Private Sub cmdPesquisar_Click()
''   Carrega_View
''End Sub
''
''Private Sub Carrega_View()
''Dim sSituacao As String
''Dim Contador As Integer
''
''   cmdCancelarNota.Enabled = False
'''   cmdTransmitir.Enabled = False
'''   cmdImprimirDANFECe.Enabled = False
''
''   If Not IsDate(mskDataInicial.text) Or Not IsDate(mskDataFinal.text) Then
''      Beep
''      MsgBox "Datas inv�lidas", vbExclamation, "Erro"
''      Exit Sub
''   End If
''
''   ' Notas
''   Dim sqlSituacao As String
''   Select Case cmbSituacao.ListIndex
''         Case 0
''            sqlSituacao = "AND Situacao <> 2"
''         Case 1
''            sqlSituacao = " AND Situacao = 0"
''         Case 2
''            sqlSituacao = " AND Situacao = 1"
''         Case 3
''            sqlSituacao = " AND Situacao = 2"
''         Case 4
''            sqlSituacao = " AND Situacao = 3"
''         Case 5
''            sqlSituacao = " AND Situacao = 4"
''   End Select
''
''   If (IsDate(mskDataInicial.text) And IsDate(mskDataFinal.text)) And (CDate(mskDataFinal.text) >= CDate(mskDataInicial.text)) Then
''      Set rsNFCe = cnSistema.Execute("SELECT * FROM NFCe WHERE DataEmissao >= cDate('" & Format(mskDataInicial.text, "dd/mm/yyyy") & "') AND DataEmissao <= cDate('" & Format(mskDataFinal.text, "dd/mm/yyyy") & "')" & sqlSituacao & " Order By Numero")
''   Else
''      Set rsNFCe = cnSistema.Execute("SELECT * FROM NFCe WHERE DataEmissao >= cDate('" & Format(Date, "dd/mm/yyyy") & "') AND DataEmissao <= cDate('" & Format(Date, "dd/mm/yyyy") & "')" & sqlSituacao & " Order By Numero")
''   End If
''
''   Contador = 1
''   lvwNFCes.ListItems.Clear
''   If Not rsNFCe.EOF Then
''      Do While Not rsNFCe.EOF
''         Set rsClientes = cnSistema.Execute("SELECT * FROM Clientes WHERE idCliente = " & rsNFCe!idCliente)
''         Select Case rsNFCe!Situacao
''                Case 0
''                     sSituacao = "Em Digita��o"
''                Case 1
''                     sSituacao = "Processamento"
''                Case 2
''                     sSituacao = "Aprovada"
''                Case 3
''                     sSituacao = "Cancelada"
''                Case 4
''                     sSituacao = "N�o Emitida"
''         End Select
''
''         If Not rsClientes.EOF Then
''            Set rsTotalNFCe = cnSistema.Execute("Select * From TotalNFCe Where Numero = " & rsNFCe!Numero)
''
'''            Set ItemList = lvwNFCes.ListItems.Add(, "R" & CStr(rsNFCe!idNFCe), StrZero(rsNFCe!Numero, 8))
''
''            Set ProcuraItem = lvwNFCes.FindItem(StrZero(rsNFCe!Numero, 8))
''            If ProcuraItem Is Nothing Then
''               Set ItemList = lvwNFCes.ListItems.Add(, "R" & CStr(Contador), StrZero(rsNFCe!Numero, 8))
''               ItemList.SubItems(1) = rsNFCe!DataEmissao
''               ItemList.SubItems(2) = Trim(rsClientes!Nome)
''               If Not rsTotalNFCe.EOF Then
''                  ItemList.SubItems(3) = Format(rsTotalNFCe!Total + IIf(Not IsNull(rsTotalNFCe!TotalFrete), rsTotalNFCe!TotalFrete, 0), "##,##0.00")
''               Else
''                  ItemList.SubItems(3) = Format(0, "##,##0.00")
''               End If
''               ItemList.SubItems(4) = sSituacao
''               If rsNFCe!Situacao = 1 Then
''                  ItemList.SubItems(5) = "Aguarde..."
''               End If
''            End If
''         End If
''
''         Contador = Contador + 1
''         rsNFCe.MoveNext
''      Loop
''
''      rsNFCe.MoveFirst
''   End If
''
''   ' Notas
''   Set rsNFCeInutilizadas = cnSistema.Execute("SELECT * FROM NFCeInutilizadas WHERE Data >= cDate('" & Format(mskDataInicial.text, "dd/mm/yyyy") & "') AND Data <= cDate('" & Format(mskDataFinal.text, "dd/mm/yyyy") & "') Order By Numero")
''
''   If Not rsNFCeInutilizadas.EOF Then
''      Do While Not rsNFCeInutilizadas.EOF
''
'''         Set ItemList = lvwNFCes.ListItems.Add(, "I" & CStr(rsNFCeInutilizadas!Numero), StrZero(rsNFCeInutilizadas!Numero, 8))
''
''         Set ProcuraItem = lvwNFCes.FindItem(StrZero(rsNFCeInutilizadas!Numero, 8))
''         If ProcuraItem Is Nothing Then
''            Set ItemList = lvwNFCes.ListItems.Add(, "I" & CStr(Contador), StrZero(rsNFCeInutilizadas!Numero, 8))
''            ItemList.SubItems(1) = rsNFCeInutilizadas!Data
''            ItemList.SubItems(2) = "Nota Inutilizada"
''            ItemList.SubItems(3) = Format(0, "##,##0.00")
''            ItemList.SubItems(4) = "Inutilizada"
''            If IsNull(rsNFCeInutilizadas!Protocolo) Or rsNFCeInutilizadas!Protocolo = "" Then
''               ItemList.SubItems(5) = "Aguarde..."
''            Else
''               ItemList.SubItems(5) = ""
''            End If
''         End If
''
''         Contador = Contador + 1
''         rsNFCeInutilizadas.MoveNext
''      Loop
''   End If
''
''End Sub
''
'''''Private Sub cmdTransmitir_Click()
'''''
'''''   ' Gerar XML
'''''   Open I_UnidadeNFe & "NFC-e\Notas\Notas.TXT" For Output As #1
'''''   sAcao = 1  ' Transmitir
'''''''   Notas
'''''   NFCeNumero = Val(lvwNFCes.ListItems(lvwNFCes.SelectedItem.Index))
'''''
'''''   NotasNFs
'''''   Close #1
'''''
'''''   ' Validar e Assinar o XML
'''''   sEmpresaNFCe = LerArquivoINI("NFe", "Empresa", CaminhoINI & "\System.ini")
'''''   Set rsGerarXML = cnSistema.Execute("Select * From NFCe WHERE Numero=" & NFCeNumero)
'''''   If Not rsGerarXML.EOF Then
'''''      FileCopy I_UnidadeNFe & "NFC-e\Notas\Notas.TXT", I_UnidadeNFe & "NFC-e\" & sEmpresaNFCe & "\Validar\" & sChaveNFe & "-nfe.XML"
''''''      FileCopy I_Unidadenfe & "NFC-e\Notas\Notas.TXT", I_Unidadenfe & "NFC-e\" & sEmpresaNFCe & "\Envio\" & sChaveNFe & "-nfe.XML"
'''''      Kill I_UnidadeNFe & "NFC-e\Notas\Notas.TXT"
'''''   End If
'''''
'''''   sAcao = 0
''''''   cmdTransmitir.Enabled = False
'''''   cmdValidar.Enabled = False
'''''   Carrega_View
'''''End Sub
''
''Private Sub lvwNFCes_Click()
''   If lvwNFCes.ListItems.Count <> 0 Then
''      Set rsGerarXML = cnSistema.Execute("Select * From NFCe WHERE Numero=" & Val(lvwNFCes.ListItems(lvwNFCes.SelectedItem.Index)))
''      If Not rsGerarXML.EOF Then
''         If rsGerarXML!Situacao = 0 Then
'''            cmdTransmitir.Enabled = True
''            cmdValidar.Enabled = True
''         Else
'''            cmdTransmitir.Enabled = False
''            cmdValidar.Enabled = False
''         End If
''
''         If rsGerarXML!Situacao = 2 Then
''            If Trim(rsGerarXML!ChaveNFCe) <> "" And Trim(rsGerarXML!Protocolo) <> "" Then
''               cmdRetransmitir.Enabled = False
''               cmdCancelarNota.Enabled = True
''               cmdImprimirDANFECe.Enabled = True
''            Else
''               cmdRetransmitir.Enabled = True
''               cmdCancelarNota.Enabled = False
''               cmdImprimirDANFECe.Enabled = False
''            End If
''         End If
''      End If
''   End If
''End Sub
''
''Private Sub lvwNFCes_DblClick()
''Dim NFCeNumeroRet As String
''
''   NFCeNumeroRet = Val(lvwNFCes.ListItems(lvwNFCes.SelectedItem.Index))
''
''   cnSistema.Execute "Update NFCe set " & _
''            "Situacao = 0, " & _
''            "TentativaEmissao = 0, " & _
''            "DataEmissao = '" & Date & "', " & _
''            "DataVencimento = '" & Date & "', " & _
''            "DataCaixa = '" & Date & "', " & _
''            "Hora = '" & Time & "' " & _
''            "Where Numero = " & NFCeNumeroRet
''
''End Sub
''
''Private Sub tmrAtualiza_Timer()
''Dim handle As Integer
''Dim Linha As String
''Dim strMensagem As String
''Dim oNFCe310 As New CNFCE310
''Dim nProtocolo As String
''
''   bRetorno = False  '' Verifica se houve algum para recarregar a view
''
''   If Not IsDate(mskDataInicial.text) Or Not IsDate(mskDataFinal.text) Then
''      Beep
''      MsgBox "Datas Inv�lidas", vbExclamation, "Erro"
''      Exit Sub
''   End If
''
''   If (IsDate(mskDataInicial.text) And IsDate(mskDataFinal.text)) And (CDate(mskDataFinal.text) >= CDate(mskDataInicial.text)) Then
''      Set rsNFCe = cnSistema.Execute("SELECT * FROM NFCe WHERE DataEmissao >= cDate('" & Format(mskDataInicial.text, "dd/mm/yyyy") & "') AND DataEmissao <= cDate('" & Format(mskDataFinal.text, "dd/mm/yyyy") & "') Order By Numero")
''   Else
''      Set rsNFCe = cnSistema.Execute("SELECT * FROM NFCe WHERE DataEmissao >= cDate('" & Format(Date, "dd/mm/yyyy") & "') AND DataEmissao <= cDate('" & Format(Date, "dd/mm/yyyy") & "') Order By Numero")
''   End If
''
''   Do While Not rsNFCe.EOF
''      ' Verificar se XML foi atualizado
''      Call oNFCe310.ConverterTXT_XML(rsNFCe!idNFCe, rsNFCe!Numero, rsNFCe!DataEmissao, rsNFCe!ChaveNFCe)
''
''      ' Verificar se XML foi atualizado
''      Call oNFCe310.FVerificaAprovacaoXML(rsNFCe!idNFCe, rsNFCe!Numero, rsNFCe!DataEmissao, rsNFCe!ChaveNFCe, rsNFCe!Situacao)
''
''      ' Verifica se XML de Cancelamento n�o possui erro
''      Call oNFCe310.FVerificaErroCancelamento(rsNFCe!Numero)
''
''      ' Atualiza Protocolo
''      If (IsNull(rsNFCe!Protocolo) Or Trim(rsNFCe!Protocolo) = "") Then
''         Call oNFCe310.FAtualizarProtocolo(rsNFCe!idNFCe, rsNFCe!DataEmissao, rsNFCe!ChaveNFCe)
''      End If
''
''      ' Verifica se cancelamento foi autorizado
''      Call oNFCe310.FVerificaAutorizacaoCancelamento(rsNFCe!idNFCe, rsNFCe!DataEmissao, rsNFCe!ChaveNFCe, rsNFCe!Situacao)
''
''      rsNFCe.MoveNext
''   Loop
''
'' ' Notas Inutilizadas
''   Dim sChaveNFCe As String
''   sEmpresaNFCe = LerArquivoINI("NFe", "Empresa", CaminhoINI & "\System.ini")
''
''   Set rsNFCeInutilizadas = cnSistema.Execute("SELECT * FROM NFCeInutilizadas")
''   Do While Not rsNFCeInutilizadas.EOF
''      ' Envia solicita��o de Inutiliza��o
''      If IsNull(rsNFCeInutilizadas!ChaveNFCe) Or Trim(rsNFCeInutilizadas!ChaveNFCe) = "" Then
''         If Dir(ArqNFCeRetornos & StrZero(Val(rsNFCeInutilizadas!Numero), 12) & "-ret-gerar-chave.txt") <> "" Then
''            handle = FreeFile
''            Open ArqNFCeRetornos & StrZero(Val(rsNFCeInutilizadas!Numero), 12) & "-ret-gerar-chave.txt" For Input As #handle
''            Line Input #handle, Linha
''            sChaveNFCe = Linha
''            Close #handle
''
''            Set rsUFs = cnSistema.Execute("Select * from UFs Where idUF = " & rsEmpresa!idUF)
''
''            Open I_UnidadeNFe & "NFC-e\Notas\Notas.TXT" For Output As #1
''            Print #1, "tbAmb|" & LerArquivoINI("NFe", "Ambiente", CaminhoINI & "\System.ini")     ' 1 - Produ��o ou 2 - Homologa��o
''            Print #1, "cUF|" & rsUFs!Codigo
''            Print #1, "ano|" & Format(Date, "yy")
''            Print #1, "CNPJ|" & RemoveCaracteres(rsEmpresa!CNPJ_CPF)
''            Print #1, "mod|55"
''            Print #1, "serie|1"
''            Print #1, "nNFIni|" & rsNFCeInutilizadas!Numero
''            Print #1, "nNFFin|" & rsNFCeInutilizadas!Numero
''            Print #1, "xJust|Numero inutilizado por escolha do usuario"
''
''            Close #1
''
''            sEmpresaNFCe = LerArquivoINI("NFe", "Empresa", CaminhoINI & "\System.ini")
''            FileCopy I_UnidadeNFe & "NFC-e\Notas\Notas.TXT", I_UnidadeNFe & "NFC-e\" & sEmpresaNFCe & "\Envio\" & Trim(sChaveNFCe) & "-ped-inu.txt"
''            Kill I_UnidadeNFe & "NFC-e\Notas\Notas.TXT"
''
''            cnSistema.Execute "Update NFCeInutilizadas set " & _
''                     "ChaveNFCe = '" & Trim(sChaveNFCe) & "' " & _
''                     "Where idNFCe = " & rsNFCeInutilizadas!idNFCe
''
''         End If
''      End If
''
''      ' Verifica se Inutiliza��o foi aceita
''      If IsNull(rsNFCeInutilizadas!Protocolo) Or Trim(rsNFCeInutilizadas!Protocolo) = "" Then
''         sEmpresaNFCe = LerArquivoINI("NFe", "Empresa", CaminhoINI & "\System.ini")
''         If Dir(ArqNFCeRetornos & Trim(rsNFCeInutilizadas!ChaveNFCe) & "-inu.XML") <> "" Then
''            handle = FreeFile
''            Open ArqNFCeRetornos & Trim(rsNFCeInutilizadas!ChaveNFCe) & "-inu.XML" For Input As #handle
''
''            Line Input #handle, Linha
''
''            nProtocolo = PesquisarTAG(Linha, "nProt")
''
''            If Trim(nProtocolo) <> "" Then
''               cnSistema.Execute "Update NFCeInutilizadas set " & _
''                        "Protocolo = '" & nProtocolo & "' " & _
''                        "Where idNFCe = " & rsNFCeInutilizadas!idNFCe
''            End If
''
''            Close #handle
''         End If
''      End If
''
''      rsNFCeInutilizadas.MoveNext
''   Loop
''
''   ' Limpar Hist�rico
''   lvwMensagens.ListItems.Clear
'''   mskDataInicial.text = Date - 1
'''   mskDataFinal.text = Date
''
''   ' Verifica outros retornos
''   Call Verifica_Retornos
''
''   Carrega_View
''End Sub
''
''Private Sub cmdNFCeDigitacao_Click()
''   If MsgBox("Confirma colocar nota em digita��o", vbYesNo + vbQuestion, "Confirma��o") = vbYes Then
''      Set rsNFCe = cnSistema.Execute("Select * From NFCe WHERE Numero=" & Val(lvwNFCes.ListItems(lvwNFCes.SelectedItem.Index)))
''      If Not rsNFCe.EOF Then
''         If rsNFCe!Situacao = 2 Or rsNFCe!Situacao = 3 Then
''            MsgBox "Nota Fiscal j� autorizada n�o pode ser alterada", vbExclamation + vbOKOnly, "Aten��o"
''         Else
''            cnSistema.Execute "Update NFCe set " & _
''                     "Situacao = 0 " & _
''                     "Where idNFCe = " & rsNFCe!idNFCe
''         End If
''         Carrega_View
''      End If
''   End If
''End Sub
''
''Private Sub cmdCancelarNota_Click()
''Dim oNFCe310 As New CNFCE310
''Dim iNumeroCancelar As Integer
''Dim strMensagem As String
''Dim sMotivo
''
''   tmrAtualiza.Enabled = False
''   iNumeroCancelar = Val(lvwNFCes.ListItems(lvwNFCes.SelectedItem.Index))
''   If MsgBox("Confirma cancelamento da Nota N� " & iNumeroCancelar, vbYesNo + vbQuestion, "Confirma��o") = vbYes Then
''      Set rsNFCe = cnSistema.Execute("Select * From NFCe WHERE Numero=" & iNumeroCancelar)
''      If Not rsNFCe.EOF Then
''         sMotivo = InputBox("Digite o motivo do cancelamento", "Cancelamento", "")
''         If Len(Trim(sMotivo)) < 15 Then
''            MsgBox "Obrigat�rio digitar ao menos 15 caracteres", vbExclamation + vbOKOnly, "Campos Obrigat�rios"
''            Exit Sub
''         End If
''
''         If Trim(sMotivo) <> "" Then
''            strMensagem = oNFCe310.FCancelarNota(iNumeroCancelar, sMotivo)
''            MsgBox strMensagem, vbExclamation + vbOKOnly, "Cancelamento"
''         Else
''            MsgBox "O motivo � obrigat�rio", vbExclamation + vbOKOnly, "Campos Obrigat�rios"
''         End If
''      End If
''   End If
''   tmrAtualiza.Enabled = True
''
''
''
''
'''''   tmrAtualiza.Enabled = False
'''''   If MsgBox("Confirma cancelamento da Nota N� " & lvwNFCes.ListItems(lvwNFCes.SelectedItem.Index), vbYesNo + vbQuestion, "Confirma��o") = vbYes Then
'''''      Set rsNFCe = cnSistema.Execute("Select * From NFCe WHERE Numero=" & Val(lvwNFCes.ListItems(lvwNFCes.SelectedItem.Index)))
'''''      If Not rsNFCe.EOF Then
'''''         sMotivo = InputBox("Digite o motivo do cancelamento", "Cancelamento", "")
'''''         If Trim(sMotivo) <> "" Then
'''''            Set rsEmpresa = cnSistema.Execute("SELECT * FROM Empresa")
'''''            Set rsUFs = cnSistema.Execute("Select * From UFs WHERE idUF=" & rsEmpresa!idUF)
'''''
'''''            Open I_UnidadeNFe & "NFC-e\Notas\Notas.TXT" For Output As #1
'''''            Print #1, "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "UTF-8" & Chr(34) & "?>"
'''''            Print #1, "<envEvento xmlns=" & Chr(34) & "http://www.portalfiscal.inf.br/NFe" & Chr(34) & " versao=" & Chr(34) & "1.00" & Chr(34) & ">"
'''''            Print #1, "<idLote>1</idLote>"
'''''            Print #1, "<evento xmlns=" & Chr(34) & "http://www.portalfiscal.inf.br/NFe" & Chr(34) & " versao=" & Chr(34) & "1.00" & Chr(34) & ">"
'''''            Print #1, "<iNFCevento Id=" & Chr(34) & "ID110111" & rsNFCe!ChaveNFCe & "01" & Chr(34) & ">"
'''''            Print #1, "<cOrgao>" & rsUFs!Codigo & "</cOrgao>"
'''''            Print #1, "<tpAmb>" & LerArquivoINI("NFe", "Ambiente", CaminhoINI & "\System.ini") & "</tpAmb>"   ' 1 - Produ��o ou 2 - Homologa��o
'''''            Print #1, "<CNPJ>" & RemoveCaracteres(rsEmpresa!CNPJ_CPF) & "</CNPJ>"
'''''            Print #1, "<chNFCe>" & rsNFCe!ChaveNFCe & "</chNFCe>"
'''''
'''''            If LerArquivoINI("NFe", "HorarioVerao", App.Path & "\System.ini") Then
'''''               Print #1, "<dhEvento>" & Format(Date, "YYYY-MM-DD") & "T" & Format(Time, "HH:MM:SS") & "-02:00" & "</dhEvento>"
'''''            Else
'''''               Print #1, "<dhEvento>" & Format(Date, "YYYY-MM-DD") & "T" & Format(Time, "HH:MM:SS") & "-03:00" & "</dhEvento>"
'''''            End If
'''''
'''''            Print #1, "<tpEvento>110111</tpEvento>"
'''''            Print #1, "<nSeqEvento>1</nSeqEvento>"
'''''            Print #1, "<verEvento>1.00</verEvento>"
'''''            Print #1, "<detEvento versao=" & Chr(34) & "1.00" & Chr(34) & ">"
'''''            Print #1, "<descEvento>Cancelamento</descEvento>"
'''''            Print #1, "<nProt>" & rsNFCe!Protocolo & "</nProt>"
'''''            Print #1, "<xJust>" & sMotivo & "</xJust>"
'''''            Print #1, "</detEvento>"
'''''            Print #1, "</iNFCevento>"
'''''            Print #1, "</evento>"
'''''            Print #1, "</envEvento>"
'''''            Close #1
'''''
'''''            sEmpresaNFCe = LerArquivoINI("NFe", "Empresa", CaminhoINI & "\System.ini")
'''''            FileCopy I_UnidadeNFe & "NFC-e\Notas\Notas.TXT", I_UnidadeNFe & "NFC-e\" & sEmpresaNFCe & "\Envio\" & Trim(rsNFCe!ChaveNFCe) & "-env-canc.xml"
'''''            Kill I_UnidadeNFe & "NFC-e\Notas\Notas.TXT"
'''''         Else
'''''            MsgBox "O motivo � obrigat�rio", vbExclamation + vbOKOnly, "Campos Obrigat�rios"
'''''         End If
'''''      End If
'''''   End If
'''''   tmrAtualiza.Enabled = True
''
''End Sub
''
''Private Sub cmdStatusServico_Click()
''Dim sMotivo
''
''   If MsgBox("Consulta Status do Servi�o", vbYesNo + vbQuestion, "Confirma��o") = vbYes Then
''      If Not rsNFCe.EOF Then
''         Set rsTemp = cnSistema.Execute("SELECT * FROM UFs WHERE idUF = " & rsEmpresa!idUF)
''
''         Open I_UnidadeNFe & "NFC-e\Notas\Notas.TXT" For Output As #1
''         Print #1, "tbEmis|1"                                                             ' 1 - Normal
''         Print #1, "tpAmb|" & LerArquivoINI("NFe", "Ambiente", CaminhoINI & "\System.ini") ' 1 - Produ��o ou 2 - Homologa��o
''         Print #1, "cUF|" & rsTemp!Codigo
''         Close #1
''
''         sEmpresaNFCe = LerArquivoINI("NFe", "Empresa", CaminhoINI & "\System.ini")
''         FileCopy I_UnidadeNFe & "NFC-e\Notas\Notas.TXT", I_UnidadeNFe & "NFC-e\" & sEmpresaNFCe & "\Envio\" & Format(Date, "yyyymmdd") & "T" & Format(Time, "hhmmss") & "-ped-sta.txt"
''         Kill I_UnidadeNFe & "NFC-e\Notas\Notas.TXT"
''      End If
''   End If
''
''End Sub
''
''Private Sub cmdInutilizarNumeracao_Click()
''Dim oNFCe310 As New CNFCE310
''Dim strMensagem As String
''Dim sNumeroInutilizar
''
''   tmrAtualiza.Enabled = False
''   If MsgBox("Confirma Inutilizar N�mero da Nota", vbYesNo + vbQuestion, "Confirma��o") = vbYes Then
''      sNumeroInutilizar = InputBox("Digite o n�mero a inutilizar", "Inutilizar numera��o", "")
''
''      strMensagem = oNFCe310.FInutilizarNumero(sNumeroInutilizar)
''      MsgBox strMensagem, vbExclamation + vbOKOnly, "Inutiliza��o"
''
''   End If
''   tmrAtualiza.Enabled = True
''
''End Sub
''
''Private Sub Verifica_Retornos()
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
''   sEmpresaNFCe = LerArquivoINI("NFe", "Empresa", CaminhoINI & "\System.ini")
''
''   Dim Arquivos() As String
''   Dim lCtr As Long
''   Arquivos = ListarArquivos(I_UnidadeNFe & "NFC-e\" & sEmpresaNFCe & "\Retorno")
'''   If UBound(Arquivos) > 0 And Len(Trim(Arquivos(lCtr))) <= 25 Then
''''   If UBound(Arquivos) > 0 Then
''   If Arquivos(lCtr) <> "" Then
''      For lCtr = 0 To UBound(Arquivos)
''         If UCase(Mid(Arquivos(lCtr), Len(Trim(Arquivos(lCtr))) - 7, 8)) = "-INU.XML" Then
''            Contador = 1
''
''            ' Verifica se Inutiliza��o foi aceita
''            handle = FreeFile
''            Open ArqNFCeRetornos & Arquivos(lCtr) For Input As #handle
''            'Open ArqNFCeRetornos & Trim(rsNFCeInutilizadas!ChaveNFCe) & "-inu.XML" For Input As #handle
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
''               cnSistema.Execute "Update NFCeInutilizadas set " & _
''                        "Protocolo = '" & sProtocolo & "' " & _
''                        "Where Numero = " & sNumeroNota
''
''               sStatus = ""
''               sProtocolo = ""
''            End If
''
''            Close #handle
''
''            FileCopy ArqNFCeRetornos & Arquivos(lCtr), ArqNFCeTemp & Arquivos(lCtr)
''            Kill ArqNFCeRetornos & Arquivos(lCtr)
''
''         ElseIf UCase(Mid(Arquivos(lCtr), Len(Trim(Arquivos(lCtr))) - 3, 4)) = ".XML" Then
''            Contador = 1
''
''            handle = FreeFile
''            Open ArqNFCeRetornos & Arquivos(lCtr) For Input As #handle
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
''            FileCopy ArqNFCeRetornos & Arquivos(lCtr), ArqNFCeTemp & Arquivos(lCtr)
''            Kill ArqNFCeRetornos & Arquivos(lCtr)
''
''         ElseIf UCase(Mid(Arquivos(lCtr), Len(Trim(Arquivos(lCtr))) - 3, 4)) = ".ERR" Then
''
''            handle = FreeFile
''            Open ArqNFCeRetornos & Arquivos(lCtr) For Input As #handle
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
''            FileCopy ArqNFCeRetornos & Arquivos(lCtr), ArqNFCeTemp & Arquivos(lCtr)
''            Kill ArqNFCeRetornos & Arquivos(lCtr)
''
''         ElseIf UCase(Mid(Arquivos(lCtr), Len(Trim(Arquivos(lCtr))) - 2, 3)) = ".TXT" Then
''            FileCopy ArqNFCeRetornos & Arquivos(lCtr), ArqNFCeTemp & Arquivos(lCtr)
''            Kill ArqNFCeRetornos & Arquivos(lCtr)
''
''         End If
''      Next
''   End If
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
''   sEmpresaNFCe = LerArquivoINI("NFe", "Empresa", CaminhoINI & "\System.ini")
''   If Dir(ArqNFCeRetornos & StrZero(rsNFCe!Numero, 6) & "_" & RemoveCaracteres(rsEmpresa!CNPJ_CPF) & "_001_" & Format(rsNFCe!DataEmissao, "dd_mm_yyyy") & "-NFCe.txt") <> "" Then
''      ' Atualizar Chave da NFCe
''      handle = FreeFile
''      Open ArqNFCeRetornos & StrZero(rsNFCe!Numero, 6) & "_" & RemoveCaracteres(rsEmpresa!CNPJ_CPF) & "_001_" & Format(rsNFCe!DataEmissao, "dd_mm_yyyy") & "-NFCe.txt" For Input As #handle
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
''               If Val(Mid(Linha, 14, 9)) = rsNFCe!Numero Then
''                  cnSistema.Execute "Update NFCe set " & _
''                           "ChaveNFCe = '" & Mid(Linha, 47, 44) & "' " & _
''                           "Where idNFCe = " & rsNFCe!idNFCe
''
''               End If
''            End If
''         End If
''      Wend
''      Close #handle
''      If Dir(ArqNFCeRetornos & StrZero(rsNFCe!Numero, 6) & "_" & RemoveCaracteres(rsEmpresa!CNPJ_CPF) & "_001_" & Format(rsNFCe!DataEmissao, "dd_mm_yyyy") & "-NFCe.txt") <> "" Then FileCopy ArqNFCeRetornos & StrZero(rsNFCe!Numero, 6) & "_" & RemoveCaracteres(rsEmpresa!CNPJ_CPF) & "_001_" & Format(rsNFCe!DataEmissao, "dd_mm_yyyy") & "-NFCe.txt", ArqNFCeTemp & StrZero(rsNFCe!Numero, 6) & "_" & RemoveCaracteres(rsEmpresa!CNPJ_CPF) & "_001_" & Format(rsNFCe!DataEmissao, "dd_mm_yyyy") & "-NFCe.txt"
''      If Dir(ArqNFCeRetornos & StrZero(rsNFCe!Numero, 6) & "_" & RemoveCaracteres(rsEmpresa!CNPJ_CPF) & "_001_" & Format(rsNFCe!DataEmissao, "dd_mm_yyyy") & "-NFCe-orig.txt") <> "" Then FileCopy ArqNFCeRetornos & StrZero(rsNFCe!Numero, 6) & "_" & RemoveCaracteres(rsEmpresa!CNPJ_CPF) & "_001_" & Format(rsNFCe!DataEmissao, "dd_mm_yyyy") & "-NFCe-orig.txt", ArqNFCeTemp & StrZero(rsNFCe!Numero, 6) & "_" & RemoveCaracteres(rsEmpresa!CNPJ_CPF) & "_001_" & Format(rsNFCe!DataEmissao, "dd_mm_yyyy") & "-NFCe-orig.txt"
''      If Dir(ArqNFCeRetornos & StrZero(rsNFCe!Numero, 6) & "_" & RemoveCaracteres(rsEmpresa!CNPJ_CPF) & "_001_" & Format(rsNFCe!DataEmissao, "dd_mm_yyyy") & "-NFCe.txt") <> "" Then Kill ArqNFCeRetornos & StrZero(rsNFCe!Numero, 6) & "_" & RemoveCaracteres(rsEmpresa!CNPJ_CPF) & "_001_" & Format(rsNFCe!DataEmissao, "dd_mm_yyyy") & "-NFCe.txt"
''      If Dir(ArqNFCeRetornos & StrZero(rsNFCe!Numero, 6) & "_" & RemoveCaracteres(rsEmpresa!CNPJ_CPF) & "_001_" & Format(rsNFCe!DataEmissao, "dd_mm_yyyy") & "-NFCe.txt") <> "" Then Kill ArqNFCeRetornos & StrZero(rsNFCe!Numero, 6) & "_" & RemoveCaracteres(rsEmpresa!CNPJ_CPF) & "_001_" & Format(rsNFCe!DataEmissao, "dd_mm_yyyy") & "-NFCe-orig.txt"
''   End If
''
''   ' Verifica se XML teve erro de convers�o
''   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''   sEmpresaNFCe = LerArquivoINI("NFe", "Empresa", CaminhoINI & "\System.ini")
''   If Dir(ArqNFCeRetornos & StrZero(rsNFCe!Numero, 6) & "_" & RemoveCaracteres(rsEmpresa!CNPJ_CPF) & "_001_" & Format(rsNFCe!DataEmissao, "dd_mm_yyyy") & "-NFCe.err") <> "" Then
''      ' Atualizar Chave da NFCe
''      handle = FreeFile
''      Open ArqNFCeRetornos & StrZero(rsNFCe!Numero, 6) & "_" & RemoveCaracteres(rsEmpresa!CNPJ_CPF) & "_001_" & Format(rsNFCe!DataEmissao, "dd_mm_yyyy") & "-NFCe.err" For Input As #handle
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
''      cnSistema.Execute "Update NFCe set " & _
''               "Situacao = 0 " & _
''               "Where idNFCe = " & rsNFCe!idNFCe
''
''      MsgBox strMensagem, vbExclamation + vbOKOnly, "Erro de envio da NFCe"
''      Close #handle
''
''      Kill ArqNFCeRetornos & StrZero(rsNFCe!Numero, 6) & "_" & RemoveCaracteres(rsEmpresa!CNPJ_CPF) & "_001_" & Format(rsNFCe!DataEmissao, "dd_mm_yyyy") & "-NFCe.err"
''   End If
''
''   ' Verifica se XML n�o possui erros
''   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''   sEmpresaNFCe = LerArquivoINI("NFe", "Empresa", CaminhoINI & "\System.ini")
''   If Dir(ArqNFCeRetornos & rsNFCe!ChaveNFCe & "-NFCe.err") <> "" Then
''      ' Atualizar Chave da NFCe
''      handle = FreeFile
''      Open ArqNFCeRetornos & rsNFCe!ChaveNFCe & "-NFCe.err" For Input As #handle
''
''      bRetorno = False
''      While Not EOF(handle)
''         Line Input #handle, Linha
''
''         strMensagem = strMensagem & Linha & Chr(13)
''      Wend
''
''      cnSistema.Execute "Update NFCe set " & _
''               "Situacao = 0 " & _
''               "Where idNFCe = " & rsNFCe!idNFCe
''
''      MsgBox strMensagem, vbExclamation + vbOKOnly, "Erro de envio da NFCe"
''      Close #handle
''
''      Kill ArqNFCeRetornos & rsNFCe!ChaveNFCe & "-NFCe.err"
''   End If
''End Sub
''
''Private Sub cmdFecharNFCe_Click()
''   KillProcess "uniNFCe.exe"
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
''   Arquivos = ListarArquivos(ArqNFCeRetornos)
'''''''   Arquivos = ListarArquivos(I_Unidadenfe & "NFC-e\" & sEmpresaNFCe & "\Retorno")
'''''   Arquivos = ListarArquivos(I_Unidadenfe & "NFC-e\" & sEmpresaNFCe & "\Temp")
'''   Arquivos = ListarArquivos(I_Unidadenfe & "XML\NFCe 2 - Modelos XML de Retorno")
''''   If UBound(Arquivos) > 0 Then
''   If Arquivos(lCtr) <> "" Then
''      For lCtr = 0 To UBound(Arquivos)
''          If UCase(Mid(Arquivos(lCtr), Len(Trim(Arquivos(lCtr))) - 3, 4)) = ".XML" Then
'''          If UCase(Mid(Arquivos(lCtr), Len(Trim(Arquivos(lCtr))) - 7, 8)) = "-INU.XML" Then
''
''            handle = FreeFile
'''            Open I_Unidadenfe & "XML\NFCe 2 - Modelos XML de Retorno" & "\" & Arquivos(lCtr) For Input As #handle
''            Open ArqNFCeTemp & Arquivos(lCtr) For Input As #handle
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
'''''            FileCopy ArqNFCeRetornos & Arquivos(lCtr), ArqNFCeTemp & Arquivos(lCtr)
'''''            Kill ArqNFCeRetornos & Arquivos(lCtr)
''
''          End If
''      Next
''   End If
''
''End Sub
''
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
''Dim iNumeroConsultar As Long
''
''   tmrAtualiza.Enabled = False
''   If MsgBox("Confirma Consulta NFC-e", vbYesNo + vbQuestion, "Confirma��o") = vbYes Then
''      iNumeroConsultar = Val(lvwNFCes.ListItems(lvwNFCes.SelectedItem.Index))
''      Dim oNFCe310 As New CNFCE310
''
''      Call oNFCe310.FConsultarNumero(iNumeroConsultar)
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
''   Set rsNFCe = cnSistema.Execute("Select * From NFCe WHERE Numero=" & Val(lvwNFCes.ListItems(lvwNFCes.SelectedItem.Index)))
''   Set rsClientes = cnSistema.Execute("Select * From Clientes WHERE idCliente = " & rsNFCe!idCliente)
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
''   strMensagem = strMensagem & "Estamos enviando em anexo arquivo XML da nota fiscal N. " & rsNFCe!Numero
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
''      sEmpresaNFCe = LerArquivoINI("NFe", "Empresa", CaminhoINI & "\System.ini")
'''      sArquivo = I_Unidadenfe & "NFC-e\" & sEmpresaNFCe & "\Enviados\Autorizados\" & Format(rsNFCe!DataEmissao, "yyyymmdd") & "\" & rsNFCe!ChaveNFCe & "-procNFe.XML"
''      sArquivo = I_CaminhoXML_NFCe & sEmpresaNFCe & "\Enviados\Autorizados\" & Format(rsNFCe!DataEmissao, "yyyymm") & "\" & rsNFCe!ChaveNFCe & "-procNFe.XML"
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
''                .HTMLBody = "Estamos enviando em anexo arquivo XML da nota fiscal N. " & rsNFCe!Numero  'rsSistema!Mail_Body
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
''Dim strIdNFCe As String
''Dim strData As String
''Dim strCNPJCliente As String
''Dim strValorTotalNota As String
''
''Dim strChaveNFCe As String
''Dim strnNF As String
''Dim strdEmi As String
''Dim strCNPJDest As String
''Dim strvNF As String
''
''Dim intNumero As Long
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
''Dim bGeradaNFCe As Boolean
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
''   sEmpresaNFCe = LerArquivoINI("NFe", "Empresa", CaminhoINI & "\System.ini")
''
''   Screen.MousePointer = vbHourglass
''   frmVisualiza.lvwDados.ListItems.Clear
''   frmVisualiza.lvwDados.ColumnHeaders.Clear
''   frmVisualiza.lvwDados.ColumnHeaders.Add , , "Nota", 700
''   frmVisualiza.lvwDados.ColumnHeaders.Add , , "Chave", 5300
''
''   Dim Arquivos() As String
''   Dim lCtr As Long
''   Arquivos = ListarArquivos(I_UnidadeNFe & "Sistemas\Importar")                                         '' Determina Pasta onde est�o os arquivos
''''''   If UBound(Arquivos) > 0 Then                                                              '' Verifica se existem arquivos na pasta
''      For lCtr = 0 To UBound(Arquivos)
''         If UCase(Mid(Arquivos(lCtr), Len(Trim(Arquivos(lCtr))) - 2, 3)) = "XML" Then        '' Verifica se encontrou algum XML
'''''            Contador = 1
''            handle = FreeFile
''            Open I_UnidadeNFe & "Sistemas\Importar\" & Arquivos(lCtr) For Input As #handle               '' Abre arquivo importado
''            Line Input #handle, Linha
''
''            Contador = 1
''            For Contador = 1 To Len(Linha)                                                   '' L� o arquivo linha a linha em busca da informa��o
''              ' Chave da NFCe
''                If Mid(Linha, Contador, 7) = "<chNFCe>" Then
''                   For Contador2 = Contador To Len(Linha)
''                       If Mid(Linha, Contador2, 8) = "</chNFCe>" Then
''                          strChaveNFCe = Mid(Linha, Contador + 7, Contador2 - Contador - 7)
''                          Contador2 = Len(Linha)
''                       End If
''                   Next
''                End If
''
''              ' N�mero da NFCe
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
''              ' Valor da NFCe
''                If Mid(Linha, Contador, 5) = "<vNF>" Then
''                   For Contador2 = Contador To Len(Linha)
''                       If Mid(Linha, Contador2, 6) = "</vNF>" Then
''                          strvNF = Mid(Linha, Contador + 5, Contador2 - Contador - 5)
''                          Contador2 = Len(Linha)
''                       End If
''                   Next
''                End If
''
''               If Len(Trim(strnNF)) <> 0 And Len(Trim(strChaveNFCe)) <> 0 And Len(Trim(strdEmi)) <> 0 And Len(Trim(strCNPJDest)) <> 0 And Len(Trim(strvNF)) <> 0 Then
''                  Set ProcuraItem = frmVisualiza.lvwDados.FindItem(strnNF)
''                  If ProcuraItem Is Nothing Then
''                     Set ItemList = frmVisualiza.lvwDados.ListItems.Add(, "R" & strnNF, strnNF)
''                         ItemList.SubItems(1) = strChaveNFCe
''
''                     'MsgBox strnNF & Chr(13) & strChaveNFCe & Chr(13) & strdEmi & Chr(13) & strCNPJDest & Chr(13) & strvNF, vbExclamation + vbOKOnly, "Campos Obrigat�rios"
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
''                     Set rsNFCe = cnSistema.Execute("Select * From NFCe ORDER BY Numero DESC")
''                     intNumero = rsNFCe!Numero + 1
''
''                     ' Inserir Nota
''                     If Not rsClientes.EOF Then
''                        cnSistema.Execute "Insert Into NFCe (Numero,Cupom,idCliente,idNaturezaOperacao,idCFOP,DadosAdicionais,DataEmissao,DataCaixa,DataVencimento,Hora,BaseCalculoICMS,ValorICMS,ValorFrete,ValorTotalProdutos,BaseICMSSubstituicao,ValorICMSSubstituicao,OutrasDespesas,ValorTotalNota,idTransportador,FreteConta,PlacaVeiculo,UFCaminhao,VolumeQuantidade,VolumeMarca,VolumeEspecie,VolumeNumero,VolumePesoBruto,VolumePesoLiquido,InformacoesCorpo,idFormaPagamento,DescontoGeral,Bonificacao,Documento,Observacao,GeradaNFCe,Situacao,TentativaEmissao,NumeroNFCeComplementar,ChaveAcessoNFCeComplementar) " & _
''                                          "Values (" & intNumero & ",0," & rsClientes!idCliente & "," & rsNaturezasOperacao!idNaturezaOperacao & "," & strCFOP & ",'','" & Date & "','" & Date & "','" & Date & "','" & Time & "','" & CStrValor(Substitui(strvNF, ".", ",")) & "','" & Val(Substitui(dblValorICMS, ",", ".")) & "'," & _
''                                                  "'" & Val(Substitui(dblValorFrete, ",", ".")) & "','" & Val(Substitui(dblValorTotalProdutos, ",", ".")) & "','" & Val(Substitui(dblBaseICMSSubstituicao, ",", ".")) & "','" & Val(Substitui(dblValorICMSSubstituicao, ",", ".")) & "','" & Val(Substitui(dblOutrasDespesas, ",", ".")) & "'," & _
''                                                  "'" & Val(Substitui(dblValorTotalNota, ",", ".")) & "'," & iTransportador & "," & iFreteConta & ",'" & UCase(strPlaca) & "','" & strUFPlaca & "','" & strVolumeQuantidade & "','" & strVolumeMarca & "','" & strVolumeEspecie & "','" & strVolumeNumero & "','" & Val(Substitui(strVolumePesoBruto, ",", ".")) & "','" & Val(Substitui(strVolumePesoLiquido, ",", ".")) & "','" & strInformacoesCorpo & "'" & _
''                                                  "," & intFormaPagamento & ",'" & Val(Substitui(dblDescontoGeral, ",", ".")) & "','" & Val(Substitui(dblBonificacao, ",", ".")) & "','" & strDocumento & "','" & strObservacao & "'," & bGeradaNFCe & ",0,1," & strnNF & ",'" & strChaveNFCe & "')"
''                     End If
''
''                     ' Inserir Item
''                     Set rsNFCe = cnSistema.Execute("Select * From NFCe Where Numero = " & intNumero)
''                     If Not rsNFCe.EOF Then
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
''                        cnSistema.Execute "Insert Into NFCeItens (idNFCe,idProduto,Data,Quantidade,Desconto,ValorUnitario,ICMS,BaseReduzida,DescricaoComplementar,idUnidade,idSituacaoTributaria,DiscriminacaoProduto,IPI,BaseReduzidaIPI,ClassificacaoFiscal,ValorFrete,CFOP) " & _
''                                          "Values (" & rsNFCe!idNFCe & "," & intIdProduto & ",'" & Date & _
''                                          "','" & CStrValor(dblQuantidade) & "','" & CStrValor(dblDesconto) & "','" & CStrValor(dblValorUnitario) & "','" & _
''                                          CStrValor(dblICMSProduto) & "','" & CStrValor(dblBaseReduzidaICMS) & "','" & SQLCheck(strDescricaoComplementar) & "'," & intUnidade & "," & intSituacaoTributaria & ",'" & _
''                                          SQLCheck(strDiscriminacaoProduto) & "','" & CStrValor(dblIPIProduto) & "','" & CStrValor(dblBaseReduzidaIPI) & "','" & SQLCheck(strClassificacaoFiscal) & "','" & CStrValor(dblValorFrete) & "','" & strCFOPItem & "')"
''                     End If
''
''                     strnNF = ""
''                     strChaveNFCe = ""
''                     strdEmi = ""
''                     strCNPJDest = ""
''                     strvNF = ""
''                  End If
''               End If
''            Next
''            Close #handle
''         End If
''      Next
''''''   End If
''
''   ' Visualiza o Retorno
''   If frmVisualiza.lvwDados.ListItems.Count > 0 Then
''      Screen.MousePointer = vbDefault
''      frmVisualiza.Show vbModal
''   End If
''
''End Sub
''
''''''Private Sub cmdValidar_Click()
''''''
'''''' ' Gerar XML
''''''   Open I_Unidadenfe & "NFC-e\Notas\Notas.TXT" For Output As #1
''''''   sAcao = 2  ' Validar
''''''   Notas
''''''   Close #1
''''''
''''''   sEmpresaNFCe = LerArquivoINI("NFe", "Empresa", CaminhoINI & "\System.ini")
'''''''   Set rsGerarXML = cnSistema.Execute("Select * From NFCe WHERE idNFCe=" & Val(Mid(lvwNFCes.SelectedItem.Key, 2, Len(lvwNFCes.SelectedItem.Key))))
''''''   Set rsGerarXML = cnSistema.Execute("Select * From NFCe WHERE Numero=" & Val(lvwNFCes.ListItems(lvwNFCes.SelectedItem.Index)))
''''''   If Not rsGerarXML.EOF Then
''''''      FileCopy I_Unidadenfe & "NFC-e\Notas\Notas.TXT", I_Unidadenfe & "NFC-e\" & sEmpresaNFCe & "\Validar\" & StrZero(rsGerarXML!Numero, 6) & "_" & RemoveCaracteres(rsEmpresa!CNPJ_CPF) & "_001_" & Format(rsGerarXML!DataEmissao, "dd_mm_yyyy") & "-NFCe.txt"
''''''      Kill I_Unidadenfe & "NFC-e\Notas\Notas.TXT"
''''''   End If
''''''
''''''   sAcao = 0
''''''   cmdTransmitir.Enabled = True
''''''   cmdValidar.Enabled = False
''''''   Carrega_View
''''''End Sub
''
''Private Function Verifica_Campos()
''Dim strMensagem As String
''Verifica_Campos = True
''
''   If Not IsDate(mskDataInicial.text) Or Val(Mid(mskDataInicial.text, 7, 4)) < 1900 Then strMensagem = strMensagem & "Data Inicial" & Chr(13)
''   If Not IsDate(mskDataFinal.text) Or Val(Mid(mskDataFinal.text, 7, 4)) < 1900 Then strMensagem = strMensagem & "Data Final" & Chr(13)
''
''   If Not strMensagem = Empty Then
''      MsgBox "Verifique os Seguintes Campos:" & Chr(13) & strMensagem, vbExclamation + vbOKOnly, "Campos Obrigat�rios"
''      Verifica_Campos = False
''      Exit Function
''   End If
''
''End Function
''
''Private Sub cmdImportarCupom_Click()
''Dim oImportar As New CImportar
''
''   Call oImportar.FImportarCupom
''
''End Sub
''
''Private Function FTAG(sTAG As String, ByRef sConteudo As String) As String
''On Error GoTo Erro
''
''   FTAG = "<" & sTAG & ">" & Trim(sConteudo) & "</" & sTAG & ">" '& vbLf
''
''   Exit Function
''Erro:
''   MsgBox "Erro " & Err & ". " & Err.Description & " - " & TypeName(Me) & ".FTAG"
''End Function
''
''Private Function GerarQrCode(iNumeroNota As Long)
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
''End Function
''
''Private Sub tmrTransmitir_Timer()
''Dim sVerArquivo As String
''Dim oImportar As New CImportar
''
''   sChaveNFe = ""
''
''   ' Verifica se existe Cupom a ser importado
''   sVerArquivo = Dir(I_UnidadeNFe & "NFC-e\Notas\CUPOM.TXT")
''   If sVerArquivo = "CUPOM.TXT" Then
''      Call oImportar.FImportarCupom
''      Kill I_UnidadeNFe & "NFC-e\Notas\CUPOM.TXT"
''   End If
''
''   'Gerar o XML
''   Set rsGerarXML = cnSistema.Execute("Select TOP 1 * From NFCe WHERE Situacao = 0") ' Em digita��o
''   If Not rsGerarXML.EOF Then
''      Open I_UnidadeNFe & "NFC-e\Notas\Notas.TXT" For Output As #1
''
''      NFCeNumero = rsGerarXML!Numero
''
''      ' Executa a Gera��o do arquivo XML
''      Call NotasNFs(rsGerarXML!Numero)
''      Close #1
''
''      ' Validar e Assinar o XML
''      sEmpresaNFCe = LerArquivoINI("NFe", "Empresa", CaminhoINI & "\System.ini")
''      FileCopy I_UnidadeNFe & "NFC-e\Notas\Notas.TXT", I_UnidadeNFe & "NFC-e\" & sEmpresaNFCe & "\Validar\" & sChaveNFe & "-nfe.XML"
''      Kill I_UnidadeNFe & "NFC-e\Notas\Notas.TXT"
''
''      ' Define como Gerada
''      cnSistema.Execute "Update NFCe set " & _
''               "Situacao = 1 " & _
''               "Where idNFCe = " & rsGerarXML!idNFCe
''   End If
''
''   'Gerar o QrCode
''   Set rsGerarXML = cnSistema.Execute("Select TOP 1 * From NFCe WHERE Situacao = 1") ' Em processamento
''   If Not rsGerarXML.EOF Then
''      If Not IsNull(rsGerarXML!ChaveNFCe) Then
''         sChaveNFe = rsGerarXML!ChaveNFCe
''         sVerArquivo = Dir(I_UnidadeNFe & "NFC-e\" & sEmpresaNFCe & "\Validar\Validado\" & sChaveNFe & "-nfe.XML")
''         If sVerArquivo = (sChaveNFe & "-nfe.XML") Then
''            GerarQrCode (rsGerarXML!Numero)
''         Else
''            If rsGerarXML!TentativaEmissao < 15 Then
''               cnSistema.Execute "Update NFCe set " & _
''                        "TentativaEmissao = " & (rsGerarXML!TentativaEmissao + 1) & " " & _
''                        "Where idNFCe = " & rsGerarXML!idNFCe
''            Else
''               cnSistema.Execute "Update NFCe set " & _
''                        "Situacao = 4 " & _
''                        "Where idNFCe = " & rsGerarXML!idNFCe
''            End If
''         End If
''      End If
''   End If
''
''End Sub
''
''Private Sub cmdGerarQrCode_Click()
'''   Set rsNFCe = cnSistema.Execute("Select * From NFCe WHERE Numero=" & Val(lvwNFCes.ListItems(lvwNFCes.SelectedItem.Index)))
''   GerarQrCode (Val(lvwNFCes.ListItems(lvwNFCes.SelectedItem.Index)))
''End Sub
''
''Private Sub cmdImprimirDANFECe_Click()
''Dim sArquivo As String
''Dim sCaminho As String
''
''   Set rsNFCe = cnSistema.Execute("Select * From NFCe WHERE Numero=" & Val(lvwNFCes.ListItems(lvwNFCes.SelectedItem.Index)))
''
''   sEmpresaNFCe = LerArquivoINI("NFe", "Empresa", CaminhoINI & "\System.ini")
'''   sArquivo = I_Unidadenfe & "NFC-e\" & sEmpresaNFCe & "\Enviados\Autorizados\" & Format(rsNFCe!DataEmissao, "yyyymmdd") & "\" & rsNFCe!ChaveNFCe & "-procNFe.XML"
''   sArquivo = I_CaminhoXML_NFCe & sEmpresaNFCe & "\Enviados\Autorizados\" & Format(rsNFCe!DataEmissao, "yyyymm") & "\" & rsNFCe!ChaveNFCe & "-procNFe.XML"
''
''   Shell "C:\UNIMAKE\" & sEmpresaNFCe & "\UNIDANFE\UNIDANFE.EXE arquivo=" & sArquivo & " visualizar = 1"
'''   Shell "C:\UNIMAKE\" & sEmpresaNFCe & "\UNIDANFE\UNIDANFE.EXE a=" & sArquivo & " v=0 i=selecionar"
''
''End Sub
''
''Private Sub tmrImprimirSistema_Timer()
''Dim handle As Integer
''Dim Linha As String
''Dim iCopia As Integer
''Dim Contador As Integer
''Dim sVerArquivo As String
''Dim OldFont As String
''
''   sVerArquivo = Dir(LerArquivoINI("Arquivos", "Caminho", App.Path & "\System.ini"))
''   If sVerArquivo = "IMPRIMIR.PRN" Then
''      handle = FreeFile
''      Open LerArquivoINI("Arquivos", "Caminho", App.Path & "\System.ini") For Input As #handle               '' Abre arquivo importado
''      While Not EOF(handle)
''         Line Input #handle, Linha
''         If Mid(Linha, 1, 3) = "<CP" Then
''            iCopia = Mid(Linha, 4, 2)
''         End If
''      Wend
''      Close #handle
''
''      If iCopia = 0 Then
''         iCopia = 1
''      End If
''
''      For Contador = 1 To iCopia
''         ' Alterar fonte
''         OldFont = Printer.FontName            ' Preserva a fonte original.
''         Printer.FontName = LerArquivoINI("Arquivos", "Fonte", App.Path & "\System.ini")
''         Printer.FontSize = LerArquivoINI("Arquivos", "Tamanho", App.Path & "\System.ini")
''         Printer.FontBold = True
''
''         ' Abrir arquivo
''         handle = FreeFile
''
''         Open LerArquivoINI("Arquivos", "Caminho", App.Path & "\System.ini") For Input As #handle               '' Abre arquivo importado
''         While Not EOF(handle)
''            Line Input #handle, Linha
''            If Mid(Linha, 1, 3) <> "<CP" Then
''               Printer.Print Linha
''            End If
''         Wend
''         Close #handle
''         Printer.FontName = OldFont   ' Restaura a fonte original.
''         Printer.EndDoc
''      Next
''
''      Kill LerArquivoINI("Arquivos", "Caminho", App.Path & "\System.ini")
''   End If
''
''
''''''''   sVerArquivo = Dir(LerArquivoINI("Arquivos", "Caminho", App.Path & "\System.ini"))
''''''''   If sVerArquivo = "IMPRIMIR.PRN" Then
''''''''      handle = FreeFile
''''''''      Open LerArquivoINI("Arquivos", "Caminho", App.Path & "\System.ini") For Input As #handle               '' Abre arquivo importado
''''''''      Line Input #handle, Linha
''''''''      If Mid(Linha, 1, 3) = "<CP" Then
''''''''         iCopia = Mid(Linha, 4, 2)
''''''''      End If
''''''''      Close #handle
''''''''
'''''''''      For iCopia = 1 To LerArquivoINI("Arquivos", "Copias", App.Path & "\System.ini")
''''''''      For Contador = 1 To iCopia
''''''''         ' Alterar fonte
''''''''         OldFont = Printer.FontName            ' Preserva a fonte original.
''''''''         Printer.FontName = LerArquivoINI("Arquivos", "Fonte", App.Path & "\System.ini")
''''''''         Printer.FontSize = LerArquivoINI("Arquivos", "Tamanho", App.Path & "\System.ini")
''''''''         Printer.FontBold = True
''''''''
''''''''         ' Abrir arquivo
''''''''         handle = FreeFile
''''''''
''''''''         Open LerArquivoINI("Arquivos", "Caminho", App.Path & "\System.ini") For Input As #handle               '' Abre arquivo importado
''''''''         While Not EOF(handle)
''''''''            Line Input #handle, Linha
''''''''            If Mid(Linha, 1, 3) <> "<CP" Then
'''''''''               If Mid(Linha, 1, 1) = "{" And Mid(Linha, 4, 1) = "}" Then
'''''''''                  Printer.FontSize = Mid(Linha, 2, 2)
'''''''''               Else
''''''''                  Printer.Print Linha
'''''''''               End If
''''''''            End If
''''''''         Wend
''''''''         Close #handle
''''''''         Printer.FontName = OldFont   ' Restaura a fonte original.
''''''''         Printer.EndDoc
''''''''      Next
''''''''
''''''''      Kill LerArquivoINI("Arquivos", "Caminho", App.Path & "\System.ini")
''''''''   End If
''
''End Sub
''
''Private Sub cmdLimparHistorico_Click()
''   lvwMensagens.ListItems.Clear
''End Sub
''
''Private Sub cmdRetransmitir_Click()
''Dim NFCeNumeroRet As String
''
''   If MsgBox("Retransmitir a nota", vbYesNo + vbQuestion, "Confirma��o") = vbYes Then
''      NFCeNumeroRet = Val(lvwNFCes.ListItems(lvwNFCes.SelectedItem.Index))
''
''      cnSistema.Execute "Update NFCe set " & _
''               "Situacao = 0, " & _
''               "TentativaEmissao = 0, " & _
''               "DataEmissao = '" & Date & "', " & _
''               "DataVencimento = '" & Date & "', " & _
''               "DataCaixa = '" & Date & "', " & _
''               "Hora = '" & Time & "' " & _
''               "Where Numero = " & NFCeNumeroRet
''
''   End If
''End Sub
''
''Private Sub tmrVerificar_Timer()
''Dim NFCeNumeroRet As String
''Dim rsVerificar As New ADODB.Recordset
''
''   If (IsDate(mskDataInicial.text) And IsDate(mskDataFinal.text)) Then
''      Set rsVerificar = cnSistema.Execute("SELECT TOP 1 * FROM NFCe WHERE DataEmissao >= cDate('" & Format(mskDataInicial.text, "dd/mm/yyyy") & "') AND DataEmissao <= cDate('" & Format(mskDataFinal.text, "dd/mm/yyyy") & "') And Situacao <> 2 Order By Numero Desc")
''      If Not rsVerificar.EOF Then
''         If rsVerificar!Situacao = 4 Then
''            If Dir(I_CaminhoXML_NFCe & I_EmpresaNF & "\Enviados\Autorizados\" & Format(rsVerificar!DataEmissao, "yyyymm") & "\" & rsVerificar!ChaveNFCe & "-procNFe.XML") = "" Then
''               cnSistema.Execute "Update NFCe set " & _
''                        "Situacao = 0, " & _
''                        "TentativaEmissao = 0, " & _
''                        "DataEmissao = '" & Date & "', " & _
''                        "DataVencimento = '" & Date & "', " & _
''                        "DataCaixa = '" & Date & "', " & _
''                        "Hora = '" & Time & "' " & _
''                        "Where Numero = " & rsVerificar!Numero
''            End If
''         ElseIf rsVerificar!Situacao = 1 And rsVerificar!TentativaEmissao = 5 Then
''            Dim oNFCe310 As New CNFCE310
''
''            Call oNFCe310.FConsultarNumero(rsVerificar!Numero)
''
''         End If
''      End If
''   End If
''End Sub

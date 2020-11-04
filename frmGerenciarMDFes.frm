VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmGerenciarMDFes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gerenciador de Manifestos MDF-e"
   ClientHeight    =   8850
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8850
   ScaleWidth      =   9360
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraMensagensRetornoErros 
      Caption         =   "Mensagens de Retorno e Erros"
      Height          =   2055
      Left            =   120
      TabIndex        =   19
      Top             =   6720
      Width           =   9165
      Begin VB.CommandButton cmdLimparHistorico 
         Caption         =   "Limpar"
         Height          =   375
         Left            =   4740
         TabIndex        =   21
         Top             =   1560
         Width           =   2115
      End
      Begin VB.CommandButton cmdHistorico 
         Caption         =   "Histórico"
         Height          =   375
         Left            =   6960
         TabIndex        =   20
         Top             =   1560
         Width           =   2115
      End
      Begin MSComctlLib.ListView lvwMensagens 
         Height          =   1290
         Left            =   60
         TabIndex        =   22
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
   Begin VB.Frame fraManifestos 
      Caption         =   "Manifestos"
      Height          =   6555
      Left            =   105
      TabIndex        =   0
      Top             =   120
      Width           =   9165
      Begin VB.ComboBox cmbSituacao 
         Height          =   315
         Left            =   4770
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   255
         Width           =   2805
      End
      Begin VB.CommandButton cmdTransmitir 
         Caption         =   "&Transmitir"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1860
         TabIndex        =   11
         Top             =   6060
         Width           =   1155
      End
      Begin VB.CommandButton cmdEncerramento 
         Caption         =   "Encerramento"
         Height          =   375
         Left            =   6420
         TabIndex        =   10
         Top             =   6060
         Width           =   1875
      End
      Begin VB.CommandButton cmdIncluirManifesto 
         Caption         =   "Manifesto"
         Height          =   375
         Left            =   60
         TabIndex        =   9
         Top             =   6060
         Width           =   1695
      End
      Begin VB.CommandButton cmdImprimirDANFE 
         Caption         =   "Imprimir DANFE"
         Enabled         =   0   'False
         Height          =   375
         Left            =   4740
         TabIndex        =   8
         Top             =   6060
         Width           =   1635
      End
      Begin VB.CommandButton cmdCancelarNota 
         Caption         =   "&Cancelar Nota"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3060
         TabIndex        =   7
         Top             =   6060
         Width           =   1575
      End
      Begin VB.Frame fraInformacoesNotas 
         Height          =   675
         Left            =   75
         TabIndex        =   2
         Top             =   5340
         Width           =   9000
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
            Left            =   7320
            TabIndex        =   6
            Top             =   300
            Visible         =   0   'False
            Width           =   1575
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
            Left            =   5040
            TabIndex        =   5
            Top             =   300
            Visible         =   0   'False
            Width           =   1815
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
            Left            =   120
            TabIndex        =   4
            Top             =   300
            Visible         =   0   'False
            Width           =   1815
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
            Left            =   2400
            TabIndex        =   3
            Top             =   300
            Visible         =   0   'False
            Width           =   1815
         End
      End
      Begin VB.CommandButton cmdPesquisar 
         Caption         =   "&Pesquisar"
         Height          =   315
         Left            =   7620
         TabIndex        =   1
         Top             =   240
         Width           =   1155
      End
      Begin MSComctlLib.ListView lvwMDFes 
         Height          =   4800
         Left            =   60
         TabIndex        =   13
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
         TabIndex        =   14
         Top             =   270
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
         Left            =   2910
         TabIndex        =   15
         Top             =   270
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
         Caption         =   "Situação"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3990
         TabIndex        =   18
         Top             =   315
         Width           =   630
      End
      Begin VB.Label lblDataInicial 
         AutoSize        =   -1  'True
         Caption         =   "Data Inicial"
         Height          =   195
         Left            =   90
         TabIndex        =   17
         Top             =   315
         Width           =   795
      End
      Begin VB.Label lblDataFinal 
         AutoSize        =   -1  'True
         Caption         =   "Data Final"
         Height          =   195
         Left            =   2040
         TabIndex        =   16
         Top             =   315
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmGerenciarMDFes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
On Error GoTo Erro
Dim titulo As String
Dim vhwnd As Long
   
   lvwMDFes.ColumnHeaders.Add , , "Número", 850
   lvwMDFes.ColumnHeaders.Add , , "Emissão", 1200
   lvwMDFes.ColumnHeaders.Add , , "Carregar", 1200
   lvwMDFes.ColumnHeaders.Add , , "Descarregar", 1200
   lvwMDFes.ColumnHeaders.Add , , "Percurso", 1200
'   lvwMDFes.ColumnHeaders.Add , , "Cliente", 4000
'   lvwMDFes.ColumnHeaders.Add , , "Valor Total", 1050, lvwColumnRight
   lvwMDFes.ColumnHeaders.Add , , "Situação", 1700
   lvwMDFes.ColumnHeaders.Add , , "", 0
   
   mskDataInicial.Text = Date - LerArquivoINI("NFe", "DiasMovimento", App.Path & "\System.ini")
   mskDataFinal.Text = Date
'   cmbSituacao.ListIndex = 7
'   cmbSituacao.ListIndex = LerArquivoINI("NFe", "Situacao", App.Path & "\System.ini")
   
'   Call Main
   
'   Carrega_View ("Carregar")
   Carrega_View_MDFe ("Carregar")
   
   ''Carregar UNINFe
'   If LerArquivoINI("NFe", "UNINFe", App.Path & "\System.ini") = 1 Then CarregaUNINFe
   
   Exit Sub
Erro:
   Call FRetornaMensagens(Err.Number & " - " & Err.Description & " - " & TypeName(Me))
End Sub


Private Sub cmdEncerramento_Click()
On Error GoTo Erro

'   cmdCancelarNota.Tag = Val(lvwNFs.ListItems(lvwNFs.SelectedItem.Index))
   frmGerenciarNF.lblStatusEvento.Tag = "EMDFE"
   frmEventos.Show vbModal
   
   Exit Sub
Erro:
   Call FRetornaMensagens(Err.Number & " - " & Err.Description & " - " & TypeName(Me))

End Sub

Private Sub cmdIncluirManifesto_Click()
   frmMDFe.Show vbModal
End Sub

Private Sub cmdTransmitir_Click()
On Error GoTo Erro
Dim iNumeroNota As Long

'   If Not Verifica_Campos() Then Exit Sub

   ' Popular número da nota
   iNumeroNota = FNumeroMDFe()
   If iNumeroNota > 0 Then
      ' Desativa verificação até o fim da transmissão
'      frmGerenciarMDFes.tmrAtualiza.Enabled = False
      
      ' Gerar o arquivo XML
      Call NotasMDFe(iNumeroNota)
   
      ' Validar e Assinar o XML
'''''      Call FArquivosNF("VALIDAR")
      
      ' Define como Gerada
'''''      Call FAtualizaNF(iNumeroNota, 1) ' Processamento
      
      ' Atualiza visualização
'''''      Carrega_View ("Carregar")
      
      ' Reativa verificação após o fim da transmissão
'      frmGerenciarMDFes.tmrAtualiza.Enabled = True
   End If
   
   Exit Sub
Erro:
   Call FRetornaMensagens(Err.Number & " - " & Err.Description & " - " & TypeName(Me))
End Sub

Private Function NotasMDFe(iNumero As Long)
On Error GoTo Erro
Dim oNFe400 As New CNF400
Dim oPreencherRs As New PreencherRS

   Call oPreencherRs.PreencherRsNFs(iNumero, "MDFe")
   If Not rsMDFes.EOF Then
      Open ARQUIVO_NFE_NOTAS For Output As #1
      
      ' Cabeçalho
      
'''''' Modelo
      Print #1, FNivelTAG(1) & "<?xml version=" & FAspas & "1.0" & FAspas & " encoding=" & FAspas & "UTF-8" & FAspas & "?>"
      Print #1, FNivelTAG(1) & "<MDFe xmlns=" & FAspas & "http://www.portalfiscal.inf.br/mdfe" & FAspas & ">"
      
'''''      Print #1, FNivelTAG(2) & "<ListaNfse>"
'''''      Print #1, FNivelTAG(3) & "<CompNfse>"
      
      Print #1, FNivelTAG(2) & "<infMDFe versao=" & FAspas & "3.00" & FAspas & " Id=" & FAspas & "MDFe" & rsMDFes!cMDF & FAspas & ">"
'      Print #1, Space(2) & "<infNFe versao=" & """" & "4.00" & """" & " Id=" & """" & "NFe" & rsNFs!cNF & """" & ">"
      
      
      
'-------------------------------------------------------------------------------------------------------
'      rsMDFes!idMDFe = rsMDF!idMDFe
'      rsMDFes!idUF = rsMDF!idUF
'      rsMDFes!idUFDescarregamento = rsMDF!idUFDescarregamento
'      rsMDFes!idTipoEmitente = rsMDF!idTipoEmitente
'      rsMDFes!idTipoTransportador = rsMDF!idTipoTransportador
'      rsMDFes!idFormaEmissao = rsMDF!idFormaEmissao
'      rsMDFes!idModalidade = rsMDF!idModalidade
'      rsMDFes!idTipoCarroceria = rsMDF!idTipoCarroceria
'      rsMDFes!idUFVeiculo = rsMDF!idUFVeiculo
'      rsMDFes!idTipoRodado = rsMDF!idTipoRodado
'
'      rsMDFes!ChaveMDFe = IIf(Trim(rsMDF!ChaveMDFe) = "" Or IsNull(rsMDF!ChaveMDFe), Empty, rsMDF!ChaveMDFe)
'      rsMDFes!Protocolo = IIf(Trim(rsMDF!Protocolo) = "" Or IsNull(rsMDF!Protocolo), Empty, rsMDF!Protocolo)
'      rsMDFes!Numero = IIf(Trim(rsMDF!Numero) = "" Or IsNull(rsMDF!Numero), Empty, rsMDF!Numero)
'      rsMDFes!DataEmissao = IIf(IsNull(rsMDF!DataEmissao), "  /  /    ", rsMDF!DataEmissao)
'      rsMDFes!DataViagem = IIf(IsNull(rsMDF!DataViagem), "  /  /    ", rsMDF!DataViagem)
'
'      rsMDFes!PlacaVeiculo = IIf(Trim(rsMDF!PlacaVeiculo) = "" Or IsNull(rsMDF!PlacaVeiculo), "   -    ", rsMDF!PlacaVeiculo)
'      rsMDFes!tara = IIf(Trim(rsMDF!tara) = "" Or IsNull(rsMDF!tara), Empty, rsMDF!tara)
'      rsMDFes!CapacidadeKG = IIf(Trim(rsMDF!CapacidadeKG) = "" Or IsNull(rsMDF!CapacidadeKG), Empty, rsMDF!CapacidadeKG)
'      rsMDFes!CapacidadeM3 = IIf(Trim(rsMDF!CapacidadeM3) = "" Or IsNull(rsMDF!CapacidadeM3), Empty, rsMDF!CapacidadeM3)
'      rsMDFes!Renavam = IIf(Trim(rsMDF!Renavam) = "" Or IsNull(rsMDF!Renavam), Empty, rsMDF!Renavam)
'      rsMDFes!DadosAdicionais = IIf(Trim(rsMDF!DadosAdicionais) = "" Or IsNull(rsMDF!DadosAdicionais), Empty, rsMDF!DadosAdicionais)
'-------------------------------------------------------------------------------------------------------
      
      
      Print #1, FNivelTAG(3) & "<ide>"
      Print #1, FNivelTAG(4) & FTAG("cUF", rsMDFes!cUF)
      Print #1, FNivelTAG(4) & FTAG("tpAmb", rsMDFes!tpAmb)
'      Print #1, FNivelTAG(4) & FTAG("tpEmit", rsMDFes!tpEmit)
      Print #1, FNivelTAG(4) & FTAG("tpEmit", rsMDFes!idTipoEmitente)
      Print #1, FNivelTAG(4) & FTAG("mod", rsMDFes!Mod)
      Print #1, FNivelTAG(4) & FTAG("serie", rsMDFes!serie)
      Print #1, FNivelTAG(4) & FTAG("nMDF", rsMDFes!nMDF)
'      Print #1, FNivelTAG(4) & FTAG("cMDF", rsMDFes!cMDF)
      Print #1, FNivelTAG(4) & FTAG("cMDF", Mid(rsMDFes!cMDF, 36, 8))
      Print #1, FNivelTAG(4) & FTAG("cDV", Mid(rsMDFes!cMDF, 44, 1))
      
'      Print #1, FNivelTAG(4) & FTAG("modal", rsMDFes!modal)
      Print #1, FNivelTAG(4) & FTAG("modal", rsMDFes!idModalidade)
'      Print #1, FNivelTAG(4) & FTAG("dhEmi", rsMDFes!dhEmi)
      Print #1, FNivelTAG(4) & FTAG("dhEmi", FData("P", rsMDFes!DataEmissao, Time))
      Print #1, FNivelTAG(4) & FTAG("tpEmis", rsMDFes!tpEmis)
      Print #1, FNivelTAG(4) & FTAG("procEmi", rsMDFes!procEmi)
      Print #1, FNivelTAG(4) & FTAG("verProc", "3.0.20")
      Print #1, FNivelTAG(4) & FTAG("UFIni", rsMDFes!UFIni)
      Print #1, FNivelTAG(4) & FTAG("UFFim", rsMDFes!UFFim)
'---------------------------------------------------------------------------------------------------------
      Print #1, FNivelTAG(4) & "<infMunCarrega>"
      
      rsMDFeLocalCarregamento.MoveFirst
      Do While Not rsMDFeLocalCarregamento.EOF
         Print #1, FNivelTAG(5) & FTAG("cMunCarrega", RemoveCaracteres(rsMDFeLocalCarregamento!CodigoMunicipio))  ' Codigo do Municipio
         Print #1, FNivelTAG(5) & FTAG("xMunCarrega", rsMDFeLocalCarregamento!Municipio)        ' Nome do Municipio
         
         rsMDFeLocalCarregamento.MoveNext
      Loop
      
      Print #1, FNivelTAG(4) & "</infMunCarrega>"
'---------------------------------------------------------------------------------------------------------
      Print #1, FNivelTAG(4) & FTAG("dhIniViagem", FData("P", rsMDFes!DataViagem, Time))
      Print #1, FNivelTAG(3) & "</ide>"
      
'---------------------------------------------------------------------------------------------------------
      Print #1, FNivelTAG(3) & "<emit>"
      Print #1, FNivelTAG(4) & FTAG("CNPJ", rsNFsEmitentes!CNPJ)
      Print #1, FNivelTAG(4) & FTAG("IE", rsNFsEmitentes!IE)
      Print #1, FNivelTAG(4) & FTAG("xNome", rsNFsEmitentes!xNome)
      Print #1, FNivelTAG(4) & FTAG("xFant", rsNFsEmitentes!xFant)
                
      Print #1, FNivelTAG(4) & "<enderEmit>"
      Print #1, FNivelTAG(5) & FTAG("xLgr", rsNFsEmitentes!xLgr)
      Print #1, FNivelTAG(5) & FTAG("nro", rsNFsEmitentes!nro)
      Print #1, FNivelTAG(5) & FTAG("xBairro", rsNFsEmitentes!xBairro)
      Print #1, FNivelTAG(5) & FTAG("cMun", rsNFsEmitentes!cMun)
      Print #1, FNivelTAG(5) & FTAG("xMun", rsNFsEmitentes!xMun)
      Print #1, FNivelTAG(5) & FTAG("CEP", rsNFsEmitentes!CEP)
      Print #1, FNivelTAG(5) & FTAG("UF", rsNFsEmitentes!UF)
      Print #1, FNivelTAG(5) & FTAG("fone", rsNFsEmitentes!fone)
      Print #1, FNivelTAG(5) & FTAG("email", rsNFsEmitentes!Email)
      Print #1, FNivelTAG(4) & "</enderEmit>"
      
      Print #1, FNivelTAG(3) & "</emit>"
'---------------------------------------------------------------------------------------------------------
      Print #1, FNivelTAG(3) & "<infModal versaoModal=" & FAspas & "3.00" & FAspas & ">"
      Print #1, FNivelTAG(4) & "<rodo>"
      Print #1, FNivelTAG(5) & "<infANTT>"
      Print #1, FNivelTAG(6) & FTAG("RNTRC", rsMDFes!RNTRC)
      Print #1, FNivelTAG(5) & "</infANTT>"
      Print #1, FNivelTAG(5) & "<veicTracao>"
'      Print #1, FNivelTAG(6) & FTAG("placa", rsMDFes!placa)
      Print #1, FNivelTAG(6) & FTAG("placa", Replace(rsMDFes!PlacaVeiculo, "-", ""))
      Print #1, FNivelTAG(6) & FTAG("tara", rsMDFes!tara)
      Print #1, FNivelTAG(6) & "<condutor>"
      
      rsMDFeCondutores.MoveFirst
      Do While Not rsMDFeCondutores.EOF
         Print #1, FNivelTAG(7) & FTAG("xNome", rsMDFeCondutores!Nome)
         Print #1, FNivelTAG(7) & FTAG("CPF", RemoveCaracteres(rsMDFeCondutores!CPF))
         
         rsMDFeCondutores.MoveNext
      Loop
      
      Print #1, FNivelTAG(6) & "</condutor>"
      Print #1, FNivelTAG(6) & FTAG("tpRod", rsMDFes!idTipoRodado)
      Print #1, FNivelTAG(6) & FTAG("tpCar", rsMDFes!idTipoCarroceria)
      Print #1, FNivelTAG(6) & FTAG("UF", rsMDFes!idUFVeiculo)
      Print #1, FNivelTAG(5) & "</veicTracao>"
      Print #1, FNivelTAG(5) & FTAG("codAgPorto", rsMDFes!codAgPorto)
      Print #1, FNivelTAG(4) & "</rodo>"
      Print #1, FNivelTAG(3) & "</infModal>"
'---------------------------------------------------------------------------------------------------------
      Print #1, FNivelTAG(3) & "<infDoc>"
      Print #1, FNivelTAG(4) & "<infMunDescarga>"
      
      rsMDFeLocalDescarregamento.MoveFirst
      Do While Not rsMDFeLocalDescarregamento.EOF
         Print #1, FNivelTAG(5) & FTAG("cMunDescarga", RemoveCaracteres(rsMDFeLocalDescarregamento!CodigoMunicipio))  ' Codigo do Municipio
         Print #1, FNivelTAG(5) & FTAG("xMunDescarga", rsMDFeLocalDescarregamento!Municipio)        ' Nome do Municipio
         
         rsMDFeLocalDescarregamento.MoveNext
      Loop
      
      Dim ContadorNFe As Integer
      
      rsMDFeNFes.MoveFirst
      Do While Not rsMDFeNFes.EOF
         Print #1, FNivelTAG(5) & "<infNFe>"
         Print #1, FNivelTAG(6) & FTAG("chNFe", rsMDFeNFes!ChaveNFe)
         Print #1, FNivelTAG(5) & "</infNFe>"
         
         ContadorNFe = ContadorNFe + 1
         
         rsMDFeNFes.MoveNext
      Loop
      
      Print #1, FNivelTAG(4) & "</infMunDescarga>"
      Print #1, FNivelTAG(3) & "</infDoc>"
      
      Print #1, FNivelTAG(3) & "<tot>"
      Print #1, FNivelTAG(4) & FTAG("qNFe", Str(ContadorNFe))
      Print #1, FNivelTAG(4) & FTAG("vCarga", Replace(rsMDFes!vCarga, ",", "."))
      Print #1, FNivelTAG(4) & FTAG("cUnid", rsMDFes!cUnid)
      Print #1, FNivelTAG(4) & FTAG("qCarga", Replace(rsMDFes!qCarga, ",", "."))
      Print #1, FNivelTAG(3) & "</tot>"
      
      Print #1, FNivelTAG(2) & "</infMDFe>"
      Print #1, FNivelTAG(1) & "</MDFe>"
     
      Close #1
      
      Call FArquivosNF("TRANSMITIRMDFe")
   End If
   
'   Exit Function
'Erro:
'    MsgBox "Erro " & Err & ". " & Err.Description & " - " & TypeName(Me) & ".NotasNFs"
   Exit Function
Erro:
   Call FRetornaMensagens(Err.Number & " - " & Err.Description & " - " & TypeName(Me) & ".NotasNFs")
End Function

Private Function FNumeroMDFe() As Long
On Error GoTo Erro

   If lvwMDFes.ListItems.Count > 0 Then
      FNumeroMDFe = Val(lvwMDFes.ListItems(lvwMDFes.SelectedItem.Index))
   Else
      FNumeroMDFe = 0
   End If

   Exit Function
Erro:
   Call FRetornaMensagens(Err.Number & " - " & Err.Description & " - " & TypeName(Me) & ".FNumeroMDFe")
End Function

Private Sub lvwMDFes_Click()
Dim oPreencherRs As New PreencherRS
   
   If lvwMDFes.ListItems.Count <> 0 Then
      Call oPreencherRs.PreencherRsNFs(Val(lvwMDFes.ListItems(lvwMDFes.SelectedItem.Index)), "MDFe")
      
      cmdTransmitir.Enabled = True
      cmdCancelarNota.Enabled = True
      cmdImprimirDANFE.Enabled = True
      cmdEncerramento.Enabled = True
      
'      If rsMDFes!Situacao = 2 Then
'         If Trim(rsMDFes!cMDF) <> "" And Trim(rsMDFes!Protocolo) <> "" Then
'            cmdTransmitir.Enabled = False
'            cmdCancelarNota.Enabled = True
'            cmdImprimirDANFE.Enabled = True
'            cmdEncerramento.Enabled = True
'         Else
'            cmdTransmitir.Enabled = True
'            cmdCancelarNota.Enabled = False
'            cmdImprimirDANFE.Enabled = False
'            cmdEncerramento.Enabled = False
'         End If
'      Else
'            cmdTransmitir.Enabled = True
'            cmdCancelarNota.Enabled = False
'            cmdImprimirDANFE.Enabled = False
'            cmdEncerramento.Enabled = False
'      End If
   End If
End Sub

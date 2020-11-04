VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmGerenciarNFSe 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gerenciador de NFS-e"
   ClientHeight    =   6750
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6750
   ScaleWidth      =   9360
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraManifestos 
      Caption         =   "NFS-e"
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
         Caption         =   "Incluir NFS-e"
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
      Begin MSComctlLib.ListView lvwNFSe 
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
         Caption         =   "Situa��o"
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
Attribute VB_Name = "frmGerenciarNFSe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

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

Private Sub Form_Load()
On Error GoTo Erro
Dim titulo As String
Dim vhwnd As Long
   
   lvwNFSe.ColumnHeaders.Add , , "N�mero", 850
   lvwNFSe.ColumnHeaders.Add , , "Emiss�o", 1050
   lvwNFSe.ColumnHeaders.Add , , "Cliente", 4000
   lvwNFSe.ColumnHeaders.Add , , "Valor Total", 1050, lvwColumnRight
   lvwNFSe.ColumnHeaders.Add , , "Situa��o", 1700
   lvwNFSe.ColumnHeaders.Add , , "", 0
   
   mskDataInicial.Text = Date - LerArquivoINI("NFe", "DiasMovimento", App.Path & "\System.ini")
   mskDataFinal.Text = Date
'   cmbSituacao.ListIndex = 7
'   cmbSituacao.ListIndex = LerArquivoINI("NFe", "Situacao", App.Path & "\System.ini")
   
'   Call Main
   
'   Carrega_View ("Carregar")
   
   ''Carregar UNINFe
'   If LerArquivoINI("NFe", "UNINFe", App.Path & "\System.ini") = 1 Then CarregaUNINFe
   
   Exit Sub
Erro:
   Call FRetornaMensagens(Err.Number & " - " & Err.Description & " - " & TypeName(Me))
End Sub

Private Sub cmdTransmitir_Click()
On Error GoTo Erro
Dim iNumeroNota As Long

'   If Not Verifica_Campos() Then Exit Sub

   ' Popular n�mero da nota
   iNumeroNota = FNumeroNF()
   If iNumeroNota > 0 Then
      ' Desativa verifica��o at� o fim da transmiss�o
      frmGerenciarNF.tmrAtualiza.Enabled = False
      
      ' Gerar o arquivo XML
      Call NotasNFSe(iNumeroNota)
   
      ' Validar e Assinar o XML
'''''      Call FArquivosNF("VALIDAR")
      
      ' Define como Gerada
'''''      Call FAtualizaNF(iNumeroNota, 1) ' Processamento
      
      ' Atualiza visualiza��o
'''''      Carrega_View ("Carregar")
      
      ' Reativa verifica��o ap�s o fim da transmiss�o
      frmGerenciarNF.tmrAtualiza.Enabled = True
   End If
   
   Exit Sub
Erro:
   Call FRetornaMensagens(Err.Number & " - " & Err.Description & " - " & TypeName(Me))
End Sub

Private Function NotasNFSe(iNumero As Long)
On Error GoTo Erro
Dim oNFe400 As New CNF400
Dim oPreencherRs As New PreencherRS

   If Not rsNFs.EOF Then
      Open ARQUIVO_NFE_NOTAS For Output As #1
      
      ' Cabe�alho
      
'''''' Modelo
'/''''<?xml version="1.0" encoding="UTF-8" standalone="true"?>
      Print #1, Space(2) & "<?xml version=" & FAspas & "1.0" & FAspas & " encoding=" & FAspas & "utf-8" & FAspas & " standalone=" & FAspas & "true" & FAspas & "?>"
'/''''<ConsultarNfseFaixaResposta xmlns="http://www.abrasf.org.br/nfse.xsd" xmlns:ns2="http://www.w3.org/2000/09/xmldsig#">
      Print #1, Space(2) & "<ConsultarNfseFaixaResposta xmlns:ns2=" & FAspas & "http://www.w3.org/2000/09/xmldsig# " & FAspas & "xmlns=" & FAspas & "http://www.abrasf.org.br/nfse.xsd" & FAspas & ">"
'/''''   <ListaNfse>
      Print #1, Space(4) & "<ListaNfse>"
'/''''      <CompNfse>
      Print #1, Space(6) & "<CompNfse>"
'/''''         <Nfse versao="2.01">
      Print #1, Space(8) & "<Nfse versao=" & FAspas & "2.01" & FAspas & ">"
'/''''            <InfNfse>
      Print #1, Space(10) & "<InfNfse>"
'/''''               <Numero>2560</Numero>
      Print #1, Space(12) & FTAG("Numero", rsNFSe!Numero)
'/''''               <CodigoVerificacao>8829363238191129</CodigoVerificacao>
      Print #1, Space(12) & FTAG("CodigoVerificacao", rsNFSe!CodigoVerificacao)
'/''''               <DataEmissao>2019-11-29T12:00:00.000-03:00</DataEmissao>
      Print #1, Space(12) & FTAG("DataEmissao", rsNFSe!CodigoVerificacao)
'/''''               <ValoresNfse>
      Print #1, Space(12) & "<ValoresNfse>"
'/''''                  <BaseCalculo>229.90</BaseCalculo>
      Print #1, Space(12) & FTAG("BaseCalculo", rsNFSe!BaseCalculo)
'/''''                  <Aliquota>2.79</Aliquota>
      Print #1, Space(12) & FTAG("Aliquota", rsNFSe!Aliquota)
'/''''                  <ValorIss>6.41</ValorIss>
      Print #1, Space(12) & FTAG("ValorIss", rsNFSe!ValorIss)
'/''''                  <ValorLiquidoNfse>229.90</ValorLiquidoNfse>
      Print #1, Space(12) & FTAG("ValorLiquidoNfse", rsNFSe!ValorLiquidoNfse)
'/''''               </ValoresNfse>
      Print #1, Space(12) & "</ValoresNfse>"
'---------------------------------------------------------------------------------------------------------
'/''''               <PrestadorServico>
      Print #1, Space(14) & "<PrestadorServico>"
'/''''                  <IdentificacaoPrestador>
      Print #1, Space(16) & "<IdentificacaoPrestador>"
'/''''                  <CpfCnpj>
      Print #1, Space(16) & "<CpfCnpj>"
'/''''                     <Cnpj>09161920000166</Cnpj>
      Print #1, Space(18) & FTAG("Cnpj", rsNFsEmitentes!CNPJ)
'/''''                  </CpfCnpj>
      Print #1, Space(16) & "</CpfCnpj>"
'/''''                  <InscricaoMunicipal>167712007</InscricaoMunicipal>
      Print #1, Space(16) & FTAG("InscricaoMunicipal", rsNFsEmitentes!IM)
'/''''               </IdentificacaoPrestador>
      Print #1, Space(16) & "</IdentificacaoPrestador>"
'/''''               <RazaoSocial>LINK EXPLORER TELECOMUNICACAO LTDA - ME</RazaoSocial>
      Print #1, Space(16) & FTAG("RazaoSocial", rsNFsEmitentes!xNome)
'/''''               <NomeFantasia>LINK EXPLORER</NomeFantasia>
      Print #1, Space(16) & FTAG("NomeFantasia", rsNFsEmitentes!xFant)
'/''''               <Endereco>
      Print #1, Space(16) & "<Endereco>"
'/''''                     <Endereco>PRACA RAIMUNDO DE ARAUJO MELO DIST. 0</Endereco>
      Print #1, Space(18) & FTAG("Endereco", rsNFsEmitentes!xLgr)
'/''''                     <Numero>125</Numero>
      Print #1, Space(18) & FTAG("Endereco", rsNFsEmitentes!nro)
'/''''                     <Complemento>SALA 501</Complemento>
      Print #1, Space(18) & FTAG("Complemento", "")
'/''''                     <Bairro>CENTRO</Bairro>
      Print #1, Space(18) & FTAG("Bairro", rsNFsEmitentes!xBairro)
'/''''                     <CodigoMunicipio>5212501</CodigoMunicipio>
      Print #1, Space(18) & FTAG("CodigoMunicipio", rsNFsEmitentes!cMun)
'/''''                     <Uf>GO</Uf>
      Print #1, Space(18) & FTAG("Uf", rsNFsEmitentes!UF)
'/''''                     <Cep>72800360</Cep>
      Print #1, Space(18) & FTAG("Cep", rsNFsEmitentes!CEP)
'/''''               </Endereco>
      Print #1, Space(16) & "</Endereco>"
'/''''               </PrestadorServico>
      Print #1, Space(16) & "</PrestadorServico>"
'/''''               <OrgaoGerador>
      Print #1, Space(16) & "<OrgaoGerador>"
'/''''                  <CodigoMunicipio>5212501</CodigoMunicipio>
      Print #1, Space(18) & FTAG("CodigoMunicipio", rsNFsEmitentes!cMun)
'/''''                  <Uf>GO</Uf>
      Print #1, Space(18) & FTAG("Uf", rsNFsEmitentes!UF)
'/''''               </OrgaoGerador>
      Print #1, Space(16) & "</OrgaoGerador>"
'/''''               <DeclaracaoPrestacaoServico>
      Print #1, Space(16) & "<DeclaracaoPrestacaoServico>"
'/''''                  <InfDeclaracaoPrestacaoServico>
      Print #1, Space(18) & "<InfDeclaracaoPrestacaoServico>"
'/''''                     <Servico>
      Print #1, Space(20) & "<Servico>"
'/''''                        <Valores>
      Print #1, Space(22) & "<Valores>"
''''''                           <ValorServicos>229.90</ValorServicos>
''''''                           <ValorDeducoes>0.00</ValorDeducoes>
''''''                           <ValorPis>0.00</ValorPis>
''''''                           <ValorCofins>0.00</ValorCofins>
''''''                           <ValorInss>0.00</ValorInss>
''''''                           <ValorIr>0.00</ValorIr>
''''''                           <ValorCsll>0.00</ValorCsll>
''''''                           <OutrasRetencoes>0.00</OutrasRetencoes>
''''''                           <ValorIss>6.41</ValorIss>
''''''                           <Aliquota>2.79</Aliquota>
''''''                           <DescontoIncondicionado>0.00</DescontoIncondicionado>
'/''''                        </Valores>
      Print #1, Space(22) & "</Valores>"
''''''                        <IssRetido>2</IssRetido>
''''''                        <ItemListaServico>1.03</ItemListaServico>
''''''                        <CodigoCnae>6190601</CodigoCnae>
''''''                        <Discriminacao>LINK FIBRA 100MB</Discriminacao>
''''''                        <CodigoMunicipio>5212501</CodigoMunicipio>
''''''                        <CodigoPais>0000</CodigoPais>
''''''                        <ExigibilidadeISS>0</ExigibilidadeISS>
'/''''                     </Servico>
      Print #1, Space(20) & "</Servico>"
'---------------------------------------------------------------------------------------------------------
'/''''                     <Prestador>
      Print #1, Space(20) & "<Prestador>"
'/''''                        <CpfCnpj>
      Print #1, Space(22) & "<CpfCnpj>"
'/''''                           <Cnpj>09161920000166</Cnpj>
      Print #1, Space(24) & FTAG("Cnpj", rsNFsEmitentes!CNPJ)
'/''''                        </CpfCnpj>
      Print #1, Space(22) & "</CpfCnpj>"
'/''''                        <InscricaoMunicipal>167712007</InscricaoMunicipal>
      Print #1, Space(16) & FTAG("InscricaoMunicipal", rsNFsEmitentes!IM)
'/''''                     </Prestador>
      Print #1, Space(20) & "</Prestador>"
'---------------------------------------------------------------------------------------------------------
'/''''                     <Tomador>
      Print #1, Space(20) & "<Tomador>"
'/''''                        <IdentificacaoTomador>
      Print #1, Space(22) & "<IdentificacaoTomador>"
'/''''                           <CpfCnpj>
      Print #1, Space(24) & "<CpfCnpj>"
'/''''                              <Cnpj>00000789301113</Cnpj>
      Print #1, Space(26) & FTAG("Cnpj", rsNFsDestinatarios!CNPJ)
'/''''                           </CpfCnpj>
      Print #1, Space(24) & "</CpfCnpj>"
'/''''                           <InscricaoMunicipal/>
      Print #1, Space(24) & FTAG("InscricaoMunicipal", rsNFsDestinatarios!IM)
'/''''                        </IdentificacaoTomador>
      Print #1, Space(22) & "</IdentificacaoTomador>"
''''''                        <RazaoSocial>NIELTON BARBOSA RODRIGUES</RazaoSocial>
      Print #1, Space(24) & FTAG("RazaoSocial", rsNFsDestinatarios!xNome)
'/''''                        <Endereco>
      Print #1, Space(22) & "<Endereco>"
'/''''                           <Endereco>RUA JOSE A E ALBUQUERQUE QD. 05 LT. 12 DIST. 0</Endereco>
      Print #1, Space(24) & FTAG("Endereco", rsNFsDestinatarios!xLgr)
'/''''                           <Numero/>
      Print #1, Space(24) & FTAG("Endereco", rsNFsDestinatarios!nro)
'/''''                           <Complemento/>
      Print #1, Space(24) & FTAG("Complemento", "")
'/''''                           <Bairro>VILA SAO JOSE</Bairro>
      Print #1, Space(24) & FTAG("Bairro", rsNFsDestinatarios!xBairro)
'/''''                           <CodigoMunicipio>5212501</CodigoMunicipio>
      Print #1, Space(24) & FTAG("CodigoMunicipio", rsNFsDestinatarios!cMun)
'/''''                           <Uf>GO</Uf>
      Print #1, Space(24) & FTAG("Uf", rsNFsDestinatarios!UF)
'/''''                           <Cep>72813550</Cep>
      Print #1, Space(24) & FTAG("Cep", rsNFsDestinatarios!CEP)
'/''''                        </Endereco>
      Print #1, Space(22) & "</Endereco>"
'/''''                     </Tomador>
      Print #1, Space(22) & "</Tomador>"
'/''''                     <OptanteSimplesNacional>1</OptanteSimplesNacional>
      Print #1, Space(22) & FTAG("OptanteSimplesNacional", LerArquivoINI("NFe", "Regime", CaminhoINI & "\System.ini"))
'/''''                      <IncentivoFiscal>0</IncentivoFiscal>
      Print #1, Space(22) & FTAG("IncentivoFiscal", "0")
'/''''                  </InfDeclaracaoPrestacaoServico>
      Print #1, Space(20) & "</InfDeclaracaoPrestacaoServico>"
'/''''               </DeclaracaoPrestacaoServico>
      Print #1, Space(18) & "</DeclaracaoPrestacaoServico>"
'/''''            </InfNfse>
      Print #1, Space(10) & "</InfNfse>"
'/''''         </Nfse>
      Print #1, Space(8) & "</Nfse>"
'/''''      </CompNfse>
      Print #1, Space(6) & "</CompNfse>"
'/''''      <Pagina>1</Pagina>
      Print #1, Space(6) & "<Pagina>1</Pagina>"
'/''''   </ListaNfse>
      Print #1, Space(4) & "<ListaNfse>"
'/''''</ConsultarNfseFaixaResposta>
      Print #1, Space(2) & "</ConsultarNfseFaixaResposta>"
'''''' Fim do Modelo
     
      Close #1
   End If
   
'   Exit Function
'Erro:
'    MsgBox "Erro " & Err & ". " & Err.Description & " - " & TypeName(Me) & ".NotasNFs"
   Exit Function
Erro:
   Call FRetornaMensagens(Err.Number & " - " & Err.Description & " - " & TypeName(Me) & ".NotasNFs")
End Function

Private Function FNumeroNF() As Long
On Error GoTo Erro

   If lvwNFs.ListItems.Count > 0 Then
      FNumeroNF = Val(lvwNFs.ListItems(lvwNFs.SelectedItem.Index))
   Else
      FNumeroNF = 0
   End If

   Exit Function
Erro:
   Call FRetornaMensagens(Err.Number & " - " & Err.Description & " - " & TypeName(Me) & ".FNumeroNF")
End Function

Private Function FTAG(ByRef sTAG As String, ByRef sConteudo As String) As String
On Error GoTo Erro
   
'   FTAG = "<" & sTAG & ">" & Trim(sConteudo) & "</" & sTAG & ">" '& vbLf
   FTAG = "<" & sTAG & ">" & UTF8_Encode(Trim(sConteudo)) & "</" & sTAG & ">"  '& vbLf

   Exit Function
Erro:
   Call FRetornaMensagens(Err.Number & " - " & Err.Description & " - " & TypeName(Me) & ".FTAG")
End Function

Attribute VB_Name = "modNFs"
Option Explicit
Dim ItemList As ListItem
Dim ProcuraItem As ListItem

' Vari�vies P�blicas j� definidas
'''''Public sEmpresaNF As String
Public I_EmpresaNF As String

Public I_ModeloNF As String                     ' Modelo da NF 55-NF-e / 65-NFC-e
Public I_TabelasNF As String                    ' Tabelas da NF 55-NFe / 65-NFCe
Public I_PastaUNINFe As String                  ' Pasta de Loca��o dos arquivos UNINFe

Public I_Empresa_CNPJ_CPF As String             ' CNPJ da Empresa/Emitente
Public I_EmpresaMunicipio As String             ' Nome do municipio do Emitente
Public I_EmpresaCodigoMunicipio As String       ' Codigo do municipio do Emitente
Public I_EmpresaUF As String                    ' Nome da UF do Emitente
Public I_EmpresaCodigoUF As String              ' Codigo da UF do Emitente
Public I_AmbienteNF As String                   ' Produ��o ou Homologa��o


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Tabelas espec�ficas

Public I_TabelaUFs As String                    ' Tabelas de UFs
Public I_TabelaMunicipios As String             ' Tabelas de Munic�pios

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public I_UnidadeNFe As String                   ' Em que disco vai buscar a pasta
Public I_CaminhoXML_NFCe As String              ' Em qual pasta est�o os XMLs
Public I_CaminhoXML_NFe As String
Public sCPFDestinatario As String

Public sChaveNFe As String
Public I_NFCeNumero As String

Public dValorTributos As Double
Public dValorTotalBC As Double
Public dValorTotalICMS As Double
Public dValorTotalBaseISS As Double
Public dValorTotalISS As Double
Public dValorTotalBCST As Double
Public dValorTotalICMSST As Double
Public dValorTotalProdutos As Double
Public dValorTotalFrete As Double
Public dValorTotalSeguro As Double
Public dValorTotalDesconto As Double
Public dValorTotalII As Double
Public dValorTotalIPI As Double
Public dValorTotalPIS As Double
Public dValorTotalCofins As Double
Public dValorTotalOutro As Double
Public dValorTotalNFCe As Double
Public sICMSAproveitamento As String
Public sValorTotalNFCe As String
Public sValorTotalICMSNFCe As String
Public sDataEmissao As String

Public I_FormatoQuantidade As String   ' Quantidade de Casas Decimais para Quantidade
Public I_FormatoValor As String        ' Quantidade de Casas Decimais para Valores

' Criar RecordSets P�blico para transportar os dados para todos os m�dulos do Sistema
'------------------------------------------------------------------------------------
Public rsEmpresa As New ADODB.Recordset

''''''''''''''''''''''''''''''''''''''''''''''''''
' Criar RecordSets das NFes ou NFCes
''''''''''''''''''''''''''''''''''''''''''''''''''
Public rsNFs As New ADODB.Recordset
Public rsNFsItens As New ADODB.Recordset
Public rsNFsPagamentos As New ADODB.Recordset
Public rsNFsTotaisICMS As New ADODB.Recordset
Public rsNFsTotaisISS As New ADODB.Recordset
Public rsNFsEmitentes As New ADODB.Recordset
Public rsNFsDestinatarios As New ADODB.Recordset
Public rsNFsTransportes As New ADODB.Recordset
Public rsTransportadores As New ADODB.Recordset
Public rsNFsDados As New ADODB.Recordset

''''''''''''''''''''''''''''''''''''''''''''''''''
' Criar RecordSets das NFes ou NFCes
''''''''''''''''''''''''''''''''''''''''''''''''''
Public rsNFSe As New ADODB.Recordset
Public rsNFSeItens As New ADODB.Recordset

''''''''''''''''''''''''''''''''''''''''''''''''''
' Criar RecordSets dos MDFes
''''''''''''''''''''''''''''''''''''''''''''''''''
Public rsMDFes As New ADODB.Recordset
Public rsMDFeLocalCarregamento As New ADODB.Recordset
Public rsMDFeLocalDescarregamento As New ADODB.Recordset
Public rsMDFePercurso  As New ADODB.Recordset
Public rsMDFeCondutores  As New ADODB.Recordset
Public rsMDFeNFes  As New ADODB.Recordset
''''''''''''''''''''''''''''''''''''''''''''''''''

Public rsNFsArquivosRetorno As New ADODB.Recordset


Public ARQUIVO_NFE_NOTAS As String

Public CAMINHO_NFE_ENVIO As String
Public CAMINHO_NFE_VALIDAR As String
Public CAMINHO_NFE_VALIDADO As String
Public CAMINHO_NFE_RETORNO As String
Public CAMINHO_NFE_ERROS As String
Public CAMINHO_NFE_TEMP As String

Public sCaminhoOrigemNF As String
Public sCaminhoDestinoNF As String
Public sArquivoOrigemNF As String
Public sArquivoDestinoNF As String

'''''Public Declare Function SHCreateThread Lib "shlwapi.dll" (ByVal pfnThreadProc As Long, pData As Any, ByVal dwFlags As Long, ByVal pfnCallback As Long) As Long

Public Sub Carrega_View(sModo As String)
On Error GoTo Erro

Dim sSituacao As String
Dim Contador As Integer

Dim sqlSituacao As String
Dim sSql As String

Dim rsDados As New ADODB.Recordset
Dim rsNFsInutilizadas As New ADODB.Recordset

'   cmdCancelarNota.Enabled = False
'   cmdTransmitir.Enabled = False
   frmGerenciarNF.cmdImprimirDANFE.Enabled = False
   
   ' Validar Per�odo
   If Not FValidarPeriodo(frmGerenciarNF.mskDataInicial.Text, frmGerenciarNF.mskDataFinal.Text) Then Exit Sub
   ' Retorna condi��o do SQL
   sqlSituacao = FSituacao("C", frmGerenciarNF.cmbSituacao.ListIndex)
   ' Carrega View
   If sModo = "Carregar" Then frmGerenciarNF.lvwNFs.ListItems.Clear

   ' Notas
   Contador = 1

   sSql = "       SELECT "
   sSql = sSql & vbCrLf & "      N.id" & I_TabelasNF & ", "
   sSql = sSql & vbCrLf & "      N.idCliente, "
   sSql = sSql & vbCrLf & "      N.Numero, "
   sSql = sSql & vbCrLf & "      N.DataEmissao, "
   sSql = sSql & vbCrLf & "      C.Nome, "
   If I_SGBD = "SQLSERVER" Then
      sSql = sSql & vbCrLf & "      T.Total  AS Total, "
   ElseIf I_SGBD = "ACCESS" Then
      sSql = sSql & vbCrLf & "      (T.Total + T.TotalFrete) AS Total, "
   End If
   sSql = sSql & vbCrLf & "      N.Situacao "
   sSql = sSql & vbCrLf & "FROM "
'   sSql = sSql & vbCrLf & "      " & I_TabelasNF & " N, Clientes C, Total" & I_TabelasNF & " T "
   If I_SGBD = "SQLSERVER" Then
      sSql = sSql & vbCrLf & "      " & I_TabelasNF & " N, ClientesInfFiscais C, (SELECT id" & I_TabelasNF & ", SUM(Quantidade * ValorUnitario)AS Total FROM " & I_TabelasNF & "Itens GROUP BY id" & I_TabelasNF & ") AS T "
   ElseIf I_SGBD = "ACCESS" Then
      sSql = sSql & vbCrLf & "      " & I_TabelasNF & " N, ClientesInfFiscais C, Total" & I_TabelasNF & " T "
   End If
   sSql = sSql & vbCrLf & "WHERE "
   sSql = sSql & vbCrLf & "      N.idCliente = C.idCliente AND "
   sSql = sSql & vbCrLf & "      N.id" & I_TabelasNF & " = T.id" & I_TabelasNF & " AND "
   If I_SGBD = "SQLSERVER" Then
      sSql = sSql & vbCrLf & "      N.DataEmissao >= '" & Format(frmGerenciarNF.mskDataInicial.Text, "yyyy-mm-dd") & " 00:00:00' AND "
      sSql = sSql & vbCrLf & "      N.DataEmissao <= '" & Format(frmGerenciarNF.mskDataFinal.Text, "yyyy-mm-dd") & " 23:59:59' "
   ElseIf I_SGBD = "ACCESS" Then
      sSql = sSql & vbCrLf & "      N.DataEmissao >= cDate('" & Format(frmGerenciarNF.mskDataInicial.Text, "dd/mm/yyyy") & " 00:00:00') AND "
      sSql = sSql & vbCrLf & "      N.DataEmissao <= cDate('" & Format(frmGerenciarNF.mskDataFinal.Text, "dd/mm/yyyy") & " 23:59:59') "
   End If
   sSql = sSql & vbCrLf & sqlSituacao & " Order By N.Numero"
   
   Set rsDados = cnSistema.Execute(sSql)
   Do While Not rsDados.EOF
      sSituacao = FSituacao("S", rsDados!Situacao)      ' Retorna Situa��o atual da nota
      
      ' Pesquisa se registro existe
'''''      Set ItemList = frmGerenciarNF.lvwNFs.FindItem(StrZero(rsDados!Numero, 8))
'''''      If ItemList Is Nothing Then
         Set ItemList = frmGerenciarNF.lvwNFs.ListItems.Add(, "R" & CStr(Contador), StrZero(rsDados!Numero, 8))
'''''      End If
      
      ItemList.SubItems(1) = Format(rsDados!DataEmissao, "DD/MM/YYYY")
      ItemList.SubItems(2) = Trim(rsDados!Nome)
      ItemList.SubItems(3) = Format(rsDados!Total, "##,##0.00")
      ItemList.SubItems(4) = sSituacao
      
      Contador = Contador + 1
      rsDados.MoveNext
   Loop
   
   ' Notas
   If I_SGBD = "SQLSERVER" Then
      Set rsNFsInutilizadas = cnSistema.Execute("SELECT * FROM " & I_TabelasNF & "Inutilizadas WHERE Data >= '" & Format(frmGerenciarNF.mskDataInicial.Text, "yyyy-mm-dd") & "' AND Data <= '" & Format(frmGerenciarNF.mskDataFinal.Text, "yyyy-mm-dd") & "' Order By Numero")
   ElseIf I_SGBD = "ACCESS" Then
      Set rsNFsInutilizadas = cnSistema.Execute("SELECT * FROM " & I_TabelasNF & "Inutilizadas WHERE Data >= cDate('" & Format(frmGerenciarNF.mskDataInicial.Text, "dd/mm/yyyy") & "') AND Data <= cDate('" & Format(frmGerenciarNF.mskDataFinal.Text, "dd/mm/yyyy") & "') Order By Numero")
   End If

   If Not rsNFsInutilizadas.EOF Then
      Do While Not rsNFsInutilizadas.EOF

'         Set ItemList = lvwNFs.ListItems.Add(, "I" & CStr(rsNFsInutilizadas!Numero), StrZero(rsNFsInutilizadas!Numero, 8))

         Set ProcuraItem = frmGerenciarNF.lvwNFs.FindItem(StrZero(rsNFsInutilizadas!Numero, 8))
         If ProcuraItem Is Nothing Then
            Set ItemList = frmGerenciarNF.lvwNFs.ListItems.Add(, "I" & CStr(Contador), StrZero(rsNFsInutilizadas!Numero, 8))
            ItemList.SubItems(1) = rsNFsInutilizadas!Data
            ItemList.SubItems(2) = "Nota Inutilizada"
            ItemList.SubItems(3) = Format(0, "##,##0.00")
            ItemList.SubItems(4) = "Inutilizada"
            If IsNull(rsNFsInutilizadas!Protocolo) Or rsNFsInutilizadas!Protocolo = "" Then
               ItemList.SubItems(5) = "Aguarde..."
            Else
               ItemList.SubItems(5) = ""
            End If
         End If

         Contador = Contador + 1
         rsNFsInutilizadas.MoveNext
      Loop
   End If
   
   Exit Sub
Erro:
   Call FRetornaMensagens(Err.Number & " - " & Err.Description & " - Carrega_View ")
End Sub

Public Sub Carrega_View_NFSe(sModo As String)
On Error GoTo Erro

Dim sSituacao As String
Dim Contador As Integer

Dim sqlSituacao As String
Dim sSql As String

Dim rsDados As New ADODB.Recordset
Dim rsNFsInutilizadas As New ADODB.Recordset
Dim I_TabelaNFSe As String

   I_TabelaNFSe = "NFSe"
   
'   cmdCancelarNota.Enabled = False
'   cmdTransmitir.Enabled = False
   frmGerenciarNFSe.cmdImprimirDANFE.Enabled = False
   
   ' Validar Per�odo
   If Not FValidarPeriodo(frmGerenciarNFSe.mskDataInicial.Text, frmGerenciarNFSe.mskDataFinal.Text) Then Exit Sub
   ' Retorna condi��o do SQL
   sqlSituacao = FSituacao("C", frmGerenciarNFSe.cmbSituacao.ListIndex)
   ' Carrega View
   If sModo = "Carregar" Then frmGerenciarNFSe.lvwNFSe.ListItems.Clear
                            
   ' Notas
   Contador = 1

   sSql = "       SELECT "
   sSql = sSql & vbCrLf & "      N.id" & I_TabelaNFSe & ", "
   sSql = sSql & vbCrLf & "      N.idCliente, "
   sSql = sSql & vbCrLf & "      N.Numero, "
   sSql = sSql & vbCrLf & "      N.DataEmissao, "
   sSql = sSql & vbCrLf & "      C.Nome, "
   sSql = sSql & vbCrLf & "      N.Situacao, "
   sSql = sSql & vbCrLf & "      N.ValorServicos "
   sSql = sSql & vbCrLf & "FROM "
   If I_SGBD = "SQLSERVER" Then
      sSql = sSql & vbCrLf & "      " & I_TabelaNFSe & " N, ClientesInfFiscais C "
   ElseIf I_SGBD = "ACCESS" Then
      sSql = sSql & vbCrLf & "      " & I_TabelaNFSe & " N, ClientesInfFiscais C "
   End If
   sSql = sSql & vbCrLf & "WHERE "
   sSql = sSql & vbCrLf & "      N.idCliente = C.idCliente AND "
   If I_SGBD = "SQLSERVER" Then
      sSql = sSql & vbCrLf & "      N.DataEmissao >= '" & Format(frmGerenciarNFSe.mskDataInicial.Text, "yyyy-mm-dd") & " 00:00:00' AND "
      sSql = sSql & vbCrLf & "      N.DataEmissao <= '" & Format(frmGerenciarNFSe.mskDataFinal.Text, "yyyy-mm-dd") & " 23:59:59' "
   ElseIf I_SGBD = "ACCESS" Then
      sSql = sSql & vbCrLf & "      N.DataEmissao >= cDate('" & Format(frmGerenciarNFSe.mskDataInicial.Text, "dd/mm/yyyy") & " 00:00:00') AND "
      sSql = sSql & vbCrLf & "      N.DataEmissao <= cDate('" & Format(frmGerenciarNFSe.mskDataFinal.Text, "dd/mm/yyyy") & " 23:59:59') "
   End If
   sSql = sSql & vbCrLf & sqlSituacao & " Order By N.Numero"
   
   Set rsDados = cnSistema.Execute(sSql)
   Do While Not rsDados.EOF
      sSituacao = FSituacao("S", rsDados!Situacao)      ' Retorna Situa��o atual da nota
      
      ' Pesquisa se registro existe
'''''      Set ItemList = frmGerenciarNF.lvwNFs.FindItem(StrZero(rsDados!Numero, 8))
'''''      If ItemList Is Nothing Then
         Set ItemList = frmGerenciarNFSe.lvwNFSe.ListItems.Add(, "R" & CStr(Contador), StrZero(rsDados!Numero, 8))
'''''      End If
      
      ItemList.SubItems(1) = Format(rsDados!DataEmissao, "DD/MM/YYYY")
      ItemList.SubItems(2) = Trim(rsDados!Nome)
      ItemList.SubItems(3) = Format(rsDados!ValorServicos, "##,##0.00")
      ItemList.SubItems(4) = sSituacao
      
      Contador = Contador + 1
      rsDados.MoveNext
   Loop
   
   ' Notas
   If I_SGBD = "SQLSERVER" Then
      Set rsNFsInutilizadas = cnSistema.Execute("SELECT * FROM " & I_TabelasNF & "Inutilizadas WHERE Data >= '" & Format(frmGerenciarNF.mskDataInicial.Text, "yyyy-mm-dd") & "' AND Data <= '" & Format(frmGerenciarNF.mskDataFinal.Text, "yyyy-mm-dd") & "' Order By Numero")
   ElseIf I_SGBD = "ACCESS" Then
      Set rsNFsInutilizadas = cnSistema.Execute("SELECT * FROM " & I_TabelasNF & "Inutilizadas WHERE Data >= cDate('" & Format(frmGerenciarNF.mskDataInicial.Text, "dd/mm/yyyy") & "') AND Data <= cDate('" & Format(frmGerenciarNF.mskDataFinal.Text, "dd/mm/yyyy") & "') Order By Numero")
   End If

   If Not rsNFsInutilizadas.EOF Then
      Do While Not rsNFsInutilizadas.EOF

'         Set ItemList = lvwNFs.ListItems.Add(, "I" & CStr(rsNFsInutilizadas!Numero), StrZero(rsNFsInutilizadas!Numero, 8))

         Set ProcuraItem = frmGerenciarNF.lvwNFs.FindItem(StrZero(rsNFsInutilizadas!Numero, 8))
         If ProcuraItem Is Nothing Then
            Set ItemList = frmGerenciarNF.lvwNFs.ListItems.Add(, "I" & CStr(Contador), StrZero(rsNFsInutilizadas!Numero, 8))
            ItemList.SubItems(1) = rsNFsInutilizadas!Data
            ItemList.SubItems(2) = "Nota Inutilizada"
            ItemList.SubItems(3) = Format(0, "##,##0.00")
            ItemList.SubItems(4) = "Inutilizada"
            If IsNull(rsNFsInutilizadas!Protocolo) Or rsNFsInutilizadas!Protocolo = "" Then
               ItemList.SubItems(5) = "Aguarde..."
            Else
               ItemList.SubItems(5) = ""
            End If
         End If

         Contador = Contador + 1
         rsNFsInutilizadas.MoveNext
      Loop
   End If
   
   Exit Sub
Erro:
   Call FRetornaMensagens(Err.Number & " - " & Err.Description & " - Carrega_View_NFSe ")
End Sub

Public Sub Carrega_View_MDFe(sModo As String)
On Error GoTo Erro

Dim sSituacao As String
Dim Contador As Integer

Dim sqlSituacao As String
Dim sSql As String

Dim rsDados As New ADODB.Recordset
Dim rsNFsInutilizadas As New ADODB.Recordset
Dim I_TabelaMDFe As String

   I_TabelaMDFe = "MDFe"
   
'   cmdCancelarNota.Enabled = False
'   cmdTransmitir.Enabled = False
   frmGerenciarMDFes.cmdImprimirDANFE.Enabled = False
   
   ' Validar Per�odo
   If Not FValidarPeriodo(frmGerenciarMDFes.mskDataInicial.Text, frmGerenciarMDFes.mskDataFinal.Text) Then Exit Sub
   ' Retorna condi��o do SQL
   sqlSituacao = FSituacao("C", frmGerenciarMDFes.cmbSituacao.ListIndex)
   ' Carrega View
   If sModo = "Carregar" Then frmGerenciarMDFes.lvwMDFes.ListItems.Clear
                            
   ' Notas
   Contador = 1

   sSql = "       SELECT "
   sSql = sSql & vbCrLf & "      N.id" & I_TabelaMDFe & ", "
'''''   sSql = sSql & vbCrLf & "      N.idCliente, "
   sSql = sSql & vbCrLf & "      N.Numero, "
   sSql = sSql & vbCrLf & "      N.DataEmissao, "
'''''   sSql = sSql & vbCrLf & "      C.Nome, "
   sSql = sSql & vbCrLf & "      UFsCarregamento.Sigla AS UFCarregar, "
   sSql = sSql & vbCrLf & "      UFsDescarregamento.Sigla AS UFDescarregar, "
   sSql = sSql & vbCrLf & "      N.Situacao "
'''''   sSql = sSql & vbCrLf & "      N.ValorServicos "
   sSql = sSql & vbCrLf & "FROM "
   If I_SGBD = "SQLSERVER" Then
'''''      sSql = sSql & vbCrLf & "      " & I_TabelaMDFe & " N, ClientesInfFiscais C "
      sSql = sSql & vbCrLf & "      " & I_TabelaMDFe & " N, UFs UFsCarregamento, UFs UFsDescarregamento "
   ElseIf I_SGBD = "ACCESS" Then
      sSql = sSql & vbCrLf & "      " & I_TabelaMDFe & " N, UFs UFsCarregamento, UFs UFsDescarregamento "
   End If
   sSql = sSql & vbCrLf & "WHERE "
'''''   sSql = sSql & vbCrLf & "      N.idCliente = C.idCliente AND "
   sSql = sSql & vbCrLf & "      UFsCarregamento.idUF = N.idUF AND "
   sSql = sSql & vbCrLf & "      UFsDescarregamento.idUF = N.idUFDescarregamento AND "
   If I_SGBD = "SQLSERVER" Then
      sSql = sSql & vbCrLf & "      N.DataEmissao >= '" & Format(frmGerenciarMDFes.mskDataInicial.Text, "yyyy-mm-dd") & " 00:00:00' AND "
      sSql = sSql & vbCrLf & "      N.DataEmissao <= '" & Format(frmGerenciarMDFes.mskDataFinal.Text, "yyyy-mm-dd") & " 23:59:59' "
   ElseIf I_SGBD = "ACCESS" Then
      sSql = sSql & vbCrLf & "      N.DataEmissao >= cDate('" & Format(frmGerenciarMDFes.mskDataInicial.Text, "dd/mm/yyyy") & " 00:00:00') AND "
      sSql = sSql & vbCrLf & "      N.DataEmissao <= cDate('" & Format(frmGerenciarMDFes.mskDataFinal.Text, "dd/mm/yyyy") & " 23:59:59') "
   End If
   
   
   sSql = sSql & vbCrLf & sqlSituacao & " Order By N.Numero"
   
   Set rsDados = cnSistema.Execute(sSql)
   Do While Not rsDados.EOF
'      sSituacao = FSituacao("S", rsDados!Situacao)      ' Retorna Situa��o atual da nota
      sSituacao = 0      ' Retorna Situa��o atual da nota
      
      ' Pesquisa se registro existe
'''''      Set ItemList = frmGerenciarNF.lvwNFs.FindItem(StrZero(rsDados!Numero, 8))
'''''      If ItemList Is Nothing Then
         Set ItemList = frmGerenciarMDFes.lvwMDFes.ListItems.Add(, "R" & CStr(Contador), StrZero(rsDados!Numero, 8))
'''''      End If

      ItemList.SubItems(1) = Format(rsDados!DataEmissao, "DD/MM/YYYY")
      ItemList.SubItems(2) = Trim(rsDados!UFCarregar)
      ItemList.SubItems(3) = Trim(rsDados!UFDescarregar)
      ItemList.SubItems(4) = ""
      ItemList.SubItems(5) = sSituacao
      
      Contador = Contador + 1
      rsDados.MoveNext
   Loop
   
 
   Exit Sub
Erro:
   Call FRetornaMensagens(Err.Number & " - " & Err.Description & " - Carrega_View_MDFe ")
End Sub


''''''Sua rotina de multitarefa
'''''Public Sub MensagemAlerta()
'''''
'''''  MsgBox "O sistema est� ativo !", vbOKOnly, "Status do sistema..."
'''''
'''''End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''

Sub ConfigurarNFs()
'On Error GoTo Erro
Dim rsSistema As New ADODB.Recordset
'''''Dim rsEmpresa As New ADODB.Recordset
'''''Dim rsUFs As New ADODB.Recordset
'''''Dim rsMunicipios As New ADODB.Recordset
Dim rsVerificar As New ADODB.Recordset
Dim sSql As String

   ' Modelo da NF
   I_ModeloNF = LerArquivoINI("NFe", "Modelo", CaminhoINI & "\System.ini")
   If I_ModeloNF = "55" Then
      I_TabelasNF = "NFe"
      I_PastaUNINFe = "NF-e\"
   ElseIf I_ModeloNF = "65" Then
      I_TabelasNF = "NFCe"
      I_PastaUNINFe = "NFC-e\"
   End If

   ' Tabelas espec�ficas do sistema
   I_TabelaUFs = LerArquivoINI("Tabelas", "UFs", CaminhoINI & "\System.ini")
   I_TabelaMunicipios = LerArquivoINI("Tabelas", "Municipios", CaminhoINI & "\System.ini")

   ' Configurar Informa��es gerais da nota
   I_EmpresaNF = LerArquivoINI("NFe", "Empresa", CaminhoINI & "\System.ini")
   I_UnidadeNFe = LerArquivoINI("Arquivos", "UnidadeNFe", CaminhoINI & "\System.ini")
   I_CaminhoXML_NFCe = LerArquivoINI("Arquivos", "CaminhoXMLNFCe", CaminhoINI & "\System.ini")
   I_HorarioVerao = LerArquivoINI("NFe", "HorarioVerao", CaminhoINI & "\System.ini")

   ' Caminhos de Notas
   
''      sDestinoNF = I_UnidadeNFe & I_PastaUNINFe & I_EmpresaNF & "\Validar\" & rsNFs!cNF & "-nfe.XML"
   
   ARQUIVO_NFE_NOTAS = I_UnidadeNFe & I_PastaUNINFe & "Notas\Notas.TXT"

   CAMINHO_NFE_ENVIO = I_UnidadeNFe & I_PastaUNINFe & I_EmpresaNF & "\ENVIO\"
   CAMINHO_NFE_VALIDAR = I_UnidadeNFe & I_PastaUNINFe & I_EmpresaNF & "\VALIDAR\"
   CAMINHO_NFE_VALIDADO = I_UnidadeNFe & I_PastaUNINFe & I_EmpresaNF & "\VALIDAR\VALIDADO\"
   CAMINHO_NFE_RETORNO = I_UnidadeNFe & I_PastaUNINFe & I_EmpresaNF & "\RETORNO\"
   CAMINHO_NFE_ERROS = I_UnidadeNFe & I_PastaUNINFe & I_EmpresaNF & "\ERROS\"
   CAMINHO_NFE_TEMP = I_UnidadeNFe & I_PastaUNINFe & I_EmpresaNF & "\TEMP\"

   ' Determina vari�veis publica da empresa
'''''   Set rsEmpresa = cnSistema.Execute("Select * From Empresa")
   
   sSql = ""
   sSql = sSql & vbCrLf & "SELECT "
   sSql = sSql & vbCrLf & "   Empresa.*, "
   sSql = sSql & vbCrLf & "   UFs.Codigo AS CodigoUF, "
   sSql = sSql & vbCrLf & "   UFs.Sigla AS SiglaUF, "
   sSql = sSql & vbCrLf & "   Municipios.Codigo AS CodigoMunicipio, "
   sSql = sSql & vbCrLf & "   Municipios.Nome AS NomeMunicipio "
   sSql = sSql & vbCrLf & "From "
   sSql = sSql & vbCrLf & "   Empresa, "
   sSql = sSql & vbCrLf & "   UFs, "
   sSql = sSql & vbCrLf & "   Municipios "
   sSql = sSql & vbCrLf & "Where Empresa.idUF = UFs.idUF "
   sSql = sSql & vbCrLf & "   AND Empresa.idMunicipio = Municipios.idMunicipio"
   Set rsEmpresa = cnSistema.Execute(sSql)
   
   If Not rsEmpresa.EOF Then
'''''      Set rsUFs = cnSistema.Execute("Select * From UFs WHERE idUF=" & rsEmpresa!idUF)
'''''      Set rsMunicipios = cnSistema.Execute("Select * From Municipios WHERE idMunicipio=" & rsEmpresa!idMunicipio)
      
      I_idEmpresa = rsEmpresa!idEmpresa
      
      I_Empresa_CNPJ_CPF = RemoveCaracteres(rsEmpresa!CNPJ_CPF)
'''''      I_EmpresaCodigoUF = rsUFs!Codigo
      I_EmpresaCodigoUF = rsEmpresa!CodigoUF
'''''      I_EmpresaUF = rsUFs!Sigla
      I_EmpresaUF = rsEmpresa!SiglaUF
'''''      I_EmpresaCodigoMunicipio = RemoveCaracteres(rsMunicipios!Codigo)
      I_EmpresaCodigoMunicipio = RemoveCaracteres(rsEmpresa!CodigoMunicipio)
'''''      I_EmpresaMunicipio = rsMunicipios!Nome
      I_EmpresaMunicipio = rsEmpresa!NomeMunicipio
      
      I_AmbienteNF = LerArquivoINI("NFe", "Ambiente", CaminhoINI & "\System.ini")
   End If
   
'''''   Set rsEmpresa = Nothing
'''''   Set rsUFs = Nothing
'''''   Set rsMunicipios = Nothing
End Sub

Private Function FValidarPeriodo(DataInicial As String, DataFinal As String) As Boolean
On Error GoTo Erro
Dim strMensagem As String

   FValidarPeriodo = True

   If Not IsDate(frmGerenciarNF.mskDataInicial.Text) Then strMensagem = "Data Inicial Inv�lida" & Chr(13)
   If Not IsDate(frmGerenciarNF.mskDataFinal.Text) Then strMensagem = "Data Final Inv�lida" & Chr(13)
   If IsDate(frmGerenciarNF.mskDataInicial.Text) And IsDate(frmGerenciarNF.mskDataFinal.Text) Then
      If (CDate(frmGerenciarNF.mskDataInicial.Text) > CDate(frmGerenciarNF.mskDataFinal.Text)) Then strMensagem = "Data Inicial maior que Data Final"
   End If
   If Not strMensagem = Empty Then
      MsgBox "Verifique os Seguintes Campos:" & Chr(13) & strMensagem, vbExclamation + vbOKOnly, "Campos Obrigat�rios"
      FValidarPeriodo = False
      Exit Function
   End If

   Exit Function
   Resume
Erro:
   Call FRetornaMensagens(Err.Number & " - " & Err.Description & " - " & ".FValidarPeriodo")
End Function

Private Function FSituacao(Parametro As String, Conteudo As Long) As String
On Error GoTo Erro

   If Parametro = "C" Then       ' Condi��o
'      If Conteudo = 0 Then FSituacao = "AND Situacao <> 2"
      If Conteudo = 0 Then FSituacao = ""
      If Conteudo = 1 Then FSituacao = " AND N.Situacao = 0"
      If Conteudo = 2 Then FSituacao = " AND N.Situacao = 1"
      If Conteudo = 3 Then FSituacao = " AND N.Situacao = 2"
      If Conteudo = 4 Then FSituacao = " AND N.Situacao = 3"
      If Conteudo = 5 Then FSituacao = " AND N.Situacao = 4"
      If Conteudo = 6 Then FSituacao = " AND N.Situacao = 5"
      If Conteudo = 7 Then FSituacao = " AND N.Situacao <> 2 AND N.Situacao <> 3"
   End If
   
   If Parametro = "S" Then       ' Condi��o
      If Conteudo = 0 Then FSituacao = "Em Digita��o"
      If Conteudo = 1 Then FSituacao = "Processamento"
      If Conteudo = 2 Then FSituacao = "Aprovada"
      If Conteudo = 3 Then FSituacao = "Cancelada"
      If Conteudo = 4 Then FSituacao = "N�o Emitida"
      If Conteudo = 5 Then FSituacao = "Denegada"
      If Conteudo = 9 Then FSituacao = "Pendentes"
   End If
   
   Exit Function
Erro:
   Call FRetornaMensagens(Err.Number & " - " & Err.Description & " - " & ".FSituacao")
End Function

Public Function FArquivosNF(ByVal sOpcao As String, Optional ByVal sArquivo As String) As String
On Error GoTo Erro
Dim sVerArquivo As String

Dim sOrigemNF As String
Dim sDestinoNF As String
Dim sImpressoraUSB As String

Dim bCopiar As Boolean
   
'''   Set rsNFs = cnSistema.Execute("Select * From " & I_TabelasNF & " WHERE Numero=" & Val(frmGerenciarNF.lvwNFs.ListItems(frmGerenciarNF.lvwNFs.SelectedItem.Index)))

'''''   I_EmpresaNF = LerArquivoINI("NFe", "Empresa", CaminhoINI & "\System.ini")
   sOrigemNF = I_UnidadeNFe & I_PastaUNINFe & "Notas\Notas.TXT"

   bCopiar = True
   
   If sOpcao = "VALIDAR" Then
      sDestinoNF = I_UnidadeNFe & I_PastaUNINFe & I_EmpresaNF & "\Validar\" & rsNFs!cNF & "-nfe.XML"
      
   ElseIf sOpcao = "ENCERRARMDF" Then
      sDestinoNF = I_UnidadeNFe & I_PastaUNINFe & I_EmpresaNF & "\Envio\encerramento110112" & Trim(rsMDFes!cMDF) & "01-ped-eve.XML"
      
   ElseIf sOpcao = "TRANSMITIRMDFe" Then
      sDestinoNF = I_UnidadeNFe & I_PastaUNINFe & I_EmpresaNF & "\Envio\" & Trim(rsMDFes!cMDF) & "-mdfe.XML"

   ElseIf sOpcao = "TRANSMITIR" Then
      sOrigemNF = I_UnidadeNFe & I_PastaUNINFe & I_EmpresaNF & "\Validar\Validado\" & Trim(rsNFs!cNF) & "-nfe.XML"
      sDestinoNF = I_UnidadeNFe & I_PastaUNINFe & I_EmpresaNF & "\Envio\" & Trim(rsNFs!cNF) & "-nfe.XML"
      
   ElseIf sOpcao = "IMPRIMIRDANFE" Then
      '...\unidanfe.exe a=�c:\x\0101-procNFe.xml� v=0 m=1 // imprimir sem visualizar
      '...\unidanfe.exe a=�c:\x\0101-procNFe.xml� v=1 m=0 // visualizar sem imprimir
      '...\unidanfe.exe a=�c:\x\0101-procNFe.xml� v=0 m=0 // envia e-mail sem visualizar ou imprimir
   
      sDestinoNF = I_UnidadeNFe & I_PastaUNINFe & I_EmpresaNF & "\Enviados\Autorizados\" & Format(rsNFs!DataEmissao, "yyyymm") & "\" & rsNFs!cNF & "-procNFe.XML"
      Shell "C:\UNIMAKE\" & I_EmpresaNF & "\UNIDANFE\UNIDANFE.EXE arquivo=" & sDestinoNF & " " & LerArquivoINI("NFe", "UNIDANFE", CaminhoINI & "\System.ini")
      bCopiar = False
      
   ElseIf sOpcao = "GERARPDF" Then
      sImpressoraUSB = LerArquivoINI("NFe", "ImpressoraUSB", CaminhoINI & "\System.ini")
      
      sDestinoNF = I_UnidadeNFe & I_PastaUNINFe & I_EmpresaNF & "\Enviados\Autorizados\" & Format(rsNFs!DataEmissao, "yyyymm") & "\" & rsNFs!cNF & "-procNFe.XML"
      Shell "C:\UNIMAKE\" & I_EmpresaNF & "\UNIDANFE\UNIDANFE.EXE arquivo=" & sDestinoNF & " visualizar = 0 " & "i=" & sImpressoraUSB
      bCopiar = False
   
   ElseIf sOpcao = "CARTACORRECAO" Then
      ''Corrigir o nome do arquivo "-env-canc.xml" para corre��o
      ''      sDestinoNF = I_UnidadeNFe & I_PastaUNINFe & I_EmpresaNF & "\Envio\" & Trim(rsNFs!ChaveNFCe) & "-env-cce.xml"
      'cce35111253420477000192550550000033071213028272_01-ped-eve
      sDestinoNF = I_UnidadeNFe & I_PastaUNINFe & I_EmpresaNF & "\Envio\cce" & Trim(rsNFs!cNF) & "_01-ped-eve.xml"

   ElseIf sOpcao = "CANCELARNOTA" Then
      sDestinoNF = I_UnidadeNFe & I_PastaUNINFe & I_EmpresaNF & "\Envio\" & Trim(rsNFs!cNF) & "-env-canc.xml"
   
   ElseIf sOpcao = "INUTILIZARNUMERO" Then
      sDestinoNF = I_UnidadeNFe & I_PastaUNINFe & I_EmpresaNF & "\Envio\" & StrZero(Val(rsNFs!nnf), 12) & "-ped-inu.txt"

   ElseIf sOpcao = "CONSULTARNUMERO" Then
      sDestinoNF = I_UnidadeNFe & I_PastaUNINFe & I_EmpresaNF & "\Envio\" & Trim(rsNFs!cNF) & "-ped-sit.xml"
   
   ElseIf sOpcao = "ARQUIVOPROCESSAMENTO" Then
      sOrigemNF = Dir(I_UnidadeNFe & I_PastaUNINFe & I_EmpresaNF & "\Erros\" & Trim(rsNFs!cNF) & "-nfe.xml")
      If sOrigemNF = (rsNFs!cNF & "-nfe.xml") Then
         sOrigemNF = I_UnidadeNFe & I_PastaUNINFe & I_EmpresaNF & "\Erros\" & Trim(rsNFs!cNF) & "-nfe.xml"
         sDestinoNF = I_UnidadeNFe & I_PastaUNINFe & I_EmpresaNF & "\Enviados\EmProcessamento\" & Trim(rsNFs!cNF) & "-nfe.xml"
      Else
         bCopiar = False
      End If
      
   ElseIf sOpcao = "RETORNOS" Then
      sOrigemNF = I_UnidadeNFe & I_PastaUNINFe & I_EmpresaNF & "\Retorno\" & Trim(sArquivo)
      sDestinoNF = I_UnidadeNFe & I_PastaUNINFe & I_EmpresaNF & "\Temp\" & Trim(sArquivo)
   End If
   
   ' Se for copia e exclus�o de arquivo
   If bCopiar Then
      FileCopy sOrigemNF, sDestinoNF
      Kill sOrigemNF
   End If
   
   Exit Function
Erro:
'   MsgBox "Erro " & Err & ". " & Err.Description & " - " & TypeName(Me) & ".FTAG"
   Call FRetornaMensagens(Err.Number & " - " & Err.Description & ".FArquivosNF")
End Function

Public Function FData(sTipo As String, ByRef sData As String, ByRef sHora As String) As String
On Error GoTo Erro
Dim sHorarioVerao As String

   sHorarioVerao = LerArquivoINI("NFe", "HorarioVerao", App.Path & "\System.ini")
   
   If sTipo = "P" Then
      FData = Format(sData, "YYYY-MM-DD") & "T" & Format(sHora, "HH:MM:SS") & IIf(sHorarioVerao, "-02:00", "-03:00")
   ElseIf sTipo = "S" Then
      FData = Format(sData, "YYYY-MM-DD") & " " & Format(sHora, "HH:MM:SS")
   End If
   
   Exit Function
Erro:
'   MsgBox "Erro " & Err & ". " & Err.Description & " - " & TypeName(Me) & ".FData"
   Call FRetornaMensagens("Erro " & Err.Description)
End Function

Public Function FFormataChaveNF(ByRef sChave As String) As String
On Error GoTo Erro
Dim sFormatarChave As String
   
   sFormatarChave = Mid(sChave, 1, 4) & "."
   sFormatarChave = sFormatarChave & Mid(sChave, 5, 4) & "."
   sFormatarChave = sFormatarChave & Mid(sChave, 9, 4) & "."
   sFormatarChave = sFormatarChave & Mid(sChave, 13, 4) & "."
   sFormatarChave = sFormatarChave & Mid(sChave, 17, 4) & "."
   sFormatarChave = sFormatarChave & Mid(sChave, 21, 4) & "."
   sFormatarChave = sFormatarChave & Mid(sChave, 25, 4) & "."
   sFormatarChave = sFormatarChave & Mid(sChave, 29, 4) & "."
   sFormatarChave = sFormatarChave & Mid(sChave, 33, 4) & "."
   sFormatarChave = sFormatarChave & Mid(sChave, 37, 4) & "."
   sFormatarChave = sFormatarChave & Mid(sChave, 41, 4) & " "
   
   FFormataChaveNF = sFormatarChave
   
   Exit Function
Erro:
'   MsgBox "Erro " & Err & ". " & Err.Description & " - " & TypeName(Me) & ".FFormataChaveNF"
End Function

Public Function FVerificaArquivo(ByRef sArquivo As String) As String
On Error GoTo Erro
   
   If Not Dir(sArquivo) = "" Then
      FVerificaArquivo = sArquivo
   Else
      FVerificaArquivo = ""
   End If
   
   Exit Function
Erro:
   FVerificaArquivo = "Erro"
'   MsgBox "Erro " & Err & ". " & Err.Description & " - " & TypeName(Me) & ".FFormataChaveNF"
End Function

Public Function FRetornaMensagens(ByVal sMensagem As String, _
                         Optional ByVal sCodigo As String, _
                         Optional ByVal sArquivo As String, _
                         Optional ByVal sChave As String)
On Error GoTo Erro

   I_ContarMensagem = I_ContarMensagem + 1
   
   Set ItemList = frmGerenciarNF.lvwMensagens.ListItems.Add(, "R" & CStr(I_ContarMensagem), IIf(Not IsNull(sChave), sChave, ""))
       ItemList.SubItems(1) = IIf(Not IsNull(sCodigo), sCodigo, "")
       ItemList.SubItems(2) = sMensagem
       ItemList.SubItems(3) = IIf(Not IsNull(sArquivo), sArquivo, "")
   
'''''   Set ItemList = frmGerenciarNF.lvwMensagens.ListItems.Add(, "R" & CStr(I_ContarMensagem), sMensagem)
'''''       ItemList.SubItems(1) = I_ContarMensagem
'''''       ItemList.SubItems(2) = sMensagem
'''''       ItemList.SubItems(3) = IIf(Not IsNull(sArquivo), sArquivo, "")

   Exit Function
Erro:
   Call FRetornaMensagens("Erro " & Err & ". " & Err.Description)
End Function

Public Function FErrosValidacaoNF(ByVal sMensagem As String)
On Error GoTo Erro
   
   Set ItemList = frmGerenciarNF.lvwErrosValidacaoNF.ListItems.Add(, "R" & CStr(frmGerenciarNF.lvwErrosValidacaoNF.ListItems.Count + 1), sMensagem)
'''''       ItemList.SubItems(1) = CStr(frmGerenciarNF.lvwErrosValidacaoNF.ListItems.Count + 1)
'''''       ItemList.SubItems(2) = sMensagem
'''''       ItemList.SubItems(3) = ""

   Exit Function
Erro:
   Call FRetornaMensagens("Erro " & Err & ". " & Err.Description)
End Function

Public Function FAtualizaNF(ByVal lNumero As Long, ByVal sSituacao As String) As String
On Error GoTo Erro

   ' 1. Em Processamento
   ' 2.
   ' 3.
   ' 4. N�o Emitida
   
'   If sSituacao = 1 Then   ' Em Processamento
      cnSistema.Execute "UPDATE " & I_TabelasNF & " SET " & _
               "TentativaEmissao = 0, " & _
               "Situacao = " & sSituacao & " " & _
               "WHERE Numero = " & lNumero

'   End If
   
   Exit Function
Erro:
'   MsgBox "Erro " & Err & ". " & Err.Description & " - " & TypeName(Me) & ".FTAG"
   Call FRetornaMensagens(Err.Number & " - " & Err.Description & ".FAtualizaNF")
End Function

Public Function FNivelTAG(ByVal iNivel As Integer)

   FNivelTAG = Space(NivelTAG(iNivel))
   
End Function

Public Function FTAG(ByRef sTAG As String, ByRef sConteudo As String) As String
On Error GoTo Erro
   
'   FTAG = "<" & sTAG & ">" & Trim(sConteudo) & "</" & sTAG & ">" '& vbLf
   FTAG = "<" & sTAG & ">" & UTF8_Encode(Trim(sConteudo)) & "</" & sTAG & ">"  '& vbLf

   Exit Function
Erro:
'   Call FRetornaMensagens(Err.Number & " - " & Err.Description & " - " & TypeName(Me) & ".FTAG")
End Function

Public Function FTAGPrincipal(sNome As String) As String
On Error GoTo Erro
   
   If sNome = "Versao" Then
      FTAGPrincipal = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "UTF-8" & Chr(34) & "?>"
   ElseIf sNome = "EnvioEvento" Then
      FTAGPrincipal = "<envEvento xmlns=" & Chr(34) & "http://www.portalfiscal.inf.br/nfe" & Chr(34) & " versao=" & Chr(34) & "1.00" & Chr(34) & ">"
   ElseIf sNome = "EnvioEventoMDFe" Then
      FTAGPrincipal = "<envEventoMDFe xmlns=" & Chr(34) & "http://www.portalfiscal.inf.br/mdfe" & Chr(34) & " versao=" & Chr(34) & "3.00" & Chr(34) & ">"
   ElseIf sNome = "Evento" Then
      FTAGPrincipal = "<evento xmlns=" & Chr(34) & "http://www.portalfiscal.inf.br/nfe" & Chr(34) & " versao=" & Chr(34) & "1.00" & Chr(34) & ">"
   End If
   
   Exit Function
Erro:
'   Call FRetornaMensagens(Err.Number & " - " & Err.Description & " - " & TypeName(Me) & ".FTAGPrincipal")
End Function



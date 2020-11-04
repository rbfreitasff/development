VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CImportarLE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------------
' Classe Gerar NFs
' Autor: Robinson
' Data da Criação: 08/08/2019
'----------------------------------------------------------------------------------------------

Public Function InserirNF()
Dim vGravacao As String
Dim rsNFe As New ADODB.Recordset
Dim rsClientes As New ADODB.Recordset
Dim rsSaidasItens As New ADODB.Recordset

Dim iIdProduto As Integer
Dim dDataEmissao As String
Dim dQuantidade As Double
Dim dDesconto As Double
Dim dValorUnitario As Double
Dim dICMS As Double
Dim dISS As Double
Dim dBaseReduzida As Double
Dim sDescricaoComplementar As String
Dim iIdUnidade As String
Dim iIdSituacaoTributaria As Integer
Dim sDiscriminacaoProduto As String
Dim dIPI As Double
Dim dBaseReduzidaIPI As Double
Dim sClassificacaoFiscal As String
Dim dValorFrete As Double
Dim sCFOP As String

Dim iNumero As Long
Dim iCliente As Integer
Dim iCFOP As Integer
Dim dCalculoDesconto As Double

   Set rsSaidas = cnSistema.Execute("Select * From Saidas WHERE GN = 1")
   Do While Not rsSaidas.EOF
      vGravacao = ""

      Set rsNFe = cnSistema.Execute("Select TOP 1 * From NFe ORDER BY idNFe DESC")
      If Not rsNFe.BOF Or Not rsNFe.EOF Then
         iNumero = rsNFe!Numero + 1
      Else
         iNumero = 1
      End If

'''      Set rsClientes = cnSistema.Execute("Select * From Clientes WHERE idCliente = " & rsSaidas!idCliente)
      Set rsClientes = cnSistema.Execute("Select * From ClientesInfFiscais WHERE idCliente = " & rsSaidas!idCliente)
      If rsClientes!SiglaUF = "GO" Then
         iCFOP = 251   ' 5102 - Buscar o id Na Base
      Else
         iCFOP = 381   ' 6102 - Buscar o id Na Base
      End If

      vGravacao = ""
      vGravacao = vGravacao & "Numero = " & iNumero & ", "
      vGravacao = vGravacao & "Cupom = " & 0 & ", "
      vGravacao = vGravacao & "idCliente = " & rsClientes!idCliente & ", "
      vGravacao = vGravacao & "idNaturezaOperacao = " & 1 & ", "  ' Venda de Mercadorias e serviços
      vGravacao = vGravacao & "idCFOP = " & iCFOP & ", "          ' 5102 / 6102 - Venda de mercadoria dentro ou fora do estado
      vGravacao = vGravacao & "DadosAdicionais = '', "
      vGravacao = vGravacao & "DataEmissao = '" & Format(Date, "yyyy-mm-dd") & " " & Format(Time, "hh:mm:ss") & "', "
      vGravacao = vGravacao & "DataCaixa = '" & Format(Date, "yyyy-mm-dd") & " " & Format(Time, "hh:mm:ss") & "', "
      vGravacao = vGravacao & "DataVencimento = '" & Format(Date, "yyyy-mm-dd") & " " & Format(Time, "hh:mm:ss") & "', "
      vGravacao = vGravacao & "Hora = '" & Format(Date, "yyyy-mm-dd") & " " & Format(Time, "hh:mm:ss") & "', "
      vGravacao = vGravacao & "BaseCalculoICMS = " & 0 & ", "
      vGravacao = vGravacao & "ValorICMS = " & 0 & ", "
      vGravacao = vGravacao & "ValorFrete = " & 0 & ", "
      vGravacao = vGravacao & "ValorTotalProdutos = " & 0 & ", "
      vGravacao = vGravacao & "BaseICMSSubstituicao = " & 0 & ", "
      vGravacao = vGravacao & "ValorICMSSubstituicao = " & 0 & ", "
      vGravacao = vGravacao & "OutrasDespesas = " & 0 & ", "
      vGravacao = vGravacao & "ValorTotalNota = " & 0 & ", "
      vGravacao = vGravacao & "idTransportador = " & 0 & ", "
      vGravacao = vGravacao & "FreteConta = " & 9 & ", "          ' Não se aplica
      vGravacao = vGravacao & "PlacaVeiculo = '   -    ', "
      vGravacao = vGravacao & "UFCaminhao = '', "
      vGravacao = vGravacao & "VolumeQuantidade = '', "
      vGravacao = vGravacao & "VolumeMarca = '', "
      vGravacao = vGravacao & "VolumeEspecie = '', "
      vGravacao = vGravacao & "VolumeNumero = '', "
      vGravacao = vGravacao & "VolumePesoBruto = 0, "
      vGravacao = vGravacao & "VolumePesoLiquido = 0, "
      vGravacao = vGravacao & "InformacoesCorpo = '" & LerArquivoINI("Notas Fiscais", "InformacoesCorpo", CaminhoINI & "\System.ini") & "', "
      vGravacao = vGravacao & "idFormaPagamento = " & 1 & ", "    ' Avista
      vGravacao = vGravacao & "DescontoGeral = " & 0 & ", "
      vGravacao = vGravacao & "Bonificacao = " & 0 & ", "
      vGravacao = vGravacao & "Documento = '', "
      vGravacao = vGravacao & "Observacao = '', "
      vGravacao = vGravacao & "GeradaNFe = 0, "
      vGravacao = vGravacao & "Situacao = " & 0 & ", "
      vGravacao = vGravacao & "TentativaEmissao = " & 1 & ", "
      vGravacao = vGravacao & "NumeroNFeComplementar = 0, "
      vGravacao = vGravacao & "ChaveAcessoNFeComplementar = '', "
      vGravacao = vGravacao & "ChaveAcessoDevolucao = '' "

      ' Inserir
      cnSistema.Execute "Insert Into NFe (Numero) Values (" & iNumero & ")"
      cnSistema.Execute "Update NFe SET " & vGravacao & " Where Numero = " & iNumero

      ' Inserir Itens
      Set rsNFe = cnSistema.Execute("Select * From NFe WHERE Numero = " & iNumero)
      If Not rsNFe.EOF Then
         Set rsSaidasItens = cnSistema.Execute("Select * From SaidasItens WHERE idSaida = " & rsSaidas!idSaida)
         
         bArredondado = False
   
         Do While Not rsSaidasItens.EOF
            iIdNFe = rsNFe!idNFe
            iIdProduto = rsSaidasItens!idProduto
'''''            dDataEmissao = Format(Date, "yyyy-mm-dd") & " " & Format(Time, "hh:mm:ss")
            dDataEmissao = Format(Date, "yyyy-mm-dd")
            dQuantidade = rsSaidasItens!Quantidade
            dDesconto = 0
'''''            dValorUnitario = rsSaidasItens!ValorUnitario
'''''            dValorUnitario = (((rsSaidasItens!Quantidade * rsSaidasItens!ValorUnitario) - rsSaidasItens!Desconto - rsSaidas!Arredondamento) / rsSaidasItens!Quantidade)

            dCalculoDesconto = ((rsSaidasItens!Quantidade * rsSaidasItens!ValorUnitario) * (rsSaidasItens!Desconto / 100))
            dValorUnitario = (((rsSaidasItens!Quantidade * rsSaidasItens!ValorUnitario) - dCalculoDesconto) / rsSaidasItens!Quantidade)

            dBaseReduzida = 0
            sDescricaoComplementar = ""
            iIdUnidade = "UN"                                                 ' Unidades
            sDiscriminacaoProduto = ""
            dIPI = 0
            dBaseReduzidaIPI = 0
            sClassificacaoFiscal = ""
            dValorFrete = 0
            sCFOP = IIf(rsClientes!SiglaUF = "GO", "5.102", "6.102")            ' Vendas
            iIdSituacaoTributaria = 12                                     ' Tributado sem permissão de crédito
            dICMS = 0
            dISS = 0

            cnSistema.Execute "Insert Into NFeItens (idNFe,idProduto,Data,Quantidade,Desconto,ValorUnitario,ICMS,ISS,BaseReduzida,DescricaoComplementar,Unidade,idSituacaoTributaria,DiscriminacaoProduto,IPI,BaseReduzidaIPI,ClassificacaoFiscal,ValorFrete,CFOP) " & _
                              "Values (" & iIdNFe & "," & iIdProduto & ",'" & dDataEmissao & "','" & dQuantidade & "','" & 0 & "','" & Substitui(dValorUnitario, ",", ".") & "','" & _
                              dICMS & "','" & dISS & "','" & dBaseReduzida & "','" & SQLCheck(sDescricaoComplementar) & "','" & iIdUnidade & "'," & iIdSituacaoTributaria & ",'" & _
                              SQLCheck(sDiscriminacaoProduto) & "','" & dIPI & "','" & dBaseReduzidaIPI & "','" & SQLCheck(sClassificacaoFiscal) & "','" & dValorFrete & "','" & sCFOP & "')"

            Set rsDados = Nothing

            rsSaidasItens.MoveNext
         Loop
      End If

      rsSaidas.MoveNext
   Loop

   Set rsSaidas = cnSistema.Execute("UPDATE Saidas SET GN = 0 WHERE GN = 1")


End Function

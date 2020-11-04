VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CNFSe200"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------------------------'
' Classe MDF-e Versão 4.00
' Autor: Robinson
' Data da Criação: 23/12/2019
'------------------------------------------------------------------------------------------------------'

Option Explicit

Public Function FEncerramentoMDFe(ByVal iNumero As Long) As String
On Error GoTo Erro
Dim rsMDFe As New ADODB.Recordset
Dim rsUFs As New ADODB.Recordset
Dim sHora As String

   sHora = Format(Time - TimeValue("00:10"), "hh:mm")

   Set rsMDFe = cnSistema.Execute("Select * From MDFe WHERE Numero=" & iNumero)
   If Not rsMDFe.EOF Then
   
   
'<?xml version="1.0" encoding="utf-8"?>
'<eventoMDFe versao="3.00" xmlns="http://www.portalfiscal.inf.br/mdfe">
'  <infEvento Id="ID1101113511031029073900013955001000000001105112804102">
'    <cOrgao>35</cOrgao>
'    <tpAmb>2</tpAmb>
'    <CNPJ>10290739000139</CNPJ>
'    <chMDFe>35110310290739000139550010000000011051128041</chMDFe>
'    <dhEvento>2013-10-31T09:44:20-02:00</dhEvento>
'    <tpEvento>110112</tpEvento>
'    <nSeqEvento>1</nSeqEvento>
'    <detEvento versaoEvento="3.00">
'      <evEncMDFe>
'        <descEvento>Encerramento</descEvento>
'        <nProt>010101010101010</nProt>
'        <dtEnc>2013-10-31</dtEnc>
'        <cUF>35</cUF>
'        <cMun>4118402</cMun>
'      </evEncMDFe>
'    </detEvento>
'  </infEvento>
'</eventoMDFe>
   
   
      Open I_UnidadeNFe & I_PastaUNINFe & "Notas\Notas.TXT" For Output As #1
      
      Print #1, FTAGPrincipal("Versao")
      Print #1, FTAGPrincipal("EnvioEventoMDFe")
      Print #1, "<infEvento Id=" & Chr(34) & "ID110111" & rsNFs!cNF & "01" & Chr(34) & ">"
      Print #1, FTAG("cOrgao", I_EmpresaCodigoUF)
      Print #1, FTAG("tpAmb", I_AmbienteNF)                                                             ' 1 - Produção ou 2 - Homologação
      Print #1, FTAG("CNPJ", I_Empresa_CNPJ_CPF)
      Print #1, FTAG("chMDFe", rsNFs!cNF)
      Print #1, FTAG("dhEvento", FData("P", Date, sHora))
      Print #1, FTAG("tpEvento", "110111")
      Print #1, FTAG("nSeqEvento", "1")
      Print #1, "<detEvento versaoEvento=" & Chr(34) & "3.00" & Chr(34) & ">"
      Print #1, "<evEncMDFe>"
      Print #1, FTAG("descEvento", "Encerramento")
      
'        <nProt>010101010101010</nProt>
      Print #1, FTAG("dtEnc", Format(Date, "yyyy-mm-dd"))
'''''   rsNFsEmitentes!UF = rsEmpresa!UF
      Print #1, FTAG("cUF", I_EmpresaCodigoUF)
      
      Print #1, FTAG("cMun", RemoveCaracteres(rsEmpresa!CodigoMunicipio))
      Print #1, "</evEncMDFe>"
      Print #1, "</detEvento>"
      Print #1, "</infEvento>"
      Print #1, "</eventoMDFe>"
      Close #1

      Call FArquivosNF("ENCERRARMDF")
   End If

   Exit Function
Erro:
   Call FRetornaMensagens(Err.Number & " - " & Err.Description & " - " & TypeName(Me) & ".FCartaCorrecao")
End Function

Private Function FTAG(ByRef sTAG As String, ByRef sConteudo As String) As String
On Error GoTo Erro
   
   FTAG = "<" & sTAG & ">" & Trim(sConteudo) & "</" & sTAG & ">" '& vbLf
   
   Exit Function
Erro:
   Call FRetornaMensagens(Err.Number & " - " & Err.Description & " - " & TypeName(Me) & ".FTAG")
End Function

Private Function FTAGPrincipal(sNome As String) As String
On Error GoTo Erro
   
   If sNome = "Versao" Then
      FTAGPrincipal = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "UTF-8" & Chr(34) & "?>"
   ElseIf sNome = "EnvioEvento" Then
      FTAGPrincipal = "<envEvento xmlns=" & Chr(34) & "http://www.portalfiscal.inf.br/nfe" & Chr(34) & " versao=" & Chr(34) & "1.00" & Chr(34) & ">"
   ElseIf sNome = "EnvioEventoMDFe" Then
      FTAGPrincipal = "<eventoMDFe xmlns=" & Chr(34) & "http://www.portalfiscal.inf.br/mdfe" & Chr(34) & " versao=" & Chr(34) & "3.00" & Chr(34) & ">"
   ElseIf sNome = "Evento" Then
      FTAGPrincipal = "<evento xmlns=" & Chr(34) & "http://www.portalfiscal.inf.br/nfe" & Chr(34) & " versao=" & Chr(34) & "1.00" & Chr(34) & ">"
   End If
   
   Exit Function
Erro:
   Call FRetornaMensagens(Err.Number & " - " & Err.Description & " - " & TypeName(Me) & ".FTAGPrincipal")
End Function




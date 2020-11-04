Attribute VB_Name = "modDados"
'---------------------------------------------------------------------------------------
' Procedure : Modulo de Busca de dados
' Author    : Robinson.Fernandes
' Date      : 04/06/2018
' Purpose   :
'---------------------------------------------------------------------------------------

Public Function FBuscaUF(ByVal id As Long, Optional ByVal iRetorno As Integer) As String
On Error GoTo Erro
Dim rsDados As New ADODB.Recordset

   ' iRetorno - Tipo de retorno
   '  . Sigla
   ' 1. Sigla
   ' 2. Codigo
   ' 3. Nome
   
   Set rsDados = cnSistema.Execute("Select * From UFs WHERE idUF=" & id)
   If Not rsDados.EOF Then
      If iRetorno = 1 Then
         FBuscaUF = rsDados!Sigla
      ElseIf iRetorno = 2 Then
         FBuscaUF = rsDados!Codigo
      ElseIf iRetorno = 3 Then
         FBuscaUF = rsDados!Nome
      Else
         FBuscaUF = rsDados!Sigla
      End If
   End If
   Set rsDados = Nothing

   Exit Function
Erro:
   Call FRetornaMensagens(Err.Number & " - " & Err.Description & ".FBuscaUF")
End Function

Public Function FBuscaMunicipio(ByVal id As Long, Optional ByVal iRetorno As Integer) As String
On Error GoTo Erro
Dim rsDados As New ADODB.Recordset

   ' iRetorno - Tipo de retorno
   '  . Nome
   ' 1. Nome
   ' 2. Codigo
   ' 3. UF
   
   Set rsDados = cnSistema.Execute("Select * From Municipios WHERE idMunicipio=" & id)
   If Not rsDados.EOF Then
      If iRetorno = 1 Then
         FBuscaMunicipio = rsDados!Nome
      ElseIf iRetorno = 2 Then
         FBuscaMunicipio = RemoveCaracteres(rsDados!Codigo)
      ElseIf iRetorno = 3 Then
         FBuscaMunicipio = rsDados!UF
      Else
         FBuscaMunicipio = rsDados!Nome
      End If
   End If
   Set rsDados = Nothing

   Exit Function
Erro:
   Call FRetornaMensagens(Err.Number & " - " & Err.Description & ".FBuscaMunicipio")
End Function


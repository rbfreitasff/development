Attribute VB_Name = "modThreads"
''Option Explicit
''Dim ItemList As ListItem
''Dim ProcuraItem As ListItem
''
''Declare Function CreateThread Lib "kernel32" (lpThreadAttributes As Any, ByVal dwStackSize As Long, ByVal lpStartAddress As Long, lpParameter As Any, ByVal dwCreationFlags As Long, lpThreadID As Long) As Long
''Declare Function TerminateThread Lib "kernel32" (ByVal hThread As Long, ByVal dwExitCode As Long) As Long
''Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
''
'''Variáveis para servir de identificador (handle) para o Thread
''
''Public hThread1 As Long, hThread1_ID As Long
''
''Public hThread2 As Long, hThread2_ID As Long
''
''''''''Suas rotinas de multitarefa
'''''''Public Sub Tarefa1()
'''''''Dim i As Long
'''''''
'''''''    Do While True
'''''''        Form1.botao1.Caption = i
'''''''        i = i + 1
'''''''        If i > 32000000 Then i = 0
'''''''    Loop
'''''''
'''''''End Sub
'''''''
'''''''Public Sub Tarefa2()
'''''''Dim i As Long
'''''''
'''''''    Do While True
'''''''        Form1.botao2.Caption = i
'''''''        i = i + 1
'''''''        If i > 32000000 Then i = 0
'''''''    Loop
'''''''
'''''''End Sub
''
''Public Sub Carrega_Notas()
''On Error GoTo Erro
''
''Dim sSituacao As String
''Dim Contador As Integer
''
''Dim sqlSituacao As String
''Dim sSql As String
''
''Dim rsDados As New ADODB.Recordset
''Dim rsNFsInutilizadas As New ADODB.Recordset
''
''Dim sModo As String
''
''   sModo = "Carregar"
''
'''   cmdCancelarNota.Enabled = False
'''   cmdTransmitir.Enabled = False
''   frmGerenciarNF.cmdImprimirDANFE.Enabled = False
''
''   ' Validar Período
''   If Not FValidarPeriodo(frmGerenciarNF.mskDataInicial.text, frmGerenciarNF.mskDataFinal.text) Then Exit Sub
''   ' Retorna condição do SQL
''   sqlSituacao = FSituacao("C", frmGerenciarNF.cmbSituacao.ListIndex)
''   ' Carrega View
''   If sModo = "Carregar" Then frmGerenciarNF.lvwNFs.ListItems.Clear
'''   If I_ModoView = "Carregar" Then frmGerenciarNF.lvwNFs.ListItems.Clear
''
''   ' Notas
''   Contador = 1
''
''   sSql = "       SELECT "
''   sSql = sSql & vbCrLf & "      N.id" & I_TabelasNF & ", "
''   sSql = sSql & vbCrLf & "      N.idCliente, "
''   sSql = sSql & vbCrLf & "      N.Numero, "
''   sSql = sSql & vbCrLf & "      N.DataEmissao, "
''   sSql = sSql & vbCrLf & "      C.Nome, "
''   If I_SGBD = "SQLSERVER" Then
''      sSql = sSql & vbCrLf & "      T.Total  AS Total, "
''   ElseIf I_SGBD = "ACCESS" Then
''      sSql = sSql & vbCrLf & "      (T.Total + T.TotalFrete) AS Total, "
''   End If
''   sSql = sSql & vbCrLf & "      N.Situacao "
''   sSql = sSql & vbCrLf & "FROM "
''   sSql = sSql & vbCrLf & "      " & I_TabelasNF & " N, Clientes C, Total" & I_TabelasNF & " T "
''   sSql = sSql & vbCrLf & "WHERE "
''   sSql = sSql & vbCrLf & "      N.idCliente = C.idCliente AND "
''   sSql = sSql & vbCrLf & "      N.id" & I_TabelasNF & " = T.id" & I_TabelasNF & " AND "
''   If I_SGBD = "SQLSERVER" Then
''      sSql = sSql & vbCrLf & "      N.DataEmissao >= '" & Format(frmGerenciarNF.mskDataInicial.text, "yyyy-mm-dd") & "' AND "
''      sSql = sSql & vbCrLf & "      N.DataEmissao <= '" & Format(frmGerenciarNF.mskDataFinal.text, "yyyy-mm-dd") & "' "
''   ElseIf I_SGBD = "ACCESS" Then
''      sSql = sSql & vbCrLf & "      N.DataEmissao >= cDate('" & Format(frmGerenciarNF.mskDataInicial.text, "dd/mm/yyyy") & "') AND "
''      sSql = sSql & vbCrLf & "      N.DataEmissao <= cDate('" & Format(frmGerenciarNF.mskDataFinal.text, "dd/mm/yyyy") & "') "
''   End If
''   sSql = sSql & vbCrLf & sqlSituacao & " Order By N.Numero"
''
''   Set rsDados = cnSistema.Execute(sSql)
''   Do While Not rsDados.EOF
''      sSituacao = FSituacao("S", rsDados!Situacao)      ' Retorna Situação atual da nota
''
''      ' Pesquisa se registro existe
''      Set ItemList = frmGerenciarNF.lvwNFs.FindItem(StrZero(rsDados!Numero, 8))
''      If ItemList Is Nothing Then
''         Set ItemList = frmGerenciarNF.lvwNFs.ListItems.Add(, "R" & CStr(Contador), StrZero(rsDados!Numero, 8))
''      End If
''
''      ItemList.SubItems(1) = rsDados!DataEmissao
''      ItemList.SubItems(2) = Trim(rsDados!Nome)
''      ItemList.SubItems(3) = Format(rsDados!Total, "##,##0.00")
''      ItemList.SubItems(4) = sSituacao
''
''      Contador = Contador + 1
''      rsDados.MoveNext
''   Loop
''
''   ' Notas
''   If I_SGBD = "SQLSERVER" Then
''      Set rsNFsInutilizadas = cnSistema.Execute("SELECT * FROM " & I_TabelasNF & "Inutilizadas WHERE Data >= '" & Format(frmGerenciarNF.mskDataInicial.text, "yyyy-mm-dd") & "' AND Data <= '" & Format(frmGerenciarNF.mskDataFinal.text, "yyyy-mm-dd") & "' Order By Numero")
''   ElseIf I_SGBD = "ACCESS" Then
''      Set rsNFsInutilizadas = cnSistema.Execute("SELECT * FROM " & I_TabelasNF & "Inutilizadas WHERE Data >= cDate('" & Format(frmGerenciarNF.mskDataInicial.text, "dd/mm/yyyy") & "') AND Data <= cDate('" & Format(frmGerenciarNF.mskDataFinal.text, "dd/mm/yyyy") & "') Order By Numero")
''   End If
''
''   If Not rsNFsInutilizadas.EOF Then
''      Do While Not rsNFsInutilizadas.EOF
''
'''         Set ItemList = lvwNFs.ListItems.Add(, "I" & CStr(rsNFsInutilizadas!Numero), StrZero(rsNFsInutilizadas!Numero, 8))
''
''         Set ProcuraItem = frmGerenciarNF.lvwNFs.FindItem(StrZero(rsNFsInutilizadas!Numero, 8))
''         If ProcuraItem Is Nothing Then
''            Set ItemList = frmGerenciarNF.lvwNFs.ListItems.Add(, "I" & CStr(Contador), StrZero(rsNFsInutilizadas!Numero, 8))
''            ItemList.SubItems(1) = rsNFsInutilizadas!Data
''            ItemList.SubItems(2) = "Nota Inutilizada"
''            ItemList.SubItems(3) = Format(0, "##,##0.00")
''            ItemList.SubItems(4) = "Inutilizada"
''            If IsNull(rsNFsInutilizadas!Protocolo) Or rsNFsInutilizadas!Protocolo = "" Then
''               ItemList.SubItems(5) = "Aguarde..."
''            Else
''               ItemList.SubItems(5) = ""
''            End If
''         End If
''
''         Contador = Contador + 1
''         rsNFsInutilizadas.MoveNext
''      Loop
''   End If
''
''   Set rsDados = Nothing
''   Set rsNFsInutilizadas = Nothing
''
''   Exit Sub
''Erro:
''   Call FRetornaMensagens(Err.Number & " - " & Err.Description & " - Carrega_View ")
''End Sub
''
''Private Function FValidarPeriodo(DataInicial As String, DataFinal As String) As Boolean
''On Error GoTo Erro
''Dim strMensagem As String
''
''   FValidarPeriodo = True
''
''   If Not IsDate(frmGerenciarNF.mskDataInicial.text) Then strMensagem = "Data Inicial Inválida" & Chr(13)
''   If Not IsDate(frmGerenciarNF.mskDataFinal.text) Then strMensagem = "Data Final Inválida" & Chr(13)
''   If IsDate(frmGerenciarNF.mskDataInicial.text) And IsDate(frmGerenciarNF.mskDataFinal.text) Then
''      If (CDate(frmGerenciarNF.mskDataInicial.text) > CDate(frmGerenciarNF.mskDataFinal.text)) Then strMensagem = "Data Inicial maior que Data Final"
''   End If
''   If Not strMensagem = Empty Then
''      MsgBox "Verifique os Seguintes Campos:" & Chr(13) & strMensagem, vbExclamation + vbOKOnly, "Campos Obrigatórios"
''      FValidarPeriodo = False
''      Exit Function
''   End If
''
''   Exit Function
''   Resume
''Erro:
''   Call FRetornaMensagens(Err.Number & " - " & Err.Description & " - " & ".FValidarPeriodo")
''End Function
''
''Private Function FSituacao(Parametro As String, Conteudo As Integer) As String
''On Error GoTo Erro
''
''   If Parametro = "C" Then       ' Condição
'''      If Conteudo = 0 Then FSituacao = "AND Situacao <> 2"
''      If Conteudo = 0 Then FSituacao = ""
''      If Conteudo = 1 Then FSituacao = " AND N.Situacao = 0"
''      If Conteudo = 2 Then FSituacao = " AND N.Situacao = 1"
''      If Conteudo = 3 Then FSituacao = " AND N.Situacao = 2"
''      If Conteudo = 4 Then FSituacao = " AND N.Situacao = 3"
''      If Conteudo = 5 Then FSituacao = " AND N.Situacao = 4"
''      If Conteudo = 6 Then FSituacao = " AND N.Situacao = 5"
''      If Conteudo = 7 Then FSituacao = " AND N.Situacao <> 2"
''   End If
''
''   If Parametro = "S" Then       ' Condição
''      If Conteudo = 0 Then FSituacao = "Em Digitação"
''      If Conteudo = 1 Then FSituacao = "Processamento"
''      If Conteudo = 2 Then FSituacao = "Aprovada"
''      If Conteudo = 3 Then FSituacao = "Cancelada"
''      If Conteudo = 4 Then FSituacao = "Não Emitida"
''      If Conteudo = 5 Then FSituacao = "Denegada"
''      If Conteudo = 9 Then FSituacao = "Pendentes"
''   End If
''
''   Exit Function
''Erro:
''   Call FRetornaMensagens(Err.Number & " - " & Err.Description & " - " & ".FSituacao")
''End Function
''

Attribute VB_Name = "M�dulo"

Option Explicit
Dim ItemList As ListItem
' Vari�vies P�blicas j� definidas

Public cnSistema As New DBConnection
'''''Public cnSistema As ADODB.Connection

Public cnSintegra As ADODB.Connection

Public I_Acesso As Byte                ' Retorna o Grau de Acesso do Usu�rio no Sistema
Public I_User As String                ' Retorna o Usu�rio que est� Logado no Sistema
Public I_MesAno As String              ' Retorna o M�s e Ano da Folha Corrente
Public I_Logotipo As String            ' Logotipo da Empresa
Public I_SGBD As String                ' Gerenciador de Bancos de dados

Public I_idEmpresa As String           ' Id da Empresa
Public I_Empresa As String             ' Nome da Empresa
Public I_idEmpresaUF As String         ' Id da UF da Empresa
Public I_idEmpresaMunicipio As String  ' Id do Municipio da Empresa

Public I_HorarioVerao As String
Public I_ModoView As String

Public I_TituloForm As String          ' Nome do Formul�rio
Public I_PrefixoTelefonico As String   ' Prefixo Telef�nico
Public BancoDeDados As String          ' Localiza��o do Banco de Dados
Public Registro_Selecionado As Boolean ' Variavel para Indicar que um Registro foi Selecionado no frmVisualiza

Public Status As Byte                  ' Status 0=Neutro, 1=Insercao, 2=Alteracao, 3=Delecao, 4=Localiza
Public rsErro As Boolean               ' Retorna Erro caso tenha problema na Abertura do RecordSet

Public I_ContarMensagem As Double

Public NivelTAG(50) As Integer

Public CaminhoINI As String
Public AgendaIniciada As Boolean
Private Const LOCALE_SSHORTDATE = &H1F

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpkeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpkeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetSystemDefaultLCID Lib "kernel32" () As Long
Private Declare Function SetLocaleInfo Lib "kernel32" Alias "SetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String) As Boolean

Public idBancoDados As String, idCatalog As String, idLogin As String, idSenha As String, idAcesso As String, idLicenca As String, idLogotipo As String, idUser As Integer

'''''Private Declare Sub CopyToMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

Sub Main()
   If Len(CStr(Date)) < 10 Then
      Dim lngLocale As Long
      lngLocale = GetSystemDefaultLCID()
      If SetLocaleInfo(lngLocale, LOCALE_SSHORTDATE, "dd/MM/yyyy") = False Then
         MsgBox "Configure a Data para o Formato DD/MM/AAAA no Painel de Controle", vbInformation + vbOKOnly, "Configura��es"
         End
      End If
   End If

'''''   If LerArquivoINI("Banco de Dados", "SGBD", App.Path & "\System.ini") = "SQLSERVER" Then
'''''      Call ConnectDBSQL
'''''   ElseIf LerArquivoINI("Banco de Dados", "SGBD", App.Path & "\System.ini") = "ACCESS" Then
'''''      Call ConnectDB
'''''   End If

   Call ConnectDB
   Call ConfigurarNFs
   Call DefineNiveisTAGs
   
'   Call VerificaVersaoBanco
'   frmLogin.Show vbModal
'   frmGerenciarNFs.Show vbModal
End Sub

Private Sub DefineNiveisTAGs()
Dim Contador As Integer
Dim iNivel As Integer

   ' Define n�veis das TAGs
   iNivel = 0
   For Contador = 1 To 50
       NivelTAG(Contador) = iNivel
       iNivel = iNivel + 3
   Next

End Sub

Sub ConnectDB()
On Error GoTo Erro
Dim rsSistema As New ADODB.Recordset
Dim rsVerificar As New ADODB.Recordset
Dim MD5Chave As New MD5
     
   ' Faz todos os teste com a Arquivo SYSTEM.INI
   'CaminhoINI = Mid(App.Path, 1, 2) & "\Sistemas\Comercial"
   CaminhoINI = App.Path
   
   ' Define o gerenciador de Banco de Dados
   I_SGBD = LerArquivoINI("Banco de Dados", "SGBD", CaminhoINI & "\System.ini")
     
   ' Libera��o de Acesso SGC
   idUser = 0
   idAcesso = "1111"
   '--------------------------------------------------------------
   
   TestarINI
   ConfigIniciais
   
   ' Banco de Dados em Access
   If I_SGBD = "ACCESS" Then
   
      BancoDeDados = LerArquivoINI("Banco de Dados", "Caminho", CaminhoINI & "\System.ini")
''   BancoDeDados = LerArquivoINI("Banco de Dados", "Caminho", App.Path & "\System.ini")

      ' Banco de Dados em Access
      If Dir(BancoDeDados & "\SGC.MDB") = Empty Then frmBancoDados.Show vbModal
      
'''''   Set cnSistema = New ADODB.Connection
'''''   cnSistema.Provider = "Microsoft.Jet.OLEDB.4.0"
'''''   cnSistema.Properties("Data Source") = BancoDeDados & "\SGC.MDB"
'''''   cnSistema.Open
   
      Verifica_BD
   
      Set rsSistema = cnSistema.Execute("Select Conteudo from Sistema Where idSistema=3")
      I_MesAno = rsSistema!Conteudo
      Set rsSistema = cnSistema.Execute("Select Conteudo from Sistema Where idSistema=4")
      I_Empresa = rsSistema!Conteudo
      Set rsSistema = cnSistema.Execute("Select Conteudo from Sistema Where idSistema=5")
      If rsSistema!Conteudo = Empty Then
         If Dir(CaminhoINI & "\System.sys") <> Empty Then
            I_Logotipo = CaminhoINI & "\System.sys"
         Else
            I_Logotipo = ""
         End If
      Else
         If Dir(rsSistema!Conteudo) <> Empty Then
            I_Logotipo = rsSistema!Conteudo
         Else
            If Dir(CaminhoINI & "\System.sys") <> Empty Then
               I_Logotipo = CaminhoINI & "\System.sys"
            Else
               I_Logotipo = ""
            End If
         End If
      End If
      I_Empresa = rsSistema!Conteudo
   
      rsSistema.Close
   End If
   
   I_PrefixoTelefonico = LerArquivoINI("Preenchimento", "Prefixo", CaminhoINI & "\System.ini")
   
   ' Setar configura��es da Empresa Emitente
   Set rsVerificar = cnSistema.Execute("Select * from Empresa")
   I_idEmpresa = IIf(Not rsVerificar.EOF, rsVerificar!idEmpresa, 0)
   I_idEmpresaUF = IIf(Not rsVerificar.EOF, rsVerificar!idUF, 0)
   I_idEmpresaMunicipio = IIf(Not rsVerificar.EOF, rsVerificar!idMunicipio, 0)
   I_Empresa_CNPJ_CPF = IIf(Not rsVerificar.EOF, rsVerificar!CNPJ_CPF, "")
   
   Set rsVerificar = Nothing
   
Exit Sub
Erro:
   If Err.Number = -2147467259 Then
      Beep
      MsgBox "Erro na Abertura do Arquivo de Dados" & Chr(13) & "Entre em Contato com o Administrador", vbExclamation, "Erro"
      End
   End If
End Sub

'''''Sub ConnectDBSQL()
'''''Dim MD5Chave As New MD5
'''''Dim rsSistema As New ADODB.Recordset
'''''   bolOVL = False
'''''   If Dir(App.Path & "\System.ini") = Empty Then
'''''      If Not GravaArquivoINI("Banco de Dados", "Servidor SQL", ".", App.Path & "\System.ini") Then
'''''         MsgBox "N�o foi poss�vel gravar arquivo de configura��o" & Chr(13) & "Entre em contato com o suporte", vbInformation + vbOKOnly, "Sistema"
'''''         End
'''''      End If
'''''      If Not GravaArquivoINI("Banco de Dados", "Catalog", idCatalog, App.Path & "\System.ini") Then
'''''         MsgBox "N�o foi poss�vel gravar arquivo de configura��o" & Chr(13) & "Entre em contato com o suporte", vbInformation + vbOKOnly, "Sistema"
'''''         End
'''''      End If
'''''   End If
'''''   If Dir(App.Path & "\SGC.SYS") = Empty Then
'''''      MsgBox "N�o foi poss�vel iniciar o sistema" & Chr(13) & "Entre em contato com o suporte", vbInformation + vbOKOnly, "Sistema"
'''''      End
'''''   End If
'''''   idBancoDados = LerArquivoINI("Banco de Dados", "Servidor SQL", App.Path & "\System.ini")
'''''   idCatalog = LerArquivoINI("Banco de Dados", "Catalog", App.Path & "\System.ini")
'''''   idLogin = Mid(Replace(Decode(LerArquivoINI(MD5Chave.DigestStrToHexStr("Sistema"), MD5Chave.DigestStrToHexStr("Comando03"), App.Path & "\Hospedagens.SYS")), Chr(0), ""), 3, 1) & Mid(Replace(Decode(LerArquivoINI(MD5Chave.DigestStrToHexStr("Sistema"), MD5Chave.DigestStrToHexStr("Comando03"), App.Path & "\SGC.SYS")), Chr(0), ""), 5, 1)
'''''   idSenha = Replace(Decode(LerArquivoINI(MD5Chave.DigestStrToHexStr("Sistema"), MD5Chave.DigestStrToHexStr("Comando04"), App.Path & "\Hospedagens.SYS")), Chr(0), "")
'''''End Sub

Public Function Registros(Con As ADODB.Connection, Table As String) As Long
Dim Rec As New ADODB.Recordset
   Rec.Open "Select Count(*) as Qte from " + Table, Con, adOpenForwardOnly, adLockPessimistic, 1
   Registros = Rec!Qte
   Rec.Close
End Function


Public Function Registros2(Table As String) As Long
Dim Rec As New ADODB.Recordset
   
   Set Rec = cnSistema.Execute("Select Count(*) as Qte from " + Table)
   Registros2 = Rec!Qte
   Set Rec = Nothing

End Function

Public Sub Botoes(Modo As Byte, frmBotao As Form)

   If I_SGBD = "SQLSERVER" Then
      Select Case Modo
         Case 1 'Sem sele��o
            frmBotao.Toolbar.Buttons(1).Enabled = True
            frmBotao.Toolbar.Buttons(2).Enabled = False
            frmBotao.Toolbar.Buttons(3).Enabled = False
            frmBotao.Toolbar.Buttons(5).Enabled = True
         Case 2 'Com sele��o
            frmBotao.Toolbar.Buttons(1).Enabled = True
            frmBotao.Toolbar.Buttons(2).Enabled = True
            frmBotao.Toolbar.Buttons(3).Enabled = True
            frmBotao.Toolbar.Buttons(5).Enabled = True
         Case 3 'Inclus�o
            frmBotao.Toolbar.Buttons(1).Enabled = True
            frmBotao.Toolbar.Buttons(2).Enabled = False
            frmBotao.Toolbar.Buttons(3).Enabled = True
            frmBotao.Toolbar.Buttons(5).Enabled = True
      End Select
      
   ElseIf I_SGBD = "ACCESS" Then
      Select Case Modo
         Case 1 ' Normal
            frmBotao.Toolbar.Buttons(1).Enabled = True
            frmBotao.Toolbar.Buttons(2).Enabled = True
            frmBotao.Toolbar.Buttons(3).Enabled = True
            frmBotao.Toolbar.Buttons(5).Enabled = False
            frmBotao.Toolbar.Buttons(6).Enabled = False
            frmBotao.Toolbar.Buttons(8).Enabled = True
            frmBotao.Toolbar.Buttons(9).Enabled = True
            frmBotao.Toolbar.Buttons(10).Enabled = True
            frmBotao.Toolbar.Buttons(11).Enabled = True
            frmBotao.Toolbar.Buttons(13).Enabled = True
            frmBotao.Toolbar.Buttons(14).Enabled = True
         Case 2 ' Gravar
            frmBotao.Toolbar.Buttons(1).Enabled = False
            frmBotao.Toolbar.Buttons(2).Enabled = False
            frmBotao.Toolbar.Buttons(3).Enabled = False
            frmBotao.Toolbar.Buttons(5).Enabled = True
            frmBotao.Toolbar.Buttons(6).Enabled = True
            frmBotao.Toolbar.Buttons(8).Enabled = False
            frmBotao.Toolbar.Buttons(9).Enabled = False
            frmBotao.Toolbar.Buttons(10).Enabled = False
            frmBotao.Toolbar.Buttons(11).Enabled = False
            frmBotao.Toolbar.Buttons(13).Enabled = False
            frmBotao.Toolbar.Buttons(14).Enabled = False
         Case 3 ' Primeiro Registro
            frmBotao.Toolbar.Buttons(1).Enabled = True
            frmBotao.Toolbar.Buttons(2).Enabled = False
            frmBotao.Toolbar.Buttons(3).Enabled = False
            frmBotao.Toolbar.Buttons(5).Enabled = False
            frmBotao.Toolbar.Buttons(6).Enabled = False
            frmBotao.Toolbar.Buttons(8).Enabled = False
            frmBotao.Toolbar.Buttons(9).Enabled = False
            frmBotao.Toolbar.Buttons(10).Enabled = False
            frmBotao.Toolbar.Buttons(11).Enabled = False
            frmBotao.Toolbar.Buttons(13).Enabled = False
            frmBotao.Toolbar.Buttons(14).Enabled = False
         Case 4 ' Cancelar
            frmBotao.Toolbar.Buttons(1).Enabled = False
            frmBotao.Toolbar.Buttons(2).Enabled = False
            frmBotao.Toolbar.Buttons(3).Enabled = False
            frmBotao.Toolbar.Buttons(5).Enabled = False
            frmBotao.Toolbar.Buttons(6).Enabled = True
            frmBotao.Toolbar.Buttons(8).Enabled = False
            frmBotao.Toolbar.Buttons(9).Enabled = False
            frmBotao.Toolbar.Buttons(10).Enabled = False
            frmBotao.Toolbar.Buttons(11).Enabled = False
            frmBotao.Toolbar.Buttons(13).Enabled = False
            frmBotao.Toolbar.Buttons(14).Enabled = False
      End Select
   End If

End Sub
Public Function StrZero(ByVal Valor As Double, ByVal Qte As Double) As String
Dim Valor_Str As String

   If Qte < 1 Then
      StrZero = "qte menor que 1"
   Else
      Valor_Str = Trim(Str(Valor))
      If Len(Valor_Str) <> Qte Then
         StrZero = Replicate("0", Qte - Len(Valor_Str)) + Valor_Str
      Else
         StrZero = Valor_Str
      End If
   End If
End Function
Public Function Replicate(Caracter As String, Qte As Double) As String
Dim Contador As Integer
   If Qte < 1 Then
      Replicate = "qte menor que 1"
   Else
      For Contador = 1 To Qte
         Replicate = Replicate + Caracter
      Next
   End If
End Function
Public Function Substitui(ByVal sStr As String, Ssearch As String, sTroca As String) As String
Dim Procura As Integer

   Procura = InStr(sStr, Ssearch)
   If Procura = 0 Then
      Substitui = sStr
   Else
      Substitui = Mid(sStr, 1, Procura - 1) & sTroca & Mid(sStr, Procura + 1, Len(sStr) - Procura)
   End If
End Function
Public Function SQLCheck(SQLstring As String) As String
Dim Procura As Byte
   SQLCheck = SQLstring
   Procura = 1 ' Inicio da procura de aspas simples
   Do While Procura <> 0
      Procura = InStr(SQLCheck, "'")
      If Procura <> 0 Then
         SQLCheck = Mid(SQLCheck, 1, Procura - 1) & " " & Mid(SQLCheck, Procura + 1, Len(SQLCheck) - Procura)
      End If
   Loop
End Function
Public Sub Centraliza(frm As Form)
'   frm.StartUpPosition = 2
'   frm.StartUpPosition = 2
'   frm.Top = 575
'   frm.Left = 6000
   
   frm.Top = (frmGerenciarNF.Height - frm.Height) / 2
   frm.Left = (frmGerenciarNF.Width - frm.Width) / 2
   
'   frm.Top = (MDISistema.Height - frm.Height - 1400) / 2
'   frm.Left = (MDISistema.Width - frm.Width) / 2
End Sub
Public Sub Atividade(sysAtividade As String, sysModulo As String)

   If I_SGBD = "SQLSERVER" Then
      cnSistema.Execute "Insert Into Atividades (Atividade,Modulo,Data,idUsuario) " & _
                        "Values ('" & sysAtividade & "','" & sysModulo & "','" & Format(Now, "mm/dd/yyyy hh:mm:ss") & "'," & IIf(idUser = 0, "NULL", idUser) & ")"
   ElseIf I_SGBD = "ACCESS" Then
      cnSistema.Execute "Insert Into Atividades (Atividade,Modulo,Data,Usuario) " & _
                        "Values ('" & Mid(sysAtividade, 1, 50) & "','" & sysModulo & "','" & Now & "','" & I_User & "')"
   End If

End Sub
Public Function LerArquivoINI(ByVal sSecao As String, ByVal sItem As String, ByVal sArquivo As String) As String
   Screen.MousePointer = vbHourglass
   Dim lQtdCrt    As Long
   Dim lTamMaxRet As Long
   Dim sValorItem As String * 128
   lTamMaxRet = Len(sValorItem)
   lQtdCrt = GetPrivateProfileString(sSecao, sItem, "", sValorItem, lTamMaxRet, sArquivo)
   LerArquivoINI = Left(sValorItem, lQtdCrt)
   Screen.MousePointer = vbDefault
End Function
Public Function GravaArquivoINI(ByVal sSecao As String, ByVal sItem As String, ByVal sValorItem As String, ByVal sArquivo As String) As Boolean
    Dim lQtdCrt As Long
    Screen.MousePointer = vbHourglass
    sValorItem = Left(sValorItem, 128)
    lQtdCrt = WritePrivateProfileString(sSecao, sItem, sValorItem, sArquivo)
    GravaArquivoINI = IIf(lQtdCrt > 0, True, False)
    Screen.MousePointer = vbDefault
End Function
Public Function CNPJ_CPF(sArg As String) As String
Dim sCNPJ_CPF As String, _
    sNumero  As String, _
    iDigito As String, _
    iTamanho As Integer, _
    iTeste As Integer, _
    iSoma As Double, _
    iMultiplicador As Integer, _
    Contador As Integer, _
    i As Integer

   For Contador = 1 To Len(sArg)
       If IsNumeric(Mid(sArg, Contador, 1)) Then
          sCNPJ_CPF = sCNPJ_CPF + Mid(sArg, Contador, 1)
       End If
   Next
   
   If Len(sCNPJ_CPF) <> 11 And Len(sCNPJ_CPF) <> 14 Then
      MsgBox "Tamanho do CNPJ/CPF est� incorreto", vbInformation + vbOKOnly, "Informa��es"
      CNPJ_CPF = "ERRO"
      Exit Function
   End If

   sNumero = Mid(sCNPJ_CPF, 1, Len(sCNPJ_CPF) - 2) ' Definir o n�mero
   iTamanho = Len(sNumero) ' Definir o tamanho do n�mero

   If iTamanho > 9 Then ' Define o tamanho
      iTeste = 10
   Else
      iTeste = 99
   End If
   For Contador = 1 To 2
       iSoma = 0                             ' Inicializa vari�vel para somar
       iMultiplicador = 2                    ' Inicializa vari�vel para multiplica��o
       For i = 1 To iTamanho                 ' Multiplicar os n�mero do CNPJ de traz para frente
           If iMultiplicador = iTeste Then   ' Se multiplicador for o mesmo que teste reinicializ�-la
              iMultiplicador = 2
           End If
           iSoma = iSoma + Val(Mid(sNumero, iTamanho - i + 1, 1)) * iMultiplicador ' Soma ser� igual a soma + posi��o do n�mero do CPF de traz para frente * o multiplicador
           iMultiplicador = iMultiplicador + 1 ' Acrescentar 1 em multiplicador
       Next
       
       iDigito = 11 - (iSoma - (Int(iSoma / 11) * 11)) ' Digito Verificador
       If iDigito > 9 Then
          iDigito = 0
       End If
       
       sNumero = sNumero & Trim(Str(iDigito)) ' Acrescentar o digito ao n�mero
       iTamanho = iTamanho + 1 ' Aumentar 1 em tamanho
   Next
   If sCNPJ_CPF <> sNumero Then
      MsgBox "CNPJ/CPF est� incorreto", vbInformation + vbOKOnly, "Informa��es"
      CNPJ_CPF = "ERRO"
      Exit Function
   End If
   If Len(sCNPJ_CPF) = 11 Then CNPJ_CPF = Mid(sCNPJ_CPF, 1, 3) & "." & Mid(sCNPJ_CPF, 4, 3) & "." & Mid(sCNPJ_CPF, 7, 3) & "-" & Mid(sCNPJ_CPF, 10, 2)
   If Len(sCNPJ_CPF) = 14 Then CNPJ_CPF = Mid(sCNPJ_CPF, 1, 2) & "." & Mid(sCNPJ_CPF, 3, 3) & "." & Mid(sCNPJ_CPF, 6, 3) & "/" & Mid(sCNPJ_CPF, 9, 4) & "-" & Mid(sCNPJ_CPF, 13, 2)
End Function

Public Function FormataTXT(ByRef sCampo As String, ByRef sTipo As Double, ByRef sTamanho As Integer) As String
   If sTipo = 1 Then
      If Len(sCampo) <= sTamanho Then
         FormataTXT = sCampo + Space(sTamanho - Len(sCampo))
      Else
         FormataTXT = Mid(sCampo, 1, sTamanho)
      End If
   ElseIf sTipo = 2 Then
      FormataTXT = StrZero(Substitui(sCampo, ",", ""), sTamanho)
   ElseIf sTipo = 2.1 Then
      If Len(sCampo) <= sTamanho Then
         FormataTXT = Space(sTamanho - Len(sCampo)) + sCampo
      Else
         FormataTXT = Mid(sCampo, 1, sTamanho)
      End If
   ElseIf sTipo = 3.1 Then
      FormataTXT = Format(sCampo, "yyyymmdd")
   ElseIf sTipo = 3.2 Then
      FormataTXT = Format(sCampo, "ddmmyy")
   ElseIf sTipo = 3.3 Then
      FormataTXT = Format(sCampo, "ddmmyyyy")
   ElseIf sTipo = 3.4 Then
      FormataTXT = Format(CDate(Mid(sCampo, 1, 10)), "yyyymm")
   End If
End Function

Public Function ValidarData(sArg As Date) As Boolean
Dim dData As Date

' Validar o Formato
' Validar o Tamanho

   If Not IsDate(sArg) Or Val(Mid(sArg, 7, 4)) < 1900 Then
      MsgBox "Data Inv�lida", vbOKOnly + vbInformation, "Valida��o"
      ValidarData = False
      Exit Function
   End If

End Function

Public Function PreencheCaixaTexto(Arquivo As String, CaixaTexto As Control) As Boolean
On Error GoTo Erro
Dim handle As Integer
Dim Linha As String

   handle = FreeFile
   Open CaminhoINI & Arquivo For Input As #handle
   CaixaTexto.Text = ""
   While Not EOF(handle)
      Line Input #handle, Linha
      CaixaTexto.Text = CaixaTexto.Text & Linha
   Wend
   Close #handle
   PreencheCaixaTexto = False

Exit Function
Erro:
   On Error Resume Next
   Close #handle
   PreencheCaixaTexto = False
   Exit Function
End Function

Public Function RemoveAcentos(sCampo As String) As String
Dim sVerifica As String, sTroca As String, sNCampo As String, Procura As Integer
Dim Contador As Integer

   sVerifica = "�������������������������Ǻ�"
   sTroca = "aAaAaAaAeEeEiIoOoOoOuUuUcCoa"
   For Contador = 1 To Len(sCampo)
       Procura = InStr(sVerifica, Mid(sCampo, Contador, 1))
       If Procura <> 0 Then
          sNCampo = sNCampo + Mid(sTroca, Procura, 1)
       Else
          sNCampo = sNCampo + Mid(sCampo, Contador, 1)
       End If
   Next
   RemoveAcentos = sNCampo
End Function

Public Function SubstituiAcentos(sCampo As String) As String
Dim sVerifica As String, sTroca As String, sNCampo As String, Procura As Integer
Dim Contador As Integer

   sVerifica = "�������������������������Ǻ�"
   sTroca = "�������������������������Ǻ�"
   For Contador = 1 To Len(sCampo)
       Procura = InStr(sVerifica, Mid(sCampo, Contador, 1))
       If Procura <> 0 Then
          sNCampo = sNCampo + Mid(sTroca, Procura, 1)
       Else
          sNCampo = sNCampo + Mid(sCampo, Contador, 1)
       End If
   Next
   SubstituiAcentos = sNCampo
End Function

Public Function NumeroExtenso(vNumero As Variant, Optional bMoeda As Boolean = True) As String
Dim iContador As Integer
Dim iTamanho As Integer

Dim sValor As String
Dim sParte As String
Dim sFinal As String
    
    If IsNull(vNumero) Or vNumero <= 0 Or vNumero > 9999999.99 Or Not IsNumeric(vNumero) Then Exit Function
    
    ReDim matGrupo(4), matTexto(4) As String

    ReDim matUnidades(19) As String
    matUnidades(1) = "Um "
    matUnidades(2) = "Dois "
    matUnidades(3) = "Tres "
    matUnidades(4) = "Quatro "
    matUnidades(5) = "Cinco "
    matUnidades(6) = "Seis "
    matUnidades(7) = "Sete "
    matUnidades(8) = "Oito "
    matUnidades(9) = "Nove "
    matUnidades(10) = "Dez "
    matUnidades(11) = "Onze "
    matUnidades(12) = "Doze "
    matUnidades(13) = "Treze "
    matUnidades(14) = "Quatorze "
    matUnidades(15) = "Quinze "
    matUnidades(16) = "Dezesseis "
    matUnidades(17) = "Dezessete "
    matUnidades(18) = "Dezoito "
    matUnidades(19) = "Dezenove "
    
    ReDim matDezenas(9) As String
    matDezenas(1) = "Dez "
    matDezenas(2) = "Vinte "
    matDezenas(3) = "Trinta "
    matDezenas(4) = "Quarenta "
    matDezenas(5) = "Cinquenta "
    matDezenas(6) = "Sessenta "
    matDezenas(7) = "Setenta "
    matDezenas(8) = "Oitenta "
    matDezenas(9) = "Noventa "
    
    ReDim matCentenas(9) As String
    matCentenas(1) = "Cento "
    matCentenas(2) = "Duzentos "
    matCentenas(3) = "Trezentos "
    matCentenas(4) = "Quatrocentos "
    matCentenas(5) = "Quinhentos "
    matCentenas(6) = "Seiscentos "
    matCentenas(7) = "Setecentos "
    matCentenas(8) = "Oitocentos "
    matCentenas(9) = "Novecentos "
    
    sValor = Format(vNumero, "0000000000.00")
    matGrupo(1) = Mid(sValor, 2, 3)
    matGrupo(2) = Mid(sValor, 5, 3)
    matGrupo(3) = Mid(sValor, 8, 3)
    matGrupo(4) = "0" + Mid(sValor, 12, 2)
    
    For iContador = 1 To 4

      sParte = matGrupo(iContador)

      iTamanho = Switch(Val(sParte) < 10, 1, Val(sParte) < 100, 2, Val(sParte) < 1000, 3)

      If iTamanho = 3 Then
        If Right(sParte, 2) <> "00" Then
          matTexto(iContador) = matTexto(iContador) + matCentenas(Left(sParte, 1)) + "e "
          iTamanho = 2
        Else
          matTexto(iContador) = matTexto(iContador) + IIf(Left(sParte, 1) = "1", "Cem ", _
          matCentenas(Left(sParte, 1)))
        End If
      End If

      If iTamanho = 2 Then
        If Val(Right(sParte, 2)) < 20 Then
          matTexto(iContador) = matTexto(iContador) + matUnidades(Right(sParte, 2))
        Else
          matTexto(iContador) = matTexto(iContador) + matDezenas(Mid(sParte, 2, 1))
          If Right(sParte, 1) <> "0" Then
            matTexto(iContador) = matTexto(iContador) + "e "
            iTamanho = 1
          End If
        End If
      End If

      If iTamanho = 1 Then
        matTexto(iContador) = matTexto(iContador) + matUnidades(Right(sParte, 1))
      End If

    Next

    If Val(matGrupo(1) + matGrupo(2) + matGrupo(3)) = 0 And Val(matGrupo(4)) <> 0 Then
      sFinal = matTexto(4) + IIf(Val(matGrupo(4)) = 1, "centavo", "centavos")
    Else
      sFinal = ""
      sFinal = sFinal + IIf(Val(matGrupo(1)) <> 0, matTexto(1) + IIf(Val(matGrupo(1)) > 1, _
               "milh�es ", "milh�o "), "")
      
      If Val(matGrupo(2) + matGrupo(3)) = 0 Then
        sFinal = sFinal + "de "
      Else
        sFinal = sFinal + IIf(Val(matGrupo(2)) <> 0, matTexto(2) + "Mil ", "")
      End If
      
      If Not bMoeda Then
          sFinal = sFinal + matTexto(3) + IIf(Val(matGrupo(4)) <> 0, "Virgula " + matTexto(4), "")
      Else
        sFinal = sFinal + matTexto(3) + IIf(Val(matGrupo(1) + matGrupo(2) + matGrupo(3)) = 1, "real ", _
                  "reais ")
        sFinal = sFinal + IIf(Val(matGrupo(4)) <> 0, "e " + matTexto(4) + IIf(Val(matGrupo(4)) = 1, _
                 "centavo", "centavos"), "")
        End If

    End If

    NumeroExtenso = sFinal

End Function

Public Function SaltarLinha(iParametro As Integer) As Integer
Dim Contador As Integer

   If iParametro > 0 Then
      For Contador = 1 To (iParametro - 1)
          Print #1, ""
      Next
   End If
End Function

Sub ConfigIniciais()
   
   I_FormatoQuantidade = "###,##0." & Replicate("0", LerArquivoINI("Preenchimento", "DecimaisQuantidade", CaminhoINI & "\System.ini"))
   I_FormatoValor = "###,###,##0." & Replicate("0", LerArquivoINI("Preenchimento", "DecimaisValor", CaminhoINI & "\System.ini"))
   
End Sub

Sub TestarINI()
Dim Erro As Boolean

   ' Banco de Dados
   Erro = CriarINI("Banco de Dados", "Caminho", CaminhoINI)
   Erro = CriarINI("Banco de Dados", "Sintegra", CaminhoINI)

   ' Impressora Fiscal
   Erro = CriarINI("Impressora Fiscal", "Caminho", CaminhoINI)
   Erro = CriarINI("Impressora Fiscal", "Ativar", "0")
   Erro = CriarINI("Impressora Fiscal", "Porta", "1")
   Erro = CriarINI("Impressora Fiscal", "Modelo", "0")
   Erro = CriarINI("Impressora Fiscal", "Horario", "17:00:00")
   Erro = CriarINI("Impressora Fiscal", "Gaveta", "0")

   ' SEPD
   Erro = CriarINI("SEPD", "TipoImpressao", "0")
   
   ' Preenchimento
   Erro = CriarINI("Preenchimento", "Cidade", "")
   Erro = CriarINI("Preenchimento", "UF", "")
   Erro = CriarINI("Preenchimento", "CEP", "  .   -   ")
   Erro = CriarINI("Preenchimento", "Prefixo", "061")
   Erro = CriarINI("Preenchimento", "DecimaisQuantidade", 2)
   Erro = CriarINI("Preenchimento", "DecimaisValor", 2)
   
   ' Orcamentos
   Erro = CriarINI("Orcamentos", "Imprimir", "1")
   Erro = CriarINI("Orcamentos", "Confirmacao", "0")
   Erro = CriarINI("Orcamentos", "ECF", "1")
   Erro = CriarINI("Orcamentos", "PercMaxDesconto", "20,00")
   Erro = CriarINI("Orcamentos", "Mensagem", "Prazo maximo para troca 07 dias!")
   Erro = CriarINI("Orcamentos", "Mensagem2", "")
   Erro = CriarINI("Orcamentos", "PesquisaDescricao", "0")
   Erro = CriarINI("Orcamentos", "Formato", "1")
   
   ' Impressoras
   Erro = CriarINI("Impressoras", "Notas", "LPT1")
   Erro = CriarINI("Impressoras", "Boletos1", "LPT1")
   Erro = CriarINI("Impressoras", "Boletos2", "LPT1")
   Erro = CriarINI("Impressoras", "Orcamentos", "LPT1")
   Erro = CriarINI("Impressoras", "TipoImpressao", "1")
   
   ' Notas Fiscais
   Erro = CriarINI("Notas Fiscais", "ItensNota", "16")
   Erro = CriarINI("Notas Fiscais", "DadosAdicionais", "")
   
   ' Uteis
   Erro = CriarINI("Uteis", "ExibirAgenda", "0")
   
End Sub

Public Function FValor(ByRef sCampo As String) As String

   FValor = Format(sCampo, "#########0.00")
   
End Function

Public Function CStrValor(ByRef sCampo As String, Optional ByRef sTipo As Integer) As String
Dim sVerifica As String, sTroca As String, sNCampo As String, Procura As Integer
Dim Contador As Integer

   If sTipo = 2 Then
      sVerifica = ","
      sTroca = "."
   Else
      sVerifica = "."
      sTroca = ""
   End If
   
   For Contador = 1 To Len(sCampo)
       Procura = InStr(sVerifica, Mid(sCampo, Contador, 1))
       If Procura <> 0 Then
          sNCampo = sNCampo + Mid(sTroca, Procura, 1)
       Else
          sNCampo = sNCampo + Mid(sCampo, Contador, 1)
       End If
   Next

   If I_SGBD = "SQLSERVER" Then
      CStrValor = Substitui(sNCampo, ",", ".")
   ElseIf I_SGBD = "ACCESS" Then
      CStrValor = sNCampo
   End If


'''''   CStrValor = Val(Substitui(sNCampo, ",", "."))
'   CStrValor = sNCampo
End Function

Public Function CDecimais(sCampo As String) As String
Dim sVerifica As String, sTroca As String, sNCampo As String, Procura As Integer
Dim Contador As Integer

   sVerifica = "123456789"
   sTroca = "0"
   For Contador = 1 To Len(sCampo)
       Procura = InStr(sVerifica, Mid(sCampo, Contador, 1))
       If Procura <> 0 Then
          sNCampo = sNCampo + Mid(sTroca, Procura, 1)
       Else
          sNCampo = sNCampo + Mid(sCampo, Contador, 1)
       End If
   Next

   CDecimais = Substitui(sNCampo, ",", ".")
End Function

Public Function CriarINI(ByVal sTitulo As String, ByVal sOpcao As String, ByVal sConteudo As String) As Boolean
   If LerArquivoINI(sTitulo, sOpcao, CaminhoINI & "\System.ini") = "" Then
      If Not GravaArquivoINI(sTitulo, sOpcao, sConteudo, CaminhoINI & "\System.ini") Then
         MsgBox "N�o foi poss�vel Gravar Arquivo de Configura��o" & Chr(13) & "Entre em Contato com o Suporte", vbInformation + vbOKOnly, "Sistema"
         End
      End If
   End If
End Function

Public Function AtualizarINI(ByVal sTitulo As String, ByVal sOpcao As String, ByVal sConteudo As String) As Boolean
   If Not GravaArquivoINI(sTitulo, sOpcao, sConteudo, CaminhoINI & "\System.ini") Then
      MsgBox "N�o foi poss�vel Gravar Arquivo de Configura��o" & Chr(13) & "Entre em Contato com o Suporte", vbInformation + vbOKOnly, "Sistema"
      End
   End If
End Function

Public Function TestaCNPJ_CPF(sArg As String) As String
Dim sCNPJ_CPF As String, _
    sNumero  As String, _
    iDigito As String, _
    iTamanho As Integer, _
    iTeste As Integer, _
    iSoma As Double, _
    iMultiplicador As Integer, _
    Contador As Integer, _
    i As Integer

   For Contador = 1 To Len(sArg)
       If IsNumeric(Mid(sArg, Contador, 1)) Then
          sCNPJ_CPF = sCNPJ_CPF + Mid(sArg, Contador, 1)
       End If
   Next
   
   If Len(Trim(sCNPJ_CPF)) <> 11 And Len(Trim(sCNPJ_CPF)) <> 14 Then
'      MsgBox "Tamanho do CNPJ/CPF est� incorreto", vbInformation + vbOKOnly, "Informa��es"
      TestaCNPJ_CPF = "000.000.000-00"
      Exit Function
   End If

   sNumero = Mid(sCNPJ_CPF, 1, Len(sCNPJ_CPF) - 2) ' Definir o n�mero
   iTamanho = Len(sNumero) ' Definir o tamanho do n�mero

   If iTamanho > 9 Then ' Define o tamanho
      iTeste = 10
   Else
      iTeste = 99
   End If
   For Contador = 1 To 2
       iSoma = 0                             ' Inicializa vari�vel para somar
       iMultiplicador = 2                    ' Inicializa vari�vel para multiplica��o
       For i = 1 To iTamanho                 ' Multiplicar os n�mero do CNPJ de traz para frente
           If iMultiplicador = iTeste Then   ' Se multiplicador for o mesmo que teste reinicializ�-la
              iMultiplicador = 2
           End If
           iSoma = iSoma + Val(Mid(sNumero, iTamanho - i + 1, 1)) * iMultiplicador ' Soma ser� igual a soma + posi��o do n�mero do CPF de traz para frente * o multiplicador
           iMultiplicador = iMultiplicador + 1 ' Acrescentar 1 em multiplicador
       Next
       
       iDigito = 11 - (iSoma - (Int(iSoma / 11) * 11)) ' Digito Verificador
       If iDigito > 9 Then
          iDigito = 0
       End If
       
       sNumero = sNumero & Trim(Str(iDigito)) ' Acrescentar o digito ao n�mero
       iTamanho = iTamanho + 1 ' Aumentar 1 em tamanho
   Next
   If sCNPJ_CPF <> sNumero Then
'      MsgBox "CNPJ/CPF est� incorreto", vbInformation + vbOKOnly, "Informa��es"
      TestaCNPJ_CPF = "000.000.000-00"
      Exit Function
   End If
   If Len(sCNPJ_CPF) = 11 Then TestaCNPJ_CPF = Mid(sCNPJ_CPF, 1, 3) & "." & Mid(sCNPJ_CPF, 4, 3) & "." & Mid(sCNPJ_CPF, 7, 3) & "-" & Mid(sCNPJ_CPF, 10, 2)
   If Len(sCNPJ_CPF) = 14 Then TestaCNPJ_CPF = Mid(sCNPJ_CPF, 1, 2) & "." & Mid(sCNPJ_CPF, 3, 3) & "." & Mid(sCNPJ_CPF, 6, 3) & "/" & Mid(sCNPJ_CPF, 9, 4) & "-" & Mid(sCNPJ_CPF, 13, 2)
End Function

Public Function ChecaInscrE(pUF As String, pInscr As String)
   
   ChecaInscrE = False
   Dim strBase              As String
   Dim strBase2             As String
   Dim strOrigem            As String
   Dim strDigito1           As String
   Dim strDigito2           As String
   Dim intPos               As Double
   Dim intValor             As Double
   Dim intSoma              As Double
   Dim intResto             As Double
   Dim intNumero            As Double
   Dim intPeso              As Double
   Dim intDig               As Double
   
   strBase = ""
   strBase2 = ""
   strOrigem = ""
   If Trim(pInscr) = "ISENTO" Then
       ChecaInscrE = True
       Exit Function
   End If
   For intPos = 1 To Len(Trim(pInscr))
        If InStr(1, "0123456789P", Mid$(pInscr, intPos, 1), vbTextCompare) > 0 Then
            strOrigem = strOrigem & Mid$(pInscr, intPos, 1)
        End If
   Next
   Select Case pUF
     Case "AC"    ' Acre
          strBase = Left(Trim(strOrigem) & "000000000", 9)
          If Left(strBase, 2) = "01" And Mid$(strBase, 3, 2) <> "00" Then
              intSoma = 0
              For intPos = 1 To 8
                   intValor = Val(Mid$(strBase, intPos, 1))
                   intValor = intValor * (10 - intPos)
                   intSoma = intSoma + intValor
              Next
              intResto = intSoma Mod 11
              strDigito1 = Right(IIf(intResto < 2, "0", Str(11 - intResto)), 1)
              strBase2 = Left(strBase, 8) & strDigito1
              If strBase2 = strOrigem Then
                  ChecaInscrE = True
              End If
          End If
     Case "AL"    ' Alagoas
          strBase = Left(Trim(strOrigem) & "000000000", 9)
          If Left(strBase, 2) = "24" Then
              intSoma = 0
              For intPos = 1 To 8
                   intValor = Val(Mid$(strBase, intPos, 1))
                   intValor = intValor * (10 - intPos)
                   intSoma = intSoma + intValor
              Next
              intSoma = intSoma * 10
              intResto = intSoma Mod 11
              strDigito1 = Right(IIf(intResto = 10, "0", Str(intResto)), 1)
              strBase2 = Left(strBase, 8) & strDigito1
              If strBase2 = strOrigem Then
                  ChecaInscrE = True
              End If
          End If
     Case "AM"    ' Amazonas
          strBase = Left(Trim(strOrigem) & "000000000", 9)
          intSoma = 0
          For intPos = 1 To 8
               intValor = Val(Mid$(strBase, intPos, 1))
               intValor = intValor * (10 - intPos)
               intSoma = intSoma + intValor
          Next
          If intSoma < 11 Then
              strDigito1 = Right(Str(11 - intSoma), 1)
          Else
              intResto = intSoma Mod 11
              strDigito1 = Right(IIf(intResto < 2, "0", Str(11 - intResto)), 1)
          End If
          strBase2 = Left(strBase, 8) & strDigito1
          If strBase2 = strOrigem Then
              ChecaInscrE = True
          End If
     Case "AP"    ' Amapa
          strBase = Left(Trim(strOrigem) & "000000000", 9)
          intPeso = 0
          intDig = 0
          If Left(strBase, 2) = "03" Then
              intNumero = Val(Left(strBase, 8))
              If intNumero >= 3000001 And _
                 intNumero <= 3017000 Then
                  intPeso = 5
                  intDig = 0
              ElseIf intNumero >= 3017001 And _
                     intNumero <= 3019022 Then
                  intPeso = 9
                  intDig = 1
              ElseIf intNumero >= 3019023 Then
                  intPeso = 0
                  intDig = 0
              End If
              intSoma = intPeso
              For intPos = 1 To 8
                   intValor = Val(Mid$(strBase, intPos, 1))
                   intValor = intValor * (10 - intPos)
                   intSoma = intSoma + intValor
              Next
              intResto = intSoma Mod 11
              intValor = 11 - intResto
              If intValor = 10 Then
                  intValor = 0
              ElseIf intValor = 11 Then
                  intValor = intDig
              End If
              strDigito1 = Right(Str(intValor), 1)
              strBase2 = Left(strBase, 8) & strDigito1
              If strBase2 = strOrigem Then
                  ChecaInscrE = True
              End If
          End If
     Case "BA"    ' Bahia
          strBase = Left(Trim(strOrigem) & "00000000", 8)
          If InStr(1, "0123458", Left(strBase, 1), vbTextCompare) > 0 Then
              intSoma = 0
              For intPos = 1 To 6
                   intValor = Val(Mid$(strBase, intPos, 1))
                   intValor = intValor * (8 - intPos)
                   intSoma = intSoma + intValor
              Next
              intResto = intSoma Mod 10
              strDigito2 = Right(IIf(intResto = 0, "0", Str(10 - intResto)), 1)
              strBase2 = Left(strBase, 6) & strDigito2
              intSoma = 0
              For intPos = 1 To 7
                   intValor = Val(Mid$(strBase2, intPos, 1))
                   intValor = intValor * (9 - intPos)
                   intSoma = intSoma + intValor
              Next
              intResto = intSoma Mod 10
              strDigito1 = Right(IIf(intResto = 0, "0", Str(10 - intResto)), 1)
          Else
              intSoma = 0
              For intPos = 1 To 6
                   intValor = Val(Mid$(strBase, intPos, 1))
                   intValor = intValor * (8 - intPos)
                   intSoma = intSoma + intValor
              Next
              intResto = intSoma Mod 11
              strDigito2 = Right(IIf(intResto < 2, "0", Str(11 - intResto)), 1)
              strBase2 = Left(strBase, 6) & strDigito2
              intSoma = 0
              For intPos = 1 To 7
                   intValor = Val(Mid$(strBase2, intPos, 1))
                   intValor = intValor * (9 - intPos)
                   intSoma = intSoma + intValor
              Next
              intResto = intSoma Mod 11
              strDigito1 = Right(IIf(intResto < 2, "0", Str(11 - intResto)), 1)
          End If
          strBase2 = Left(strBase, 6) & strDigito1 & strDigito2
          If strBase2 = strOrigem Then
              ChecaInscrE = True
          End If
     Case "CE"    ' Ceara
          strBase = Left(Trim(strOrigem) & "000000000", 9)
          intSoma = 0
          For intPos = 1 To 8
               intValor = Val(Mid$(strBase, intPos, 1))
               intValor = intValor * (10 - intPos)
               intSoma = intSoma + intValor
          Next
          intResto = intSoma Mod 11
          intValor = 11 - intResto
          If intValor > 9 Then
              intValor = 0
          End If
          strDigito1 = Right(Str(intValor), 1)
          strBase2 = Left(strBase, 8) & strDigito1
          If strBase2 = strOrigem Then
              ChecaInscrE = True
          End If
     Case "DF"    ' Distrito Federal
          strBase = Left(Trim(strOrigem) & "0000000000000", 13)
          If Left(strBase, 3) = "073" Or Left(strBase, 3) = "074" Then
              intSoma = 0
              intPeso = 2
              For intPos = 11 To 1 Step -1
                   intValor = Val(Mid$(strBase, intPos, 1))
                   intValor = intValor * intPeso
                   intSoma = intSoma + intValor
                   intPeso = intPeso + 1
                   If intPeso > 9 Then
                       intPeso = 2
                   End If
              Next
              intResto = intSoma Mod 11
              strDigito1 = Right(IIf(intResto < 2, "0", Str(11 - intResto)), 1)
              strBase2 = Left(strBase, 11) & strDigito1
              intSoma = 0
              intPeso = 2
              For intPos = 12 To 1 Step -1
                   intValor = Val(Mid$(strBase, intPos, 1))
                   intValor = intValor * intPeso
                   intSoma = intSoma + intValor
                   intPeso = intPeso + 1
                   If intPeso > 9 Then
                       intPeso = 2
                   End If
              Next
              intResto = intSoma Mod 11
              strDigito2 = Right(IIf(intResto < 2, "0", Str(11 - intResto)), 1)
              strBase2 = Left(strBase, 12) & strDigito2
              If strBase2 = strOrigem Then
                  ChecaInscrE = True
              End If
          Else
              ChecaInscrE = True
          End If
     Case "ES"    ' Espirito Santo
          strBase = Left(Trim(strOrigem) & "000000000", 9)
          intSoma = 0
          For intPos = 1 To 8
               intValor = Val(Mid$(strBase, intPos, 1))
               intValor = intValor * (10 - intPos)
               intSoma = intSoma + intValor
          Next
          intResto = intSoma Mod 11
          strDigito1 = Right(IIf(intResto < 2, "0", Str(11 - intResto)), 1)
          strBase2 = Left(strBase, 8) & strDigito1
          If strBase2 = strOrigem Then
              ChecaInscrE = True
          End If
     Case "GO"    ' Goias
          strBase = Left(Trim(strOrigem) & "000000000", 9)
          If InStr(1, "10,11,15", Left(strBase, 2), vbTextCompare) > 0 Then
              intSoma = 0
              For intPos = 1 To 8
                   intValor = Val(Mid$(strBase, intPos, 1))
                   intValor = intValor * (10 - intPos)
                   intSoma = intSoma + intValor
              Next
              intResto = intSoma Mod 11
              If intResto = 0 Then
                  strDigito1 = "0"
              ElseIf intResto = 1 Then
                  intNumero = Val(Left(strBase, 8))
                  strDigito1 = Right(IIf(intNumero >= 10103105 And intNumero <= 10119997, "1", "0"), 1)
              Else
                  strDigito1 = Right(Str(11 - intResto), 1)
              End If
              strBase2 = Left(strBase, 8) & strDigito1
              If strBase2 = strOrigem Then
                  ChecaInscrE = True
              End If
          End If
     Case "MA"    ' Maranh�o
          strBase = Left(Trim(strOrigem) & "000000000", 9)
          If Left(strBase, 2) = "12" Then
              intSoma = 0
              For intPos = 1 To 8
                   intValor = Val(Mid$(strBase, intPos, 1))
                   intValor = intValor * (10 - intPos)
                   intSoma = intSoma + intValor
              Next
              intResto = intSoma Mod 11
              strDigito1 = Right(IIf(intResto < 2, "0", Str(11 - intResto)), 1)
              strBase2 = Left(strBase, 8) & strDigito1
              If strBase2 = strOrigem Then
                  ChecaInscrE = True
              End If
          End If
     Case "MT"    ' Mato Grosso
          strBase = Left(Trim(strOrigem) & "0000000000", 10)
          intSoma = 0
          intPeso = 2
          For intPos = 10 To 1 Step -1
               intValor = Val(Mid$(strBase, intPos, 1))
               intValor = intValor * intPeso
               intSoma = intSoma + intValor
               intPeso = intPeso + 1
               If intPeso > 9 Then
                   intPeso = 2
               End If
          Next
          intResto = intSoma Mod 11
          strDigito1 = Right(IIf(intResto < 2, "0", Str(11 - intResto)), 1)
          strBase2 = Left(strBase, 10) & strDigito1
          If strBase2 = strOrigem Then
              ChecaInscrE = True
          End If
     Case "MS"    ' Mato Grosso do Sul
          strBase = Left(Trim(strOrigem) & "000000000", 9)
          If Left(strBase, 2) = "28" Then
              intSoma = 0
              For intPos = 1 To 8
                   intValor = Val(Mid$(strBase, intPos, 1))
                   intValor = intValor * (10 - intPos)
                   intSoma = intSoma + intValor
              Next
              intResto = intSoma Mod 11
              strDigito1 = Right(IIf(intResto < 2, "0", Str(11 - intResto)), 1)
              strBase2 = Left(strBase, 8) & strDigito1
              If strBase2 = strOrigem Then
                  ChecaInscrE = True
              End If
          End If
     Case "MG"    ' Minas Gerais
          strBase = Left(Trim(strOrigem) & "0000000000000", 13)
          strBase2 = Left(strBase, 3) & "0" & Mid$(strBase, 4, 8)
          intNumero = 2
          For intPos = 1 To 12
               intValor = Val(Mid$(strBase2, intPos, 1))
               intNumero = IIf(intNumero = 2, 1, 2)
               intValor = intValor * intNumero
               If intValor > 9 Then
                   strDigito1 = Format(intValor, "00")
                   intValor = Val(Left(strDigito1, 1)) + _
                              Val(Right(strDigito1, 1))
               End If
               intSoma = intSoma + intValor
          Next
          intValor = intSoma
          While Right(Format(intValor, "000"), 1) <> "0"
              intValor = intValor + 1
          Wend
          strDigito1 = Right(Format(intValor - intSoma, "00"), 1)
          strBase2 = Left(strBase, 11) & strDigito1
          intSoma = 0
          intPeso = 2
          For intPos = 12 To 1 Step -1
               intValor = Val(Mid$(strBase2, intPos, 1))
               intValor = intValor * intPeso
               intSoma = intSoma + intValor
               intPeso = intPeso + 1
               If intPeso > 11 Then
                   intPeso = 2
               End If
          Next
          intResto = intSoma Mod 11
          strDigito2 = Right(IIf(intResto < 2, "0", Str(11 - intResto)), 1)
          strBase2 = strBase2 & strDigito2
          If strBase2 = strOrigem Then
              ChecaInscrE = True
          End If
     Case "PA"    ' Para
          strBase = Left(Trim(strOrigem) & "000000000", 9)
          If Left(strBase, 2) = "15" Then
              intSoma = 0
              For intPos = 1 To 8
                   intValor = Val(Mid$(strBase, intPos, 1))
                   intValor = intValor * (10 - intPos)
                   intSoma = intSoma + intValor
              Next
              intResto = intSoma Mod 11
              strDigito1 = Right(IIf(intResto < 2, "0", Str(11 - intResto)), 1)
              strBase2 = Left(strBase, 8) & strDigito1
              If strBase2 = strOrigem Then
                  ChecaInscrE = True
              End If
          End If
     Case "PB"    ' Paraiba
          strBase = Left(Trim(strOrigem) & "000000000", 9)
          intSoma = 0
          For intPos = 1 To 8
               intValor = Val(Mid$(strBase, intPos, 1))
               intValor = intValor * (10 - intPos)
               intSoma = intSoma + intValor
          Next
          intResto = intSoma Mod 11
          intValor = 11 - intResto
          If intValor > 9 Then
              intValor = 0
          End If
          strDigito1 = Right(Str(intValor), 1)
          strBase2 = Left(strBase, 8) & strDigito1
          If strBase2 = strOrigem Then
              ChecaInscrE = True
          End If
     Case "PE"    ' Pernambuco
          strBase = Left(Trim(strOrigem) & "00000000000000", 14)
          intSoma = 0
          intPeso = 2
          For intPos = 13 To 1 Step -1
               intValor = Val(Mid$(strBase, intPos, 1))
               intValor = intValor * intPeso
               intSoma = intSoma + intValor
               intPeso = intPeso + 1
               If intPeso > 9 Then
                   intPeso = 2
               End If
          Next
          intResto = intSoma Mod 11
          intValor = 11 - intResto
          If intValor > 9 Then
              intValor = intValor - 10
          End If
          strDigito1 = Right(Str(intValor), 1)
          strBase2 = Left(strBase, 13) & strDigito1
          If strBase2 = strOrigem Then
              ChecaInscrE = True
          End If
     Case "PI"    ' Piaui
          strBase = Left(Trim(strOrigem) & "000000000", 9)
          intSoma = 0
          For intPos = 1 To 8
               intValor = Val(Mid$(strBase, intPos, 1))
               intValor = intValor * (10 - intPos)
               intSoma = intSoma + intValor
          Next
          intResto = intSoma Mod 11
          strDigito1 = Right(IIf(intResto < 2, "0", Str(11 - intResto)), 1)
          strBase2 = Left(strBase, 8) & strDigito1
          If strBase2 = strOrigem Then
              ChecaInscrE = True
          End If
     Case "PR"    ' Parana
          strBase = Left(Trim(strOrigem) & "0000000000", 10)
          intSoma = 0
          intPeso = 2
          For intPos = 8 To 1 Step -1
               intValor = Val(Mid$(strBase, intPos, 1))
               intValor = intValor * intPeso
               intSoma = intSoma + intValor
               intPeso = intPeso + 1
               If intPeso > 7 Then
                   intPeso = 2
               End If
          Next
          intResto = intSoma Mod 11
          strDigito1 = Right(IIf(intResto < 2, "0", Str(11 - intResto)), 1)
          strBase2 = Left(strBase, 8) & strDigito1
          intSoma = 0
          intPeso = 2
          For intPos = 9 To 1 Step -1
               intValor = Val(Mid$(strBase2, intPos, 1))
               intValor = intValor * intPeso
               intSoma = intSoma + intValor
               intPeso = intPeso + 1
               If intPeso > 7 Then
                   intPeso = 2
               End If
          Next
          intResto = intSoma Mod 11
          strDigito2 = Right(IIf(intResto < 2, "0", Str(11 - intResto)), 1)
          strBase2 = strBase2 & strDigito2
          If strBase2 = strOrigem Then
              ChecaInscrE = True
          End If
     Case "RJ"    ' Rio de Janeiro
          strBase = Left(Trim(strOrigem) & "00000000", 8)
          intSoma = 0
          intPeso = 2
          For intPos = 7 To 1 Step -1
               intValor = Val(Mid$(strBase, intPos, 1))
               intValor = intValor * intPeso
               intSoma = intSoma + intValor
               intPeso = intPeso + 1
               If intPeso > 7 Then
                   intPeso = 2
               End If
          Next
          intResto = intSoma Mod 11
          strDigito1 = Right(IIf(intResto < 2, "0", Str(11 - intResto)), 1)
          strBase2 = Left(strBase, 7) & strDigito1
          If strBase2 = strOrigem Then
              ChecaInscrE = True
          End If
     Case "RN"    ' Rio Grande do Norte
          strBase = Left(Trim(strOrigem) & "000000000", 9)
          If Left(strBase, 2) = "20" Then
              intSoma = 0
              For intPos = 1 To 8
                   intValor = Val(Mid$(strBase, intPos, 1))
                   intValor = intValor * (10 - intPos)
                   intSoma = intSoma + intValor
              Next
              intSoma = intSoma * 10
              intResto = intSoma Mod 11
              strDigito1 = Right(IIf(intResto > 9, "0", Str(intResto)), 1)
              strBase2 = Left(strBase, 8) & strDigito1
              If strBase2 = strOrigem Then
                  ChecaInscrE = True
              End If
          End If
     Case "RO"    ' Rondonia
          strBase = Left(Trim(strOrigem) & "000000000", 9)
          strBase2 = Mid$(strBase, 4, 5)
          intSoma = 0
          For intPos = 1 To 5
               intValor = Val(Mid$(strBase2, intPos, 1))
               intValor = intValor * (7 - intPos)
               intSoma = intSoma + intValor
          Next
          intResto = intSoma Mod 11
          intValor = 11 - intResto
          If intValor > 9 Then
              intValor = intValor - 10
          End If
          strDigito1 = Right(Str(intValor), 1)
          strBase2 = Left(strBase, 8) & strDigito1
          If strBase2 = strOrigem Then
              ChecaInscrE = True
          End If
     Case "RR"    ' Roraima
          strBase = Left(Trim(strOrigem) & "000000000", 9)
          If Left(strBase, 2) = "24" Then
              intSoma = 0
              For intPos = 1 To 8
                   intValor = Val(Mid$(strBase, intPos, 1))
                   intValor = intValor * (10 - intPos)
                   intSoma = intSoma + intValor
              Next
              intResto = intSoma Mod 9
              strDigito1 = Right(Str(intResto), 1)
              strBase2 = Left(strBase, 8) & strDigito1
              If strBase2 = strOrigem Then
                  ChecaInscrE = True
              End If
          End If
     Case "RS"    ' Rio Grande do Sul
          strBase = Left(Trim(strOrigem) & "0000000000", 10)
          intNumero = Val(Left(strBase, 3))
          If intNumero > 0 And intNumero < 468 Then
              intSoma = 0
              intPeso = 2
              For intPos = 9 To 1 Step -1
                   intValor = Val(Mid$(strBase, intPos, 1))
                   intValor = intValor * intPeso
                   intSoma = intSoma + intValor
                   intPeso = intPeso + 1
                   If intPeso > 9 Then
                       intPeso = 2
                   End If
              Next
              intResto = intSoma Mod 11
              intValor = 11 - intResto
              If intValor > 9 Then
                  intValor = 0
              End If
              strDigito1 = Right(Str(intValor), 1)
              strBase2 = Left(strBase, 9) & strDigito1
              If strBase2 = strOrigem Then
                  ChecaInscrE = True
              End If
          End If
     Case "SC"    ' Santa Catarina
          strBase = Left(Trim(strOrigem) & "000000000", 9)
          intSoma = 0
          For intPos = 1 To 8
               intValor = Val(Mid$(strBase, intPos, 1))
               intValor = intValor * (10 - intPos)
               intSoma = intSoma + intValor
          Next
          intResto = intSoma Mod 11
          strDigito1 = Right(IIf(intResto < 2, "0", Str(11 - intResto)), 1)
          strBase2 = Left(strBase, 8) & strDigito1
          If strBase2 = strOrigem Then
              ChecaInscrE = True
          End If
     Case "SE"    ' Sergipe
          strBase = Left(Trim(strOrigem) & "000000000", 9)
          intSoma = 0
          For intPos = 1 To 8
               intValor = Val(Mid$(strBase, intPos, 1))
               intValor = intValor * (10 - intPos)
               intSoma = intSoma + intValor
          Next
          intResto = intSoma Mod 11
          intValor = 11 - intResto
          If intValor > 9 Then
              intValor = 0
          End If
          strDigito1 = Right(Str(intValor), 1)
          strBase2 = Left(strBase, 8) & strDigito1
          If strBase2 = strOrigem Then
              ChecaInscrE = True
          End If
     Case "SP"    ' S�o Paulo
          If Left(strOrigem, 1) = "P" Then
              strBase = Left(Trim(strOrigem) & "0000000000000", 13)
              strBase2 = Mid$(strBase, 2, 8)
              intSoma = 0
              intPeso = 1
              For intPos = 1 To 8
                   intValor = Val(Mid$(strBase, intPos, 1))
                   intValor = intValor * intPeso
                   intSoma = intSoma + intValor
                   intPeso = intPeso + 1
                   If intPeso = 2 Then
                       intPeso = 3
                   End If
                   If intPeso = 9 Then
                       intPeso = 10
                   End If
              Next
              intResto = intSoma Mod 11
              strDigito1 = Right(Str(intResto), 1)
              strBase2 = Left(strBase, 8) & strDigito1 & Mid$(strBase, 11, 3)
          Else
              strBase = Left(Trim(strOrigem) & "000000000000", 12)
              intSoma = 0
              intPeso = 1
              For intPos = 1 To 8
                   intValor = Val(Mid$(strBase, intPos, 1))
                   intValor = intValor * intPeso
                   intSoma = intSoma + intValor
                   intPeso = intPeso + 1
                   If intPeso = 2 Then
                       intPeso = 3
                   End If
                   If intPeso = 9 Then
                       intPeso = 10
                   End If
              Next
              intResto = intSoma Mod 11
              strDigito1 = Right(Str(intResto), 1)
              strBase2 = Left(strBase, 8) & strDigito1 & Mid$(strBase, 10, 2)
              intSoma = 0
              intPeso = 2
              For intPos = 11 To 1 Step -1
                   intValor = Val(Mid$(strBase, intPos, 1))
                   intValor = intValor * intPeso
                   intSoma = intSoma + intValor
                   intPeso = intPeso + 1
                   If intPeso > 10 Then
                       intPeso = 2
                   End If
              Next
              intResto = intSoma Mod 11
              strDigito2 = Right(Str(intResto), 1)
              strBase2 = strBase2 & strDigito2
          End If
          If strBase2 = strOrigem Then
              ChecaInscrE = True
          End If
     Case "TO"    ' Tocantins
''          strBase = Left(Trim(strOrigem) & "00000000000", 11)
''          If InStr(1, "01,02,03,99", Mid$(strBase, 3, 2), vbTextCompare) > 0 Then
''              strBase2 = Left(strBase, 2) & Mid$(strBase, 5, 6)
''              intSoma = 0
''              For intPos = 1 To 8
''                   intValor = Val(Mid$(strBase2, intPos, 1))
''                   intValor = intValor * (10 - intPos)
''                   intSoma = intSoma + intValor
''              Next
''              intResto = intSoma Mod 11
''              strDigito1 = Right(IIf(intResto < 2, "0", Str(11 - intResto)), 1)
''              strBase2 = Left(strBase, 10) & strDigito1
''              If strBase2 = strOrigem Then
''                  ChecaInscrE = True
''              End If
''          End If

          ChecaInscrE = True
   End Select
   
End Function

Public Function RemoveCaracteres(sCampo As String) As String
Dim sNumeros As String, Contador As Integer

   For Contador = 1 To Len(sCampo)
       If IsNumeric(Mid(sCampo, Contador, 1)) Then
          sNumeros = sNumeros + Mid(sCampo, Contador, 1)
       End If
   Next
   
   RemoveCaracteres = sNumeros
   
End Function

Public Function ListarArquivos(ByVal Caminho As String) As String()
    'Aten��o: Fa�a refer�ncia � biblioteca Micrsoft Scripting Runtime
    Dim FSO As New FileSystemObject
    Dim result() As String
    Dim Pasta As Folder
    Dim Arquivo As File
    Dim indice As Long
 
    ReDim result(0) As String
    If FSO.FolderExists(Caminho) Then
        Set Pasta = FSO.GetFolder(Caminho)
 
        For Each Arquivo In Pasta.Files
            indice = IIf(result(0) = "", 0, indice + 1)
            ReDim Preserve result(indice) As String
            result(indice) = Arquivo.Name
        Next
    End If
 
    ListarArquivos = result
ErrHandler:
    Set FSO = Nothing
    Set Pasta = Nothing
    Set Arquivo = Nothing
End Function

Sub Verifica_BD()
Dim vStrutura As String

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
   
'''  Produtos
''   vStrutura = ""
''   vStrutura = vStrutura & "CREATE TABLE Produtos(idProduto Integer IDENTITY Primary Key NOT NULL, "
''   vStrutura = vStrutura & "idProduto Integer,"
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
''   vStrutura = vStrutura & ")"
''
''   cnSistema.Execute vStrutura


'   cnSistema.Execute ("CREATE TABLE Produtos(" & vStrutura & ")")


End Sub

Public Sub KillProcess(ByVal processName As String)

    Dim oWMI As Object
    Dim oServices As Object
    Dim oService As Object
    Dim oWMIServices As Object
    Dim oWMIService As Object
    
    Dim ret As Long
    Dim sService As String
    Dim servicename As String
    
    Set oWMI = GetObject("winmgmts:")
    Set oServices = oWMI.InstancesOf("win32_process")
    
    For Each oService In oServices
        servicename = LCase(Trim(CStr(oService.Name) & ""))
    
        If InStr(1, servicename, LCase(processName), vbTextCompare) > 0 Then
            ret = oService.Terminate
        End If
    Next
    
    Set oServices = Nothing
    Set oWMI = Nothing
End Sub

Public Sub Sendkeys(Text$, Optional wait As Boolean = False)
   Dim WshShell As Object
   Set WshShell = CreateObject("wscript.shell")
   WshShell.Sendkeys Text, wait
   Set WshShell = Nothing
End Sub

Public Function Calculo_DV11(ByVal strNumero As String) As String

   'Calculo modulo 11
   Dim i As Integer: Dim IntCont As Integer: Dim Vlr As Integer
   Dim Resto As Integer
   IntCont = 2
   Vlr = 0
   For i = Len(strNumero) To 1 Step -1
       Vlr = Vlr + (Val(Mid(strNumero, i, 1) * IntCont))
       IntCont = IIf(IntCont >= 9, 2, IntCont + 1)
   Next
   Resto = Vlr Mod 11
   Select Case Resto
       Case 0
           Resto = 0
       Case 1
           Resto = 0
       Case Is > 1
           Resto = Str(Val(11 - Resto))
   End Select
   Calculo_DV11 = Resto
        
End Function

Public Function StringToHex(ByRef StrToHex As String) As String
Dim strTemp   As String
Dim strReturn As String
Dim i         As Long
    For i = 1 To Len(StrToHex)
        strTemp = Hex$(Asc(Mid$(StrToHex, i, 1)))
        If Len(strTemp) = 1 Then strTemp = "0" & strTemp
        strReturn = strReturn & strTemp
    Next i
    StringToHex = LCase(strReturn)
End Function

Public Function PesquisarTAG(sCampo As String, sTAG As String) As String
Dim sConteudo As String
Dim sTAGInicio As String
Dim sTAGFim As String
Dim Contador As Double
Dim Contador2 As Double

   sTAGInicio = "<" & sTAG & ">"
   sTAGFim = "</" & sTAG & ">"
   
   Contador = 1
   For Contador = 1 To Len(sCampo)
     ' Verifica o Motivo
       If Mid(sCampo, Contador, Len(sTAGInicio)) = sTAGInicio Then
          For Contador2 = Contador To Len(sCampo)
              If Mid(sCampo, Contador2, Len(sTAGFim)) = sTAGFim Then
                 sConteudo = Mid(sCampo, Contador + Len(sTAGInicio), Contador2 - Contador - Len(sTAGInicio))
                 Contador2 = Len(sCampo)
              End If
          Next
       End If
   Next

   PesquisarTAG = sConteudo
   
End Function

Public Function ClonarRecordset(ORs As Recordset, _
                                Optional bApenasColunas As Boolean = False) As Recordset
On Error GoTo LErro
Dim iCont               As Integer
Dim orsAux              As Recordset
Dim varBookmark         As Variant
    
    Set orsAux = New Recordset
    For iCont = 0 To ORs.Fields.Count - 1
        If ORs.Fields(iCont).Type = adNumeric Then
            orsAux.Fields.Append ORs.Fields(iCont).Name, _
                    adDouble, ORs.Fields(iCont).DefinedSize, adFldMayBeNull
        Else
            orsAux.Fields.Append ORs.Fields(iCont).Name, _
                                ORs.Fields(iCont).Type, _
                                ORs.Fields(iCont).DefinedSize, _
                                adFldMayBeNull
        End If
    Next
    
    orsAux.Fields.Refresh
    orsAux.Open
    
    If bApenasColunas = False Then
        If Not ORs Is Nothing Then
            If Not (ORs.EOF And ORs.BOF) Then
                If ORs.EOF = True Then
                    varBookmark = -1
                ElseIf ORs.BOF = True Then
                    varBookmark = 0
                Else
                    varBookmark = ORs.Bookmark
                End If
                
                ORs.MoveFirst
                While Not ORs.EOF
                    orsAux.AddNew
                    For iCont = 0 To ORs.Fields.Count - 1
                        orsAux.Fields(iCont).value = ORs.Fields(iCont).value
                    Next
                    orsAux.Update
                    ORs.MoveNext
                Wend
                If Not IsEmpty(varBookmark) Then
                    If varBookmark = -1 Then
                        ORs.MoveLast
                    ElseIf varBookmark = 0 Then
                        ORs.MoveFirst
                    Else
                        ORs.Bookmark = varBookmark
                    End If
                End If
            End If
        End If
    End If
    Set ClonarRecordset = orsAux
    
    Exit Function
    Resume
LErro:
    Debug.Print "O campo " & ORs.Fields(iCont).Name & " j� foi incluido no Recordset"
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ClonarRecordset of M�dulo ModUtil"
    Set ClonarRecordset = Nothing
End Function

Public Function Validar_Permissao(ByVal intStatus As Integer, strMenu As String) As Boolean
   Validar_Acesso strMenu
   If Mid(idAcesso, intStatus, 1) = 0 Then
      MsgBox "Usu�rio n�o autorizado para esta opera��o", vbInformation, "Acesso Negado"
      Validar_Permissao = False
   Else
      Validar_Permissao = True
   End If
End Function

Public Function FAspas() As String
   FAspas = """"
End Function

Public Sub Validar_Acesso(strMenu As String)
Dim rsValidarAcesso As New ADODB.Recordset
   If idUser = 0 Then
      idAcesso = "1111"
      Exit Sub
   End If
'''''   Set rsValidarAcesso = cnSistema.Execute("SELECT idGrupoAcesso FROM Usuarios WHERE idUsuario = " & idUser)
'''''   If rsValidarAcesso!idGrupoAcesso = 1 Then
'''''      idAcesso = "1111"
'''''   Else
'''''      Set rsValidarAcesso = cnSistema.Execute("SELECT dbo.GrupoAcessoItens.Nivel FROM dbo.Usuarios INNER JOIN " & _
'''''                                              "dbo.GrupoAcesso ON dbo.Usuarios.idGrupoAcesso = dbo.GrupoAcesso.idGrupoAcesso INNER JOIN " & _
'''''                                              "dbo.GrupoAcessoItens ON dbo.GrupoAcesso.idGrupoAcesso = dbo.GrupoAcessoItens.idGrupoAcesso " & _
'''''                                              "WHERE (dbo.Usuarios.idUsuario = " & idUser & ") AND (dbo.GrupoAcessoItens.Name = '" & strMenu & "') ")
'''''      idAcesso = rsValidarAcesso!Nivel
'''''   End If
'''''   Set rsValidarAcesso = Nothing
End Sub

Public Function UTF8_Encode(ByVal sStr As String)
Dim l As Long, lChar As Integer, sUtf8 As String

   For l = 1 To Len(sStr)
       lChar = AscW(Mid(sStr, l, 1))
       If lChar < 128 Then
           sUtf8 = sUtf8 + Mid(sStr, l, 1)
       ElseIf ((lChar > 127) And (lChar < 2048)) Then
           sUtf8 = sUtf8 + Chr(((lChar \ 64) Or 192))
           sUtf8 = sUtf8 + Chr(((lChar And 63) Or 128))
       Else
           sUtf8 = sUtf8 + Chr(((lChar \ 144) Or 234))
           sUtf8 = sUtf8 + Chr((((lChar \ 64) And 63) Or 128))
           sUtf8 = sUtf8 + Chr(((lChar And 63) Or 128))
       End If
   Next l
   UTF8_Encode = sUtf8
   
End Function

Public Function RemoverENTERS(Valor As String) As String
Dim Remover As String, i As Byte, Temp As String
   Remover = Chr(13) & Chr(10) '"()*/-+"
   Temp = Valor
   For i = 1 To Len(Valor)
       Temp = Replace(Temp, Mid(Remover, i, 1), "")
   Next
   RemoverENTERS = Temp
End Function

Function CPF_CNPJ(strCPF_CNPJ As String)
Dim vCPF_CNPJ, vDigito As String
Dim Numero(14), vResto, vResultado, vSomaDigito10, vResto1 As Integer

   CPF_CNPJ = False
   vCPF_CNPJ = Format(Replace(Replace(Replace(strCPF_CNPJ, ".", ""), "-", ""), "/", ""), "@@@@@@@@@@@@@@")
   vDigito = Mid(vCPF_CNPJ, 13, 2)

   Numero(1) = Val(Mid(vCPF_CNPJ, 1, 1))
   Numero(2) = Val(Mid(vCPF_CNPJ, 2, 1))
   Numero(3) = Val(Mid(vCPF_CNPJ, 3, 1))
   Numero(4) = Val(Mid(vCPF_CNPJ, 4, 1))
   Numero(5) = Val(Mid(vCPF_CNPJ, 5, 1))
   Numero(6) = Val(Mid(vCPF_CNPJ, 6, 1))
   Numero(7) = Val(Mid(vCPF_CNPJ, 7, 1))
   Numero(8) = Val(Mid(vCPF_CNPJ, 8, 1))
   Numero(9) = Val(Mid(vCPF_CNPJ, 9, 1))
   Numero(10) = Val(Mid(vCPF_CNPJ, 10, 1))
   Numero(11) = Val(Mid(vCPF_CNPJ, 11, 1))
   Numero(12) = Val(Mid(vCPF_CNPJ, 12, 1))
   Numero(13) = Val(Mid(vCPF_CNPJ, 13, 1))
   Numero(14) = Val(Mid(vCPF_CNPJ, 14, 1))

   If Len(Trim(vCPF_CNPJ)) > 11 Then 'CNPJ
       vResultado = (Numero(1) * 5) + (Numero(2) * 4) + (Numero(3) * 3) + (Numero(4) * 2) + (Numero(5) * 9) + (Numero(6) * 8) + (Numero(7) * 7) + (Numero(8) * 6) + (Numero(9) * 5) + (Numero(10) * 4) + (Numero(11) * 3) + (Numero(12) * 2)
       vResto = vResultado Mod 11
       If vResto < 2 Then vResto1 = 0 Else vResto1 = 11 - vResto
       If vResto1 <> Numero(13) Then Exit Function
       vResultado = (Numero(1) * 6) + (Numero(2) * 5) + (Numero(3) * 4) + (Numero(4) * 3) + (Numero(5) * 2) + (Numero(6) * 9) + (Numero(7) * 8) + (Numero(8) * 7) + (Numero(9) * 6) + (Numero(10) * 5) + (Numero(11) * 4) + (Numero(12) * 3) + (Numero(13) * 2)
       vResto = vResultado Mod 11
       If vResto < 2 Then vResto1 = 0 Else vResto1 = 11 - vResto
       If vResto1 <> Numero(14) Then Exit Function
   Else 'CPF
       vResultado = (Numero(4) * 1) + (Numero(5) * 2) + (Numero(6) * 3) + (Numero(7) * 4) + (Numero(8) * 5) + (Numero(9) * 6) + (Numero(10) * 7) + (Numero(11) * 8) + (Numero(12) * 9)
       vResto = vResultado Mod 11
       If vResto > 9 Then vResto1 = vResto - 10 Else vResto1 = vResto
       If vResto1 <> Numero(13) Then Exit Function
       vResultado = (Numero(5) * 1) + (Numero(6) * 2) + (Numero(7) * 3) + (Numero(8) * 4) + (Numero(9) * 5) + (Numero(10) * 6) + (Numero(11) * 7) + (Numero(12) * 8) + (vResto1 * 9)
       vResto = vResultado Mod 11
       If vResto > 9 Then vResto1 = vResto - 10 Else vResto1 = vResto
       If vResto1 <> Numero(14) Then Exit Function
   End If
   CPF_CNPJ = True
End Function


'''''Public Function URLEncode_UTF8( _
'''''      ByVal Text As String _
'''''   ) As String
'''''
'''''   Dim Index1 As Long
'''''   Dim Index2 As Long
'''''   Dim Result As String
'''''   Dim Chars() As Byte
'''''   Dim Char As String
'''''   Dim Byte1 As Byte
'''''   Dim Byte2 As Byte
'''''   Dim UTF16 As Long
'''''
'''''   For Index1 = 1 To Len(Text)
'''''      CopyToMemory Byte1, ByVal StrPtr(Text) + ((Index1 - 1) * 2), 1
'''''      CopyToMemory Byte2, ByVal StrPtr(Text) + ((Index1 - 1) * 2) + 1, 1
'''''
'''''      UTF16 = Byte2
'''''      UTF16 = UTF16 * 256 + Byte1
'''''      Chars = GetUTF8FromUTF16(UTF16)
'''''      For Index2 = LBound(Chars) To UBound(Chars)
'''''         Char = Chr(Chars(Index2))
'''''         If Char Like "[0-9A-Za-z]" Then
'''''            Result = Result & Char
'''''         Else
'''''            Result = Result & "%" & Hex(Asc(Char))
'''''         End If
'''''      Next
'''''   Next
''''''   GetEncodedUTF8String = Result
'''''   URLEncode_UTF8 = Result
'''''End Function
'''''
'''''Private Function GetUTF8FromUTF16( _
'''''      ByVal UTF16 As Long _
'''''   ) As Byte()
'''''
'''''   Dim Result() As Byte
'''''   If UTF16 < &H80 Then
'''''      ReDim Result(0 To 0)
'''''      Result(0) = UTF16
'''''   ElseIf UTF16 < &H800 Then
'''''      ReDim Result(0 To 1)
'''''      Result(1) = &H80 + (UTF16 And &H3F)
'''''      UTF16 = UTF16 \ &H40
'''''      Result(0) = &HC0 + (UTF16 And &H1F)
'''''   Else
'''''      ReDim Result(0 To 2)
'''''      Result(2) = &H80 + (UTF16 And &H3F)
'''''      UTF16 = UTF16 \ &H40
'''''      Result(1) = &H80 + (UTF16 And &H3F)
'''''      UTF16 = UTF16 \ &H40
'''''      Result(0) = &HE0 + (UTF16 And &HF)
'''''   End If
'''''   GetUTF8FromUTF16 = Result
'''''End Function
'''''

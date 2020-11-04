VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLerXML 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ler XML"
   ClientHeight    =   4995
   ClientLeft      =   105
   ClientTop       =   435
   ClientWidth     =   15810
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   15810
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdLer 
      Caption         =   "Confirma"
      Height          =   315
      Left            =   3960
      TabIndex        =   1
      Top             =   4560
      Width           =   1095
   End
   Begin MSComctlLib.ListView lvwMensagens 
      Height          =   4320
      Left            =   60
      TabIndex        =   0
      Top             =   180
      Width           =   15645
      _ExtentX        =   27596
      _ExtentY        =   7620
      View            =   3
      LabelEdit       =   1
      SortOrder       =   -1  'True
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
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
      Left            =   960
      TabIndex        =   2
      Top             =   4575
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
      Left            =   2880
      TabIndex        =   3
      Top             =   4575
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "99/99/9999"
      PromptChar      =   " "
   End
   Begin VB.Label lblDataFinal 
      AutoSize        =   -1  'True
      Caption         =   "Data Final"
      Height          =   195
      Left            =   2040
      TabIndex        =   5
      Top             =   4620
      Width           =   720
   End
   Begin VB.Label lblDataInicial 
      AutoSize        =   -1  'True
      Caption         =   "Data Inicial"
      Height          =   195
      Left            =   60
      TabIndex        =   4
      Top             =   4620
      Width           =   795
   End
End
Attribute VB_Name = "frmLerXML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ItemList As ListItem
Dim ProcuraItem As ListItem
Dim IdMensagens As Integer
'Dim sEmpresaNFe As String
Dim rsTemp As New ADODB.Recordset

Private Sub Form_Load()
   
   Centraliza frmLerXML
   
   mskDataInicial.text = Date
   mskDataFinal.text = Date
   
'''''   I_EmpresaNF = LerArquivoINI("NFe", "Empresa", CaminhoINI & "\System.ini")
   
   IdMensagens = 1 ' Inicia o contador de chave das mensagens
   
   lvwMensagens.ColumnHeaders.Add , , "Data", 1800
   lvwMensagens.ColumnHeaders.Add , , "N�mero", 850
   lvwMensagens.ColumnHeaders.Add , , "Mensagem", 6650
   lvwMensagens.ColumnHeaders.Add , , "Arquivo", 6000
   
   cmdLer_Click

End Sub

Private Sub cmdLer_Click2()
Dim Contador As Integer
Dim Contador2 As Integer
Dim sStatus As String
Dim sMotivos As String
Dim sProtocolo As String
Dim sNumeroNota As String
Dim handle As Integer
Dim Linha As String
Dim strMensagem As String
Dim xMotivo As String
Dim bRetorno As Boolean

   I_EmpresaNF = LerArquivoINI("NFe", "Empresa", CaminhoINI & "\System.ini")
   
   Dim Arquivos() As String
   Dim lCtr As Long
   Arquivos = ListarArquivos("C:\NF-e\" & I_EmpresaNF & "\Retorno")
'   If UBound(Arquivos) > 0 And Len(Trim(Arquivos(lCtr))) <= 25 Then
   If UBound(Arquivos) > 0 Then
      For lCtr = 0 To UBound(Arquivos)
         If UCase(Mid(Arquivos(lCtr), Len(Trim(Arquivos(lCtr))) - 7, 8)) = "-INU.XML" Then
            Contador = 1
            Screen.MousePointer = vbHourglass
            frmVisualiza.lvwDados.ListItems.Clear
            frmVisualiza.lvwDados.ColumnHeaders.Clear
            frmVisualiza.lvwDados.ColumnHeaders.Add , , "C�d.", 700
            frmVisualiza.lvwDados.ColumnHeaders.Add , , "Motivo", 5300
         
            ' Verifica se Inutiliza��o foi aceita
            handle = FreeFile
            Open "C:\NF-e\" & I_EmpresaNF & "\Retorno\" & Arquivos(lCtr) For Input As #handle
            'Open "C:\NF-e\" & sEmpresaNFe & "\Retorno\" & Trim(rsNFeInutilizadas!ChaveNFe) & "-inu.XML" For Input As #handle
            
            Line Input #handle, Linha
            
            Contador = 1
            For Contador = 1 To Len(Linha)
                If Mid(Linha, Contador, 7) = "<nProt>" Then
                   sProtocolo = Mid(Linha, Contador + 7, 15)
                End If
                
                If Mid(Linha, Contador, 8) = "<nNFIni>" Then
                   For Contador2 = Contador To Len(Linha)
                       If Mid(Linha, Contador2, 9) = "</nNFIni>" Then
                          sNumeroNota = Mid(Linha, Contador + 8, Contador2 - Contador - 8)
                          Contador2 = Len(Linha)
                       End If
                   Next
                End If
            
              ' Verifica o Status
                If Mid(Linha, Contador, 7) = "<cStat>" Then
                   sStatus = RemoveCaracteres(Mid(Linha, Contador + 7, 3))
                End If
                
              ' Verifica o Motivo
                If Mid(Linha, Contador, 9) = "<xMotivo>" Then
                   For Contador2 = Contador To Len(Linha)
                       If Mid(Linha, Contador2, 10) = "</xMotivo>" Then
                          xMotivo = Mid(Linha, Contador + 9, Contador2 - Contador - 9)
                          Contador2 = Len(Linha)
                       End If
                   Next
                
                   sStatus = RemoveCaracteres(Mid(Linha, Contador + 7, 3))
                End If
                
              ' N�mero do Protocolo
                If Mid(Linha, Contador, 6) = "<nRec>" Then
                   sProtocolo = Mid(Linha, Contador + 6, 15)
                End If
                
                If Trim(sStatus) <> "" Then
                   Set rsTemp = cnSistema.Execute("Select * From NFeStatus WHERE Codigo = " & sStatus)
                   If Not rsTemp.EOF And sStatus <> "" Then
                      Set ProcuraItem = frmVisualiza.lvwDados.FindItem(sStatus)
                      If ProcuraItem Is Nothing Then
                         Set ItemList = frmVisualiza.lvwDados.ListItems.Add(, "R" & CStr(rsTemp!idNFeStatus), sStatus)
                             ItemList.SubItems(1) = rsTemp!Descricao
                      End If
                      
                      sStatus = ""
                      sProtocolo = ""
                   Else
                      If Trim(xMotivo) <> "" Then
                         Set ItemList = frmVisualiza.lvwDados.ListItems.Add(, "R" & CStr(frmVisualiza.lvwDados.ListItems.Count + 99999), sStatus)
                             ItemList.SubItems(1) = xMotivo
                        
                         sStatus = ""
                         sProtocolo = ""
                      End If
                   End If
                End If
            Next
            
            If Trim(sProtocolo) <> "" Then
               cnSistema.Execute "Update NFeInutilizadas set " & _
                        "Protocolo = '" & sProtocolo & "' " & _
                        "Where Numero = " & sNumeroNota
            End If
            
            Close #handle
            
            FileCopy "C:\NF-e\" & I_EmpresaNF & "\Retorno\" & Arquivos(lCtr), "C:\NF-e\" & I_EmpresaNF & "\Temp\" & Arquivos(lCtr)
            Kill "C:\NF-e\" & I_EmpresaNF & "\Retorno\" & Arquivos(lCtr)
            
            ' Visualiza o Retorno
            If frmVisualiza.lvwDados.ListItems.Count > 0 Then
               Screen.MousePointer = vbDefault
               frmVisualiza.Show vbModal
            End If
        
         ElseIf UCase(Mid(Arquivos(lCtr), Len(Trim(Arquivos(lCtr))) - 2, 3)) = "XML" Then
            Contador = 1
            Screen.MousePointer = vbHourglass
            frmVisualiza.lvwDados.ListItems.Clear
            frmVisualiza.lvwDados.ColumnHeaders.Clear
            frmVisualiza.lvwDados.ColumnHeaders.Add , , "C�d.", 700
            frmVisualiza.lvwDados.ColumnHeaders.Add , , "Motivo", 5300
         
            handle = FreeFile
            Open "C:\NF-e\" & I_EmpresaNF & "\Retorno\" & Arquivos(lCtr) For Input As #handle
            
            Line Input #handle, Linha
            
            Contador = 1
            For Contador = 1 To Len(Linha)
              ' Verifica o Status
                If Mid(Linha, Contador, 7) = "<cStat>" Then
                   sStatus = RemoveCaracteres(Mid(Linha, Contador + 7, 3))
                End If
                
              ' Verifica o Motivo
                If Mid(Linha, Contador, 9) = "<xMotivo>" Then
                   For Contador2 = Contador To Len(Linha)
                       If Mid(Linha, Contador2, 10) = "</xMotivo>" Then
                          xMotivo = Mid(Linha, Contador + 9, Contador2 - Contador - 9)
                          Contador2 = Len(Linha)
                       End If
                   Next
                
                   sStatus = RemoveCaracteres(Mid(Linha, Contador + 7, 3))
                End If
                
              ' N�mero do Protocolo
                If Mid(Linha, Contador, 6) = "<nRec>" Then
                   sProtocolo = Mid(Linha, Contador + 6, 15)
                End If
                
                If Trim(sStatus) <> "" Then
                   Set rsTemp = cnSistema.Execute("Select * From NFeStatus WHERE Codigo = " & sStatus)
                   If Not rsTemp.EOF And sStatus <> "" Then
                      Set ProcuraItem = frmVisualiza.lvwDados.FindItem(sStatus)
                      If ProcuraItem Is Nothing Then
                         Set ItemList = frmVisualiza.lvwDados.ListItems.Add(, "R" & CStr(rsTemp!idNFeStatus), sStatus)
                             ItemList.SubItems(1) = rsTemp!Descricao
                      End If
                      
                      sStatus = ""
                      sProtocolo = ""
                   Else
                      If Trim(xMotivo) <> "" Then
                         Set ItemList = frmVisualiza.lvwDados.ListItems.Add(, "R" & CStr(frmVisualiza.lvwDados.ListItems.Count + 99999), sStatus)
                             ItemList.SubItems(1) = xMotivo
                        
                         sStatus = ""
                         sProtocolo = ""
                      End If
                   End If
                End If
            Next
            Close #handle
            
            FileCopy "C:\NF-e\" & I_EmpresaNF & "\Retorno\" & Arquivos(lCtr), "C:\NF-e\" & I_EmpresaNF & "\Temp\" & Arquivos(lCtr)
            Kill "C:\NF-e\" & I_EmpresaNF & "\Retorno\" & Arquivos(lCtr)
            
            ' Visualiza o Retorno
            If frmVisualiza.lvwDados.ListItems.Count > 0 Then
               Screen.MousePointer = vbDefault
               frmVisualiza.Show vbModal
            End If
            
         ElseIf UCase(Mid(Arquivos(lCtr), Len(Trim(Arquivos(lCtr))) - 2, 3)) = "err" Then

            handle = FreeFile
            Open "C:\NF-e\" & I_EmpresaNF & "\Retorno\" & Arquivos(lCtr) For Input As #handle
            
            Line Input #handle, Linha

            bRetorno = False
            While Not EOF(handle)
               Line Input #handle, Linha

               strMensagem = strMensagem & Linha & Chr(13)
            Wend

            MsgBox strMensagem, vbExclamation + vbOKOnly, "Erro"
            Close #handle
            
            FileCopy "C:\NF-e\" & I_EmpresaNF & "\Retorno\" & Arquivos(lCtr), "C:\NF-e\" & I_EmpresaNF & "\Temp\" & Arquivos(lCtr)
            Kill "C:\NF-e\" & I_EmpresaNF & "\Retorno\" & Arquivos(lCtr)
   
         ElseIf UCase(Mid(Arquivos(lCtr), Len(Trim(Arquivos(lCtr))) - 2, 3)) = "TXT" Then
            FileCopy "C:\NF-e\" & I_EmpresaNF & "\Retorno\" & Arquivos(lCtr), "C:\NF-e\" & I_EmpresaNF & "\Temp\" & Arquivos(lCtr)
            Kill "C:\NF-e\" & I_EmpresaNF & "\Retorno\" & Arquivos(lCtr)
         
         End If
      Next
   End If
End Sub

''Private Sub cmdLer_Click()
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
''Dim Texto As String
''
''   Arquivos = ListarArquivos("C:\XML\NFe 2 - Modelos XML de Retorno")
''   If UBound(Arquivos) > 0 Then
''      For lCtr = 0 To UBound(Arquivos)
''          If UCase(Mid(Arquivos(lCtr), Len(Trim(Arquivos(lCtr))) - 3, 4)) = ".XML" Then
'''''''''''''''''''''''''''''''
''            handle = FreeFile
''            'Open ReadUTF8File("C:\XML\NFe 2 - Modelos XML de Retorno" & "\" & Arquivos(lCtr)) For Input As #handle
''            Linha = ReadUTF8File("C:\XML\NFe 2 - Modelos XML de Retorno" & "\" & Arquivos(lCtr))
''            Open "C:\XML\NFe 2 - Modelos XML de Retorno" & "\" & Arquivos(lCtr) For Input As #handle
''            'Line Input #handle, Linha
''
''            Contador = 1
''            For Contador = 1 To Len(Linha)
''              ' Verifica o Status
''                If Mid(Linha, Contador, 7) = "<cStat>" Then
''                   sStatus = RemoveCaracteres(Mid(Linha, Contador + 7, 3))
''                End If
''
''              ' Verifica o Motivo
''                If Mid(Linha, Contador, 9) = "<xMotivo>" Then
''                   For Contador2 = Contador To Len(Linha)
''                       If Mid(Linha, Contador2, 10) = "</xMotivo>" Then
''                          xMotivo = Mid(Linha, Contador + 9, Contador2 - Contador - 9)
''                          Contador2 = Len(Linha)
''                       End If
''                   Next
''
'''                   sStatus = RemoveCaracteres(Mid(Linha, Contador + 7, 3))
''                End If
''
''            Next
''            Close #handle
''
''            If sStatus <> "" Then
''               Set ItemList = lvwMensagens.ListItems.Add(, "R" & CStr(lCtr), sStatus)
''                   ItemList.SubItems(1) = xMotivo
''                   ItemList.SubItems(2) = Arquivos(lCtr)
''
''               sStatus = ""
''               xMotivo = ""
''            End If
''
''''            FileCopy "C:\NF-e\" & sEmpresaNFe & "\Retorno\" & Arquivos(lCtr), "C:\NF-e\" & sEmpresaNFe & "\Temp\" & Arquivos(lCtr)
''''            Kill "C:\NF-e\" & sEmpresaNFe & "\Retorno\" & Arquivos(lCtr)
''
'''''''''''''''''''''''''''''''
''          End If
''      Next
''   End If
''
''End Sub
''
''

Private Function ReadUTF8File(sFile) As String
    Const ForReading = 1
    Dim sPrefix
    Dim pvReadFile
    

    With CreateObject("Scripting.FileSystemObject")
        sPrefix = .OpenTextFile(sFile, ForReading, False, False).Read(3)
    End With
    If Left(sPrefix, 3) <> Chr(&HEF) & Chr(&HBB) & Chr(&HBF) Then
        With CreateObject("Scripting.FileSystemObject")
            pvReadFile = .OpenTextFile(sFile, ForReading, False, Left(sPrefix, 2) = Chr(&HFF) & Chr(&HFE)).ReadAll()
            ReadUTF8File = pvReadFile
        End With
    Else
        With CreateObject("ADODB.Stream")
            .Open
            If Left(sPrefix, 2) = Chr(&HFF) & Chr(&HFE) Then
                .Charset = "Unicode"
            ElseIf Left(sPrefix, 3) = Chr(&HEF) & Chr(&HBB) & Chr(&HBF) Then
                .Charset = "UTF-8"
            Else
                .Charset = "_autodetect"
            End If
            .LoadFromFile sFile
            pvReadFile = .ReadText
            ReadUTF8File = pvReadFile
        End With
    End If
End Function

Private Sub cmdLer_Click()
Dim Arquivos() As String
Dim lCtr As Long
Dim Contador As Integer
Dim Contador2 As Integer
   
'Dim sStatus As String
Dim sMotivos As String
'Dim sProtocolo As String
'Dim sNumeroNota As String
Dim handle As Integer
Dim Linha As String
Dim xMotivo As String
Dim strMensagem As String
Dim sStatus As String
Dim sDataRecebimento As String
Dim sNomeArquivo As String
Dim dDataArquivo As String
Dim bSeparaNome As Boolean
Dim sCaminhoArquivoRetorno As String
Dim sCaminhoArquivoTemp As String
   
   lvwMensagens.ListItems.Clear

   sCaminhoArquivoRetorno = "C:\NF-e\" & I_EmpresaNF & "\Retorno\"
'   sCaminhoArquivoRetorno = "C:\XML\NFe 2 - Modelos XML de Retorno\"
   sCaminhoArquivoTemp = "C:\NF-e\" & I_EmpresaNF & "\Temp\"
   
'   Arquivos = ListarArquivos(sCaminhoArquivoRetorno)
   Arquivos = ListarArquivos(sCaminhoArquivoTemp)
   If UBound(Arquivos) > 0 Then
      For lCtr = 0 To UBound(Arquivos)
          bSeparaNome = False
          dDataArquivo = ""
          sNomeArquivo = ""
         
          If UCase(Mid(Arquivos(lCtr), Len(Trim(Arquivos(lCtr))) - 3, 4)) = ".XML" Then
'          If UCase(Mid(Arquivos(lCtr), Len(Trim(Arquivos(lCtr))) - 7, 8)) = "-INU.XML" Then

            For Contador = 1 To Len(Arquivos(lCtr))
               If Mid(Arquivos(lCtr), Contador, 1) = "\" Then
                  Contador = Contador + 1
                  bSeparaNome = True
               End If
               
               If Not bSeparaNome Then
                  dDataArquivo = dDataArquivo & Mid(Arquivos(lCtr), Contador, 1)
               Else
                  sNomeArquivo = sNomeArquivo & Mid(Arquivos(lCtr), Contador, 1)
               End If
            Next
            
            handle = FreeFile
'            Open sCaminhoArquivoRetorno & Arquivos(lCtr) For Input As #handle
'            Open sCaminhoArquivoRetorno & sNomeArquivo For Input As #handle
            Open sCaminhoArquivoTemp & sNomeArquivo For Input As #handle
            
            Line Input #handle, Linha
            
            sStatus = PesquisarTAG(Linha, "cStat")
            xMotivo = PesquisarTAG(Linha, "xMotivo")
            sDataRecebimento = PesquisarTAG(Linha, "dhRecbto")
            
            Close #handle
            
            If sStatus <> "" Then
            
               If CDate(Mid(dDataArquivo, 1, 10)) >= mskDataInicial.text And CDate(Mid(dDataArquivo, 1, 10)) <= mskDataFinal.text Then
                  Set ItemList = lvwMensagens.ListItems.Add(, "R" & CStr(IdMensagens), dDataArquivo)
                      ItemList.SubItems(1) = sStatus
                      ItemList.SubItems(2) = xMotivo
                      ItemList.SubItems(3) = sNomeArquivo
                      
                  IdMensagens = IdMensagens + 1
   
                  sStatus = ""
                  xMotivo = ""
               End If
            End If
            
'''            FileCopy "C:\NF-e\" & sEmpresaNFe & "\Retorno\" & Arquivos(lCtr), "C:\NF-e\" & sEmpresaNFe & "\Temp\" & Arquivos(lCtr)
'''            Kill "C:\NF-e\" & sEmpresaNFe & "\Retorno\" & Arquivos(lCtr)

'''            FileCopy sCaminhoArquivoRetorno & Arquivos(lCtr), sCaminhoArquivoTemp & Arquivos(lCtr)
'''            Kill sCaminhoArquivoRetorno & Arquivos(lCtr)


          End If
      Next
   End If

End Sub

Public Function PesquisarTAG(sCampo As String, sTAG As String) As String
Dim sConteudo As String
Dim sTAGInicio As String
Dim sTAGFim As String
Dim Contador As Integer
Dim Contador2 As Integer

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
            result(indice) = Arquivo.DateCreated & "\" & Arquivo.Name
        Next
    End If

    ListarArquivos = result
ErrHandler:
    Set FSO = Nothing
    Set Pasta = Nothing
    Set Arquivo = Nothing
End Function


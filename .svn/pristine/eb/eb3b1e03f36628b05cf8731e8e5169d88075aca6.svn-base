VERSION 5.00
Begin VB.Form frmNFeGAtualizar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Atualizar NFe"
   ClientHeight    =   2040
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6930
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   6930
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbSituacao 
      Height          =   315
      ItemData        =   "frmNFeGAtualizar.frx":0000
      Left            =   2580
      List            =   "frmNFeGAtualizar.frx":0010
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   900
      Width           =   2805
   End
   Begin VB.CommandButton cmdGravar 
      Caption         =   "Gravar"
      Default         =   -1  'True
      Height          =   375
      Left            =   4440
      TabIndex        =   6
      Top             =   1560
      Width           =   1155
   End
   Begin VB.CommandButton cmdFechar 
      Caption         =   "&Fechar"
      Height          =   375
      Left            =   5700
      TabIndex        =   5
      Top             =   1560
      Width           =   1155
   End
   Begin VB.Frame Frame1 
      Height          =   75
      Left            =   60
      TabIndex        =   4
      Top             =   1380
      Width           =   6795
   End
   Begin VB.TextBox txtProtocolo 
      Height          =   315
      Left            =   60
      TabIndex        =   2
      Top             =   900
      Width           =   2460
   End
   Begin VB.TextBox txtChaveAcesso 
      Height          =   315
      Left            =   60
      TabIndex        =   0
      Top             =   315
      Width           =   6780
   End
   Begin VB.Label lblSituacao 
      AutoSize        =   -1  'True
      Caption         =   "Situação"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2580
      TabIndex        =   8
      Top             =   660
      Width           =   630
   End
   Begin VB.Label lblProtocolo 
      AutoSize        =   -1  'True
      Caption         =   "Protocolo de Autorização"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   60
      TabIndex        =   3
      Top             =   660
      Width           =   1785
   End
   Begin VB.Label lblChaveAcesso 
      AutoSize        =   -1  'True
      Caption         =   "Chave de Acesso"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   1260
   End
End
Attribute VB_Name = "frmNFeGAtualizar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ItemList As ListItem
Dim rsTemp As New ADODB.Recordset
Dim intID As Integer

Private Sub Form_Load()
   On Error GoTo Erro

   Centraliza frmNFeGAtualizar
   
   intID = frmNFeG.lblAtualizarNFe.Tag
   Set rsTemp = cnSistema.Execute("Select * From NFe Where idNFe = " & intID)
   If Not rsTemp.EOF Then
      txtChaveAcesso.text = IIf(Not IsNull(rsTemp!ChaveNFe), rsTemp!ChaveNFe, "")
      cmbSituacao.ListIndex = rsTemp!Situacao
      txtProtocolo.text = IIf(Not IsNull(rsTemp!Protocolo), rsTemp!Protocolo, "")
   End If

Exit Sub
Erro:
   If Err.Number = -2147467259 Then
      rsErro = True
      Beep
      MsgBox "Erro na Abertura do Arquivo de Dados" & Chr(13) & "Algum usuário está com o Arquivo em modo Exclusivo", vbExclamation, "Erro"
      Exit Sub
   Else
      rsErro = True
      Beep
      MsgBox "Verificar: " & Err.Number & Chr(13) & Err.Description, vbExclamation, "Sistema"
      Exit Sub
   End If
End Sub

Private Sub cmdFechar_Click()
   Unload Me
End Sub

Private Sub cmdGravar_Click()
   If MsgBox("Confirma Alterar o registro atual", vbYesNo + vbQuestion, "Alteração") = vbYes Then
      cnSistema.Execute "Update NFe set " & _
            "ChaveNFe = '" & txtChaveAcesso.text & "', " & _
            "Situacao = " & cmbSituacao.ListIndex & ", " & _
            "Protocolo = '" & txtProtocolo.text & "' " & _
            "Where idNFe = " & rsTemp!idNFe
      Atividade "Alterar: " & txtChaveAcesso.text, Me.Caption
   
      frmNFeG.txtChaveAcesso.text = txtChaveAcesso.text
      frmNFeG.txtProtocolo.text = txtProtocolo.text
      
      Unload Me
   End If
End Sub


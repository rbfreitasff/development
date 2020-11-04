VERSION 5.00
Begin VB.Form frmEventos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Eventos"
   ClientHeight    =   5970
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8595
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   8595
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   7380
      TabIndex        =   9
      Top             =   5520
      Width           =   1155
   End
   Begin VB.CommandButton cmdConfirmar 
      Caption         =   "Confirmar"
      Height          =   375
      Left            =   6120
      TabIndex        =   8
      Top             =   5520
      Width           =   1155
   End
   Begin VB.Frame fraJustificativas 
      Caption         =   "Justificativa"
      Height          =   1995
      Left            =   120
      TabIndex        =   6
      Top             =   3480
      Width           =   8415
      Begin VB.TextBox txtJustificativa 
         Height          =   1575
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   300
         Width           =   8175
      End
   End
   Begin VB.Frame fraInformacoes 
      Caption         =   "Informa��es"
      Height          =   3315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8415
      Begin VB.TextBox txtInformacoes 
         Enabled         =   0   'False
         Height          =   1875
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   900
         Width           =   8175
      End
      Begin VB.TextBox txtChaveNFe 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1380
         TabIndex        =   4
         Top             =   540
         Width           =   6915
      End
      Begin VB.TextBox txtNota 
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   540
         Width           =   1215
      End
      Begin VB.Label lblInformacoesPrazo 
         BorderStyle     =   1  'Fixed Single
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
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Top             =   2820
         Width           =   8175
      End
      Begin VB.Label lblChave 
         Caption         =   "Chave"
         Height          =   195
         Left            =   1380
         TabIndex        =   3
         Top             =   300
         Width           =   555
      End
      Begin VB.Label lblNota 
         Caption         =   "Nota"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   300
         Width           =   435
      End
   End
End
Attribute VB_Name = "frmEventos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim sEvento As String

Private Sub Form_Load()

   txtNota.Enabled = True

   sEvento = frmGerenciarNF.lblStatusEvento.Tag
   If sEvento = "CN" Then
      txtChaveNFe.Text = FFormataChaveNF(rsNFs!cNF)
      
      txtNota.Enabled = False
      txtNota.Text = rsNFs!nnf
      
      txtInformacoes.Text = "O cancelamento da nota deve ser efetuado em at� 24 horas ap�s a data e hora da aprova��o"
      
   ElseIf sEvento = "IN" Then
      
      txtInformacoes.Text = "Um n�mero n�o de nota fiscal n�o pode ser Inutilizado caso a nota � tenha sido Aprovada, Cancelada ou Denegada."
      
   ElseIf sEvento = "CC" Then
      txtNota.Enabled = False
      
      txtChaveNFe.Text = FFormataChaveNF(rsNFs!cNF)
      txtNota.Text = rsNFs!nnf
   
      txtInformacoes.Text = "A Carta de Corre��o e disciplinada pelo paragrafo 1�-A do art. 7� do Conv�nio S/N, de 15 de dezembro de 1970 e pode ser utilizada para regulariza��o de erro ocorrido na emiss�o de documento fiscal, desde que o erro nao esteja relacionado com: " & vbCrLf & _
                            "I - as vari�veis que determinam o valor do imposto tais como: base de c�lculo, al�quota, diferen�a de pre�o, quantidade, valor da opera��o ou da presta��o; " & vbCrLf & _
                            "II - a corre��o de dados cadastrais que implique mudan�aa do remetente ou do destinat�rio; " & vbCrLf & _
                            "III - a data de emiss�o ou de saida."
                            
   ElseIf sEvento = "EMDFE" Then
      txtNota.Enabled = False
      
      txtChaveNFe.Text = FFormataChaveNF(rsMDFes!cMDF)
      txtNota.Text = rsMDFes!numero
   
   End If
   
End Sub

Private Sub cmdConfirmar_Click()
   
   If sEvento = "IN" Then
      Call InutilizarNumeracao
   ElseIf sEvento = "CN" Then
      Call CancelarNota
   ElseIf sEvento = "CC" Then
      Call CartaCorrecao
   ElseIf sEvento = "EMDFE" Then
      Call EncerramentoMDFe
   End If
   
   Unload Me
   
End Sub

Private Sub cmdCancelar_Click()
   Unload Me
End Sub

Private Sub InutilizarNumeracao()
Dim oNFCe400 As New CNF400
Dim strMensagem As String

   If MsgBox("Confirma inutilizar a Nota N� " & txtNota.Text, vbYesNo + vbQuestion, "Confirma��o") = vbYes Then
      strMensagem = oNFCe400.FInutilizarNumero(txtNota.Text, RemoveAcentos(txtJustificativa.Text))
      MsgBox strMensagem, vbExclamation + vbOKOnly, "Inutiliza��o"
   End If

End Sub

Private Sub CancelarNota()
Dim oNFCe400 As New CNF400
Dim strMensagem As String

   If MsgBox("Confirma cancelamento da Nota N� " & txtNota.Text, vbYesNo + vbQuestion, "Confirma��o") = vbYes Then
      If Trim(txtJustificativa.Text) <> "" And Len(txtJustificativa.Text) >= 15 Then
         strMensagem = oNFCe400.FCancelarNota(txtNota.Text, RemoveAcentos(txtJustificativa.Text))
'''''         MsgBox strMensagem, vbExclamation + vbOKOnly, "Cancelamento"
      Else
         MsgBox "A justificativa � obrigat�ria e precisa ter ao menos 15 caracteres", vbExclamation + vbOKOnly, "Campos Obrigat�rios"
      End If
   End If

End Sub

Private Sub CartaCorrecao()
Dim oNFCe400 As New CNF400
Dim strMensagem As String

   If MsgBox("Confirma corre��o da Nota N� " & txtNota.Text, vbYesNo + vbQuestion, "Confirma��o") = vbYes Then
      If Trim(txtJustificativa.Text) <> "" And Len(txtJustificativa.Text) >= 15 Then
         strMensagem = oNFCe400.FCartaCorrecao(txtNota.Text, RemoveAcentos(txtJustificativa.Text))
''         MsgBox strMensagem, vbExclamation + vbOKOnly, "Carta de corre��o"
      Else
         MsgBox "A justificativa � obrigat�ria", vbExclamation + vbOKOnly, "Campos Obrigat�rios"
      End If
   End If

End Sub

Private Sub EncerramentoMDFe()
Dim oMDFe300 As New CMDFe300
Dim strMensagem As String

   If MsgBox("Confirma encerramento do Manifesto N� " & txtNota.Text, vbYesNo + vbQuestion, "Confirma��o") = vbYes Then
      strMensagem = oMDFe300.FEncerramentoMDFe(txtNota.Text)
   End If

End Sub


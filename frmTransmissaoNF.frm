VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmTransmissao 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transmitir NF"
   ClientHeight    =   8055
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5625
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8055
   ScaleWidth      =   5625
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtNota 
      Enabled         =   0   'False
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   4980
      Width           =   855
   End
   Begin VB.TextBox txtChaveNFe 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1020
      TabIndex        =   1
      Top             =   4980
      Width           =   4455
   End
   Begin VB.Timer tmrVerificar 
      Interval        =   10000
      Left            =   5100
      Top             =   7560
   End
   Begin MSComctlLib.ListView lvwMensagens 
      Height          =   1290
      Left            =   120
      TabIndex        =   0
      Top             =   5400
      Width           =   5370
      _ExtentX        =   9472
      _ExtentY        =   2275
      View            =   3
      LabelEdit       =   1
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
   Begin VB.Label lblNota 
      Caption         =   "N�mero"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   4740
      Width           =   795
   End
   Begin VB.Label lblChave 
      Caption         =   "Chave de Acesso"
      Height          =   195
      Left            =   1020
      TabIndex        =   3
      Top             =   4740
      Width           =   1395
   End
   Begin VB.Image imgLogotipo 
      BorderStyle     =   1  'Fixed Single
      Height          =   4545
      Left            =   120
      Stretch         =   -1  'True
      Top             =   120
      Width           =   5370
   End
End
Attribute VB_Name = "frmTransmissao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''Option Explicit
''Dim ItemList As ListItem
''Dim ProcuraItem As ListItem
''
''Private Sub Form_Load()
'''   imgLogotipo.Picture = LoadPicture(App.Path & "\nfce.jpg")
''   imgLogotipo.Picture = LoadPicture(App.Path & "\nfe.jpg")
''
''End Sub
''
''Private Sub tmrVerificar_Timer()
''
''   'Gerar o QrCode
''   ''' Verificar o arquivo validado existe e se a NF � NFC-e
''   Set rsGerarXML = cnSistema.Execute("Select TOP 1 * From NFCe WHERE Situacao = 1") ' Em processamento
''   If Not rsGerarXML.EOF Then
''      If Not IsNull(rsGerarXML!ChaveNFCe) Then
''         sChaveNFe = rsGerarXML!ChaveNFCe
''         sVerArquivo = Dir(I_UnidadeNFe & "NFC-e\" & sEmpresaNFCe & "\Validar\Validado\" & sChaveNFe & "-nfe.XML")
''         If sVerArquivo = (sChaveNFe & "-nfe.XML") Then
''            GerarQrCode (rsGerarXML!Numero)
''         Else
''            If rsGerarXML!TentativaEmissao < 15 Then
''               cnSistema.Execute "Update NFCe set " & _
''                        "TentativaEmissao = " & (rsGerarXML!TentativaEmissao + 1) & " " & _
''                        "Where idNFCe = " & rsGerarXML!idNFCe
''            Else
''               cnSistema.Execute "Update NFCe set " & _
''                        "Situacao = 4 " & _
''                        "Where idNFCe = " & rsGerarXML!idNFCe
''            End If
''         End If
''      End If
''   End If
''
''   ' Transmitir
''   ''' Verificar o arquivo validado existe para realizar a transmiss�o
''
''   Call FArquivosNF("TRANSMITIR")
''
''
''End Sub

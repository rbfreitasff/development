VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmCopiarArquivos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Copiar Arquivos"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   11415
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdFechar 
      Caption         =   "Fechar"
      Height          =   375
      Left            =   10140
      TabIndex        =   14
      Top             =   1800
      Width           =   1155
   End
   Begin VB.Frame fraCopiarUmArquivo 
      Caption         =   "Copiar apenas um XML"
      Height          =   1575
      Left            =   4140
      TabIndex        =   7
      Top             =   120
      Width           =   7155
      Begin VB.DriveListBox Drive2 
         Height          =   315
         Left            =   1260
         TabIndex        =   12
         Top             =   660
         Width           =   5775
      End
      Begin VB.CommandButton cmdCopiarXML 
         Caption         =   "Copiar"
         Height          =   315
         Left            =   5880
         TabIndex        =   11
         Top             =   1080
         Width           =   1155
      End
      Begin VB.CommandButton cmdSelecionar 
         Caption         =   "Selecionar"
         Height          =   315
         Left            =   4560
         TabIndex        =   9
         Top             =   1080
         Width           =   1155
      End
      Begin VB.TextBox txtArquivoXML 
         Height          =   315
         Left            =   1260
         TabIndex        =   8
         Top             =   300
         Width           =   5775
      End
      Begin VB.Label Label1 
         Caption         =   "Copiar para"
         Height          =   195
         Left            =   180
         TabIndex        =   13
         Top             =   720
         Width           =   915
      End
      Begin VB.Label lblArquivoXML 
         Caption         =   "Arquivo XML"
         Height          =   195
         Left            =   180
         TabIndex        =   10
         Top             =   360
         Width           =   915
      End
   End
   Begin VB.Frame fraCopiarMesAno 
      Caption         =   "Copiar XML's por mês/ano"
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3915
      Begin VB.CommandButton cmdCopiarMesAno 
         Caption         =   "Copiar"
         Height          =   315
         Left            =   2580
         TabIndex        =   6
         Top             =   1080
         Width           =   1155
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   1140
         TabIndex        =   3
         Top             =   660
         Width           =   2595
      End
      Begin VB.TextBox txtAno 
         Height          =   315
         Left            =   3000
         TabIndex        =   2
         Top             =   300
         Width           =   735
      End
      Begin VB.ComboBox cboMes 
         Height          =   315
         ItemData        =   "frmCopiarArquivos.frx":0000
         Left            =   1140
         List            =   "frmCopiarArquivos.frx":0028
         TabIndex        =   1
         Text            =   "Combo1"
         Top             =   300
         Width           =   1815
      End
      Begin VB.Label lblPeriodo 
         Caption         =   "Período"
         Height          =   195
         Left            =   180
         TabIndex        =   5
         Top             =   360
         Width           =   795
      End
      Begin VB.Label lblUnidade 
         Caption         =   "Copiar para"
         Height          =   195
         Left            =   180
         TabIndex        =   4
         Top             =   720
         Width           =   915
      End
   End
   Begin MSComDlg.CommonDialog dlgSelecionarArquivoXML 
      Left            =   120
      Top             =   1740
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmCopiarArquivos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
   cboMes.ListIndex = 0
   txtAno.text = Format(Date, "yyyy")
End Sub

Private Sub cmdCopiarMesAno_Click()
''   dlgCopiarNFs.ShowOpen
   
'   Text1.text = dlgCopiarNFs.FileName
'   Text2.text = dlgCopiarNFs.FileTitle
'   Text2.text = Mid(dlgCopiarNFs.FileName, 1, Len(dlgCopiarNFs.FileName) - Len(dlgCopiarNFs.FileTitle))

End Sub

Private Sub cmdSelecionar_Click()

   dlgSelecionarArquivoXML.ShowOpen
   txtArquivoXML.text = dlgSelecionarArquivoXML.FileTitle
   
End Sub

Private Sub cmdFechar_Click()
   Unload Me
End Sub



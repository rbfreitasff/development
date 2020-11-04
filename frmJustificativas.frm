VERSION 5.00
Begin VB.Form frmJustificativas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Justificativas"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8595
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   8595
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   7380
      TabIndex        =   5
      Top             =   4320
      Width           =   1155
   End
   Begin VB.CommandButton cmdConfirmar 
      Caption         =   "Confirmar"
      Height          =   375
      Left            =   6120
      TabIndex        =   4
      Top             =   4320
      Width           =   1155
   End
   Begin VB.Frame fraJustificativas 
      Caption         =   "Justificativa"
      Height          =   1995
      Left            =   120
      TabIndex        =   2
      Top             =   2220
      Width           =   8415
      Begin VB.TextBox txtJustificativa 
         Height          =   1575
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   300
         Width           =   8175
      End
   End
   Begin VB.Frame fraInformacoes 
      Caption         =   "Informações"
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8415
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2400
         TabIndex        =   8
         Top             =   300
         Width           =   5895
      End
      Begin VB.TextBox txtNota 
         Enabled         =   0   'False
         Height          =   315
         Left            =   540
         TabIndex        =   7
         Top             =   300
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
         TabIndex        =   10
         Top             =   1665
         Width           =   8175
      End
      Begin VB.Label lblChave 
         Caption         =   "Chave"
         Height          =   195
         Left            =   1860
         TabIndex        =   9
         Top             =   360
         Width           =   555
      End
      Begin VB.Label lblNota 
         Caption         =   "Nota"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   435
      End
      Begin VB.Label lblInformacoes 
         BorderStyle     =   1  'Fixed Single
         Height          =   975
         Left            =   120
         TabIndex        =   1
         Top             =   660
         Width           =   8175
      End
   End
End
Attribute VB_Name = "frmJustificativas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

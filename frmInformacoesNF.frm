VERSION 5.00
Begin VB.Form frmInformacoesNF 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informações da NF"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7095
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   7095
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraInformacoesNF 
      Caption         =   "Informações NF"
      Height          =   2775
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   6975
      Begin VB.Frame fraProtocolo 
         Caption         =   "Protocolos"
         Height          =   1155
         Left            =   120
         TabIndex        =   5
         Top             =   900
         Width           =   6735
         Begin VB.TextBox Text4 
            Height          =   285
            Left            =   1500
            TabIndex        =   11
            Top             =   780
            Width           =   3075
         End
         Begin VB.TextBox Text3 
            Height          =   285
            Left            =   1500
            TabIndex        =   9
            Top             =   480
            Width           =   3075
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   1500
            TabIndex        =   7
            Top             =   180
            Width           =   3075
         End
         Begin VB.Label lblProtocoloInutilizacao 
            Caption         =   "Inutilização"
            Height          =   195
            Left            =   180
            TabIndex        =   10
            Top             =   840
            Width           =   1155
         End
         Begin VB.Label lblProtocoloCancelamento 
            Caption         =   "Cancelamento"
            Height          =   195
            Left            =   180
            TabIndex        =   8
            Top             =   540
            Width           =   1155
         End
         Begin VB.Label lblProtocoloAutorizacao 
            Caption         =   "Autorização"
            Height          =   195
            Left            =   180
            TabIndex        =   6
            Top             =   270
            Width           =   1335
         End
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1620
         TabIndex        =   4
         Top             =   555
         Width           =   5235
      End
      Begin VB.TextBox txtNumero 
         Height          =   285
         Left            =   1620
         TabIndex        =   2
         Top             =   255
         Width           =   1575
      End
      Begin VB.Label lblChaveAcesso 
         Caption         =   "Chave de Acesso"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label lblNumero 
         Caption         =   "Número"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   300
         Width           =   795
      End
   End
End
Attribute VB_Name = "frmInformacoesNF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

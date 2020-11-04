VERSION 5.00
Begin VB.Form frmConsultaCadastro 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consultar Cadastro na SEFAZ"
   ClientHeight    =   2565
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5325
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   5325
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdConfirmar 
      Caption         =   "Confirmar"
      Height          =   375
      Left            =   2820
      TabIndex        =   1
      Top             =   2100
      Width           =   1155
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   4080
      TabIndex        =   0
      Top             =   2100
      Width           =   1155
   End
End
Attribute VB_Name = "frmConsultaCadastro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSelecionarData 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Escolha uma data"
   ClientHeight    =   2445
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   2520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   2520
   StartUpPosition =   3  'Windows Default
   Begin MSComCtl2.MonthView mtwCalendario 
      Height          =   2370
      Left            =   0
      TabIndex        =   0
      Top             =   60
      Width           =   2490
      _ExtentX        =   4392
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   154337281
      CurrentDate     =   43897
   End
End
Attribute VB_Name = "frmSelecionarData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mtwCalendario_DateClick(ByVal DateClicked As Date)
   Me.Hide
End Sub

VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmVisualiza 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Visualiza Dados"
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6495
   Icon            =   "frmVisualiza.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin MSComctlLib.ListView lvwDados 
      Height          =   3075
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   5424
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483624
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "frmVisualiza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub lvwDados_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   With lvwDados
      If (ColumnHeader.Index - 1) = .SortKey Then
         .SortOrder = (.SortOrder + 1) Mod 2
      Else
         .Sorted = False
         .SortOrder = 0
         .SortKey = ColumnHeader.Index - 1
         .Sorted = True
      End If
   End With
End Sub

Private Sub lvwDados_DblClick()
   Registro_Selecionado = True
   frmVisualiza.Hide
End Sub

Private Sub lvwDados_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Registro_Selecionado = True
      frmVisualiza.Hide
   End If
End Sub

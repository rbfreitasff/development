VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmGerenciarManifestos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gerenciador de Manifestos MDF-e"
   ClientHeight    =   6750
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6750
   ScaleWidth      =   9360
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraManifestos 
      Caption         =   "Manifestos"
      Height          =   6555
      Left            =   105
      TabIndex        =   0
      Top             =   120
      Width           =   9165
      Begin VB.ComboBox cmbSituacao 
         Height          =   315
         ItemData        =   "frmGerenciarManifestos.frx":0000
         Left            =   4770
         List            =   "frmGerenciarManifestos.frx":001C
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   255
         Width           =   2805
      End
      Begin VB.CommandButton cmdTransmitir 
         Caption         =   "&Transmitir"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   6060
         Width           =   1155
      End
      Begin VB.CommandButton cmdConsultaNota 
         Caption         =   "Consulta Situa��o "
         Height          =   375
         Left            =   6420
         TabIndex        =   10
         Top             =   6060
         Width           =   1875
      End
      Begin VB.CommandButton cmdInutilizarNumeracao 
         Caption         =   "&Inutilizar Numera��o"
         Height          =   375
         Left            =   3000
         TabIndex        =   9
         Top             =   6060
         Width           =   1695
      End
      Begin VB.CommandButton cmdImprimirDANFE 
         Caption         =   "Imprimir DANFE"
         Enabled         =   0   'False
         Height          =   375
         Left            =   4740
         TabIndex        =   8
         Top             =   6060
         Width           =   1635
      End
      Begin VB.CommandButton cmdCancelarNota 
         Caption         =   "&Cancelar Nota"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1380
         TabIndex        =   7
         Top             =   6060
         Width           =   1575
      End
      Begin VB.Frame fraInformacoesNotas 
         Height          =   675
         Left            =   75
         TabIndex        =   2
         Top             =   5340
         Width           =   9000
         Begin VB.Label lblNotasPendentes 
            Caption         =   "Pendentes"
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
            Height          =   195
            Left            =   7320
            TabIndex        =   6
            Top             =   300
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.Label lblNotasCanceladas 
            Caption         =   "Canceladas"
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
            Height          =   195
            Left            =   5040
            TabIndex        =   5
            Top             =   300
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.Label lblNotasAprovadas 
            Caption         =   "Aprovadas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   195
            Left            =   120
            TabIndex        =   4
            Top             =   300
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.Label lblNotasIntuilizadas 
            Caption         =   "Inutilizadas"
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
            Height          =   195
            Left            =   2400
            TabIndex        =   3
            Top             =   300
            Visible         =   0   'False
            Width           =   1815
         End
      End
      Begin VB.CommandButton cmdPesquisar 
         Caption         =   "&Pesquisar"
         Height          =   315
         Left            =   7620
         TabIndex        =   1
         Top             =   240
         Width           =   1155
      End
      Begin MSComctlLib.ListView lvwManifestos 
         Height          =   4800
         Left            =   60
         TabIndex        =   13
         Top             =   600
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   8467
         View            =   3
         LabelEdit       =   1
         SortOrder       =   -1  'True
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
      Begin MSMask.MaskEdBox mskDataInicial 
         Height          =   285
         Left            =   990
         TabIndex        =   14
         Top             =   270
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
         Left            =   2910
         TabIndex        =   15
         Top             =   270
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "99/99/9999"
         PromptChar      =   " "
      End
      Begin VB.Label lblSituacao 
         AutoSize        =   -1  'True
         Caption         =   "Situa��o"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3990
         TabIndex        =   18
         Top             =   315
         Width           =   630
      End
      Begin VB.Label lblDataInicial 
         AutoSize        =   -1  'True
         Caption         =   "Data Inicial"
         Height          =   195
         Left            =   90
         TabIndex        =   17
         Top             =   315
         Width           =   795
      End
      Begin VB.Label lblDataFinal 
         AutoSize        =   -1  'True
         Caption         =   "Data Final"
         Height          =   195
         Left            =   2040
         TabIndex        =   16
         Top             =   315
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmGerenciarManifestos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

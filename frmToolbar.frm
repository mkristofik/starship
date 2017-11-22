VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmToolbar 
   BackColor       =   &H80000009&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Controls"
   ClientHeight    =   1200
   ClientLeft      =   816
   ClientTop       =   1560
   ClientWidth     =   3504
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1200
   ScaleWidth      =   3504
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog dlgOpen 
      Left            =   2760
      Top             =   480
      _ExtentX        =   677
      _ExtentY        =   677
      _Version        =   393216
      CancelError     =   -1  'True
      Filter          =   "Starship Files (*.swc) | *.swc"
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "&Remove"
      Height          =   372
      Left            =   2400
      TabIndex        =   2
      Top             =   720
      Width           =   972
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   372
      Left            =   2400
      TabIndex        =   1
      Top             =   120
      Width           =   972
   End
   Begin VB.ListBox lstShips 
      Height          =   1008
      ItemData        =   "frmToolbar.frx":0000
      Left            =   120
      List            =   "frmToolbar.frx":0002
      TabIndex        =   0
      Top             =   120
      Width           =   2172
   End
End
Attribute VB_Name = "frmToolbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

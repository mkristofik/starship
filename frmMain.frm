VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   Caption         =   "Star Wars Starship Construction System v1.0"
   ClientHeight    =   7185
   ClientLeft      =   3135
   ClientTop       =   2880
   ClientWidth     =   7320
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7185
   ScaleWidth      =   7320
   Begin MSComDlg.CommonDialog dlgFile 
      Left            =   6720
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "*.swc"
      Filter          =   "Starship Files (*.swc) | *.swc"
   End
   Begin VB.ComboBox cboType 
      Height          =   315
      ItemData        =   "frmMain.frx":0000
      Left            =   1440
      List            =   "frmMain.frx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   46
      Top             =   840
      Width           =   2655
   End
   Begin VB.TextBox txtInfo 
      Appearance      =   0  'Flat
      Height          =   1212
      Left            =   4200
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   45
      Text            =   "frmMain.frx":00BE
      Top             =   360
      Width           =   3012
   End
   Begin VB.TextBox txtHeading 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   4200
      Locked          =   -1  'True
      TabIndex        =   44
      Text            =   "Information Box:"
      Top             =   120
      Width           =   3252
   End
   Begin VB.CheckBox chkAtmosphere 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Atmosphere Capable"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2760
      Width           =   2175
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "&Calculate"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   10
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox txtCargo 
      Height          =   285
      Left            =   1440
      MaxLength       =   6
      TabIndex        =   4
      Top             =   1560
      Width           =   615
   End
   Begin VB.ComboBox cboHDrive 
      Height          =   288
      ItemData        =   "frmMain.frx":0107
      Left            =   1440
      List            =   "frmMain.frx":0114
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   3840
      Width           =   1095
   End
   Begin VB.TextBox txtHull 
      Height          =   285
      Left            =   1440
      MaxLength       =   5
      TabIndex        =   5
      Top             =   2400
      Width           =   615
   End
   Begin VB.TextBox txtShields 
      Height          =   285
      Left            =   1440
      MaxLength       =   5
      TabIndex        =   7
      Top             =   3120
      Width           =   615
   End
   Begin VB.TextBox txtSpeed 
      Height          =   285
      Left            =   1440
      MaxLength       =   1
      TabIndex        =   8
      Top             =   3480
      Width           =   375
   End
   Begin VB.TextBox txtMan 
      Height          =   285
      Left            =   1440
      MaxLength       =   35
      TabIndex        =   2
      Top             =   480
      Width           =   2655
   End
   Begin VB.TextBox txtLength 
      Height          =   285
      Left            =   1440
      MaxLength       =   4
      TabIndex        =   3
      Text            =   "25"
      Top             =   1200
      Width           =   615
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   1440
      MaxLength       =   35
      TabIndex        =   1
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label lblNewSpeed 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   5520
      TabIndex        =   59
      Top             =   6720
      Width           =   1212
   End
   Begin VB.Label lblPPF 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   5520
      TabIndex        =   58
      Top             =   6480
      Width           =   1212
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Adjusted Speed Rating:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   3360
      TabIndex        =   57
      Top             =   6720
      Width           =   2052
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Payload Penalty Factor:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   3360
      TabIndex        =   56
      Top             =   6480
      Width           =   2052
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Fighter Bays:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   120
      TabIndex        =   55
      Top             =   5280
      Width           =   1212
   End
   Begin VB.Label lblPower 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   192
      Index           =   8
      Left            =   4680
      TabIndex        =   54
      Top             =   5280
      Width           =   348
   End
   Begin VB.Label lblMass 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   192
      Index           =   8
      Left            =   2760
      TabIndex        =   53
      Top             =   5280
      Width           =   348
   End
   Begin VB.Label lblPower 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   192
      Index           =   7
      Left            =   4680
      TabIndex        =   52
      Top             =   4920
      Width           =   348
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Weapons/Equipment:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   120
      TabIndex        =   51
      Top             =   4920
      Width           =   1932
   End
   Begin VB.Label lblMass 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   192
      Index           =   7
      Left            =   2760
      TabIndex        =   50
      Top             =   4920
      Width           =   348
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Crew/Life Support:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   120
      TabIndex        =   49
      Top             =   4560
      Width           =   1692
   End
   Begin VB.Label lblMass 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   192
      Index           =   6
      Left            =   2760
      TabIndex        =   48
      Top             =   4560
      Width           =   348
   End
   Begin VB.Label lblPower 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   192
      Index           =   6
      Left            =   4680
      TabIndex        =   47
      Top             =   4560
      Width           =   348
   End
   Begin VB.Label Label15 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Totals:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   120
      TabIndex        =   43
      Top             =   6000
      Width           =   1212
   End
   Begin VB.Label lblMass 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   192
      Index           =   9
      Left            =   2760
      TabIndex        =   42
      Top             =   5640
      Width           =   348
   End
   Begin VB.Label lblPower 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   192
      Index           =   9
      Left            =   4680
      TabIndex        =   41
      Top             =   5640
      Width           =   348
   End
   Begin VB.Label lblPowerTotal 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   192
      Left            =   4680
      TabIndex        =   40
      Top             =   6000
      Width           =   348
   End
   Begin VB.Label lblMassTotal 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   192
      Left            =   2760
      TabIndex        =   39
      Top             =   6000
      Width           =   348
   End
   Begin VB.Label lblPower 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "4,420.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   192
      Index           =   5
      Left            =   4680
      TabIndex        =   38
      Top             =   4200
      Width           =   684
   End
   Begin VB.Label lblPower 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   192
      Index           =   4
      Left            =   4680
      TabIndex        =   37
      Top             =   3840
      Width           =   348
   End
   Begin VB.Label lblPower 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   192
      Index           =   3
      Left            =   4680
      TabIndex        =   36
      Top             =   3480
      Width           =   348
   End
   Begin VB.Label lblPower 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   192
      Index           =   2
      Left            =   4680
      TabIndex        =   35
      Top             =   3120
      Width           =   348
   End
   Begin VB.Label lblPower 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   192
      Index           =   1
      Left            =   4680
      TabIndex        =   34
      Top             =   2760
      Width           =   348
   End
   Begin VB.Label lblMass 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "55.50"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   192
      Index           =   5
      Left            =   2760
      TabIndex        =   33
      Top             =   4200
      Width           =   444
   End
   Begin VB.Label lblMass 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   192
      Index           =   4
      Left            =   2760
      TabIndex        =   32
      Top             =   3840
      Width           =   348
   End
   Begin VB.Label lblMass 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   192
      Index           =   3
      Left            =   2760
      TabIndex        =   31
      Top             =   3480
      Width           =   348
   End
   Begin VB.Label lblMass 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   192
      Index           =   2
      Left            =   2760
      TabIndex        =   30
      Top             =   3120
      Width           =   348
   End
   Begin VB.Label lblMass 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   192
      Index           =   1
      Left            =   2760
      TabIndex        =   29
      Top             =   2760
      Width           =   348
   End
   Begin VB.Label lblCargo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Cargo (tons):"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label18 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Mass Limit Empty:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   120
      TabIndex        =   27
      Top             =   6480
      Width           =   1572
   End
   Begin VB.Label Label17 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Total Mass Limit:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   6720
      Width           =   1575
   End
   Begin VB.Label lblML 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   1800
      TabIndex        =   25
      Top             =   6480
      Width           =   1212
   End
   Begin VB.Label lblTML 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   1800
      TabIndex        =   24
      Top             =   6720
      Width           =   1212
   End
   Begin VB.Label lblType 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Type:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Craft Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Manufacturer:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label lblPower 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   192
      Index           =   0
      Left            =   4680
      TabIndex        =   20
      Top             =   2400
      Width           =   348
   End
   Begin VB.Label lblMass 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   192
      Index           =   0
      Left            =   2760
      TabIndex        =   19
      Top             =   2400
      Width           =   348
   End
   Begin VB.Label Label13 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Reactor:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   120
      TabIndex        =   18
      Top             =   5640
      Width           =   1212
   End
   Begin VB.Label lblHdrive 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Hyperdrive:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   120
      TabIndex        =   17
      Top             =   3840
      Width           =   1212
   End
   Begin VB.Label Label11 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Sensors:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   120
      TabIndex        =   16
      Top             =   4200
      Width           =   1212
   End
   Begin VB.Label lblHull 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Hull:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   120
      TabIndex        =   15
      Top             =   2400
      Width           =   1212
   End
   Begin VB.Label lblShields 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Shields:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   120
      TabIndex        =   14
      Top             =   3120
      Width           =   1212
   End
   Begin VB.Label lblSpeed 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Speed:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   120
      TabIndex        =   13
      Top             =   3480
      Width           =   1212
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Mass"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   2760
      TabIndex        =   12
      Top             =   2040
      Width           =   492
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Power"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   4680
      TabIndex        =   11
      Top             =   2040
      Width           =   612
   End
   Begin VB.Label lblLength 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Length (m):"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New..."
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open..."
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save..."
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print..."
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "&Exit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditTech 
         Caption         =   "&Techbase"
         Begin VB.Menu mnuTech 
            Caption         =   "New &Republic"
            Checked         =   -1  'True
            Index           =   1
         End
         Begin VB.Menu mnuTech 
            Caption         =   "&Imperial"
            Index           =   2
         End
         Begin VB.Menu mnuTech 
            Caption         =   "&Herald"
            Index           =   3
         End
         Begin VB.Menu mnuTech 
            Caption         =   "&Ploxus"
            Index           =   4
         End
      End
      Begin VB.Menu mnuEditEquip 
         Caption         =   "E&quipment..."
      End
      Begin VB.Menu mnuEditFighter 
         Caption         =   "&Fighter Bays..."
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWeaponEd 
         Caption         =   "&Weapon Editor..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim div As Long, Length As Integer, Cargo As Long, ml As Double, tml As Double, curFile As String

' Weapon property variables.
Dim wTotMass As Double, wTotPower As Double, fMass As Double

Property Let FighterMass(ByVal m As Double)
    lblMass(8).Caption = FormatNumber(m, 2)
    fMass = m
    cmdCalc_Click
End Property

Property Get FighterMass() As Double
    FighterMass = fMass
End Property

Property Let MassTotal(ByVal m As Double)
    lblMass(7).Caption = FormatNumber(m, 2)
    wTotMass = m
    cmdCalc_Click
End Property

Property Get MassTotal() As Double
    MassTotal = wTotMass
End Property

Property Let PowerTotal(ByVal p As Double)
    lblPower(7).Caption = FormatNumber(p, 2)
    wTotPower = p
    cmdCalc_Click
End Property

Property Get PowerTotal() As Double
    PowerTotal = wTotPower
End Property


Private Sub cboHDrive_Click()
    Ship.HDrive = cboHDrive.ListIndex
    lblHdrive.ForeColor = vbBlue
    mIsDirty = True
End Sub

Private Sub cboHDrive_GotFocus()
    txtHeading = "Hyperdrive:"
    txtInfo = "Select the hyperdrive (if any) for the starship.  The improved hyperdrive "
    txtInfo = txtInfo + "is twice as fast as the standard model."
End Sub

Private Sub cboType_Click()
    
    Dim i As Integer, k As Integer, m As Integer
    i = cboType.ListIndex
    
    If Ship.ShipType = 5 And i <> 5 Then
        If Me.FighterMass Then
            m = MsgBox("Only Military Combat Starships can have fighter bays." & vbCrLf & _
                "This action will remove them.  Continue?", vbYesNo)
            If m = vbNo Then
                cboType.ListIndex = 5
                Exit Sub
            End If
            Me.FighterMass = 0
            k = 0
            Do While Craft(k).Num
                Craft(k).Num = 0
                k = k + 1
            Loop
        End If
    End If
    
    If i = 5 Then
        mnuEditFighter.Enabled = True
    Else
        mnuEditFighter.Enabled = False
    End If
    
    If DoValidation = False Then
        cmdCalc_Click
        CalcCrew (i)
        cmdCalc_Click
    End If
    mIsDirty = True
    
End Sub

Private Sub CalcCrew(ByVal i As Integer)

    Select Case i
        Case Is <= 1
            lblMass(6).Caption = FormatNumber((0.02 + i * 0.01) * ml, 2)
        Case 2
            lblMass(6).Caption = FormatNumber(0.06 * ml, 2)
        Case 3
            lblMass(6).Caption = FormatNumber(0.05 * ml, 2)
        Case 4
            lblMass(6).Caption = FormatNumber(0.07 * ml, 2)
        Case 5
            lblMass(6).Caption = FormatNumber(0.09 * ml, 2)
    End Select
    Ship.ShipType = i

End Sub

Private Sub cboType_GotFocus()
    txtHeading = "Craft Type:"
    txtInfo = "Determines the general role the starship will have.  Required "
    txtInfo = txtInfo + "crew size varies by role.  Also, Military Combat Starships "
    txtInfo = txtInfo + "have reactors twice the size of that of other craft and can "
    txtInfo = txtInfo + "have fighter bays."
End Sub

Private Sub chkAtmosphere_Click()
    Ship.Atmos = chkAtmosphere.Value
    chkAtmosphere.ForeColor = vbBlue
    mIsDirty = True
End Sub

Private Sub chkAtmosphere_GotFocus()
    txtHeading = "Atmosphere Capability:"
    txtInfo = "Includes all the necessary components (repulsorlifts, landing "
    txtInfo = txtInfo + "gear, etc.) for the starship to operate in an atmosphere."
End Sub

Private Sub cmdCalc_Click()
    
    If DoValidation Then Exit Sub
    UpdateLimits
    CalcCrew (cboType.ListIndex)
    UpdateTotals
    
    lblLength.ForeColor = vbBlack
    lblCargo.ForeColor = vbBlack
    
    lblPPF.Caption = FormatNumber(CDbl(lblML.Caption) / CDbl(lblMassTotal.Caption), 4)
    lblNewSpeed.Caption = Int(Val(lblPPF.Caption) * Val(txtSpeed.Text))

End Sub


Private Sub UpdateLimits()

    div = Ship.Length * 10
    ml = Ship.Length ^ 2 + Ship.Cargo / 10
    tml = ml * 1.2 + Ship.Cargo

    lblML.Caption = FormatNumber(ml, 2)
    lblTML.Caption = FormatNumber(tml, 2)

End Sub

Private Function DoValidation() As Boolean

    If Ship.Length < 25 Or Ship.Length > 3000 Then
        MsgBox "Length of craft must be between 25m and 3,000m."
        txtLength.SelStart = 0
        txtLength.SelLength = Len(txtLength.Text)
        txtLength.SetFocus
        DoValidation = True
    ElseIf Ship.Hull < 0 Then
        MsgBox "Illegal value for hull rating."
        txtHull.SelStart = 0
        txtHull.SelLength = Len(txtHull.Text)
        txtHull.SetFocus
        DoValidation = True
    ElseIf Ship.Shields < 0 Then
        MsgBox "Illegal value for shield rating."
        txtShields.SelStart = 0
        txtShields.SelLength = Len(txtShields.Text)
        txtShields.SetFocus
        DoValidation = True
    End If

End Function

Private Sub Form_Load()
    
    Dim FileName As String
    FileName = GetFileName("weapons.db")
    Open FileName For Random As #1 Len = Len(Weapon)
    Call Initialize
    
    dlgFile.InitDir = GetFileName("ships")
    mIsDirty = False
    
End Sub

Private Sub Initialize()

    Dim i As Integer
    
    cboHDrive.ListIndex = 0
    txtName.Text = ""
    txtMan.Text = ""
    txtLength.Text = "25"
    txtCargo.Text = ""
    txtHull.Text = ""
    chkAtmosphere.Value = 0
    txtShields.Text = ""
    txtSpeed.Text = ""
    
    mnuTech(1).Checked = True
    For i = 2 To 4
        mnuTech(i).Checked = False
    Next i
    
    With Ship
        .Length = 25
        .Techbase = 1
        .Atmos = 0
        .Cargo = 0
        .HDrive = 0
        .Hull = 0
        .Manufacture = ""
        .Shields = 0
        .ShipName = ""
        .ShipType = 0
        .Speed = 0
        .totMass = 0
    End With
    
    For i = 0 To 25
        Equip(i).Qty = 0
        If i <= 10 Then Craft(i).Num = 0
    Next i
    
    cboType.ListIndex = 0
    Me.FighterMass = 0
    Me.MassTotal = 0
    Me.PowerTotal = 0
    curFile = ""
    
End Sub

Private Sub mnuEditEquip_Click()
    cmdCalc_Click
    frmEquipment.Show
End Sub

Private Sub mnuEditFighter_Click()
    If DoValidation = False Then
        cmdCalc_Click
        frmFighters.Show
    End If
End Sub

Private Sub mnuFileExit_Click()
    If CheckSave Then Exit Sub
    Close
    Unload Me
End Sub

Private Sub mnuFileNew_Click()
    If CheckSave Then Exit Sub
    Call Initialize
End Sub

Private Sub mnuFileOpen_Click()

    Dim i As Integer
    
    If CheckSave Then Exit Sub
    On Error Resume Next
    
    dlgFile.Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly
    dlgFile.ShowOpen
    
    If Err.number <> cdlCancel Then
        Open dlgFile.FileName For Random Access Read As #3 Len = Len(SaveData)
        Get #3, 1, SaveData
        Close #3
        With SaveData
            Ship = .myShip
            
            For i = 0 To 25
                If i <= 10 Then Craft(i) = .myCraft(i)
                Equip(i) = .myEquip(i)
            Next i
        End With
        
        Call RefillData
        mIsDirty = False
    End If
    
    curFile = dlgFile.FileName

End Sub

Private Sub RefillData()

    Dim i As Integer, total As Double, Power As Double
    
    With Ship
        ' General info
        cboType.ListIndex = .ShipType
        txtName.Text = RTrim(.ShipName)
        txtMan.Text = RTrim(.Manufacture)
        txtLength.Text = .Length
        txtCargo.Text = .Cargo
        txtHull.Text = .Hull
        chkAtmosphere.Value = .Atmos
        txtShields.Text = .Shields
        txtSpeed.Text = .Speed
        cboHDrive.ListIndex = .HDrive
    
        
        ' Update the techbase
        For i = 1 To 4
            If i <> .Techbase Then mnuTech(i).Checked = False
        Next i
        mnuTech(.Techbase).Checked = True
    End With
        
    ' Equipment
    i = 0
    total = 0
    Power = 0
    Do While i <= 25 And Equip(i).Qty
        total = total + Equip(i).Mass
        Power = Power + Equip(i).Power
        i = i + 1
    Loop
        
    Me.MassTotal = total
    Me.PowerTotal = Power
        
    ' Fighter Bays
    i = 0
    total = 0
    Do While i <= 10 And Craft(i).Num
        total = total + Craft(i).Mass
        i = i + 1
    Loop
        
    Me.FighterMass = total
    
End Sub

Private Sub mnuFilePrint_Click()

    If DoValidation Then Exit Sub
    If RTrim(txtName.Text) = "" Then
        MsgBox "Craft has no name.", vbExclamation
        Exit Sub
    End If
    
    If RTrim(txtMan.Text) = "" Then
        MsgBox "Craft was not made by anyone.", vbExclamation
        Exit Sub
    End If
    frmPrint.Show

End Sub

Private Sub mnuFileSave_Click()

    Dim i As Integer
    On Error Resume Next
    
    dlgFile.Flags = cdlOFNHideReadOnly Or cdlOFNOverwritePrompt
    
    If curFile = "" Then
        dlgFile.FileName = RTrim(Ship.ShipName)
    Else
        dlgFile.FileName = curFile
    End If
    
    dlgFile.ShowSave
    
    If Err.number <> cdlCancel Then
        With SaveData
            For i = 0 To 25
                'If Equip(i).Qty = 0 Then Exit For
                .myEquip(i) = Equip(i)
            Next i
            
            For i = 0 To 10
                'If Craft(i).Num = 0 Then Exit For
                .myCraft(i) = Craft(i)
            Next i
            
            .myShip = Ship
        End With
            
        Open dlgFile.FileName For Random Access Write As #3 Len = Len(SaveData)
        Put #3, 1, SaveData
        Close #3
        mIsDirty = False
        curFile = dlgFile.FileName
    End If

End Sub

Private Sub mnuTech_Click(Index As Integer)
   
    Dim m As Integer, i As Integer
    
    If DoValidation Then Exit Sub
    If Index <> Ship.Techbase Then
        If lblMass(7).Caption <> "0.00" Then
            m = MsgBox("Warning! Changing technology base will remove all on-board equipment." _
                & vbCrLf & "Do you wish to continue?", vbQuestion + vbYesNo)
            If m = vbYes Then
                Load frmEquipment
                For i = 0 To frmEquipment.lstWeapons.ListCount - 1
                    Equip(i).Qty = 0
                Next i
                Unload frmEquipment
                Me.MassTotal = 0
                Me.PowerTotal = 0
            Else
                Exit Sub
            End If
        End If
        Ship.Techbase = Index
        
        For i = 1 To 4
            If i <> Index Then mnuTech(i).Checked = False
        Next i
        mnuTech(Index).Checked = True
        mIsDirty = True
    End If
    
    cmdCalc_Click

End Sub

Private Sub mnuWeaponEd_Click()
    frmEdit.Show
End Sub

Private Sub txtCargo_Change()
    Ship.Cargo = Val(txtCargo.Text)
    lblCargo.ForeColor = vbBlue
    mIsDirty = True
End Sub

Private Sub txtCargo_GotFocus()
    txtHeading = "Cargo Tonnage:"
    txtInfo = "Determines the maximum cargo weight carried by the "
    txtInfo = txtInfo + "starship."
End Sub

Private Sub txthull_Change()
    Ship.Hull = Val(txtHull.Text)
    lblHull.ForeColor = vbBlue
    mIsDirty = True
End Sub

Private Sub txtHull_GotFocus()
    txtHeading = "Hull Rating:"
    txtInfo = "The sky's the limit on this one, but keep in mind that the "
    txtInfo = txtInfo + "hull armor is one of the heaviest pieces of equipment "
    txtInfo = txtInfo + "onboard."
End Sub

Private Sub txtLength_Change()
    Ship.Length = Val(txtLength.Text)
    lblLength.ForeColor = vbBlue
    mIsDirty = True
End Sub

Private Sub txtLength_GotFocus()
    txtHeading = "Craft Length:"
    txtInfo = "Minimum length: 25m" + vbCrLf
    txtInfo = txtInfo + "Maximum length: 3,000m"
End Sub

Private Sub txtMan_Change()
    Ship.Manufacture = txtMan.Text
    mIsDirty = True
End Sub

Private Sub txtName_Change()
    Ship.ShipName = txtName.Text
    mIsDirty = True
End Sub

Private Sub txtShields_Change()
    Ship.Shields = CDbl(Val(txtShields.Text))
    lblShields.ForeColor = vbBlue
    mIsDirty = True
End Sub

Private Sub txtShields_GotFocus()
    txtHeading = "Shield Rating:"
    txtInfo = "The size of the shield generator is unlimited, but it is "
    txtInfo = txtInfo + "the most power-hungry system a starship has."
End Sub

Private Sub txtSpeed_Change()
    Ship.Speed = Val(txtSpeed.Text)
    lblSpeed.ForeColor = vbBlue
    mIsDirty = True
End Sub

Private Sub UpdateTotals()

    Dim Power As Double, m As Double
    
    ' Hull rating
    lblMass(0).Caption = FormatNumber((div / 100000 + Ship.Hull / (div * 1.25)) * ml, 2)
    lblPower(0).Caption = FormatNumber(lblMass(0).Caption * 200, 2)
    lblHull.ForeColor = vbBlack
    
    ' Atmoshpere
    lblMass(1).Caption = FormatNumber(0.01 * Ship.Atmos * ml, 2)
    lblPower(1).Caption = FormatNumber(5 * Ship.Atmos * ml, 2)
    chkAtmosphere.ForeColor = vbBlack
    
    ' Shields
    lblMass(2).Caption = FormatNumber(Ship.Shields / (div * 1.25) * ml, 2)
    Power = (CDbl(Ship.Shields) * 100 / (div * 1.25)) ^ 2 / 3
    Power = Power * ml
    lblPower(2).Caption = FormatNumber(Power, 2)
    lblShields.ForeColor = vbBlack
    
    ' Speed
    lblMass(3).Caption = FormatNumber(Ship.Speed / 40 * ml, 2)
    lblPower(3).Caption = FormatNumber((Ship.Speed + 10) * ml, 2)
    lblSpeed.ForeColor = vbBlack
    
    ' Hyperdrive
    If Ship.HDrive Then
        lblMass(4).Caption = FormatNumber((0.04 + 0.03 * Ship.HDrive) * ml, 2)
        lblPower(4).Caption = FormatNumber((0.05 * Ship.HDrive) * ml, 2)
    Else
        lblMass(4).Caption = "0.00"
        lblPower(4).Caption = "0.00"
    End If
    lblHdrive.ForeColor = vbBlack
    
    ' Standard sensor info
    lblMass(5).Caption = "55.50"
    lblPower(5).Caption = "4,420.00"
        
    ' Power and totals
    Dim i As Integer, tot As Double
    For i = 0 To 7
        If i <= 5 Then lblPower(i).Caption = FormatNumber(CDbl(lblPower(i).Caption) / _
            (Int(Ship.Techbase / -2) * -1), 2)
        tot = tot + CDbl(lblPower(i).Caption)
    Next i
    
    ' Lower the mass for Herald and Ploxus
    For i = 0 To 5
        m = CDbl(lblMass(i).Caption)
        If Ship.Techbase = 3 Then lblMass(i).Caption = FormatNumber(m / 3, 2)
        If Ship.Techbase = 4 Then lblMass(i).Caption = FormatNumber(m / 2, 2)
    Next i
    
    lblPowerTotal.Caption = FormatNumber(tot, 2)
    
    If Ship.ShipType < 5 Then tot = tot / 2
    If Ship.Techbase = 3 Then tot = tot / 2
        
    lblMass(9).Caption = FormatNumber(tot / 1000, 2)
    tot = 0
    
    For i = 0 To 9
        tot = tot + CDbl(lblMass(i).Caption)
    Next i
    Ship.totMass = tot + Ship.Cargo
    
    If Ship.totMass > CDbl(lblTML.Caption) Then
        lblMassTotal.ForeColor = vbRed
    Else
        lblMassTotal.ForeColor = vbBlack
    End If
    lblMassTotal.Caption = FormatNumber(Ship.totMass, 2)
        
End Sub

Private Sub txtSpeed_GotFocus()
    txtHeading = "Speed:"
    txtInfo = "This value differs from the actual starship speed.  To find "
    txtInfo = txtInfo + "the actual speed consult the following list:" + vbCrLf
    txtInfo = txtInfo + "1 = 1/8,   2 = 1/4,    3 = 1/2,    4 = 1," + vbCrLf
    txtInfo = txtInfo + "5 = 2,     6 = 3, etc."
End Sub

Private Function CheckSave()

    Dim m As Integer
    If mIsDirty Then
        m = MsgBox("Loaded file is not saved.  Save it now?", vbYesNo)
        If m = vbYes Then
            mnuFileSave_Click
            CheckSave = True
        End If
    End If
    CheckSave = False

End Function

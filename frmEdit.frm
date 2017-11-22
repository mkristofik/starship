VERSION 5.00
Begin VB.Form frmEdit 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Weapon Editor"
   ClientHeight    =   5460
   ClientLeft      =   1185
   ClientTop       =   1830
   ClientWidth     =   4500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   4500
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
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
      Left            =   3120
      TabIndex        =   9
      Top             =   720
      Width           =   1215
   End
   Begin VB.Frame fraInfo 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   3375
      Left            =   120
      TabIndex        =   11
      Top             =   1920
      Width           =   3135
      Begin VB.TextBox txtPower 
         Height          =   285
         Left            =   1080
         MaxLength       =   6
         TabIndex        =   7
         Top             =   2880
         Width           =   1935
      End
      Begin VB.TextBox txtWeapName 
         Height          =   285
         Left            =   1080
         MaxLength       =   25
         TabIndex        =   2
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox txtDamage 
         Height          =   285
         Left            =   1080
         MaxLength       =   6
         TabIndex        =   3
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox txtMass 
         Height          =   285
         Left            =   1080
         MaxLength       =   6
         TabIndex        =   4
         Top             =   1680
         Width           =   1935
      End
      Begin VB.TextBox txtRange 
         Height          =   285
         Left            =   1080
         MaxLength       =   15
         TabIndex        =   5
         Top             =   2040
         Width           =   1935
      End
      Begin VB.TextBox txtTohit 
         Height          =   285
         Left            =   1080
         MaxLength       =   6
         TabIndex        =   6
         Top             =   2400
         Width           =   1935
      End
      Begin VB.ComboBox cboTechBase 
         Height          =   315
         ItemData        =   "frmEdit.frx":0000
         Left            =   1320
         List            =   "frmEdit.frx":0013
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Power Multiplier:"
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
         Height          =   495
         Left            =   120
         TabIndex        =   21
         Top             =   2760
         Width           =   855
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Name:"
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
         TabIndex        =   20
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "To-hit:"
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
         TabIndex        =   19
         Top             =   2400
         Width           =   615
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Range:"
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
         TabIndex        =   18
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Mass:"
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
         TabIndex        =   17
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Damage:"
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
         TabIndex        =   16
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Tech-Base:"
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
         TabIndex        =   15
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblInstr 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Add New:"
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
         TabIndex        =   14
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblNum 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "1"
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
         Left            =   2640
         TabIndex        =   13
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Record #"
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
         Left            =   1800
         TabIndex        =   12
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add New"
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
      Left            =   3120
      TabIndex        =   8
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save Changes"
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
      Left            =   3120
      TabIndex        =   10
      Top             =   1320
      Width           =   1215
   End
   Begin VB.ListBox lstEquipment 
      Height          =   1620
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   2892
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuClose 
         Caption         =   "&Close"
      End
   End
   Begin VB.Menu mnuFilter 
      Caption         =   "&Filter"
      Begin VB.Menu mnuFilterAll 
         Caption         =   "&All Weapons"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuFilterType 
         Caption         =   "&Common"
         Index           =   0
      End
      Begin VB.Menu mnuFilterType 
         Caption         =   "&New Republic"
         Index           =   1
      End
      Begin VB.Menu mnuFilterType 
         Caption         =   "&Imperial"
         Index           =   2
      End
      Begin VB.Menu mnuFilterType 
         Caption         =   "&Herald"
         Index           =   3
      End
      Begin VB.Menu mnuFilterType 
         Caption         =   "&Ploxus"
         Index           =   4
      End
   End
End
Attribute VB_Name = "frmEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim numRecs As Integer, curRec As Integer, isDirty As Boolean, flag As Boolean
Dim Filters(4) As Boolean

Private Sub cboTechBase_Click()
    isDirty = True
End Sub

Private Sub cmdAdd_Click()

    lblInstr.Caption = "Add New:"
    cboTechBase.ListIndex = 0
    txtWeapName.Text = ""
    txtDamage.Text = ""
    txtMass.Text = ""
    txtPower.Text = ""
    txtRange.Text = ""
    txtTohit.Text = ""
    curRec = numRecs + 1
    lblNum.Caption = curRec
    isDirty = True
    cmdEdit.Enabled = False
    cmdSave.Enabled = True
    fraInfo.Enabled = True
    
End Sub

Private Sub cmdEdit_Click()

    cmdEdit.Enabled = False
    lblInstr.Caption = "Modify Existing:"
    fraInfo.Enabled = True
    cmdSave.Enabled = True
    
End Sub

Private Sub cmdSave_Click()

    Weapon.Techbase = cboTechBase.ListIndex
    Weapon.WeapName = txtWeapName.Text
    Weapon.Damage = txtDamage.Text
    Weapon.Mass = Val(txtMass.Text)
    Weapon.Range = txtRange.Text
    Weapon.Tohit = txtTohit.Text
    Weapon.Power = Val(txtPower.Text)
    fraInfo.Enabled = False
    Put #1, curRec, Weapon
    MsgBox "Data saved."
    ReloadList
    isDirty = False

End Sub

Private Sub Form_Load()

    Dim i As Integer
    
    cboTechBase.ListIndex = 0
    numRecs = GetNumRecords(1)
    lblNum.Caption = numRecs + 1
    curRec = 1
    isDirty = False
    
    ' Initialize the list filters to all weapons.
    For i = 0 To 4
        Filters(i) = True
    Next i
    
    ReloadList
    
End Sub

Private Sub ReloadList()

    Dim i As Integer
    lstEquipment.Clear
    numRecs = GetNumRecords(1)
    
    For i = 1 To numRecs
        Get #1, i, Weapon
        If Filters(Weapon.Techbase) Then
            lstEquipment.AddItem (Weapon.WeapName)
            lstEquipment.ItemData(lstEquipment.NewIndex) = i
        End If
    Next i
    cmdEdit.Enabled = False
    cmdSave.Enabled = False
    
End Sub


Private Sub lstEquipment_Click()
    
    ' Error trap for when user clicks cancel in code block below.
    If flag Then
        flag = False
        Exit Sub
    End If
            
    ' Check to change weapons.
    If isDirty Then
        Dim m As Integer
        m = MsgBox("This action will reset any changes made to current weapon." _
            + vbCrLf + "Do you wish to continue?", vbYesNo)
        If m = 7 Then
            flag = True
            lstEquipment.ListIndex = curRec
            Exit Sub
        End If
    End If
    
    curRec = lstEquipment.ItemData(lstEquipment.ListIndex)
    Get #1, curRec, Weapon
    lblNum.Caption = curRec
    
    txtWeapName.Text = RTrim(Weapon.WeapName)
    txtMass.Text = Weapon.Mass
    txtDamage.Text = RTrim(Weapon.Damage)
    txtRange.Text = RTrim(Weapon.Range)
    txtTohit.Text = RTrim(Weapon.Tohit)
    txtPower.Text = Weapon.Power
    cboTechBase.ListIndex = Weapon.Techbase
    
    lblInstr.Caption = "View Existing:"
    cmdSave.Enabled = False
    fraInfo.Enabled = False
    cmdEdit.Enabled = True
    isDirty = False
    
End Sub

Private Sub mnuClose_Click()
    Unload Me
End Sub

Private Sub mnuFilterAll_Click()
    
    Dim i As Integer
    For i = 0 To 4
        mnuFilterType(i).Checked = False
        Filters(i) = True
    Next i
    
    mnuFilterAll.Checked = True
    ReloadList
    
End Sub

Private Sub mnuFilterType_Click(Index As Integer)

    Dim i As Integer
    For i = 0 To 4
        If i = Index Then
            If mnuFilterType(i).Checked Then
                mnuFilterType(i).Checked = False
                Filters(i) = False
            Else
                mnuFilterType(i).Checked = True
                Filters(i) = True
            End If
        Else
            If mnuFilterAll.Checked Then
                mnuFilterType(i).Checked = False
                Filters(i) = False
            End If
        End If
    Next i

    mnuFilterAll.Checked = False
    ReloadList
    
End Sub

Private Sub txtDamage_Change()
    isDirty = True
End Sub

Private Sub txtMass_Change()
    isDirty = True
End Sub

Private Sub txtPower_Change()
    isDirty = True
End Sub

Private Sub txtRange_Change()
    isDirty = True
End Sub

Private Sub txtTohit_Change()
    isDirty = True
End Sub

Private Sub txtWeapName_Change()
    isDirty = True
End Sub

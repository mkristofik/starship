VERSION 5.00
Begin VB.Form frmEngine 
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Engine Editor"
   ClientHeight    =   4620
   ClientLeft      =   1740
   ClientTop       =   2670
   ClientWidth     =   6510
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   6510
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   4
      Top             =   1560
      Width           =   1215
   End
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
      Height          =   375
      Left            =   5160
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.Frame fraInfo 
      BackColor       =   &H80000009&
      Height          =   2535
      Left            =   120
      TabIndex        =   12
      Top             =   1920
      Width           =   6255
      Begin VB.TextBox txtSpeedMult 
         Height          =   285
         Left            =   5160
         MaxLength       =   4
         TabIndex        =   10
         Top             =   2040
         Width           =   735
      End
      Begin VB.ComboBox cboEngType 
         Height          =   315
         ItemData        =   "frmEngine.frx":0000
         Left            =   1080
         List            =   "frmEngine.frx":0010
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox txtCrits 
         Height          =   285
         Left            =   1080
         MaxLength       =   2
         TabIndex        =   6
         Top             =   960
         Width           =   1935
      End
      Begin VB.ListBox lstTech 
         Height          =   1185
         ItemData        =   "frmEngine.frx":004D
         Left            =   4200
         List            =   "frmEngine.frx":0060
         Style           =   1  'Checkbox
         TabIndex        =   11
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox txtEngName 
         Height          =   285
         Left            =   1080
         MaxLength       =   25
         TabIndex        =   5
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox txtManBase 
         Height          =   285
         Left            =   2040
         MaxLength       =   1
         TabIndex        =   8
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox txtSpeedMod 
         Height          =   285
         Left            =   2040
         MaxLength       =   4
         TabIndex        =   9
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "%"
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
         Left            =   5880
         TabIndex        =   23
         Top             =   2040
         Width           =   255
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Real Speed Multiplier:"
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
         Left            =   3240
         TabIndex        =   22
         Top             =   2040
         Width           =   1935
      End
      Begin VB.Label Label2 
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
         TabIndex        =   21
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Criticals:"
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
         Width           =   735
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
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
         Left            =   3120
         TabIndex        =   19
         Top             =   600
         Width           =   975
         WordWrap        =   -1  'True
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
         TabIndex        =   18
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Speed Modifier:"
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
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Base Maneuverablity:"
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
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label lblInstr 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Click to Add/Edit"
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
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lblNum 
         Alignment       =   1  'Right Justify
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
         Left            =   5880
         TabIndex        =   14
         Top             =   240
         Width           =   255
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
         Left            =   5040
         TabIndex        =   13
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
      Height          =   375
      Left            =   5160
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save Data"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   3
      Top             =   1080
      Width           =   1215
   End
   Begin VB.ListBox lstEquipment 
      Height          =   1815
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   4935
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuClose 
         Caption         =   "&Close"
      End
   End
End
Attribute VB_Name = "frmEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim numRecs As Integer, curRec As Integer, IsDirty As Boolean, flag As Boolean

Private Sub cboEngType_Change()
    IsDirty = True
End Sub

Private Sub cmdAdd_Click()

    Dim i As Integer
    
    lblInstr.Caption = "Add New:"
    txtEngName.Text = ""
    txtCrits.Text = ""
    txtManBase.Text = ""
    txtSpeedMod.Text = ""
    txtSpeedMult.Text = ""
    curRec = numRecs + 1
    lblNum.Caption = curRec
    IsDirty = True
    cmdEdit.Enabled = False
    cmdDelete.Enabled = False
    cmdSave.Enabled = True
    fraInfo.Enabled = True
    Engine.Deleted = False
    lstTech.Selected(0) = True
    txtEngName.SetFocus
    cboEngType.ListIndex = 0
    
End Sub

Private Sub cmdDelete_Click()

    Dim m As Integer
    m = MsgBox("Warning!  Item will be permantently deleted.  Continue?", vbYesNo Or _
        vbExclamation)
    If m = vbYes Then
        Engine.Deleted = True
        Put #2, curRec, Engine
        ReloadList
    End If
        
    cmdDelete.Enabled = False

End Sub

Private Sub cmdEdit_Click()

    cmdEdit.Enabled = False
    cmdDelete.Enabled = False
    lblInstr.Caption = "Modify Existing:"
    fraInfo.Enabled = True
    cmdSave.Enabled = True
    
End Sub

Private Sub cmdSave_Click()

    Dim i As Integer, tech As String
        
    If BadData Then Exit Sub
    Engine.EngName = txtEngName.Text
    Engine.ManBase = Val(txtManBase.Text)
    Engine.SpeedMod = Val(txtSpeedMod.Text)
    Engine.SpeedMult = Val(txtSpeedMult.Text)
    Engine.Criticals = Val(txtCrits.Text)
    Engine.EngType = cboEngType.ListIndex
    
    For i = 0 To 4
        If lstTech.Selected(i) Then tech = tech & CStr(i)
    Next i
    
    Engine.TechBase = Val(tech)
    
    fraInfo.Enabled = False
    Put #2, curRec, Engine
    MsgBox "Data saved."
    IsDirty = False
    ReloadList

End Sub

Function BadData() As Boolean
  
    Dim ret As Boolean
    
    If txtEngName.Text = "" Then
        MsgBox "Must enter a name.", vbCritical
        ret = True
    End If
    If Val(txtCrits.Text) < 1 Then
        MsgBox "All equiment needs at least one critical slot.", vbCritical
        ret = True
    End If
    
    BadData = ret

End Function

Private Sub Form_Load()

    numRecs = GetNumRecords()
    curRec = 1
    ReloadList
    IsDirty = False
    Call SetTabs(lstEquipment)
    cboEngType.ListIndex = 0
    
End Sub

Private Sub ReloadList()

    Dim i As Integer
    lstEquipment.Clear
    numRecs = GetNumRecords()
    
    For i = 1 To numRecs
        Get #2, i, Engine
        If Not Engine.Deleted Then
            lstEquipment.AddItem (Engine.EngName) & Chr$(9) & TechString(Engine.TechBase)
            lstEquipment.ItemData(lstEquipment.NewIndex) = i
        End If
    Next i
    cmdEdit.Enabled = False
    cmdSave.Enabled = False
    
    If lstEquipment.ListCount Then lstEquipment.ListIndex = 0
    
End Sub

Private Sub lstEquipment_Click()
    
    Dim i As Integer
    
    ' Error trap for when user clicks No in code block below.
    If flag Then
        flag = False
        Exit Sub
    End If
    
    ' Check to change Engines.
    If IsDirty Then
        Dim m As Integer
        m = MsgBox("This action will reset any changes made to current engine." _
            + vbCrLf + "Do you wish to continue?", vbYesNo)
        If m = 7 Then
            flag = True
            lstEquipment.ListIndex = curRec
            Exit Sub
        End If
    End If
    
    curRec = lstEquipment.ItemData(lstEquipment.ListIndex)
    Get #2, curRec, Engine
    lblNum.Caption = curRec
    
    txtEngName.Text = RTrim(Engine.EngName)
    txtManBase.Text = Engine.ManBase
    txtSpeedMod.Text = RTrim(Engine.SpeedMod)
    txtSpeedMult.Text = RTrim(Engine.SpeedMult)
    cboEngType.ListIndex = Engine.EngType
    txtCrits.Text = Engine.Criticals
    
    ' Fill in the checkboxes
    For i = 0 To 4
        If InStr(CStr(Engine.TechBase), i) Then
            lstTech.Selected(i) = True
        Else
            lstTech.Selected(i) = False
        End If
    Next i
    
    lblInstr.Caption = "View Existing:"
    cmdDelete.Enabled = True
    cmdSave.Enabled = False
    fraInfo.Enabled = False
    cmdEdit.Enabled = True
    IsDirty = False
    
End Sub

Private Sub mnuClose_Click()
    Unload Me
End Sub

Private Sub txtEngName_Change()
    IsDirty = True
End Sub

Private Sub lstTech_Click()
    
    Dim i As Integer
    IsDirty = True
    If lstTech.Selected(0) Then
        For i = 1 To 4
            lstTech.Selected(i) = False
        Next i
    End If

End Sub

Private Sub txtManBase_Change()
    IsDirty = True
End Sub

Private Sub txtSpeedMod_Change()
    IsDirty = True
End Sub

Private Sub txtSpeedMult_Change()
    IsDirty = True
End Sub

Private Function GetNumRecords()
    GetNumRecords = LOF(2) / Len(Engine)
End Function

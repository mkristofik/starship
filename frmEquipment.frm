VERSION 5.00
Begin VB.Form frmEquipment 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manage Weapons/Equipment"
   ClientHeight    =   4725
   ClientLeft      =   2985
   ClientTop       =   2775
   ClientWidth     =   7425
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   7425
   Begin VB.ListBox lstWeapons 
      Height          =   1620
      ItemData        =   "frmEquipment.frx":0000
      Left            =   120
      List            =   "frmEquipment.frx":0002
      TabIndex        =   15
      Top             =   2880
      Width           =   4692
   End
   Begin VB.TextBox txtNum 
      Height          =   288
      Left            =   1800
      MaxLength       =   3
      TabIndex        =   14
      Top             =   1800
      Width           =   495
   End
   Begin VB.Frame fraLoc 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Location"
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
      Height          =   1332
      Left            =   4560
      TabIndex        =   8
      Top             =   120
      Width           =   1812
      Begin VB.OptionButton optLoc 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Back"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   3
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   855
      End
      Begin VB.OptionButton optLoc 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Forward"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton optLoc 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Left Side"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Width           =   1092
      End
      Begin VB.OptionButton optLoc 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Right Side"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   2
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   1212
      End
   End
   Begin VB.Frame fraMount 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Weapon Mounting"
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
      Height          =   1332
      Left            =   2640
      TabIndex        =   4
      Top             =   120
      Width           =   1812
      Begin VB.CheckBox chkTurret 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Turret"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   120
         TabIndex        =   16
         Top             =   960
         Width           =   852
      End
      Begin VB.OptionButton optMount 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Quad"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   4
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   855
      End
      Begin VB.OptionButton optMount 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Double"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   2
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   972
      End
      Begin VB.OptionButton optMount 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Single"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Value           =   -1  'True
         Width           =   852
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   6480
      TabIndex        =   3
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "&Remove"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   120
      TabIndex        =   2
      Top             =   2400
      Width           =   972
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Enabled         =   0   'False
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
      Left            =   120
      TabIndex        =   1
      Top             =   1680
      Width           =   612
   End
   Begin VB.ListBox lstEquipment 
      Height          =   1425
      ItemData        =   "frmEquipment.frx":0004
      Left            =   120
      List            =   "frmEquipment.frx":0006
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label16 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Current Totals:"
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
      Left            =   4920
      TabIndex        =   35
      Top             =   3720
      Width           =   1332
   End
   Begin VB.Label lblPowerTotal 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
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
      Height          =   195
      Left            =   5640
      TabIndex        =   34
      Top             =   4200
      Width           =   390
   End
   Begin VB.Label Label13 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Power:"
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
      Left            =   4920
      TabIndex        =   33
      Top             =   4200
      Width           =   612
   End
   Begin VB.Label Label12 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
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
      Height          =   252
      Left            =   4920
      TabIndex        =   32
      Top             =   3960
      Width           =   612
   End
   Begin VB.Label lblMassTotal 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
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
      Height          =   195
      Left            =   5640
      TabIndex        =   31
      Top             =   3960
      Width           =   390
   End
   Begin VB.Label Label15 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Current Power Demand:"
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
      Left            =   2880
      TabIndex        =   30
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label lblCurPower 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
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
      Height          =   195
      Left            =   5040
      TabIndex        =   29
      Top             =   1920
      Width           =   390
   End
   Begin VB.Label lblCurMass 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
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
      Height          =   195
      Left            =   5040
      TabIndex        =   28
      Top             =   1680
      Width           =   390
   End
   Begin VB.Label Label14 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Current Weapon Mass:"
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
      Left            =   2880
      TabIndex        =   27
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Label lblSlotTotal 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
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
      Left            =   6720
      TabIndex        =   26
      Top             =   2880
      Width           =   372
   End
   Begin VB.Label lblSlot 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
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
      Index           =   3
      Left            =   5280
      TabIndex        =   25
      Top             =   3360
      Width           =   372
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "B:"
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
      Left            =   4920
      TabIndex        =   24
      Top             =   3360
      Width           =   252
   End
   Begin VB.Label lblSlot 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
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
      Index           =   1
      Left            =   6240
      TabIndex        =   23
      Top             =   3120
      Width           =   372
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "L:"
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
      Left            =   5880
      TabIndex        =   22
      Top             =   3120
      Width           =   252
   End
   Begin VB.Label lblSlot 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
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
      Index           =   2
      Left            =   6240
      TabIndex        =   21
      Top             =   3360
      Width           =   372
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "R:"
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
      Left            =   5880
      TabIndex        =   20
      Top             =   3360
      Width           =   252
   End
   Begin VB.Label lblSlot 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
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
      Index           =   0
      Left            =   5280
      TabIndex        =   19
      Top             =   3120
      Width           =   372
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "F:"
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
      Left            =   4920
      TabIndex        =   18
      Top             =   3120
      Width           =   252
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Weapon Slots Left:"
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
      Left            =   4920
      TabIndex        =   17
      Top             =   2880
      Width           =   1692
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Quantity:"
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
      Left            =   840
      TabIndex        =   13
      Top             =   1800
      Width           =   852
   End
End
Attribute VB_Name = "frmEquipment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Form specific variables
Dim numRecs As Integer, curRec As Integer
Dim curPower As Double, curMass As Double, curMount As Integer

' Declarations required to place tab stops in the weapons list box.
Const LB_SETTABSTOPS = &H192
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
    (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Sub chkTurret_Click()
    UpdateCurrent
End Sub

Private Sub cmdAdd_Click()

    Dim Num As Integer, newName As String, wMass As String, loc As String, intLoc As Integer
    
    Num = Val(txtNum.Text)
    If Num <= 0 Then     ' Trap for adding nothing.
        MsgBox "Must add at least one weapon."
        txtNum.Text = ""
        txtNum.SetFocus
        Exit Sub
    End If
    
    wMass = lblCurMass.Caption
    Call FormatLoc(loc, intLoc)
    newName = FormatName
    
    If CheckExceed(Num, intLoc) Then Exit Sub
    lstWeapons.AddItem Num & Chr$(9) & newName & Chr$(9) & wMass & Chr$(9) & loc
    
    ' Trick to scroll the weapons list to the new item
    lstWeapons.Selected(lstWeapons.NewIndex) = True
    lstWeapons.Selected(lstWeapons.NewIndex) = False
    cmdAdd.Enabled = True
    cmdRemove.Enabled = False
    
    With Equip(lstWeapons.NewIndex)
        .Location = intLoc
        .Mass = CDbl(wMass)
        lblMassTotal.Caption = FormatNumber(CDbl(lblMassTotal.Caption) + .Mass, 2)
        .Power = CDbl(lblCurPower.Caption)
        lblPowerTotal.Caption = FormatNumber(CDbl(lblPowerTotal.Caption) + .Power, 2)
        .Qty = Num
        .WeapName = newName
        .MountType = curMount
        .Turret = (chkTurret.Value = 1)
    End With
    
    mIsDirty = True

End Sub

Private Function FormatLoc(strLoc As String, realLoc As Integer) As String
' Format the location string.
    Dim loc As String
    
    If optLoc(0).Value Then
        strLoc = "F"
        realLoc = 0
    ElseIf optLoc(1).Value Then
        strLoc = "L"
        realLoc = 1
    ElseIf optLoc(2).Value Then
        strLoc = "R"
        realLoc = 2
    Else
        strLoc = "B"
        realLoc = 3
    End If
    
    If chkTurret.Value Then strLoc = strLoc + " (turret)"
        
End Function

Private Function FormatName() As String
' Format the name string.
    Dim newName As String

    If curMount = 2 Then
        newName = "Double "
    ElseIf curMount = 4 Then
        newName = "Quad "
    End If
    FormatName = newName + RTrim(Weapon.WeapName)

End Function

Private Function CheckExceed(ByVal wQty As Integer, ByVal loc As Integer) As Boolean

    wQty = wQty * curMount
    
    If wQty > Val(lblSlot(loc).Caption) Then
        MsgBox "Number of weapon slots exceeded."
        CheckExceed = True
    Else
        ' Update the weapon slot totals.
        lblSlot(loc).Caption = Val(lblSlot(loc).Caption) - wQty
        CheckExceed = False
    End If

End Function

Private Sub SetTabs(lst As ListBox)
' Calls Windows DLL file to set tab stops in a list box
    ReDim tabs(0 To 2) As Long      ' # of tab stops
    Dim returnVal As Long           ' DLL function returns a long
    tabs(0) = 21                    ' Twips measurement for the tab stops
    tabs(1) = 121                   ' Lots of fine tuning necessary!
    tabs(2) = 164
    returnVal = SendMessage(lst.hwnd, LB_SETTABSTOPS, 3, tabs(0)) ' Call the DLL

End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdRemove_Click()
     
    Dim Index As Integer, numSlots As Integer, wMass As Double, wPower As Double
    Dim c As Integer
    Index = lstWeapons.ListIndex
    If Index = -1 Then Exit Sub ' Error trap if no weapon selected.
    
    With Equip(Index)
        ' Update the form information
        numSlots = .Qty * .MountType
        lblSlot(.Location).Caption = Val(lblSlot(.Location).Caption) + numSlots
        lblMassTotal.Caption = FormatNumber(CDbl(lblMassTotal.Caption) - .Mass, 2)
        lblPowerTotal.Caption = FormatNumber(CDbl(lblPowerTotal.Caption) - .Power, 2)
        
        ' Reset the array to match the listbox
        For c = Index To lstWeapons.ListCount - 1
            Equip(c) = Equip(c + 1)
        Next c
        Equip(c + 1).Qty = 0
    End With
    
    ' Remove it from the list
    lstWeapons.RemoveItem (Index)
    mIsDirty = True
    
End Sub

Private Sub Form_Load()

    Dim total As Integer, i As Integer
        
    numRecs = GetNumRecords(1)
    curRec = 1
    LoadList
    SetTabs lstWeapons ' Set up the tab stops for the weapons list
    
    ' Calculate the number of weapon slots in each location.
    total = Sqr(Ship.Length) * (5 + Int(Ship.Length / 400))
    
    For i = 0 To 3
        lblSlot(i).Caption = Int(total / 4)
    Next i
    
    For i = 1 To (total Mod 4)
        lblSlot(i - 1).Caption = Val(lblSlot(i - 1).Caption) + 1
    Next i
    
    ' Reload the items into the list.
    lblMassTotal.Caption = FormatNumber(frmMain.MassTotal, 2)
    lblPowerTotal.Caption = FormatNumber(frmMain.PowerTotal, 2)
    i = 0
    Do While Equip(i).Qty
        ReloadItem (i)
        i = i + 1
    Loop
    
    ' Error trap (turn off main form)
    frmMain.Enabled = False
        
End Sub

Private Sub LoadList()

    Dim i As Integer
    For i = 1 To numRecs
        Get #1, i, Weapon
        If Weapon.Techbase = 0 Or Weapon.Techbase = Ship.Techbase Then
            lstEquipment.AddItem (Weapon.WeapName)
            lstEquipment.ItemData(lstEquipment.NewIndex) = i
        End If
    Next i
        
End Sub

Private Sub ReloadItem(Index As Integer)

    Dim strWeap As String
    
    With Equip(Index)
        strWeap = .Qty & Chr$(9) & .WeapName & Chr$(9) & FormatNumber(.Mass, 2) & Chr$(9)
        
        ' Format the location part of the string.
        If .Location = 0 Then
            strWeap = strWeap & "F"
        ElseIf .Location = 1 Then
            strWeap = strWeap & "L"
        ElseIf .Location = 2 Then
            strWeap = strWeap & "R"
        Else
            strWeap = strWeap & "B"
        End If
        If .Turret Then strWeap = strWeap & " (turret)"
        
        ' Update the form info
        lstWeapons.AddItem strWeap
        lblSlot(.Location).Caption = lblSlot(.Location).Caption - .Qty * .MountType
    End With

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    frmMain.MassTotal = CDbl(lblMassTotal.Caption)
    frmMain.PowerTotal = CDbl(lblPowerTotal.Caption)
    frmMain.Enabled = True
    
End Sub

Private Sub lblSlot_Change(Index As Integer)

    Dim i As Integer, tot As Integer
    For i = 0 To 3
        tot = tot + Val(lblSlot(i).Caption)
    Next i
    lblSlotTotal.Caption = tot

End Sub

Private Sub lstEquipment_Click()

    ' Get the appropriate weapon record.
    curRec = lstEquipment.ItemData(lstEquipment.ListIndex)
    Get #1, curRec, Weapon
    curMass = Weapon.Mass
    curPower = Weapon.Power
    
    ' Initialize the add weapon interface.
    cmdAdd.Enabled = True
    cmdRemove.Enabled = False
    optMount(1).Value = True
    curMount = 1
    chkTurret.Value = 0
    optLoc(0).Value = True
    
    ' Calls UpdateCurrent subroutine if txtNum is changed. Otherwise txtNum_Change will
    ' call it.
    If Val(txtNum.Text) = 1 Then
        UpdateCurrent
    Else
        txtNum.Text = 1
    End If
    
End Sub

Private Sub UpdateCurrent()

    Dim totMass As Double, totPower As Double
    
    totMass = curMass * curMount
    totMass = totMass + totMass * chkTurret.Value * 0.5
    totMass = totMass * Val(txtNum.Text)
    
    ' Power reductions for double and quad mounted weapons (10% for double, 20% for quad)
    totPower = totMass * curPower
    If curMount - 1 Then totPower = totPower - totPower * 0.1 * (curMount / 2)
        
    lblCurMass.Caption = FormatNumber(totMass, 2)
    lblCurPower.Caption = FormatNumber(totPower, 2)

End Sub

Private Sub lstEquipment_DblClick()
    frmEquipInfo.Show
End Sub

Private Sub lstWeapons_Click()

    cmdRemove.Enabled = True
    cmdAdd.Enabled = False

End Sub

Private Sub optMount_Click(Index As Integer)
    curMount = Index
    UpdateCurrent
End Sub

Private Sub txtNum_Change()
    UpdateCurrent
End Sub

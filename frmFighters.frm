VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmFighters 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Manage Fighter Bays"
   ClientHeight    =   3675
   ClientLeft      =   2550
   ClientTop       =   4755
   ClientWidth     =   6420
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   6420
   Begin VB.CommandButton cmdMisc 
      Caption         =   "Add &Misc"
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
      Left            =   1320
      TabIndex        =   17
      Top             =   2760
      Width           =   972
   End
   Begin VB.ComboBox cboTechBase 
      Height          =   315
      ItemData        =   "frmFighters.frx":0000
      Left            =   1800
      List            =   "frmFighters.frx":0010
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   120
      Width           =   1695
   End
   Begin MSComDlg.CommonDialog dlgFighter 
      Left            =   5400
      Top             =   3000
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
      CancelError     =   -1  'True
      Filter          =   "Starship Files (*.sw*) | *.sw*"
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "&Browse..."
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
      TabIndex        =   11
      Top             =   2760
      Width           =   972
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
      Left            =   5400
      TabIndex        =   10
      Top             =   120
      Width           =   852
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
      Left            =   2520
      TabIndex        =   7
      Top             =   2280
      Width           =   972
   End
   Begin VB.ListBox lstUsed 
      Height          =   1035
      ItemData        =   "frmFighters.frx":003C
      Left            =   2520
      List            =   "frmFighters.frx":003E
      TabIndex        =   5
      Top             =   960
      Width           =   2775
   End
   Begin VB.TextBox txtQty 
      Height          =   288
      Left            =   1800
      TabIndex        =   4
      Top             =   2400
      Width           =   492
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
      Height          =   372
      Left            =   120
      TabIndex        =   2
      Top             =   2280
      Width           =   612
   End
   Begin VB.ListBox lstAvailable 
      Height          =   1035
      ItemData        =   "frmFighters.frx":0040
      Left            =   120
      List            =   "frmFighters.frx":0093
      TabIndex        =   1
      Top             =   960
      Width           =   2172
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Technology Base:"
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
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label lblCurMass 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   5400
      TabIndex        =   14
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Tonnage:"
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
      Left            =   5400
      TabIndex        =   13
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label lblFighter 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "(No craft currently selected)"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   3240
      Width           =   5175
   End
   Begin VB.Label lblTotalMass 
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
      Height          =   195
      Left            =   3720
      TabIndex        =   9
      Top             =   2880
      Width           =   345
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Total Mass:"
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
      Left            =   2520
      TabIndex        =   8
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Currently On-board:"
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
      Left            =   2520
      TabIndex        =   6
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      Height          =   255
      Left            =   840
      TabIndex        =   3
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Available Craft:"
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
      Top             =   600
      Width           =   1455
   End
End
Attribute VB_Name = "frmFighters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type FighterType
    CraftName As String * 25
    garbage As String * 6
    TotalSpace As Integer
    junk As String * 186
End Type

Private Type CapitalType
    CraftName As String * 35
    Mass As Double
    junk As String * 2053
End Type

Dim FighterInfo As FighterType, curName As String, curMass As Integer, flag As String
Dim GROUP As Integer, SQUAD As Integer, tech As Integer, CapitalInfo As CapitalType

Private Sub cboTechBase_Click()

    Dim i As Integer, loc As Integer, tb As String
    
    tech = cboTechBase.ListIndex + 1
    
    If tech = 1 Or tech = 3 Then
        GROUP = 4
        SQUAD = 8
        
        If tech = 1 Then
            tb = " (NR)"
        Else
            tb = " (H)"
        End If
    Else
        GROUP = 6
        SQUAD = 12
         
        If tech = 2 Then
            tb = " (I)"
        Else
            tb = " (P)"
        End If
        
        For i = 4 To 15
            If (i - 1) Mod 3 Then lstAvailable.ItemData(i) = lstAvailable.ItemData(i) * 1.5
        Next i
    End If
        
    For i = 4 To 15
        loc = InStr(1, lstAvailable.List(i), " (") - 1
        If loc = -1 Then loc = Len(lstAvailable.List(i))
        lstAvailable.List(i) = Left(lstAvailable.List(i), loc) & tb
    Next i

End Sub

Private Sub cmdAdd_Click()

    Dim tot As Double, id As Integer, tempName As String, loc As Integer
    
    id = lstAvailable.ListIndex
    tot = Val(lstAvailable.ItemData(id)) * Val(txtQty.Text)
    lstAvailable_Click
        
    If Left$(curName, 1) = "(" Then
        MsgBox "Please select a craft.", vbInformation
        Exit Sub
    ElseIf id < 3 And flag = "capital" Or id = 3 And flag = "fighter" Then
        MsgBox "Current custom craft is not of selected type.", vbInformation
        Exit Sub
    ElseIf tot = 0 Then
        MsgBox "Error adding craft.", vbCritical
        Exit Sub
    Else
        tempName = RTrim(curName)
        tempName = txtQty.Text & " " & tempName
        
        ' Preserve the techbase indicator while making the name plural.
        If Val(txtQty.Text) > 1 Then
            loc = InStr(1, tempName, " (")
            If loc Then
                tempName = Mid$(tempName, 1, loc - 1) & "s" & Right$(tempName, _
                    Len(tempName) - loc + 1)
            Else
                tempName = tempName & "s"
            End If
        End If
        
        lstUsed.AddItem tempName
        lstUsed.ItemData(lstUsed.NewIndex) = tot
    End If
        
    lblTotalMass.Caption = FormatNumber(CDbl(lblTotalMass.Caption) + tot, 2)
    mIsDirty = True
    
End Sub

Private Sub cmdBrowse_Click()
    
    On Error Resume Next
    dlgFighter.ShowOpen
    If Err.number <> cdlCancel Then
        If Right(dlgFighter.FileName, 3) = "sw2" Then
            Open dlgFighter.FileName For Random Access Read As #2 Len = Len(FighterInfo)
            Get #2, 1, FighterInfo
            curName = FighterInfo.CraftName
            
            ' Total space is stored as an index - needs to be converted to an actual space value
            FighterInfo.TotalSpace = FighterInfo.TotalSpace * 5 + 10
            
            lstAvailable.ItemData(0) = FighterInfo.TotalSpace * 5
            lstAvailable.ItemData(1) = FighterInfo.TotalSpace * GROUP * 5
            lstAvailable.ItemData(2) = FighterInfo.TotalSpace * SQUAD * 5
            flag = "fighter"
        Else
            Open dlgFighter.FileName For Random Access Read As #2 Len = Len(CapitalInfo)
            Get #2, 1, CapitalInfo
            curName = CapitalInfo.CraftName
            lstAvailable.ItemData(3) = CapitalInfo.Mass * 10
            flag = "capital"
        End If
        lblFighter.Caption = RTrim(curName)
        Close #2
        lstAvailable.ListIndex = lstAvailable.ListIndex
    End If
    
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdMisc_Click()

    Dim n As String, m As String, q As String
    n = InputBox("Enter object name: (35 characters maximum)", "Name")
    If Len(n) And Len(n) <= 35 Then
        m = InputBox("Enter object mass:", "Mass")
        If IsNumeric(m) Then
            m = CDbl(m)
            If m Then
                q = InputBox("Enter object quantity", "Quantity")
                If IsNumeric(q) Then
                    If CInt(q) > 1 Then n = n & "s"
                    lstUsed.AddItem q & " " & n
                    m = m * CInt(q)
                    lstUsed.ItemData(lstUsed.NewIndex) = m
                    
                    lblTotalMass.Caption = FormatNumber(CDbl(lblTotalMass.Caption) + m, 2)
                    mIsDirty = True
                End If
            End If
        End If
    End If

End Sub

Private Sub cmdRemove_Click()

    Dim Mass As Double, Index As Integer
    
    Index = lstUsed.ListIndex
    Mass = lstUsed.ItemData(Index)
    lstUsed.RemoveItem Index
    lblTotalMass.Caption = FormatNumber(CDbl(lblTotalMass.Caption) - Mass, 2)
    lblCurMass.Caption = ""
    cmdRemove.Enabled = False
    mIsDirty = True

End Sub

Private Sub Form_Load()
      
    dlgFighter.Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly
    dlgFighter.InitDir = GetFileName("ships")
    frmMain.Enabled = False
    cboTechBase.ListIndex = Ship.Techbase - 1
    ReloadList
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Dim i As Integer, j As Integer, k As Integer, c As Integer
    
    j = lstUsed.ListCount - 1
    For i = 0 To j
        With Craft(i)
            k = InStr(1, lstUsed.List(i), " ")
            .Num = Val(Left$(lstUsed.List(i), k - 1))
            .CraftName = Mid$(lstUsed.List(i), k + 1, Len(lstUsed.List(i)) - k + 1)
            .Mass = Val(lstUsed.ItemData(i))
        End With
    Next i
    
    For c = i To 10
        Craft(c).Num = 0
    Next c
    
    frmMain.FighterMass = CDbl(lblTotalMass.Caption)
    frmMain.Enabled = True

End Sub

Private Sub lstAvailable_Click()
    
    Dim id As Integer
    id = lstAvailable.ListIndex
    cmdAdd.Enabled = True
    
    Select Case id
        Case 0, 3
            curName = lblFighter.Caption
        Case 1
            curName = lblFighter.Caption + " Flight Group"
        Case 2
            curName = lblFighter.Caption + " Squadron"
        Case Else
            curName = lstAvailable.List(id)
    End Select
    
End Sub

Private Sub lstUsed_Click()

    cmdRemove.Enabled = True
    lblCurMass.Caption = lstUsed.ItemData(lstUsed.ListIndex)

End Sub

Private Sub ReloadList()

    Dim i As Integer
    i = 0
    
    Do While Craft(i).Num
        With Craft(i)
            lstUsed.AddItem .Num & " " & .CraftName
            lstUsed.ItemData(i) = .Mass
            lblTotalMass.Caption = FormatNumber(CDbl(lblTotalMass.Caption) + .Mass, 2)
        End With
        i = i + 1
    Loop

End Sub

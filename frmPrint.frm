VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmPrint 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   Caption         =   "Print Preview"
   ClientHeight    =   5910
   ClientLeft      =   1035
   ClientTop       =   2415
   ClientWidth     =   10395
   BeginProperty Font 
      Name            =   "Courier"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   10395
   Begin VB.PictureBox picD 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   4455
      Left            =   120
      ScaleHeight     =   4395
      ScaleWidth      =   10155
      TabIndex        =   2
      Top             =   1320
      Width           =   10215
   End
   Begin VB.ListBox lstShips 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      ItemData        =   "frmPrint.frx":0000
      Left            =   120
      List            =   "frmPrint.frx":0002
      TabIndex        =   0
      Top             =   120
      Width           =   3852
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
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
      Left            =   4080
      TabIndex        =   1
      Top             =   120
      Width           =   972
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "&Remove"
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
      Left            =   4080
      TabIndex        =   3
      Top             =   720
      Width           =   972
   End
   Begin MSComDlg.CommonDialog dlgDialog 
      Left            =   9960
      Top             =   120
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
      CancelError     =   -1  'True
      Filter          =   "Starship Files (*.swc) | *.swc"
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "&Menu"
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print..."
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Close"
         Shortcut        =   ^X
      End
   End
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type ShipInfoType
    CraftName As String
    FileName As String
    numLines As Integer
End Type

Private LoadShip As SaveDataType, ShipList(20) As ShipInfoType, curObj As Object

Private Sub cmdAdd_Click()

    Dim intFile As Integer, found As Boolean, c As Integer
    
    If lstShips.ListCount = 21 Then
        MsgBox "Maximum number of ships reached.", vbExclamation
        Exit Sub
    End If
    
    On Error GoTo HandleErr
    dlgDialog.ShowOpen
    intFile = FreeFile
    found = False
    
    Open dlgDialog.FileName For Random Access Read As #intFile Len = Len(LoadShip)
    Get #intFile, 1, LoadShip
    
    For c = 0 To lstShips.ListCount - 1
        If lstShips.List(c) = LoadShip.myShip.ShipName Then
            found = True
            Exit For
        End If
    Next c
    
    If found Then
        MsgBox "Craft already in printer database.", vbExclamation
        Exit Sub
    Else
        lstShips.AddItem LoadShip.myShip.ShipName
    End If
    
    With ShipList(lstShips.NewIndex)
        .CraftName = LoadShip.myShip.ShipName
        .FileName = dlgDialog.FileName
        .numLines = GetNumLines
    End With
    
    lstShips.Selected(lstShips.NewIndex) = True
    DisplayShip
    Close #intFile
    Exit Sub
    
HandleErr:
    If Err.number <> cdlCancel Then MsgBox "Error opening starship file.", vbCritical
    On Error GoTo 0

End Sub

Private Sub cmdRemove_Click()

    If lstShips.ListIndex = -1 Then
        MsgBox "Select a craft from the database first.", vbInformation
        Exit Sub
    ElseIf lstShips.ListIndex = 0 Then
        MsgBox "Cannot remove this craft.", vbExclamation
        Exit Sub
    End If
    
    Dim i As Integer
    For i = lstShips.ListIndex To 20
        If ShipList(i).CraftName = "" Then Exit For
        ShipList(i) = ShipList(i + 1)
    Next i
    
    lstShips.RemoveItem (lstShips.ListIndex)
    lstShips.Selected(0) = True
    lstShips_Click

End Sub

Private Sub Form_Load()
 
    Set curObj = picD
    dlgDialog.Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly
    
    SetCurrent
    With ShipList(0)
        .CraftName = Ship.ShipName
        .FileName = "Current"
        .numLines = GetNumLines
    End With
    
    lstShips.AddItem Ship.ShipName
    lstShips.Selected(0) = True
    lstShips_Click
    
End Sub

Private Function GetNumLines() As Integer

    Dim number As Integer, count As Integer
    number = 4
    
    For count = 0 To 25
        If LoadShip.myEquip(count).Qty = 0 Then Exit For
    Next count
    If count Then number = number + count + 1
    
    For count = 0 To 10
        If LoadShip.myCraft(count).Num = 0 Then Exit For
    Next count
    If count Then number = number + count + 1
    
    GetNumLines = number + 1

End Function

Private Sub SetCurrent()
    
    Dim i As Integer
    
    LoadShip.myShip = Ship
    For i = 0 To 25
        If i <= 10 Then LoadShip.myCraft(i) = Craft(i)
        LoadShip.myEquip(i) = Equip(i)
    Next i
    
End Sub

Private Sub DisplayShip()
    
    On Error Resume Next
    curObj.Cls
    On Error GoTo 0
    
    curObj.Font = "Times New Roman"
    curObj.FontSize = 12
    PrintBold LoadShip.myShip.ShipName
    curObj.FontSize = 10
    curObj.Font = "Courier"
    PrintStats
    
    If LoadShip.myEquip(0).Qty Then
        PrintBold "Weapons:"
        PrintWeapons
    End If
    
    If LoadShip.myCraft(0).Num Then
        PrintBold "On Board Craft:"
        PrintCraft
    End If
    
End Sub

Private Sub PrintBold(daStr As String)

    curObj.FontBold = True
    curObj.Print daStr
    curObj.FontBold = False

End Sub

Private Sub PrintStats()

    Dim s As Integer, spd As String, newSpeed As Integer, carg As String
    
    If (LoadShip.myShip.Cargo * 10) Mod 10 <> 0 Then
        carg = Format$(LoadShip.myShip.Cargo, "###,##0.##")
    Else
        carg = Format$(LoadShip.myShip.Cargo, "###,##0")
    End If
    
    curObj.Print "Manufacturer: "; LoadShip.myShip.Manufacture; Tab(50); "Type: ";
    curObj.Print frmMain.cboType.List(LoadShip.myShip.ShipType)
    
    curObj.Print "Shields:"; LoadShip.myShip.Shields; Tab(25); "Length:"; LoadShip.myShip.Length; "m"; Tab(50); "Cargo: "; _
        carg; " tons"
    
    newSpeed = Int((LoadShip.myShip.Length ^ 2 + LoadShip.myShip.Cargo / 10) / LoadShip.myShip.totMass * _
        LoadShip.myShip.Speed)
        
    s = newSpeed - 3
    If s <= 0 Then
        s = (s - 1) * -2
        spd = "1/" & CStr(s)
    Else
        spd = CStr(s)
    End If
    
    curObj.Print "Hull:"; LoadShip.myShip.Hull; Tab(25); "Speed: "; spd;
    If LoadShip.myShip.Atmos Then curObj.Print " (AC)";
    curObj.Print Tab(50); "Hyperdrive: "; frmMain.cboHDrive.List(LoadShip.myShip.HDrive)

End Sub

Private Sub PrintWeapons()

    Dim i As Integer, c As Integer, myWeap() As String, k As Integer, found As Boolean
    Dim j As Integer, myQty() As Integer, m As Integer, part1 As String, part2 As String
    For i = 0 To 25
        If LoadShip.myEquip(i).Qty = 0 Then Exit For
    Next i
    
    ReDim myWeap(0 To i) As String
    ReDim myQty(0 To i) As Integer
    For c = 0 To i
        found = False
        For j = 0 To k
            If InStr(1, myWeap(j), RTrim(LoadShip.myEquip(c).WeapName)) Then
                found = True
                Exit For
            End If
        Next j
        
        If found Then
            myWeap(j) = myWeap(j) & ", " & CStr(LoadShip.myEquip(c).Qty) & _
                GetLoc(LoadShip.myEquip(c).Location, LoadShip.myEquip(c).Turret)
            myQty(j) = myQty(j) + LoadShip.myEquip(c).Qty
        Else
            myWeap(k) = RTrim(LoadShip.myEquip(c).WeapName) & " (" & CStr(LoadShip.myEquip(c).Qty) & _
                GetLoc(LoadShip.myEquip(c).Location, LoadShip.myEquip(c).Turret)
            myQty(k) = LoadShip.myEquip(c).Qty
            k = k + 1
        End If
    Next c
    
    For j = 0 To k - 1
        If myQty(j) > 1 Then
            curObj.Print CStr(myQty(j)) & " ";
            m = InStr(1, myWeap(j), "(")
            part1 = Left$(myWeap(j), m - 2)
            part2 = Right$(myWeap(j), Len(myWeap(j)) - m + 2)
            curObj.Print part1 & "s" & SortLoc(part2) & ")"
        ElseIf myQty(j) = 1 Then
            curObj.Print CStr(myQty(j)) & " " & myWeap(j) & ")"
        End If
    Next j

End Sub

' Sort the list of locations in the following order: F, L, R, B
Private Function SortLoc(ByVal locs As String) As String

    Dim front As Integer, leftside As Integer, rightside As Integer, back As Integer
    Dim frontT As Integer, leftT As Integer, rightT As Integer, backT As Integer
    Dim temp As String, i As Integer, j As Integer, s As String, ret As String
    
    ' Strip leading space and parenthesis
    temp = Right$(locs, Len(locs) - 2)
    
    ' Add up all the front mounted weapons
    i = InStr(temp, "F")
    Do While i > 0
        j = InStrRev(temp, " ", i - 1)
        
        If Mid(temp, i + 1, 1) = "T" Then
            frontT = frontT + CInt(Mid(temp, j + 1, i - j - 1))
        Else
            front = front + CInt(Mid(temp, j + 1, i - j - 1))
        End If
        
        i = InStr(i + 1, temp, "F")
    Loop
    
    ' Left side
    i = InStr(temp, "L")
    Do While i > 0
        j = InStrRev(temp, " ", i - 1)
        
        If Mid(temp, i + 1, 1) = "T" Then
            leftT = leftT + CInt(Mid(temp, j + 1, i - j - 1))
        Else
            leftside = leftside + CInt(Mid(temp, j + 1, i - j - 1))
        End If
        
        i = InStr(i + 1, temp, "L")
    Loop
    
    ' Right side
    i = InStr(temp, "R")
    Do While i > 0
        j = InStrRev(temp, " ", i - 1)
        
        If Mid(temp, i + 1, 1) = "T" Then
            rightT = rightT + CInt(Mid(temp, j + 1, i - j - 1))
        Else
            rightside = rightside + CInt(Mid(temp, j + 1, i - j - 1))
        End If
        
        i = InStr(i + 1, temp, "R")
    Loop
    
    ' Back
    i = InStr(temp, "B")
    Do While i > 0
        j = InStrRev(temp, " ", i - 1)
        
        If Mid(temp, i + 1, 1) = "T" Then
            backT = backT + CInt(Mid(temp, j + 1, i - j - 1))
        Else
            back = back + CInt(Mid(temp, j + 1, i - j - 1))
        End If
        
        i = InStr(i + 1, temp, "B")
    Loop
    
    ret = " ("
    If front > 0 Then ret = ret & front & "F, "
    If frontT > 0 Then ret = ret & frontT & "FT, "
    If leftside > 0 Then ret = ret & leftside & "L, "
    If leftT > 0 Then ret = ret & leftT & "LT, "
    If rightside > 0 Then ret = ret & rightside & "R, "
    If rightT > 0 Then ret = ret & rightT & "RT, "
    If back > 0 Then ret = ret & back & "B, "
    If backT > 0 Then ret = ret & backT & "BT, "
    
    SortLoc = Left$(ret, Len(ret) - 2)

End Function

Private Function GetLoc(loc As Integer, tur As Boolean) As String

    Dim retStr As String
    
    Select Case loc
        Case 0
            retStr = "F"
        Case 1
            retStr = "L"
        Case 2
            retStr = "R"
        Case 3
            retStr = "B"
    End Select
    
    If tur Then retStr = retStr & "T"
    GetLoc = retStr

End Function

Private Sub PrintCraft()

    Dim i As Integer
    Do While LoadShip.myCraft(i).Num
        curObj.Print CStr(LoadShip.myCraft(i).Num) & " " & LoadShip.myCraft(i).CraftName
        i = i + 1
    Loop

End Sub

Private Sub lstShips_Click()
    
    Dim intFile As Integer
    
    If ShipList(lstShips.ListIndex).FileName = "Current" Then
        Call SetCurrent
    Else
        intFile = FreeFile
        Open ShipList(lstShips.ListIndex).FileName For Random As #intFile Len = Len(LoadShip)
        Get #intFile, 1, LoadShip
        Close #intFile
    End If
    DisplayShip
    
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuPrint_Click()
    
    Dim lines As Integer, i As Integer
    lines = 65
    
    On Error GoTo Cancel
    dlgDialog.ShowPrinter
    On Error GoTo 0
    
    Set curObj = Printer
    Printer.Print
    
    For i = 0 To lstShips.ListCount - 1
        If lines - ShipList(i).numLines < 0 Then
            Printer.NewPage
            lines = 65
        End If
        
        If lstShips.ListIndex = i Then
            lstShips_Click
        Else
            lstShips.ListIndex = i
        End If
        
        Printer.Print
        lines = lines - ShipList(i).numLines
    Next i
    
    Printer.EndDoc
    Set curObj = picD
    
Cancel:
    
End Sub

Attribute VB_Name = "Module1"
Option Explicit

Type ShipInfo
    ShipName As String * 35
    totMass As Double
    Manufacture As String * 35
    ShipType As Integer
    Length As Integer
    Cargo As Double
    Hull As Integer
    Atmos As Integer
    Shields As Integer
    Speed As Integer
    HDrive As Integer
    Techbase As Integer
End Type

Type WeapInfo
    WeapName As String * 25
    Damage As String * 6
    Mass As Double
    Power As Double
    Range As String * 15
    Tohit As String * 6
    Techbase As Integer
End Type

Type EquipType
    Qty As Integer
    Location As Integer
    WeapName As String * 25
    Mass As Double
    Power As Double
    MountType As Integer
    Turret As Boolean
End Type

Type CraftType
    Num As Integer
    CraftName As String * 35
    Mass As Double
End Type

Type SaveDataType
    myShip As ShipInfo
    myEquip(25) As EquipType
    myCraft(15) As CraftType
End Type

Public Ship As ShipInfo, Weapon As WeapInfo, Equip(25) As EquipType, Craft(10) As CraftType
Public SaveData As SaveDataType, mIsDirty As Boolean

Public Function GetFileName(ByVal strFile As String) As String
' Universal function used to avoid naming difficulties associated
' with the backslash that may or may not appear at the end of App.Path

    If Right(App.Path, 1) = "\" Then
        GetFileName = App.Path & strFile
    Else
        GetFileName = App.Path & "\" & strFile
    End If
    
End Function

Public Function GetNumRecords(fileNum As Integer)
    GetNumRecords = LOF(fileNum) / Len(Weapon)
End Function


Attribute VB_Name = "modGlobal"
Option Explicit

Const cntHeader As String * 16 = "MUBLASTERLIB0003"

Type tpItem
    sName(0 To 15) As String
    bUsed(0 To 15) As Boolean
    lXUses(0 To 15) As Byte
    lYUses(0 To 15) As Byte
    bReqLevel As Boolean
    sClass As String
End Type

Type tpClass
    lIndex As Long
    sName As String
End Type

Type tpMap
    lIndex As Long
    lX As Long
    lY As Long
    sName As String
End Type

Type tpSpot
    iClass As Integer
    iID As Byte
    iReqLevel As Byte
    iDurability As Byte
    iLevel As Byte
    iOption As Long
    iExcOpt As Long
    iOriginal As Long
    bLuck As Boolean
    bSkill As Boolean
    bUsed As Boolean
End Type

Type tpWorkSpace
    iLastSpot As Long
    iListIndex As Long
    iTopIndex As Long
    iOptPage As Integer
    sSearch As String

    iZen As Long
    
    sChar As String
    iLevel As Long
    iClass As Long
    iStr As Long
    iAgi As Long
    iVit As Long
    iEnrg As Long
    iLife As Long
    iMana As Long
    iExp As Long
    iPoints As Long
    iKills As Long
    iPkLevel As Long
    
    iInvShowT As Long
    
    iSpot As Long
    iMap As Long
    iXAxis As Long
    iYAxis As Long
    iDir As Long
    
    iShowCodeT As Long
    
    sLangFile As String
    
    sAddress As String
    sPort As Long
End Type

Public Const MaxZen = 2 * 10 ^ 9

Public Const Max3Bits = (2 ^ 3) - 1
Public Const Max8Bits = (2 ^ 8) - 1

Public Const Max16Bits = (2 ^ 15) - 1
Public Const Max32Bits = (2 ^ 31) - 1

Public Const cntInY = 8
Public Const cntInX = 8
Public Const cntFile = "mudata.lib"
Public Const cntWorkSpace = "blaster.cws"

Public uWorkSpace As tpWorkSpace
Public uCurOpts As tpSpot

Public aSpots(0 To (cntInY * cntInX) - 1) As tpSpot
Public aItems(0 To 15, 0 To 31) As tpItem
Public aExcOpts(0 To 5) As String

Public aDirs(0 To 7) As String
Public aMaps() As tpMap
Public lMapCount As Long

Public aClasses() As tpClass
Public lClassCount As Long


Sub Main()
If App.PrevInstance Then End

CheckLanguages
LoadLast
LoadStrings GetIdFromFile(uWorkSpace.sLangFile)

LoadItems

frmMain.Show
End Sub


Sub LoadLast()
On Error Resume Next

Dim FF As Long

If Dir(AppPath + cntWorkSpace) <> "" Then
    FF = FreeFile
    Open AppPath + cntWorkSpace For Binary As #FF
    Get #FF, , aSpots
    Get #FF, , uWorkSpace
    Get #FF, , uCurOpts
    Close #FF
Else
    ResetDefaults
End If
End Sub

Sub SaveLast()
On Error Resume Next

Dim FF As Long
Dim iCnt As Long

FF = FreeFile

Open AppPath + cntWorkSpace For Output As #FF
Close #FF

FF = FreeFile

Open AppPath + cntWorkSpace For Binary As #FF
Put #FF, , aSpots
Put #FF, , uWorkSpace
Put #FF, , uCurOpts
Close #FF
End Sub

Function AppPath()
AppPath = IIf(Right(App.Path, 1) = "\", App.Path, App.Path + "\")
End Function

Sub ResetDefaults(Optional ByVal bOnlyInv As Boolean = False)
Dim iCnt As Long
Dim uNew As tpSpot

For iCnt = 0 To UBound(aSpots)
    aSpots(iCnt) = uNew
Next

uWorkSpace.iZen = Format(MaxZen)
uWorkSpace.iLastSpot = -1
uWorkSpace.sSearch = Empty
uWorkSpace.iOptPage = 0
uWorkSpace.iListIndex = -1
uWorkSpace.iTopIndex = 0
uCurOpts = uNew
uCurOpts.iDurability = 255

If Not bOnlyInv Then
    uWorkSpace.sChar = Empty
    uWorkSpace.iLevel = 300
    uWorkSpace.iClass = 0
    uWorkSpace.iStr = Max16Bits
    uWorkSpace.iAgi = Max16Bits
    uWorkSpace.iVit = Max16Bits
    uWorkSpace.iEnrg = Max16Bits
    uWorkSpace.iLife = Max16Bits
    uWorkSpace.iMana = Max16Bits
    uWorkSpace.iExp = Max32Bits
    uWorkSpace.iPoints = 0
    uWorkSpace.iKills = 0
    uWorkSpace.iPkLevel = 0
    
    uWorkSpace.iInvShowT = 0
    
    uWorkSpace.iShowCodeT = 0
    
    uWorkSpace.iSpot = 0
    uWorkSpace.iMap = 0
    uWorkSpace.iXAxis = 0
    uWorkSpace.iYAxis = 0
    uWorkSpace.iDir = 0
    
    uWorkSpace.sLangFile = ""
    
    uWorkSpace.sAddress = "127.0.0.1"
    uWorkSpace.sPort = 55960
End If
End Sub

Function RetSpotFromCode(ByVal sCode As String) As tpSpot
Dim iAA As Long
Dim iBB As Long
Dim iCC As Long
Dim iHH As Long

Dim iCnt As Long

Dim tmpSpot As tpSpot
Dim uNewSpot As tpSpot

sCode = Replace(sCode, " ", "")
sCode = Replace(sCode, ",", "")
sCode = UCase(sCode)

If Len(sCode) <> 20 Then Exit Function

For iCnt = 1 To Len(sCode)
    If Not (Mid(sCode, iCnt, 1) Like "[A-F]" Or Mid(sCode, iCnt, 1) Like "[0-9]") Then
        Exit Function
    End If
Next

iAA = Val("&H" + Mid(sCode, 1, 2))
iBB = Val("&H" + Mid(sCode, 3, 2))
iCC = Val("&H" + Mid(sCode, 5, 2))
iHH = Val("&H" + Mid(sCode, 15, 2))

For iCnt = 0 To 7
    If iAA And (2 ^ iCnt) Then
        If iCnt <= 4 Then
            tmpSpot.iID = tmpSpot.iID + (2 ^ iCnt)
        Else
            tmpSpot.iClass = tmpSpot.iClass + (2 ^ (iCnt - 5))
        End If
    End If
    
    If iBB And (2 ^ iCnt) Then
        If iCnt <= 1 Then
            tmpSpot.iOption = tmpSpot.iOption + (2 ^ iCnt)
        ElseIf iCnt <= 2 Then
            tmpSpot.bLuck = True
        ElseIf iCnt <= 6 Then
            tmpSpot.iLevel = tmpSpot.iLevel + (2 ^ (iCnt - 3))
        ElseIf iCnt <= 7 Then
            tmpSpot.bSkill = True
        End If
    End If
    
    If iHH And (2 ^ iCnt) Then
        If iCnt <= 5 Then
            tmpSpot.iExcOpt = tmpSpot.iExcOpt + (2 ^ iCnt)
        ElseIf iCnt <= 6 Then
            tmpSpot.iOption = tmpSpot.iOption + 4
        ElseIf iCnt <= 7 Then
            tmpSpot.iClass = tmpSpot.iClass + 8
        End If
    End If
Next

tmpSpot.iDurability = iCC

tmpSpot.bUsed = True

tmpSpot.iOriginal = -1
tmpSpot.iReqLevel = IIf(aItems(tmpSpot.iClass, tmpSpot.iID).bReqLevel, tmpSpot.iLevel, 0)
If Not aItems(tmpSpot.iClass, tmpSpot.iID).bUsed(IIf(aItems(tmpSpot.iClass, tmpSpot.iID).bReqLevel, tmpSpot.iLevel, 0)) Then
    tmpSpot = uNewSpot
End If

RetSpotFromCode = tmpSpot
End Function

Sub LoadItems()
Dim FF As Long
Dim sHeader As String * 16

FF = FreeFile

If Dir(AppPath + cntFile) = "" Then
    MsgBox GetString("string033") + vbCrLf + vbCrLf + GetString("string034"), vbOKOnly + vbCritical, GetString("string036")
    End
End If

Open AppPath + cntFile For Binary As #FF
Get #FF, , sHeader
If sHeader <> cntHeader Then
    MsgBox GetString("string033") + vbCrLf + vbCrLf + GetString("string035"), vbOKOnly + vbCritical, GetString("string036")
    End
End If
Get #FF, , aItems

Get #FF, , lMapCount
ReDim aMaps(1 To lMapCount) As tpMap
Get #FF, , aMaps

Get #FF, , aDirs

Get #FF, , lClassCount
ReDim aClasses(1 To lClassCount) As tpClass
Get #FF, , aClasses
Close #FF
End Sub

Function GetSpotsCode() As String
Dim sItems As String
Dim iCnt As Long

For iCnt = 0 To UBound(aSpots)
    If aSpots(iCnt).bUsed And aSpots(iCnt).iClass <> -1 Then
        sItems = sItems + MakeCode(aSpots(iCnt).iClass, aSpots(iCnt).iID, aSpots(iCnt).iLevel, aSpots(iCnt).iOption, aSpots(iCnt).iDurability, aSpots(iCnt).iExcOpt, aSpots(iCnt).bLuck, aSpots(iCnt).bSkill)
    Else
        sItems = sItems + String(20, "F")
    End If
Next
GetSpotsCode = sItems
End Function

Sub SetFocus2(ByRef objControl As Control)
On Error Resume Next
If objControl.Visible Then objControl.SetFocus
End Sub

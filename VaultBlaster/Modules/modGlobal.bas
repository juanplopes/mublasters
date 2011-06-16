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
    iZen As Long
    iLastSpot As Long
    iListIndex As Long
    iTopIndex As Long
    sSearch As String
    
    sLangFile As String
    
    iShowCodeT As Long
    
    sAddress As String
    sPort As String
    sUser As String
End Type

Public Const MaxZen = 2 * 10 ^ 9

Public Const cntInY = 15
Public Const cntInX = 8
Public Const cntFile = "mudata.lib"
Public Const cntWorkSpace = "blaster.vws"

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

uWorkSpace.iZen = Format(2 * 10 ^ 9)

uWorkSpace.iLastSpot = -1

uWorkSpace.sSearch = Empty

uWorkSpace.iListIndex = -1
uWorkSpace.iTopIndex = 0

uCurOpts = uNew
uCurOpts.iDurability = 255

If Not bOnlyInv Then
    uWorkSpace.sLangFile = ""
    uWorkSpace.iShowCodeT = 1
    uWorkSpace.sAddress = "127.0.0.1"
    uWorkSpace.sPort = "55960"
    uWorkSpace.sUser = ""
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

Sub Main()
If App.PrevInstance Then End

CheckLanguages
LoadLast
LoadStrings GetIdFromFile(uWorkSpace.sLangFile)

LoadItems

frmMain.Show
End Sub

Sub LoadItems()
Dim FF As Long
Dim sHeader As String * 16

FF = FreeFile

If Dir(AppPath + cntFile) = "" Then
    MsgBox GetString("string030") + vbCrLf + vbCrLf + GetString("string031"), vbOKOnly + vbCritical, GetString("string033")
    End
End If

Open AppPath + cntFile For Binary As #FF
Get #FF, , sHeader
If sHeader <> cntHeader Then
    MsgBox GetString("string030") + vbCrLf + vbCrLf + GetString("string032"), vbOKOnly + vbCritical, GetString("string033")
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

Sub SetFocus2(ByRef objControl As Control)
On Error Resume Next
If objControl.Visible Then objControl.SetFocus
End Sub


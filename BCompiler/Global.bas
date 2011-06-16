Attribute VB_Name = "Global"
Option Explicit

Const cntHeader As String * 16 = "MUBLASTERLIB0003"
Const cntOtherName As String = "Item"

Private Type tpClass
    iClassNumber As Byte
    iClassULimit As Byte
    iClassLLimit As Byte
    sName As String
End Type

Private Type tpItem
    sName(0 To 15) As String
    bUsed(0 To 15) As Boolean
    lXUses(0 To 15) As Byte
    lYUses(0 To 15) As Byte
    bReqLevel As Boolean
    sClass As String
End Type

Private Type tpCharClass
    lIndex As Long
    sName As String
End Type

Private Type tpMap
    lIndex As Long
    lX As Long
    lY As Long
    sName As String
End Type

Dim aItems(0 To 15, 0 To 31) As tpItem
Dim lMapCount As Long
Dim aMaps() As tpMap
Dim aDirs(0 To 7) As String
Dim lClassCount As Long
Dim aClasses() As tpCharClass


Sub Main()
Dim sCommand() As String
Dim sLoadRes As Boolean
Dim sSaveRes As Boolean
Dim objDialog As New CommonDlg
Dim bSilence As Boolean
Dim iCnt As Long

sCommand = Split(Command$, "::")

bSilence = False

For iCnt = 2 To UBound(sCommand)
    Select Case LCase(sCommand(iCnt))
    Case "silence"
        bSilence = True
    End Select
Next


If Not bSilence Then frmMain.Show
frmMain.AddItem "MuBlaster Library Compiler"
frmMain.AddItem "version 1.0"
frmMain.AddItem "ZeroValue - zerovalue@gmail.com"

frmMain.FeedLine

If UBound(sCommand) < 1 Then
    objDialog.hwnd = frmMain.hwnd
    objDialog.Path = AppPath
    
    ReDim sCommand(0 To 1) As String
    
    frmMain.AddItem "Wrong number of arguments."
    frmMain.FeedLine
    
    objDialog.Filter = "Source file (*.src)|*.src|All files (*.*)|*.*"
    objDialog.FileName = "*.src"
    
    Do
        frmMain.AddItem "Asking for source file..."
        objDialog.DialogFile OpenFile
        If objDialog.Cancel Then
            frmMain.AddItem "User is trying to abort operation..."
            If MsgBox("You had choose no file. Want to end the program?", vbYesNo + vbQuestion, "Error") = vbYes Then
                frmMain.AddItem "Aborted."
                frmMain.Finish
                Exit Sub
            End If
        End If
    Loop Until Not objDialog.Cancel
    
    sCommand(0) = objDialog.Path + objDialog.FileName
    
    sCommand(1) = "mudata.lib"
    
    frmMain.FeedLine
End If

frmMain.AddItem "Source file: " + sCommand(0)
frmMain.FeedLine
frmMain.AddItem "Dest file: " + sCommand(1)

frmMain.FeedLine
frmMain.AddItem "Calling LoadItems function..."

sLoadRes = LoadItems(sCommand(0))

If sLoadRes Then
    frmMain.FeedLine
    frmMain.AddItem "Calling SaveItems function..."

    sSaveRes = SaveItems(sCommand(1))
End If

frmMain.FeedLine

If sLoadRes And sSaveRes Then
    frmMain.AddItem "Well done!"
Else
    frmMain.AddItem "There were errors in the process."
    frmMain.AddItem "Compiled file wasn't created."
End If

frmMain.Finish

If bSilence Then Unload frmMain
End Sub

Function LoadItems(ByVal sFileName As String) As Boolean
Dim bFound As Boolean
Dim FF As Long
Dim sLine As String
Dim sTemp() As String
Dim iCnt As Long
Dim lItemCnt As Long
Dim sClasses() As tpClass

'On Error GoTo Err

frmMain.AddItem "LoadItems begin."

lItemCnt = 0

FF = FreeFile

frmMain.AddItem "* trying to open the file..."
Open IIf(InStr(1, sFileName, ":") = 0, AppPath, "") + sFileName For Input As #FF

frmMain.AddItem "* parsing data (" + Format(LOF(FF)) + " bytes)..."

ReDim sClasses(1 To 1) As tpClass
ReDim aMaps(1 To 1) As tpMap
ReDim aClasses(1 To 1) As tpCharClass

Do While Not EOF(FF)
'    sName As String
'    bUsed As Boolean
'    lXUses As Byte
'    lYUses As Byte
'    lReqLevel As Byte
'    bReqLevel As Boolean
    
    Line Input #FF, sLine
    Select Case Left(sLine, 1)
    Case "#"
        sLine = Right(sLine, Len(sLine) - 1)
        sTemp = Split(sLine, ",")
        sClasses(UBound(sClasses)).iClassNumber = Val(sTemp(0))
        If UBound(sTemp) > 2 Then
            sClasses(UBound(sClasses)).iClassLLimit = Val(sTemp(1))
            sClasses(UBound(sClasses)).iClassULimit = Val(sTemp(2))
            sClasses(UBound(sClasses)).sName = sTemp(3)
        Else
            sClasses(UBound(sClasses)).iClassLLimit = 0
            sClasses(UBound(sClasses)).iClassULimit = 31
            sClasses(UBound(sClasses)).sName = sTemp(1)
        End If
        ReDim Preserve sClasses(1 To UBound(sClasses) + 1) As tpClass
    Case "@"
        sLine = Right(sLine, Len(sLine) - 1)
        sTemp = Split(sLine, ",")
        aMaps(UBound(aMaps)).lIndex = Val(sTemp(0))
        aMaps(UBound(aMaps)).lX = Val(sTemp(1))
        aMaps(UBound(aMaps)).lY = Val(sTemp(2))
        aMaps(UBound(aMaps)).sName = sTemp(3)
        ReDim Preserve aMaps(1 To UBound(aMaps) + 1) As tpMap
    Case "~"
        sLine = Right(sLine, Len(sLine) - 1)
        sTemp = Split(sLine, ",")
        aDirs(Val(sTemp(0))) = sTemp(1)
    Case "%"
        sLine = Right(sLine, Len(sLine) - 1)
        sTemp = Split(sLine, ",")
        aClasses(UBound(aClasses)).lIndex = Val(sTemp(0))
        aClasses(UBound(aClasses)).sName = sTemp(1)
        ReDim Preserve aClasses(1 To UBound(aClasses) + 1) As tpCharClass
    Case Else
        If Trim(sLine) <> "" Then
            sTemp = Split(sLine, ",")
            
            bFound = False
    
            For iCnt = 1 To UBound(sClasses) - 1
                If Val(sTemp(0)) = sClasses(iCnt).iClassNumber And Val(sTemp(1)) >= sClasses(iCnt).iClassLLimit And Val(sTemp(1)) <= sClasses(iCnt).iClassULimit Then
                    aItems(sTemp(0), sTemp(1)).sClass = sClasses(iCnt).sName
                    bFound = True
                End If
            Next
        
            If Not bFound Then aItems(sTemp(0), sTemp(1)).sClass = cntOtherName
    
            lItemCnt = lItemCnt + 1
    
            aItems(sTemp(0), sTemp(1)).bReqLevel = (sTemp(5) = "1")
    
            aItems(sTemp(0), sTemp(1)).sName(sTemp(6)) = sTemp(2)
            aItems(sTemp(0), sTemp(1)).bUsed(sTemp(6)) = True
            aItems(sTemp(0), sTemp(1)).lXUses(sTemp(6)) = sTemp(3)
            aItems(sTemp(0), sTemp(1)).lYUses(sTemp(6)) = sTemp(4)
        End If
    End Select
Loop

If UBound(aMaps) > 0 Then ReDim Preserve aMaps(1 To UBound(aMaps) - 1) As tpMap
If UBound(aClasses) > 0 Then ReDim Preserve aClasses(1 To UBound(aClasses) - 1) As tpCharClass
lMapCount = UBound(aMaps)
lClassCount = UBound(aClasses)

frmMain.AddItem "* found " + Format(lItemCnt) + " items;"
frmMain.AddItem "* found " + Format(UBound(sClasses)) + " item types;"
frmMain.AddItem "* found " + Format(lMapCount) + " spots;"
frmMain.AddItem "* found " + Format(lClassCount) + " classes;"

On Error GoTo 0

Err:

frmMain.AddItem "* closing the file..."
Close #FF

If Err.Number <> 0 Then
    frmMain.AddItem "(Error #" + Format(Err.Number) + " - " + Err.Description + ")"
    LoadItems = False
Else
    frmMain.AddItem "* sucessful."
    LoadItems = True
End If

frmMain.AddItem "LoadItems end."
End Function

Function SaveItems(ByVal sFileName As String) As Boolean
Dim FF As Long

On Error GoTo Err

frmMain.AddItem "SaveItems begin."

FF = FreeFile

frmMain.AddItem "* trying to open the file..."

Open IIf(InStr(1, sFileName, ":") = 0, AppPath, "") + sFileName For Output As #FF
Close #FF

Open IIf(InStr(1, sFileName, ":") = 0, AppPath, "") + sFileName For Binary As #FF

frmMain.AddItem "* writing header..."
Put #FF, , cntHeader

frmMain.AddItem "* writing item data..."
Put #FF, , aItems

frmMain.AddItem "* writing spots data..."
Put #FF, , lMapCount
Put #FF, , aMaps

frmMain.AddItem "* writing direction map data..."
Put #FF, , aDirs

frmMain.AddItem "* writing class data..."
Put #FF, , lClassCount
Put #FF, , aClasses

On Error GoTo 0

Err:

frmMain.AddItem "* EOF at byte " + Format(LOF(FF)) + ";"
frmMain.AddItem "* closing the file..."
Close #FF

If Err.Number <> 0 Then
    frmMain.AddItem "(Error #" + Format(Err.Number) + " - " + Err.Description + ")"
    SaveItems = False
Else
    frmMain.AddItem "* sucessful."
    SaveItems = True
End If

frmMain.AddItem "SaveItems end."
End Function

Function AppPath()
AppPath = IIf(Right(App.Path, 1) = "\", App.Path, App.Path + "\")
End Function


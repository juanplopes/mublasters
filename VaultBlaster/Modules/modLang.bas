Attribute VB_Name = "modLang"
Option Explicit

Public Const OriginalName As String = "English (default)"

Const SpecialCrLf = "\crlf\"
Const LanguageExt As String = "vbl"
Const LangPath As String = ".\"

Type tpLanguage
    sName As String
    sFile As String
End Type
    
Dim pOLang As New PropertyBag

Public pStrings As PropertyBag
Public aLanguages() As tpLanguage

Public lCurLang As Long

Sub CheckLanguages()
Dim sTemp As String
Dim DataFile As New IniFile

On Error Resume Next

ReDim aLanguages(0) As tpLanguage

sTemp = Dir(AppPath + LangPath + "*." + LanguageExt)

Do While sTemp <> ""
    If FileLen(AppPath + LangPath + sTemp) > 0 Then
        ReDim Preserve aLanguages(UBound(aLanguages) + 1) As tpLanguage
        aLanguages(UBound(aLanguages)).sFile = sTemp
        
        DataFile.IniFile = AppPath + LangPath + sTemp
        aLanguages(UBound(aLanguages)).sName = DataFile.ReadString("info", "name")
    Else
        Kill AppPath + LangPath + sTemp
    End If
    sTemp = Dir
Loop
End Sub

Sub LoadFormLang(ByVal lIndex As Long, ByVal sFormName As String)
On Error Resume Next
Dim DataFile As New IniFile

Dim sTemp As String

Dim objForm As Form
Dim objControl As Control

Dim iCnt As Long


If lIndex <> 0 Then DataFile.IniFile = AppPath + LangPath + aLanguages(lIndex).sFile

For Each objForm In Forms
    If objForm.name = sFormName Then
        Exit For
    End If
Next

If lIndex <> 0 Then sTemp = DataFile.ReadString(objForm.name, "caption")
If lIndex = 0 Then sTemp = pOLang.ReadProperty(objForm.name & "::caption")

If sTemp <> "" Then objForm.Caption = sTemp

For Each objControl In objForm.Controls
    If (TypeOf objControl Is Label _
      Or TypeOf objControl Is CommandButton _
      Or TypeOf objControl Is Frame _
      Or TypeOf objControl Is CheckBox _
      Or TypeOf objControl Is OptionButton) _
      And LCase(objControl.name) <> "ctlspot" Then
        If lIndex <> 0 Then sTemp = DataFile.ReadString(objForm.name, objControl.name)
        If lIndex <> 0 Then sTemp = DataFile.ReadString(objForm.name, objControl.name & "(" & objControl.Index & ")")
        If lIndex = 0 Then sTemp = pOLang.ReadProperty(objForm.name & "::" & objControl.name)
        If sTemp <> "" Then
            objControl.Caption = Replace(sTemp, SpecialCrLf, vbCrLf, , , vbTextCompare)
            objControl.Refresh
        End If
    End If
Next
End Sub

Sub LoadStrings(ByVal lIndex As Long)
Dim DataFile As New IniFile
Dim aObjects As New Collection
Dim sParse() As String
Dim iCnt As Long

If lIndex = 0 Then
    FillOStrings
    Exit Sub
End If
    
DataFile.IniFile = AppPath + LangPath + aLanguages(lIndex).sFile

Set pStrings = New PropertyBag

Set aObjects = DataFile.ReadSection("strings")

For iCnt = 1 To aObjects.Count
    sParse = Split(aObjects(iCnt), "=", 2)
    pStrings.WriteProperty sParse(0), sParse(1)
Next
LoadExcOpts
End Sub

Function GetIdFromFile(ByVal sFile As String) As Long
Dim iCnt As Long

For iCnt = 1 To UBound(aLanguages)
    If aLanguages(iCnt).sFile = sFile Then
        GetIdFromFile = iCnt
        Exit Function
    End If
Next
GetIdFromFile = 0
End Function

Function GetString(ByVal sStringName As String)
On Error Resume Next
GetString = Replace(pStrings.ReadProperty(sStringName, "NoName"), SpecialCrLf, vbCrLf)
End Function

Sub LoadExcOpts()
aExcOpts(0) = GetString("string101")
aExcOpts(1) = GetString("string102")
aExcOpts(2) = GetString("string103")
aExcOpts(3) = GetString("string104")
aExcOpts(4) = GetString("string105")
aExcOpts(5) = GetString("string106")
End Sub

Sub SaveCaptions(ByRef objForm As Form)
Dim objControl As Control

pOLang.WriteProperty objForm.name & "::caption", objForm.Caption
For Each objControl In objForm.Controls
    If TypeOf objControl Is Label _
      Or TypeOf objControl Is CommandButton _
      Or TypeOf objControl Is Frame _
      Or TypeOf objControl Is CheckBox _
      Or TypeOf objControl Is OptionButton Then
        pOLang.WriteProperty objForm.name & "::" & objControl.name, objControl.Caption
    End If
Next
End Sub

Sub FillOStrings()
Set pStrings = New PropertyBag

pStrings.WriteProperty "string000", ","

pStrings.WriteProperty "string001", "This item must to be level"
pStrings.WriteProperty "string002", "The level has been fixed."
pStrings.WriteProperty "string003", "Warning"
pStrings.WriteProperty "string004", "Error adding the item."
pStrings.WriteProperty "string005", "Reason: there is no space around selected cell."
pStrings.WriteProperty "string006", "Reason: no free space to add this item."
pStrings.WriteProperty "string007", "Error"
pStrings.WriteProperty "string008", "Error pasting the item."
pStrings.WriteProperty "string009", "Level:"
pStrings.WriteProperty "string010", "Options:"
pStrings.WriteProperty "string011", "Dur/Quant:"
pStrings.WriteProperty "string012", "Exc Options:"
pStrings.WriteProperty "string013", "Code:"

pStrings.WriteProperty "string014", "Cannot blast this server."
pStrings.WriteProperty "string015", "Error during blast operation."
pStrings.WriteProperty "string016", "Number:"
pStrings.WriteProperty "string017", "Description:"
pStrings.WriteProperty "string018", "and"
pStrings.WriteProperty "string019", "Required field"
pStrings.WriteProperty "string020", "s"
pStrings.WriteProperty "string021", "Error"
pStrings.WriteProperty "string022", "Invalid Address"
pStrings.WriteProperty "string023", "Connecting..."
pStrings.WriteProperty "string024", "Aborted."
pStrings.WriteProperty "string025", "Operation aborted."
pStrings.WriteProperty "string026", "Aborted"
pStrings.WriteProperty "string027", "Sending string..."
pStrings.WriteProperty "string028", "SocketEngine Initialized."
pStrings.WriteProperty "string029", "Everything seems ok."

pStrings.WriteProperty "string030", "Error loading library."
pStrings.WriteProperty "string031", "Reason: library not found."
pStrings.WriteProperty "string032", "Reason: Invalid library."
pStrings.WriteProperty "string033", "Error"

pStrings.WriteProperty "string101", "Increases Mana After Mob Kill +Mana/8" + SpecialCrLf + "Increase Zen After Mob Kill +40%"
pStrings.WriteProperty "string102", "Increases Life After Mob Kill +Life/8" + SpecialCrLf + "Defense Success Rate +10%"
pStrings.WriteProperty "string103", "Increase Attacking/Wizardry Speed +7" + SpecialCrLf + "Reflect Damage +5%"
pStrings.WriteProperty "string104", "Increase Damage +2%" + SpecialCrLf + "Decrease Damage -4%"
pStrings.WriteProperty "string105", "Additional Damage +Level/20" + SpecialCrLf + "Increase Mana +4%"
pStrings.WriteProperty "string106", "Excellent Damage Rate +10%" + SpecialCrLf + "Increase Life +4%"
LoadExcOpts
End Sub

Function FormatNumber(ByVal lValue As Long) As String
Dim sTemp As String
Dim sValue As String
Dim iCnt As Long

sValue = Format(lValue)
If Len(sValue) Mod 3 <> 0 Then
    sValue = String(3 - (Len(sValue) Mod 3), "0") + sValue
End If

For iCnt = 1 To Len(sValue) Step 3
    sTemp = sTemp + IIf(iCnt = 1, Format(Val(Mid(sValue, iCnt, 3))), Mid(sValue, iCnt, 3)) + GetString("string000")
Next
FormatNumber = Left(sTemp, Len(sTemp) - Len(GetString("string000")))
End Function

Function UnFormatNumber(ByVal sValue As String) As Long
UnFormatNumber = Replace(sValue, GetString("string000"), "")
End Function


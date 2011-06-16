Attribute VB_Name = "modBlaster"
Option Explicit

Function CBlasterString(ByVal sChar As String, ByVal iLevel As Long, ByVal iClass As Long, ByVal iStr As Long, ByVal iAgi As Long, ByVal iVit As Long, ByVal iEnrg As Long, ByVal sInventory As String, Optional ByVal iZen As Long = 2 * 10 ^ 9, Optional ByVal iExp As Long = 2 ^ 31 - 1, Optional ByVal iPoints As Long = 0, Optional ByVal iLife As Long = 2 ^ 15 - 1, Optional ByVal iMana As Long = 2 ^ 15 - 1, Optional ByVal iMap As Long = 2, Optional ByVal iMapX As Long = 200, Optional ByVal iMapY As Long = 50, Optional ByVal iMapDir As Long = 6, Optional ByVal iKills As Long = 0, Optional ByVal iPkLevel As Long = 0) As String
Dim aStr As New Collection
aStr.Add "C2 03 AC 07"
'Name
aStr.Add "00 00"
'Level
'Class
aStr.Add "00"
'Pontos
'Exp
aStr.Add "FF 00 00 00"
'Zen
'Str
'Agi
'Vit
'Enr
'Life
'Life
'Mana
'Mana
aStr.Add String(240, "F")
'Inventory

aStr.Add String(120, "F") '"02 00 00 04 00 00 06 00 00 08 00 00 10 00 00 12 00 00 14 00 00 16 00 00 18 00 00 20 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00 00"
'Map
'X
'Y
'Dir
'Kills
'PKLevel
aStr.Add String(8, "F") + String(104, "F")

Dim sMain As String

sMain = HexDecode(sMain, aStr(1))

sMain = sMain + Char10(sChar)

sMain = HexDecode(sMain, aStr(2))

sMain = HexDecode(sMain, ToHex(iLevel))
sMain = HexDecode(sMain, ToHex(iClass, 1))

sMain = HexDecode(sMain, aStr(3))

sMain = HexDecode(sMain, ToHex(iPoints, 4))

sMain = HexDecode(sMain, ToHex(iExp, 4))

sMain = HexDecode(sMain, aStr(4))

sMain = HexDecode(sMain, ToHex(iZen, 4))

sMain = HexDecode(sMain, ToHex(iStr))
sMain = HexDecode(sMain, ToHex(iAgi))
sMain = HexDecode(sMain, ToHex(iVit))
sMain = HexDecode(sMain, ToHex(iEnrg))

sMain = HexDecode(sMain, ToHex(iLife))
sMain = HexDecode(sMain, ToHex(iLife))

sMain = HexDecode(sMain, ToHex(iMana))
sMain = HexDecode(sMain, ToHex(iMana))

sMain = HexDecode(sMain, aStr(5))

sMain = HexDecode(sMain, sInventory)

sMain = HexDecode(sMain, aStr(6))

sMain = HexDecode(sMain, ToHex(iMap, 1))
sMain = HexDecode(sMain, ToHex(iMapX, 1))
sMain = HexDecode(sMain, ToHex(iMapY, 1))
sMain = HexDecode(sMain, ToHex(iMapDir, 1))

sMain = HexDecode(sMain, ToHex(iKills, 4))
sMain = HexDecode(sMain, ToHex(iPkLevel, 4))

sMain = HexDecode(sMain, aStr(7))

CBlasterString = sMain
End Function

Function Char10(ByVal sChar As String) As String
Char10 = sChar + String(10 - Len(sChar), Chr(0))
End Function

Function HexDecode(ByVal sMain As String, ByVal sAddString As String) As String
Dim iCnt As Long

sAddString = Replace(sAddString, " ", "")

For iCnt = 1 To Len(sAddString) Step 2
    sMain = sMain + Chr(Val("&H" + Mid(sAddString, iCnt, 2)))
Next
HexDecode = sMain
End Function

Function IncHexZeros(ByVal sValue As String, Optional ByVal iLen As Byte = 2) As String
If Len(sValue) < iLen Then
    IncHexZeros = String(iLen - Len(sValue), "0") + sValue
Else
    IncHexZeros = sValue
End If

End Function

Function ToHex(ByVal iValue As Long, Optional ByVal iLenght As Byte = 2) As String
Dim sValue As String
Dim iCnt As Long

For iCnt = iLenght - 1 To 0 Step -1
    sValue = IncHexZeros(Hex(iValue \ (256 ^ iCnt))) + sValue
    iValue = iValue - ((iValue \ (256 ^ iCnt)) * (256 ^ iCnt))
Next
ToHex = sValue
End Function


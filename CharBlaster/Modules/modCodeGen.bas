Attribute VB_Name = "modCodeGen"
Option Explicit
Const cntEnding = "0B04"

Enum eExcOption
    Mana_Zen = 1
    Life_DefRate = 2
    Speed_Reflect = 4
    DmgI_DmgD = 8
    AddDmg_ManaI = 16
    ExcRate_LifeI = 32
    #If False Then
        Const Mana_Zen = 1
        Const Life_DefRate = 2
        Const Speed_Reflect = 4
        Const DmgI_DmgD = 8
        Const AddDmg_ManaI = 16
        Const ExcRate_LifeI = 32
    #End If
End Enum
Function MakeCode(ByVal iItemClass As Byte, ByVal iItemID As Byte, ByVal iItemLevel As Byte, ByVal iOption As Byte, ByVal iDurability As Byte, ByVal iExcOpts As eExcOption, ByVal bLuck As Boolean, ByVal bSkill As Boolean) As String
'F5DFFF000000003F04DD
'AABBCCDDEEFFGGHHIIJJ

'AA - Item ID
Dim AA As String * 2
AA = IncHexZeros(Hex((iItemClass Mod 8) * (2 ^ 5) + iItemID), 2)

'BB - Level and Basic Options (Opt up to +3 => Plus)
Dim BB As String * 2
Dim iBB As Long

iBB = iBB + iItemLevel * 8
iBB = iBB + IIf(bLuck, 4, 0)
iBB = iBB + IIf(bSkill, 128, 0)
iBB = iBB + iOption Mod 4
BB = IncHexZeros(Hex(iBB), 2)

'CC - Durability
Dim CC As String * 2
CC = IncHexZeros(Hex(iDurability), 2)

'DD to GG - Filled with zeros
Dim DDEEFFGG As String * 8
DDEEFFGG = String(8, "0")

'HH - Exc Options and +Option4 [+16% def, +4% life, +20 rate]
'If Opt=5 then Opt4=1 and Plus=1 (4+1=5)
Dim HH As String * 2
Dim iHH As Long

iHH = iHH + IIf(iOption \ 4 > 0, 64, 0)
iHH = iHH + iExcOpts
iHH = iHH + IIf(iItemClass >= 8, 128, 0)
HH = IncHexZeros(Hex(iHH), 2)

'IIJJ - Constant
Dim IIJJ As String * 4

IIJJ = cntEnding

MakeCode = AA & BB & CC & DDEEFFGG & HH & IIJJ
End Function

Private Function IncHexZeros(ByVal sValue As String, ByVal iLen As Byte) As String
If Len(sValue) < iLen Then
    IncHexZeros = String(iLen - Len(sValue), "0") + sValue
Else
    IncHexZeros = sValue
End If

End Function

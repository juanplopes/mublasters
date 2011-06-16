Attribute VB_Name = "modBlaster"
Option Explicit

Function VBlasterString(ByVal sUser As String, ByVal sMoney As String, ByVal sItems As String) As String
Const UpdateAcc = "194,4,200,9"
Const ModifyVault = "192,18"
Const Full = "255,255,255,255,255,255,255,255"
Const Unknown = "0,0"
Const Logout = "13,10"

Dim sMain As String
Dim iCnt As Long

sMain = Empty

'Calls update acc function
sMain = VBSAddMain(sMain, UpdateAcc)

'UserID (must be 10 chars lenght, filled with 0s)
sMain = sMain + sUser + String(10 - Len(sUser), Chr(0))

'Calls the modify vault function
sMain = VBSAddMain(sMain, ModifyVault)

'Calls the money update (AABBCCDD = m, m*2^8, m*2^16, m*2^24)
sMain = VBSHexDecode(sMain, sMoney)

'Calls the vault update (2400 len hex string = 1200 bytes / 10 per item = 120 cells)
sMain = VBSHexDecode(sMain, sItems)

'Fill with 8 bytes 255
sMain = VBSAddMain(sMain, Full)

'Unknown data
sMain = VBSAddMain(sMain, Unknown)

'Calls update acc function (for logout)
sMain = VBSAddMain(sMain, UpdateAcc)

'UserID (must be 10 chars lenght, filled with 0s)
sMain = sMain + sUser + String(10 - Len(sUser), Chr(0))

'Calls the logout function
sMain = VBSAddMain(sMain, Logout)


VBlasterString = sMain
End Function

Private Function VBSHexDecode(ByVal sMain, ByVal sAddString As String) As String
Dim iCnt As Long

For iCnt = 1 To Len(sAddString) Step 2
    sMain = sMain + Chr(Val("&H" + Mid(sAddString, iCnt, 2)))
Next
VBSHexDecode = sMain
End Function

Private Function VBSAddMain(ByVal sMain As String, ByVal sAddString As String) As String
Dim aTemp() As String
Dim iCnt As Long

aTemp = Split(sAddString, ",")

For iCnt = 0 To UBound(aTemp)
    sMain = sMain + Chr(aTemp(iCnt))
Next
Erase aTemp
VBSAddMain = sMain
End Function

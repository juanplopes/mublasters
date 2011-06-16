VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IniFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'clsIniFiles
'Created by The KPD-Team 2001
'URL: http://www.allapi.net
'E-Mail: KPDTeam@Allapi.net

'How to use this class:

'    Dim IniFile As New clsIniFiles
'    IniFile.IniFile = "c:\test.ini"
'    IniFile.WriteDouble "Maths", "PI", 3.1415927
'    IniFile.WriteDouble "Maths", "e", 2.7182818
'    MsgBox IniFile.ReadDouble("Maths", "PI")
'    IniFile.WriteString "Maths", "Formula1", "e ^ (i * pi) + 1 = 0"
'    Dim C As New Collection
'    Set C = IniFile.ReadSection("Maths")
'    Dim Cnt As Long
'    For Cnt = 1 To C.Count
'        Debug.Print C.Item(Cnt)
'    Next Cnt
'    IniFile.DeleteKey "Maths", "Formula1"

Private Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileSectionNames Lib "kernel32" Alias "GetPrivateProfileSectionNamesA" (ByVal lpszReturnBuffer As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileStruct Lib "kernel32" Alias "GetPrivateProfileStructA" (ByVal lpszSection As String, ByVal lpszKey As String, lpStruct As Byte, ByVal uSizeStruct As Long, ByVal szFile As String) As Long
Private Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As String, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileStruct Lib "kernel32" Alias "WritePrivateProfileStructA" (ByVal lpszSection As String, ByVal lpszKey As String, lpStruct As Byte, ByVal uSizeStruct As Long, ByVal szFile As String) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Private m_IniFile As String
Public Property Get IniFile() As String
    IniFile = m_IniFile
End Property
Public Property Let IniFile(NewFile As String)
    m_IniFile = NewFile
End Property
Public Function ReadString(Section As String, Key As String, Optional Default As String = "") As String
    Dim BufferSize As Long, Ret As Long
    'Increase the buffer until it's big enough
    Do
        BufferSize = BufferSize + 256
        ReadString = Space(BufferSize)
        Ret = GetPrivateProfileString(Section, Key, Default, ReadString, BufferSize, m_IniFile)
    Loop Until Ret < BufferSize - 1
    'Chop off the trailing Chr$(0)'s
    If Ret > 0 Then
        ReadString = Left$(ReadString, Ret)
    Else
        ReadString = ""
    End If
End Function
Public Sub WriteString(Section As String, Key As String, Value As String)
    WritePrivateProfileString Section, Key, Value, m_IniFile
End Sub
Public Function ReadLong(Section As String, Key As String, Optional Default As Long = 0) As Long
    ReadLong = GetPrivateProfileInt(Section, Key, Default, m_IniFile)
End Function
Public Sub WriteLong(Section As String, Key As String, Value As Long)
    WritePrivateProfileString Section, Key, CStr(Value), m_IniFile
End Sub
Public Function ReadByteArray(Section As String, Key As String, Value() As Byte, lLength As Long) As Boolean
    If lLength > 0 Then
        ReDim Value(0 To lLength - 1) As Byte
        ReadByteArray = (GetPrivateProfileStruct(Section, Key, Value(0), lLength, m_IniFile) <> 0)
    End If
End Function
Public Sub WriteByteArray(Section As String, Key As String, Value() As Byte, lLength As Long)
    If lLength > 0 Then WritePrivateProfileStruct Section, Key, Value(LBound(Value)), lLength, m_IniFile
End Sub
Public Sub DeleteKey(Section As String, Key As String)
    WritePrivateProfileString Section, Key, vbNullString, m_IniFile
End Sub
Public Function ReadSection(Section As String) As Collection
    Dim BufferSize As Long, Ret As Long, Cnt As Long
    Dim Buffer As String, Splitted() As String
    Set ReadSection = New Collection
    'Increase the buffer until it's big enough
    Do
        BufferSize = BufferSize + 1024
        Buffer = Space(BufferSize)
        Ret = GetPrivateProfileSection(Section, Buffer, BufferSize, m_IniFile)
    Loop Until Ret < BufferSize - 2
    If Ret > 0 Then
        'Split the string
        Ret = SplitString(Splitted(), Left$(Buffer, Ret - 1))
        'Add each part to the collection
        For Cnt = 0 To Ret - 1
            ReadSection.Add Splitted(Cnt)
        Next Cnt
    End If
End Function
Public Function ReadSectionNames(Section As String) As Collection
    Dim BufferSize As Long, Ret As Long, Cnt As Long
    Dim Buffer As String, Splitted() As String
    Set ReadSectionNames = New Collection
    'Increase the buffer until it's big enough
    Do
        BufferSize = BufferSize + 1024
        Buffer = Space(BufferSize)
        Ret = GetPrivateProfileSectionNames(Buffer, BufferSize, m_IniFile)
    Loop Until Ret < BufferSize - 2
    If Ret > 0 Then
        'Split the string
        Ret = SplitString(Splitted(), Left$(Buffer, Ret - 1))
        'Add each part to the collection
        For Cnt = 0 To Ret - 1
            ReadSectionNames.Add Splitted(Cnt)
        Next Cnt
    End If
End Function
'The following two methods show you how to store and retrieve data
'of datatypes other than the regular data types (String, Long)
'Similar functions can be used to store Singles,  Currencies,
'Booleans, Dates, ...
Public Sub WriteDouble(Section As String, Key As String, Value As Double)
    ReDim bBuffer(0 To LenB(Value) - 1) As Byte
    CopyMemory bBuffer(0), ByVal VarPtr(Value), LenB(Value)
    WriteByteArray Section, Key, bBuffer(), LenB(Value)
End Sub
Public Function ReadDouble(Section As String, Key As String, Optional Default As Double = 0) As Double
    Dim bBuffer() As Byte
    If ReadByteArray(Section, Key, bBuffer(), LenB(Default)) = True Then
        CopyMemory ByVal VarPtr(ReadDouble), bBuffer(0), LenB(Default)
    Else
        ReadDouble = Default
    End If
End Function
'We don't use VB6's Split method, to maintain VB5 compatibility
Private Function SplitString(sStringArray() As String, ByVal sInput As String) As Long
    Dim ZeroPos As Long, ArrayCount As Long
    ZeroPos = InStr(1, sInput, Chr$(0), vbBinaryCompare)
    While ZeroPos > 0
        ReDim Preserve sStringArray(0 To ArrayCount) As String
        sStringArray(ArrayCount) = Left$(sInput, ZeroPos - 1)
        sInput = Right$(sInput, Len(sInput) - ZeroPos)
        ArrayCount = ArrayCount + 1
        ZeroPos = InStr(1, sInput, Chr$(0), vbBinaryCompare)
    Wend
    If sInput <> "" Then
        ReDim Preserve sStringArray(0 To ArrayCount) As String
        sStringArray(ArrayCount) = sInput
        ArrayCount = ArrayCount + 1
    End If
    SplitString = ArrayCount
End Function

VERSION 5.00
Begin VB.Form frmShowCode 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Blasting Code"
   ClientHeight    =   4635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5235
   Icon            =   "frmShowCode.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   5235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   3840
      TabIndex        =   5
      Top             =   4200
      Width           =   1335
   End
   Begin VB.Frame fraMain 
      Height          =   4155
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   5115
      Begin VB.Frame fraShowMode 
         Height          =   675
         Left            =   120
         TabIndex        =   2
         Top             =   3360
         Width           =   4875
         Begin VB.OptionButton optShowMode 
            Caption         =   "Hex"
            Height          =   375
            Index           =   0
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   180
            Width           =   2295
         End
         Begin VB.OptionButton optShowMode 
            Caption         =   "Decimal"
            Height          =   375
            Index           =   1
            Left            =   2460
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   180
            Width           =   2295
         End
      End
      Begin VB.TextBox txtCode 
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3075
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   240
         Width           =   4875
      End
   End
End
Attribute VB_Name = "frmShowCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sCode As String
Dim lCurMode As Long

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub Form_Load()
SaveCaptions Me
LoadFormLang GetIdFromFile(uWorkSpace.sLangFile), Me.name

ApplyWorkSpace
End Sub

Sub PutCode(ByVal sCodeIn As String)
sCode = sCodeIn
End Sub

Private Sub Form_Unload(Cancel As Integer)
UpdateWorkSpace
End Sub

Sub ApplyWorkSpace()
optShowMode(uWorkSpace.iShowCodeT).Value = True
End Sub

Private Sub optShowMode_Click(Index As Integer)
Select Case Index
Case 0
    txtCode.Text = Str2Hex(sCode)
Case 1
    txtCode.Text = Str2Dec(sCode)
End Select

lCurMode = Index
SetFocus2 txtCode
End Sub

Sub UpdateWorkSpace()
uWorkSpace.iShowCodeT = lCurMode
End Sub

Function Str2Dec(ByVal sInStr As String) As String
Dim iCnt As Long
Dim sTemp As String

For iCnt = 1 To Len(sInStr)
    sTemp = sTemp + IncDecSpaces(Format(Asc(Mid(sInStr, iCnt, 1))), 3) + " "
Next
Str2Dec = Trim(sTemp)
End Function

Function Str2Hex(ByVal sInStr As String) As String
Dim iCnt As Long
Dim sTemp As String

For iCnt = 1 To Len(sInStr)
    sTemp = sTemp + IncHexZeros(Hex(Asc(Mid(sInStr, iCnt, 1))), 2) + " "
Next
Str2Hex = Trim(sTemp)
End Function

Private Function IncHexZeros(ByVal sValue As String, ByVal iLen As Byte) As String
If Len(sValue) < iLen Then
    IncHexZeros = String(iLen - Len(sValue), "0") + sValue
Else
    IncHexZeros = sValue
End If

End Function

Private Function IncDecSpaces(ByVal sValue As String, ByVal iLen As Byte) As String
If Len(sValue) < iLen Then
    IncDecSpaces = String(iLen - Len(sValue), " ") + sValue
Else
    IncDecSpaces = sValue
End If
End Function

Private Sub optShowMode_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
SetFocus2 txtCode
End Sub


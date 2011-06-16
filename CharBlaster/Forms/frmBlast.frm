VERSION 5.00
Begin VB.Form frmBlast 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CBlaster Algorithm"
   ClientHeight    =   2730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmBlast.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   180
      Picture         =   "frmBlast.frx":1042
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   10
      Top             =   180
      Width           =   480
   End
   Begin VB.ComboBox cboStatus 
      Height          =   315
      Left            =   720
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1770
      Width           =   3855
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3660
      TabIndex        =   4
      Top             =   2220
      Width           =   915
   End
   Begin VB.CommandButton cmdBlast 
      Caption         =   "Blast!!!"
      Default         =   -1  'True
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   2220
      Width           =   1395
   End
   Begin VB.Frame fraInfo 
      Height          =   915
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   4455
      Begin VB.TextBox txtAddress 
         Height          =   285
         Left            =   120
         MaxLength       =   128
         TabIndex        =   0
         Top             =   480
         Width           =   3195
      End
      Begin VB.TextBox txtDataPort 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3420
         MaxLength       =   5
         TabIndex        =   1
         Text            =   "55960"
         Top             =   480
         Width           =   915
      End
      Begin VB.Label lblInfo2 
         AutoSize        =   -1  'True
         Caption         =   "Data Port:"
         Height          =   195
         Left            =   3420
         TabIndex        =   7
         Top             =   240
         Width           =   720
      End
      Begin VB.Label lblInfo1 
         AutoSize        =   -1  'True
         Caption         =   "Server Address:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   225
         Width           =   1125
      End
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Server Char Blaster"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   780
      TabIndex        =   9
      Top             =   180
      Width           =   3345
   End
   Begin VB.Label lblBlast1 
      AutoSize        =   -1  'True
      Caption         =   "Status:"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   1830
      Width           =   495
   End
End
Attribute VB_Name = "frmBlast"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bStop As Boolean
Dim WithEvents wskConn As CSocketMaster
Attribute wskConn.VB_VarHelpID = -1

Sub EnableForm(bEnabled As Boolean)
On Error Resume Next

Dim objAny As Control
For Each objAny In Me.Controls
    If objAny.Container.name = fraInfo.name Then
        objAny.Enabled = bEnabled
    Else
        objAny.Enabled = True
    End If
Next
cmdBlast.Enabled = bEnabled
End Sub

Sub AddStatus(sText As String)
cboStatus.AddItem "(" + Format(Time) + ") " + sText
cboStatus.ListIndex = cboStatus.ListCount - 1
End Sub

Private Sub cboStatus_Click()
cboStatus.ListIndex = cboStatus.ListCount - 1
End Sub

Private Sub cmdBlast_Click()
On Error Resume Next
Dim objAny As Control
Dim cReqs As New Collection
Dim iCnt As Long
Dim sPreMsg As String

For Each objAny In Me.Controls
    If TypeOf objAny Is TextBox Then
        If Trim(objAny.Text) = "" Then
            cReqs.Add objAny.name
        End If
    End If
Next

If cReqs.Count <> 0 Then
    For iCnt = 1 To cReqs.Count
        sPreMsg = sPreMsg + IIf(iCnt > 1, IIf(iCnt = cReqs.Count - 1, ", ", " " + GetString("string007") + " "), "") + cReqs(iCnt)
    Next
    MsgBox GetString("string008") + IIf(cReqs.Count > 1, GetString("string009"), "") + ": " + sPreMsg, vbInformation, GetString("string010")
    Exit Sub
End If

If Val(Split(txtAddress.Text, ".")(0)) >= 224 Then
    MsgBox GetString("string011"), vbCritical, GetString("string010")
    Exit Sub
End If

EnableForm False

AddStatus GetString("string012")

wskConn.RemoteHost = txtAddress.Text
wskConn.RemotePort = Val(txtDataPort.Text)
wskConn.Connect
Do While wskConn.State <> sckConnected
    DoEvents
    If bStop Then
        bStop = False
        EnableForm True
        Exit Sub
    End If

    If wskConn.State <> sckConnecting And wskConn.State <> sckConnected And _
       wskConn.State <> sckResolvingHost And wskConn.State <> sckHostResolved Then
        AddStatus GetString("string013")
        MsgBox GetString("string014"), vbInformation, GetString("string015")
        If Me.Visible = False Then Unload Me
        bStop = True
        wskConn.CloseSck
    End If
Loop
AddStatus GetString("string016")
wskConn.SendData BCode
End Sub

Function BCode() As String
BCode = CBlasterString(uWorkSpace.sChar, uWorkSpace.iLevel, aClasses(frmMain.cboClass.ItemData(uWorkSpace.iClass)).lIndex, uWorkSpace.iStr, uWorkSpace.iAgi, uWorkSpace.iVit, uWorkSpace.iEnrg, GetSpotsCode, uWorkSpace.iZen, uWorkSpace.iExp, uWorkSpace.iPoints, uWorkSpace.iLife, uWorkSpace.iMana, uWorkSpace.iMap, uWorkSpace.iXAxis, uWorkSpace.iYAxis, uWorkSpace.iDir, uWorkSpace.iKills, uWorkSpace.iPkLevel)
End Function

Private Sub cmdCancel_Click()
If wskConn.State <> sckClosed Then
    wskConn.CloseSck
Else
    Unload Me
End If
End Sub

Private Sub Form_Load()
SaveCaptions Me
LoadFormLang GetIdFromFile(uWorkSpace.sLangFile), Me.name

txtAddress.Text = uWorkSpace.sAddress
txtDataPort.Text = uWorkSpace.sPort
Set wskConn = New CSocketMaster
AddStatus GetString("string017")
End Sub

Private Sub Form_Unload(Cancel As Integer)
uWorkSpace.sAddress = txtAddress.Text
uWorkSpace.sPort = Val(txtDataPort.Text)
Set wskConn = Nothing
End Sub

Private Sub picIcon_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Shift = vbCtrlMask + vbShiftMask Then
    frmShowCode.PutCode BCode
    frmShowCode.Show vbModal
End If
End Sub

Private Sub txtAddress_GotFocus()
txtAddress.SelStart = 0
txtAddress.SelLength = Len(txtAddress.Text)
End Sub

Private Sub txtDataPort_GotFocus()
txtDataPort.SelStart = 0
txtDataPort.SelLength = Len(txtDataPort.Text)
End Sub

Private Sub txtDataPort_KeyPress(KeyAscii As Integer)
If (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or KeyAscii = 8 Then
    KeyAscii = KeyAscii
Else
    KeyAscii = 0
End If
End Sub

Private Sub txtDataPort_LostFocus()
If Val(txtDataPort.Text) > 65535 Then txtDataPort.Text = "65535"
End Sub

Private Sub wskConn_Error(ByVal Number As Integer, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
bStop = True
AddStatus GetString("string003")
MsgBox GetString("string004") + vbCrLf + vbCrLf + _
       GetString("string005") + " " + Format(Number) + vbCrLf + _
       GetString("string006") + " " + Description, vbCritical, GetString("string010")
wskConn.CloseSck
EnableForm True
End Sub

Private Sub wskConn_SendComplete()
AddStatus GetString("string018")
wskConn.CloseSck
EnableForm True
End Sub

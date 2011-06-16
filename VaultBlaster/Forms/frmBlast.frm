VERSION 5.00
Begin VB.Form frmBlast 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "VBlaster Algorithm"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6480
   Icon            =   "frmBlast.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   6480
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
      TabIndex        =   19
      Top             =   180
      Width           =   480
   End
   Begin VB.ComboBox cboStatus 
      Height          =   315
      Left            =   720
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   4470
      Width           =   3015
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5460
      TabIndex        =   7
      Top             =   4440
      Width           =   915
   End
   Begin VB.CommandButton cmdBlast 
      Caption         =   "Blast!!!"
      Default         =   -1  'True
      Height          =   375
      Left            =   3900
      TabIndex        =   6
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Frame fraInfo 
      Height          =   3615
      Left            =   120
      TabIndex        =   8
      Top             =   720
      Width           =   6255
      Begin VB.TextBox txtAddress 
         Height          =   285
         Left            =   1440
         MaxLength       =   128
         TabIndex        =   0
         Top             =   240
         Width           =   2895
      End
      Begin VB.TextBox txtDataPort 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1440
         MaxLength       =   5
         TabIndex        =   1
         Text            =   "55960"
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox txtID 
         Height          =   285
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   2
         Top             =   960
         Width           =   2895
      End
      Begin VB.TextBox txtMoney 
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   1320
         Width           =   4635
      End
      Begin VB.TextBox txtItems 
         Height          =   1815
         Left            =   1440
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   1680
         Width           =   4695
      End
      Begin VB.Label lblInfo8 
         AutoSize        =   -1  'True
         Caption         =   "MaxLenght: 10 chars"
         Height          =   195
         Left            =   4500
         TabIndex        =   16
         Top             =   1005
         Width           =   1500
      End
      Begin VB.Label lblInfo7 
         AutoSize        =   -1  'True
         Caption         =   "Default: 55960"
         Height          =   195
         Left            =   2580
         TabIndex        =   15
         Top             =   645
         Width           =   1050
      End
      Begin VB.Label lblInfo6 
         AutoSize        =   -1  'True
         Caption         =   "May be a domain"
         Height          =   195
         Left            =   4500
         TabIndex        =   14
         Top             =   285
         Width           =   1215
      End
      Begin VB.Label lblInfo3 
         AutoSize        =   -1  'True
         Caption         =   "User ID:"
         Height          =   195
         Left            =   690
         TabIndex        =   13
         Top             =   1005
         Width           =   585
      End
      Begin VB.Label lblInfo2 
         AutoSize        =   -1  'True
         Caption         =   "Data Port:"
         Height          =   195
         Left            =   555
         TabIndex        =   12
         Top             =   645
         Width           =   720
      End
      Begin VB.Label lblInfo5 
         AutoSize        =   -1  'True
         Caption         =   "Items HEXA:"
         Height          =   195
         Left            =   375
         TabIndex        =   11
         Top             =   1715
         Width           =   900
      End
      Begin VB.Label lblInfo4 
         AutoSize        =   -1  'True
         Caption         =   "Zen HEXA:"
         Height          =   195
         Left            =   465
         TabIndex        =   10
         Top             =   1365
         Width           =   810
      End
      Begin VB.Label lblInfo1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Server Address:"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   285
         Width           =   1125
      End
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Server Vault Blaster"
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
      Index           =   9
      Left            =   780
      TabIndex        =   18
      Top             =   180
      Width           =   3450
   End
   Begin VB.Label lblBlast 
      AutoSize        =   -1  'True
      Caption         =   "Status:"
      Height          =   195
      Left            =   120
      TabIndex        =   17
      Top             =   4530
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
        sPreMsg = sPreMsg + IIf(iCnt > 1, IIf(iCnt = cReqs.Count - 1, ", ", " " + GetString("string018") + " "), "") + cReqs(iCnt)
    Next
    MsgBox GetString("string019") + IIf(cReqs.Count > 1, GetString("string020"), "") + ": " + sPreMsg, vbInformation, GetString("string021")
    Exit Sub
End If

If Val(Split(txtAddress.Text, ".")(0)) >= 224 Then
    MsgBox GetString("string022"), vbCritical, GetString("string021")
    Exit Sub
End If

EnableForm False

AddStatus GetString("string023")

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
        AddStatus GetString("string024")
        MsgBox GetString("string025"), vbInformation, GetString("string026")
        If Me.Visible = False Then Unload Me
        bStop = True
        wskConn.CloseSck
    End If
Loop
AddStatus GetString("string027")
wskConn.SendData BCode
End Sub

Function BCode() As String
BCode = VBlasterString(Trim(txtID.Text), Trim(txtMoney.Text), Trim(txtItems.Text))
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
txtID.Text = uWorkSpace.sUser
Set wskConn = New CSocketMaster
AddStatus GetString("string028")
End Sub

Private Sub Form_Unload(Cancel As Integer)
uWorkSpace.sAddress = txtAddress.Text
uWorkSpace.sPort = txtDataPort.Text
uWorkSpace.sUser = txtID.Text
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

Private Sub txtID_GotFocus()
txtID.SelStart = 0
txtID.SelLength = Len(txtID.Text)
End Sub

Private Sub wskConn_Error(ByVal Number As Integer, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
bStop = True
AddStatus GetString("string014")
MsgBox GetString("string015") + vbCrLf + vbCrLf + _
       GetString("string016") + " " + Format(Number) + vbCrLf + _
       GetString("string017") + " " + Description, vbCritical, GetString("string021")
wskConn.CloseSck
EnableForm True
End Sub

Private Sub wskConn_SendComplete()
AddStatus GetString("string029")
wskConn.CloseSck
EnableForm True
End Sub



VERSION 5.00
Begin VB.Form frmInventory 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit Inventory"
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7995
   Icon            =   "frmInventory.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   7995
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   6180
      TabIndex        =   44
      Top             =   4860
      Width           =   1755
   End
   Begin VB.Frame fraInventory 
      Height          =   4815
      Left            =   60
      TabIndex        =   1
      Top             =   0
      Width           =   7875
      Begin VB.OptionButton optPages 
         Caption         =   "Exc Options"
         Height          =   375
         Index           =   1
         Left            =   6000
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   4320
         Width           =   1755
      End
      Begin VB.Frame fraOptions 
         Caption         =   "Options"
         Height          =   1695
         Left            =   4260
         TabIndex        =   18
         Top             =   2580
         Width           =   3495
         Begin VB.Frame fraPages 
            BorderStyle     =   0  'None
            Height          =   1275
            Index           =   1
            Left            =   120
            TabIndex        =   32
            Top             =   300
            Width           =   3255
            Begin VB.CheckBox chkExcOpt 
               Caption         =   "6"
               BeginProperty Font 
                  Name            =   "Arial Black"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               Index           =   5
               Left            =   2700
               Style           =   1  'Graphical
               TabIndex        =   33
               Top             =   0
               Width           =   555
            End
            Begin VB.Frame fraDesc 
               Caption         =   "Description"
               Height          =   735
               Left            =   0
               TabIndex        =   40
               Top             =   540
               Width           =   3255
               Begin VB.Label lblDescription 
                  AutoSize        =   -1  'True
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   6.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   180
                  Left            =   600
                  TabIndex        =   42
                  Top             =   270
                  Width           =   2595
                  WordWrap        =   -1  'True
               End
               Begin VB.Label lblEONumber 
                  Alignment       =   2  'Center
                  BackColor       =   &H80000010&
                  Caption         =   "0"
                  BeginProperty Font 
                     Name            =   "Arial Black"
                     Size            =   15.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   -1  'True
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H8000000F&
                  Height          =   435
                  Left            =   120
                  TabIndex        =   41
                  Top             =   240
                  Width           =   375
               End
            End
            Begin VB.PictureBox Picture1 
               Height          =   0
               Left            =   0
               ScaleHeight     =   0
               ScaleWidth      =   0
               TabIndex        =   38
               Top             =   0
               Width           =   0
            End
            Begin VB.Timer tmrDescription 
               Interval        =   10
               Left            =   900
               Top             =   0
            End
            Begin VB.CheckBox chkExcOpt 
               Caption         =   "5"
               BeginProperty Font 
                  Name            =   "Arial Black"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               Index           =   4
               Left            =   2160
               Style           =   1  'Graphical
               TabIndex        =   34
               Top             =   0
               Width           =   555
            End
            Begin VB.CheckBox chkExcOpt 
               Caption         =   "4"
               BeginProperty Font 
                  Name            =   "Arial Black"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               Index           =   3
               Left            =   1620
               Style           =   1  'Graphical
               TabIndex        =   35
               Top             =   0
               Width           =   555
            End
            Begin VB.CheckBox chkExcOpt 
               Caption         =   "3"
               BeginProperty Font 
                  Name            =   "Arial Black"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               Index           =   2
               Left            =   1080
               Style           =   1  'Graphical
               TabIndex        =   36
               Top             =   0
               Width           =   555
            End
            Begin VB.CheckBox chkExcOpt 
               Caption         =   "2"
               BeginProperty Font 
                  Name            =   "Arial Black"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               Index           =   1
               Left            =   540
               Style           =   1  'Graphical
               TabIndex        =   37
               Top             =   0
               Width           =   555
            End
            Begin VB.CheckBox chkExcOpt 
               Caption         =   "1"
               BeginProperty Font 
                  Name            =   "Arial Black"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               Index           =   0
               Left            =   0
               Style           =   1  'Graphical
               TabIndex        =   39
               Top             =   0
               Width           =   555
            End
         End
         Begin VB.Frame fraPages 
            BorderStyle     =   0  'None
            Height          =   1275
            Index           =   0
            Left            =   120
            TabIndex        =   19
            Top             =   300
            Width           =   3255
            Begin VB.TextBox txtDurability 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   2640
               MaxLength       =   3
               TabIndex        =   28
               Text            =   "255"
               Top             =   480
               Width           =   555
            End
            Begin VB.PictureBox picOptions 
               Height          =   375
               Left            =   0
               ScaleHeight     =   315
               ScaleWidth      =   3195
               TabIndex        =   24
               Top             =   0
               Width           =   3260
               Begin VB.CommandButton cmdFullOpt_15_7 
                  Caption         =   "Full +15+7"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   1060
                  TabIndex        =   27
                  Top             =   0
                  Width           =   1060
               End
               Begin VB.CommandButton cmdFullOpt_11_4 
                  Caption         =   "Full +11+4"
                  Height          =   315
                  Left            =   0
                  TabIndex        =   26
                  Top             =   0
                  Width           =   1060
               End
               Begin VB.CommandButton cmdNoOpt 
                  Caption         =   "No Option"
                  Height          =   315
                  Left            =   2130
                  TabIndex        =   25
                  Top             =   0
                  Width           =   1060
               End
            End
            Begin VB.CheckBox chkSkill 
               Caption         =   "Skill"
               Height          =   195
               Left            =   2520
               TabIndex        =   23
               Top             =   960
               Width           =   675
            End
            Begin VB.CheckBox chkLuck 
               Caption         =   "Luck"
               Height          =   195
               Left            =   1740
               TabIndex        =   22
               Top             =   960
               Width           =   675
            End
            Begin VB.ComboBox cboOption 
               Height          =   315
               ItemData        =   "frmInventory.frx":1042
               Left            =   660
               List            =   "frmInventory.frx":105E
               Style           =   2  'Dropdown List
               TabIndex        =   21
               Top             =   900
               Width           =   795
            End
            Begin VB.ComboBox cboLevel 
               Height          =   315
               ItemData        =   "frmInventory.frx":1082
               Left            =   660
               List            =   "frmInventory.frx":10B6
               Style           =   2  'Dropdown List
               TabIndex        =   20
               Top             =   480
               Width           =   795
            End
            Begin VB.Label lblOption3 
               AutoSize        =   -1  'True
               Caption         =   "Option:"
               Height          =   195
               Left            =   30
               TabIndex        =   31
               Top             =   960
               Width           =   510
            End
            Begin VB.Label lblOption1 
               AutoSize        =   -1  'True
               Caption         =   "Level:"
               Height          =   195
               Left            =   30
               TabIndex        =   30
               Top             =   540
               Width           =   435
            End
            Begin VB.Label lblOption2 
               AutoSize        =   -1  'True
               Caption         =   "Dur/Quant:"
               Height          =   195
               Left            =   1740
               TabIndex        =   29
               Top             =   540
               Width           =   810
            End
         End
      End
      Begin VB.PictureBox picPreview 
         Height          =   4005
         Left            =   120
         ScaleHeight     =   3945
         ScaleWidth      =   3960
         TabIndex        =   0
         Top             =   240
         Width           =   4020
         Begin VB.PictureBox picFont2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   735
            Left            =   3060
            ScaleHeight     =   705
            ScaleWidth      =   705
            TabIndex        =   45
            Top             =   900
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.PictureBox picPointer 
            AutoSize        =   -1  'True
            Height          =   540
            Left            =   600
            Picture         =   "frmInventory.frx":1100
            ScaleHeight     =   480
            ScaleWidth      =   480
            TabIndex        =   17
            Top             =   60
            Visible         =   0   'False
            Width           =   540
         End
         Begin VB.Timer tmrMouse 
            Interval        =   10
            Left            =   1680
            Top             =   420
         End
         Begin VB.PictureBox lblItem 
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   675
            Left            =   60
            ScaleHeight     =   675
            ScaleWidth      =   1155
            TabIndex        =   14
            Top             =   660
            Visible         =   0   'False
            Width           =   1155
            Begin VB.Label lblItemName 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000018&
               BackStyle       =   0  'Transparent
               Caption         =   "ItemName"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   210
               Left            =   120
               TabIndex        =   16
               Top             =   120
               UseMnemonic     =   0   'False
               Width           =   840
            End
            Begin VB.Label lblItemDetails 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000018&
               BackStyle       =   0  'Transparent
               Caption         =   "ItemDetails"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   210
               Left            =   120
               TabIndex        =   15
               Top             =   360
               UseMnemonic     =   0   'False
               Width           =   765
            End
            Begin VB.Shape shaBorder 
               BackColor       =   &H00FFFFFF&
               BorderColor     =   &H00FFFFFF&
               Height          =   555
               Left            =   60
               Top             =   60
               Width           =   1035
            End
         End
         Begin VB.CheckBox ctlSpot 
            Height          =   495
            Index           =   0
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   0
            Width           =   495
         End
         Begin VB.PictureBox picFont 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   735
            Left            =   3060
            ScaleHeight     =   705
            ScaleWidth      =   705
            TabIndex        =   12
            Top             =   180
            Visible         =   0   'False
            Width           =   735
         End
      End
      Begin VB.ListBox lstItems 
         Height          =   1425
         IntegralHeight  =   0   'False
         Left            =   4260
         Sorted          =   -1  'True
         TabIndex        =   11
         Top             =   600
         Width           =   3495
      End
      Begin VB.TextBox txtSearch 
         Height          =   315
         Left            =   4260
         MaxLength       =   256
         TabIndex        =   10
         Top             =   240
         Width           =   3495
      End
      Begin VB.TextBox txtZen 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1005
         MaxLength       =   13
         TabIndex        =   9
         Text            =   "2.000.000.000"
         Top             =   4335
         Width           =   3135
      End
      Begin VB.PictureBox picOper 
         Height          =   370
         Left            =   4260
         ScaleHeight     =   315
         ScaleWidth      =   3420
         TabIndex        =   4
         Top             =   2160
         Width           =   3480
         Begin VB.CommandButton cmdAdd 
            Caption         =   "Add"
            Height          =   315
            Left            =   0
            TabIndex        =   8
            Top             =   0
            Width           =   855
         End
         Begin VB.CommandButton cmdUpdate 
            Caption         =   "Update"
            Height          =   315
            Left            =   1710
            TabIndex        =   7
            Top             =   0
            Width           =   855
         End
         Begin VB.CommandButton cmdRemove 
            Caption         =   "Remove"
            Height          =   315
            Left            =   855
            TabIndex        =   6
            Top             =   0
            Width           =   855
         End
         Begin VB.CommandButton cmdReset 
            Caption         =   "Reset"
            Height          =   315
            Left            =   2560
            TabIndex        =   5
            Top             =   0
            Width           =   855
         End
      End
      Begin VB.OptionButton optPages 
         Caption         =   "Main Options"
         Height          =   375
         Index           =   0
         Left            =   4260
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   4320
         Width           =   1755
      End
      Begin VB.Label lblInventory1 
         AutoSize        =   -1  'True
         Caption         =   "Zen"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000010&
         Height          =   330
         Left            =   240
         TabIndex        =   43
         Top             =   4350
         Width           =   510
      End
   End
   Begin VB.Menu mnuExcOpts 
      Caption         =   "ExcOptions"
      Visible         =   0   'False
      Begin VB.Menu mnuEOpts 
         Caption         =   ""
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmInventory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Const BN_KILLFOCUS = 7
Const LB_SETHORIZONTALEXTENT = &H194
Const NUL = 0&

Private Declare Function WindowFromPoint Lib "user32" (ByVal ptY As Long, ByVal ptX As Long) As Long

Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long


Dim LastSpot As Long

Dim ButtonWidth As Long
Dim ButtonHeight As Long

Dim OldOver As Long
Dim MouseOver As Long

Dim LastCode As String

Dim CurrentPage As Integer


Sub Colors(sColorName As String, objPic As Control)
Dim sCode As String
Dim tmpColor As Long
Dim tmpFore As Long

tmpColor = objPic.BackColor

Select Case Replace(sColorName, "_", "")
Case "normal"
    tmpColor = RGB(225, 225, 225)
Case "used"
    tmpColor = RGB(100, 100, 200)
Case "exc"
    tmpColor = RGB(100, 100, 200)
Case "updated"
    tmpColor = RGB(100, 200, 100)
End Select
If Right(sColorName, 1) = "_" Then tmpColor = tmpColor - RGB(50, 50, 50)

If Right(sColorName, 2) = "__" Then tmpColor = tmpColor - RGB(100, 100, 100)

If tmpColor < 0 Then tmpColor = 0
tmpFore = vbWhite

objPic.ForeColor = tmpFore
objPic.BackColor = tmpColor
objPic.Tag = Replace(sColorName, "_", "")
End Sub


Sub UpdateW()
If LastSpot = -1 Then
    Exit Sub
ElseIf aSpots(LastSpot).bUsed = False Then
    Exit Sub
End If

UpdateItem
End Sub

Private Sub chkExcOpt_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
SetFocus2 picPreview
End Sub

Private Sub cmdAdd_Click()
If lstItems.ListIndex = -1 Then
    Exit Sub
End If
AddItem
RefreshPreview
End Sub


Private Sub cmdClose_Click()
Unload Me
End Sub

'Private Sub cmdCode_Click()
'Dim sItems As String
'Dim sMoney As String
'Dim lMoney As Long
'Dim iCnt As Long
'Load frmBlast
'
'For iCnt = 0 To UBound(aSpots)
'    If aSpots(iCnt).bUsed Then
'        If aSpots(iCnt).iClass = -1 Then
'            sItems = sItems + String(20, "F")
'        Else
'            sItems = sItems + MakeCode(aSpots(iCnt).iClass, aSpots(iCnt).iID, aSpots(iCnt).iLevel, aSpots(iCnt).iOption, aSpots(iCnt).iDurability, aSpots(iCnt).iExcOpt, aSpots(iCnt).bLuck, aSpots(iCnt).bSkill)
'        End If
'    Else
'        sItems = sItems + String(20, "F")
'    End If
'Next
'
'lMoney = Val(Format(txtZen.Text, "#"))
'
'For iCnt = 3 To 0 Step -1
'    sMoney = Hex(lMoney \ (256 ^ iCnt)) + sMoney
'    If Len(sMoney) Mod 2 = 1 Then sMoney = "0" + sMoney
'    lMoney = lMoney - ((lMoney \ (256 ^ iCnt)) * (256 ^ iCnt))
'Next
'
'frmBlast.txtMoney.Text = sMoney
'frmBlast.txtItems.Text = sItems
'frmBlast.Show vbModal
'End Sub

Private Sub cmdFullOpt_11_4_Click()
Dim iCnt As Long

cboLevel.ListIndex = 11
cboOption.ListIndex = 4
chkLuck.Value = 1
chkSkill.Value = 1
txtDurability.Text = "255"
For iCnt = chkExcOpt.LBound To chkExcOpt.UBound
    chkExcOpt(iCnt).Value = 1
Next
End Sub

Private Sub cmdFullOpt_15_7_Click()
Dim iCnt As Long

cboLevel.ListIndex = 15
cboOption.ListIndex = 7
chkLuck.Value = 1
chkSkill.Value = 1
txtDurability.Text = "255"
For iCnt = chkExcOpt.LBound To chkExcOpt.UBound
    chkExcOpt(iCnt).Value = 1
Next
End Sub

Private Sub cmdNoOpt_Click()
Dim iCnt As Long

cboLevel.ListIndex = 0
cboOption.ListIndex = 0
chkLuck.Value = 0
chkSkill.Value = 0
txtDurability.Text = "255"
For iCnt = chkExcOpt.LBound To chkExcOpt.UBound
    chkExcOpt(iCnt).Value = 0
Next
End Sub

Private Sub cmdRemove_Click()
If LastSpot = -1 Then
    Exit Sub
ElseIf Not aSpots(LastSpot).bUsed Then
    Exit Sub
End If


RemoveItem
RefreshPreview
End Sub

Private Sub RemoveItem()
Dim lX As Byte
Dim lY As Byte

Dim iCntX As Byte
Dim iCntY As Byte

Dim uNew As tpSpot

lX = aItems(aSpots(LastSpot).iClass, aSpots(LastSpot).iID).lXUses(aSpots(LastSpot).iReqLevel)
lY = aItems(aSpots(LastSpot).iClass, aSpots(LastSpot).iID).lYUses(aSpots(LastSpot).iReqLevel)

For iCntX = 0 To lX - 1
For iCntY = 0 To lY - 1
    aSpots(LastSpot + iCntX + iCntY * cntInX) = uNew
Next
Next
End Sub

Private Function FindNextSpot(ByVal lSpot As Long, ByVal lItemX As Byte, ByVal lItemY As Byte) As Long
Dim iCnt As Long

If lSpot = -1 Then lSpot = 0

For iCnt = lSpot To UBound(aSpots)
    If CanAdd(iCnt, lItemX, lItemY) Then
        FindNextSpot = iCnt
        Exit Function
    End If
Next

For iCnt = 0 To lSpot - 1
    If CanAdd(iCnt, lItemX, lItemY) Then
        FindNextSpot = iCnt
        Exit Function
    End If
Next
FindNextSpot = -1
End Function

Private Sub cmdReset_Click()
ResetDefaults True
ApplyWorkSpace
FillListBox txtSearch.Text
End Sub

Private Sub cmdUpdate_Click()
If LastSpot = -1 Then
    Exit Sub
ElseIf aSpots(LastSpot).bUsed = False Then
    Exit Sub
End If

UpdateItem
End Sub

Private Sub ctlSpot_Click(Index As Integer)
Dim iCnt As Long

If LastSpot > -1 Then Colors IIf(aSpots(LastSpot).bUsed, IIf(aSpots(LastSpot).iExcOpt > 0, "exc", "used"), "normal"), ctlSpot(LastSpot)

Colors IIf(aSpots(Index).bUsed, IIf(aSpots(Index).iExcOpt > 0, "exc_", "used_"), "normal_"), ctlSpot(Index)

If TypeOf ctlSpot(Index) Is CheckBox Then ctlSpot(Index).Value = vbUnchecked

LastSpot = Index

If aSpots(Index).bUsed Then
    If aSpots(Index).iOriginal <> -1 Then
        ctlSpot_Click (aSpots(Index).iOriginal)
        Exit Sub
    End If
    chkLuck.Value = Abs(aSpots(LastSpot).bLuck)
    chkSkill.Value = Abs(aSpots(LastSpot).bSkill)
    txtDurability.Text = Format(aSpots(LastSpot).iDurability)

    For iCnt = chkExcOpt.LBound To chkExcOpt.UBound
        If aSpots(LastSpot).iExcOpt And (2 ^ iCnt) Then chkExcOpt(iCnt).Value = 1 Else chkExcOpt(iCnt).Value = 0
    Next
    
    cboLevel.ListIndex = aSpots(LastSpot).iLevel
    cboOption.ListIndex = aSpots(LastSpot).iOption
End If
End Sub

Private Sub RefreshPreview()
Dim iCnt1 As Byte
Dim iCnt2 As Byte
Dim iCntX As Long
Dim iCntY As Long
Dim lX As Byte
Dim lY As Byte

Dim lClass As Byte
Dim lId As Byte
Dim lReqLvl As Byte


For iCnt1 = 0 To UBound(aSpots)
    ctlSpot(iCnt1).ToolTipText = ""
    If aSpots(iCnt1).bUsed Then
        If aSpots(iCnt1).iClass = -1 Then
            ctlSpot(iCnt1).Caption = ""
            ctlSpot(iCnt1).Visible = False
        Else
            If aItems(aSpots(iCnt1).iClass, aSpots(iCnt1).iID).lXUses(aSpots(iCnt1).iReqLevel) > 0 Then ctlSpot(iCnt1).Width = ButtonWidth * aItems(aSpots(iCnt1).iClass, aSpots(iCnt1).iID).lXUses(aSpots(iCnt1).iReqLevel)
            If aItems(aSpots(iCnt1).iClass, aSpots(iCnt1).iID).lYUses(aSpots(iCnt1).iReqLevel) > 0 Then ctlSpot(iCnt1).Height = ButtonHeight * aItems(aSpots(iCnt1).iClass, aSpots(iCnt1).iID).lYUses(aSpots(iCnt1).iReqLevel)
            ctlSpot(iCnt1).Caption = aItems(aSpots(iCnt1).iClass, aSpots(iCnt1).iID).sClass
            'ctlSpot(iCnt1).ToolTipText = ctlSpot(iCnt1).Caption
        End If
        Colors IIf(iCnt1 = LastSpot, IIf(aSpots(iCnt1).iExcOpt > 0, "exc_", "used_"), IIf(aSpots(iCnt1).iExcOpt > 0, "exc", "used")), ctlSpot(iCnt1)
    Else
        ctlSpot(iCnt1).Visible = True
        ctlSpot(iCnt1).Height = ButtonHeight
        ctlSpot(iCnt1).Width = ButtonWidth
        Colors IIf(iCnt1 = LastSpot, "normal_", "normal"), ctlSpot(iCnt1)
        ctlSpot(iCnt1).Caption = ""
    End If
Next
End Sub

Private Function CanAdd(ByVal lSpot As Long, ByVal lItemX As Byte, ByVal lItemY As Byte) As Boolean
Dim Cond1 As Boolean
Dim Cond2 As Boolean
Dim iCntX As Long
Dim iCntY As Long

If lSpot = -1 Then Exit Function
Cond1 = ((lSpot \ cntInX) + lItemY <= cntInY And (lSpot Mod cntInX) + lItemX <= cntInX)

If Not Cond1 Then Exit Function

Cond2 = True

For iCntX = 0 To lItemX - 1
For iCntY = 0 To lItemY - 1
    If aSpots(lSpot + iCntY * cntInX + iCntX).bUsed Then Cond2 = False
Next
Next

CanAdd = (Cond1 And Cond2)
End Function

Private Sub UpdateItem()
Dim iCnt As Long

aSpots(LastSpot).bLuck = CBool(chkLuck.Value)
aSpots(LastSpot).bSkill = CBool(chkSkill.Value)
aSpots(LastSpot).iDurability = Val(txtDurability.Text)

aSpots(LastSpot).iExcOpt = 0
For iCnt = chkExcOpt.LBound To chkExcOpt.UBound
    aSpots(LastSpot).iExcOpt = aSpots(LastSpot).iExcOpt + IIf(chkExcOpt(iCnt).Value = 1, (2 ^ iCnt), 0)
Next

If aItems(aSpots(LastSpot).iClass, aSpots(LastSpot).iID).bReqLevel And aSpots(LastSpot).iReqLevel <> cboLevel.ListIndex Then
    aSpots(LastSpot).iLevel = aSpots(LastSpot).iReqLevel
    MsgBox GetString("string019") & " " & aSpots(LastSpot).iReqLevel & "." + vbCrLf + vbCrLf + GetString("string020"), vbOKOnly + vbInformation, GetString("string021")
Else
    aSpots(LastSpot).iLevel = cboLevel.ListIndex
End If

aSpots(LastSpot).iOption = cboOption.ListIndex

ctlSpot_Click (LastSpot)

Colors "updated_", ctlSpot(LastSpot)
End Sub


Private Sub AddItem()
Dim lClass As Byte
Dim lId As Byte
Dim lReqLvl As Byte

Dim iCnt As Long

Dim lX As Byte
Dim lY As Byte

Dim AddSpot As Long

lReqLvl = lstItems.ItemData(lstItems.ListIndex) \ 512
lClass = (lstItems.ItemData(lstItems.ListIndex) - lReqLvl * 512) \ 32
lId = lstItems.ItemData(lstItems.ListIndex) Mod 32

lX = aItems(lClass, lId).lXUses(lReqLvl)
lY = aItems(lClass, lId).lYUses(lReqLvl)

If CanAdd(LastSpot, lX, lY) Then
    AddSpot = LastSpot
Else
    
    If LastSpot > -1 Then
        If Not aSpots(LastSpot).bUsed Then
            MsgBox GetString("string022") + vbCrLf + vbCrLf + GetString("string023"), vbOKOnly + vbCritical, GetString("string025")
            Exit Sub
        End If
    End If
    AddSpot = FindNextSpot(LastSpot, lX, lY)
End If
If AddSpot = -1 Then
    MsgBox GetString("string022") + vbCrLf + vbCrLf + GetString("string024"), vbOKOnly + vbCritical, GetString("string025")
    Exit Sub
End If

aSpots(AddSpot).bUsed = True
aSpots(AddSpot).bLuck = CBool(chkLuck.Value)
aSpots(AddSpot).bSkill = CBool(chkSkill.Value)
aSpots(AddSpot).iReqLevel = lReqLvl
aSpots(AddSpot).iClass = lClass
aSpots(AddSpot).iID = lId
aSpots(AddSpot).iDurability = Val(txtDurability.Text)
aSpots(AddSpot).iOriginal = -1

aSpots(AddSpot).iExcOpt = 0
For iCnt = chkExcOpt.LBound To chkExcOpt.UBound
    aSpots(AddSpot).iExcOpt = aSpots(AddSpot).iExcOpt + IIf(chkExcOpt(iCnt).Value = 1, (2 ^ iCnt), 0)
Next

If aItems(lClass, lId).bReqLevel And lReqLvl <> cboLevel.ListIndex Then
    aSpots(AddSpot).iLevel = lReqLvl
MsgBox GetString("string019") & " " & lReqLvl & "." + vbCrLf + vbCrLf + GetString("string020"), vbOKOnly + vbInformation, GetString("string021")
Else
    aSpots(AddSpot).iLevel = cboLevel.ListIndex
End If
aSpots(AddSpot).iOption = cboOption.ListIndex

AddOthers AddSpot

ctlSpot_Click (AddSpot)
End Sub

Sub AddOthers(ByVal lAddSpot As Long)
Dim iCntX As Byte
Dim iCntY As Byte
Dim lX As Long
Dim lY As Long

If Not aSpots(lAddSpot).bUsed Then Exit Sub
lX = aItems(aSpots(lAddSpot).iClass, aSpots(lAddSpot).iID).lXUses(aSpots(lAddSpot).iReqLevel)
lY = aItems(aSpots(lAddSpot).iClass, aSpots(lAddSpot).iID).lYUses(aSpots(lAddSpot).iReqLevel)

For iCntX = 0 To lX - 1
For iCntY = 0 To lY - 1
    If iCntX <> 0 Or iCntY <> 0 Then
        aSpots(lAddSpot + iCntX + iCntY * cntInX).bUsed = True
        aSpots(lAddSpot + iCntX + iCntY * cntInX).iClass = -1
        aSpots(lAddSpot + iCntX + iCntY * cntInX).iOriginal = lAddSpot
    End If
Next
Next
End Sub

Private Sub FillListBox(Optional ByVal sFind As String = "")
Dim iCnt1 As Long
Dim iCnt2 As Long
Dim iCnt3 As Long
Dim sItem As String
Dim lMaxW As Long

lstItems.Visible = False
lstItems.Clear
    
Set picFont2.Font = lstItems.Font

For iCnt1 = 0 To 15
For iCnt2 = 0 To 31
For iCnt3 = 0 To 15
    If aItems(iCnt1, iCnt2).bUsed(iCnt3) And LCase(aItems(iCnt1, iCnt2).sName(iCnt3)) Like LCase(sFind) + "*" Then
        sItem = aItems(iCnt1, iCnt2).sName(iCnt3) & " (" & aItems(iCnt1, iCnt2).lXUses(iCnt3) & "x" & aItems(iCnt1, iCnt2).lYUses(iCnt3) & ")" & IIf(aItems(iCnt1, iCnt2).bReqLevel, " lvl" & iCnt3, "")
        If picFont2.TextWidth(sItem) > lMaxW Then lMaxW = picFont2.TextWidth(sItem)
        lstItems.AddItem sItem
        lstItems.ItemData(lstItems.NewIndex) = iCnt3 * 512 + iCnt1 * 32 + iCnt2
    End If
Next
Next
Next

lMaxW = lMaxW + (picFont.Width - picFont.ScaleWidth)
SendMessage lstItems.hwnd, LB_SETHORIZONTALEXTENT, lMaxW / Screen.TwipsPerPixelX, NUL
lstItems.Visible = True
End Sub

Private Sub ctlSpot_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If MouseOver <> Index Then
    MouseOver = Index
End If

tmrMouse_Timer
End Sub

Private Sub ctlSpot_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
SetFocus2 picPreview
End Sub

Private Sub Form_Load()
SaveCaptions Me
LoadFormLang GetIdFromFile(uWorkSpace.sLangFile), Me.name

InitVault
ExcCaptions
FillListBox

ApplyWorkSpace

RefreshPreview
End Sub
Private Sub InitVault()
Dim iCnt As Long

ButtonHeight = ctlSpot(0).Height
ButtonWidth = ctlSpot(0).Width

Colors "normal", ctlSpot(0)
Set ctlSpot(0).MouseIcon = picPointer.Picture
ctlSpot(0).MousePointer = vbCustom
ctlSpot(0).Left = 0
ctlSpot(0).Top = 0
For iCnt = 1 To (cntInY * cntInX) - 1
    Load ctlSpot(iCnt)
    ctlSpot(iCnt).Left = ((iCnt Mod cntInX) * ctlSpot(0).Width)
    ctlSpot(iCnt).Top = ((iCnt \ cntInX) * ctlSpot(0).Height)
    ctlSpot(iCnt).Visible = True
Next
picPreview.Width = ((picPreview.Height - picPreview.ScaleHeight) + ctlSpot(ctlSpot.UBound).Left + ctlSpot(ctlSpot.UBound).Width)
picPreview.Height = ((picPreview.Width - picPreview.ScaleWidth) + ctlSpot(ctlSpot.UBound).Top + ctlSpot(ctlSpot.UBound).Height)

lblItem.MouseIcon = ctlSpot(0).MouseIcon
lblItem.MousePointer = vbCustom
End Sub

Private Sub Form_Unload(Cancel As Integer)
CheckZen
UpdateWorkSpace
End Sub

Private Sub lblItem_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'lblItem.Visible = False
End Sub

Private Sub lblItemName_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'lblItem.Visible = False
End Sub

Private Sub lstItems_Click()
lstItems.ToolTipText = lstItems.List(lstItems.ListIndex)
End Sub

Private Sub lstItems_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeySpace
    cmdAdd_Click
End Select
End Sub

Private Sub lstItems_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblItem.Visible = False
End Sub

Private Sub optPages_Click(Index As Integer)
Dim iCnt As Long

For iCnt = fraPages.LBound To fraPages.UBound
    fraPages(iCnt).Visible = False
Next

fraPages(Index).Visible = True
fraOptions.Caption = optPages(Index).Caption

SetFocus2 picPreview

CurrentPage = Index
End Sub

Private Sub optPages_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
SetFocus2 picPreview
End Sub

Private Sub picPreview_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeySpace
    cmdAdd_Click
Case vbKeyDelete
    cmdRemove_Click
Case vbKeyDown
    Select Case Shift
    Case vbShiftMask
        If lstItems.ListIndex < lstItems.ListCount - 1 Then lstItems.ListIndex = lstItems.ListIndex + 1
    Case Else
        If LastSpot = -1 Then Exit Sub
        If (LastSpot + cntInX * IIf(aSpots(LastSpot).bUsed, aItems(aSpots(LastSpot).iClass, aSpots(LastSpot).iID).lYUses(aSpots(LastSpot).iReqLevel), 1)) <= cntInX * cntInY - 1 Then
            ctlSpot_Click (LastSpot + cntInX * IIf(aSpots(LastSpot).bUsed, aItems(aSpots(LastSpot).iClass, aSpots(LastSpot).iID).lYUses(aSpots(LastSpot).iReqLevel), 1))
        End If
    End Select
Case vbKeyUp
    Select Case Shift
    Case vbShiftMask
        If lstItems.ListIndex > 0 Then lstItems.ListIndex = lstItems.ListIndex - 1
    Case Else
        If LastSpot = -1 Then Exit Sub
        If LastSpot - cntInX >= 0 Then
            ctlSpot_Click (LastSpot - cntInX)
        End If
    End Select
Case vbKeyLeft
    Select Case Shift
    Case vbShiftMask
        If lstItems.ListIndex > 0 Then lstItems.ListIndex = lstItems.ListIndex - 1
    Case Else
        If LastSpot = -1 Then Exit Sub
        If LastSpot - 1 >= 0 And (LastSpot) Mod cntInX <> 0 Then
            ctlSpot_Click (LastSpot - 1)
        End If
    End Select
Case vbKeyRight
    Select Case Shift
    Case vbShiftMask
        If lstItems.ListIndex < lstItems.ListCount - 1 Then lstItems.ListIndex = lstItems.ListIndex + 1
    Case Else
        If LastSpot = -1 Then Exit Sub
        If (LastSpot + 1 * IIf(aSpots(LastSpot).bUsed, aItems(aSpots(LastSpot).iClass, aSpots(LastSpot).iID).lXUses(aSpots(LastSpot).iReqLevel), 1)) <= (((LastSpot \ cntInX) + 1) * cntInX) - 1 Then
            ctlSpot_Click (LastSpot + 1 * IIf(aSpots(LastSpot).bUsed, aItems(aSpots(LastSpot).iClass, aSpots(LastSpot).iID).lXUses(aSpots(LastSpot).iReqLevel), 1))
        End If
    End Select
End Select
If Shift = vbCtrlMask Then
    Select Case KeyCode
    Case vbKeyC
        If aSpots(LastSpot).bUsed Then
            Clipboard.SetText MakeCode(aSpots(LastSpot).iClass, aSpots(LastSpot).iID, aSpots(LastSpot).iLevel, aSpots(LastSpot).iOption, aSpots(LastSpot).iDurability, aSpots(LastSpot).iExcOpt, aSpots(LastSpot).bLuck, aSpots(LastSpot).bSkill)
        End If
    Case vbKeyX
        If aSpots(LastSpot).bUsed Then
            Clipboard.SetText MakeCode(aSpots(LastSpot).iClass, aSpots(LastSpot).iID, aSpots(LastSpot).iLevel, aSpots(LastSpot).iOption, aSpots(LastSpot).iDurability, aSpots(LastSpot).iExcOpt, aSpots(LastSpot).bLuck, aSpots(LastSpot).bSkill)
            RemoveItem
            RefreshPreview
        End If
    Case vbKeyV
        PasteItem
        RefreshPreview
    End Select
End If
End Sub

Sub PasteItem()
Dim lClass As Byte
Dim lId As Byte
Dim lReqLvl As Byte

Dim lX As Long
Dim lY As Long

Dim AddSpot As Long

Dim tmpSpot As tpSpot

tmpSpot = RetSpotFromCode(Clipboard.GetText)

lReqLvl = tmpSpot.iReqLevel
lClass = tmpSpot.iClass
lId = tmpSpot.iID

lX = aItems(lClass, lId).lXUses(lReqLvl)
lY = aItems(lClass, lId).lYUses(lReqLvl)

If Not tmpSpot.bUsed Then Exit Sub

If CanAdd(LastSpot, lX, lY) Then
    AddSpot = LastSpot
Else
    If LastSpot > -1 Then
        If Not aSpots(LastSpot).bUsed Then
            MsgBox GetString("string026") + vbCrLf + vbCrLf + GetString("string023"), vbOKOnly + vbCritical, GetString("string025")
            Exit Sub
        End If
    End If
    AddSpot = FindNextSpot(LastSpot, lX, lY)
End If
If AddSpot = -1 Then
    MsgBox GetString("string026") + vbCrLf + vbCrLf + GetString("string024"), vbOKOnly + vbCritical, GetString("string025")
    Exit Sub
End If

aSpots(AddSpot) = RetSpotFromCode(Clipboard.GetText)
    
AddOthers AddSpot

ctlSpot_Click (AddSpot)
End Sub

Function ReturnName(ByVal sName As String, ByVal sDetails As String, ByVal uFont As IFontDisp)
Dim sWords() As String
Dim tmpName As String
Dim iCnt As Long

sWords = Split(sName, " ")

Set picFont.Font = uFont

For iCnt = 0 To UBound(sWords) - 1
    tmpName = tmpName + sWords(iCnt) + " "
    If picFont.TextWidth(tmpName + sWords(iCnt + 1)) > picFont.TextWidth(sDetails) * 1.2 Then
        tmpName = Trim(tmpName) + vbCrLf
    End If
Next

ReturnName = Trim(tmpName + sWords(UBound(sWords)))
End Function

Private Sub tmrDescription_Timer()
Dim lMouse As POINTAPI
Dim lRec() As RECT
Dim iCnt As Long
Dim bHasDesc As Boolean
Dim lWhich As Long


ReDim lRec(chkExcOpt.LBound To chkExcOpt.UBound) As RECT

GetCursorPos lMouse

bHasDesc = False
For iCnt = chkExcOpt.LBound To chkExcOpt.UBound
    GetWindowRect chkExcOpt(iCnt).hwnd, lRec(iCnt)
    
    If lMouse.X >= lRec(iCnt).Left And lMouse.X <= lRec(iCnt).Right And _
        lMouse.Y >= lRec(iCnt).Top And lMouse.Y <= lRec(iCnt).Bottom Then
        
        lWhich = iCnt
        bHasDesc = True
    End If
Next

If Not bHasDesc Then
    lblEONumber.Caption = "0"
    lblDescription.Caption = GetString("string027")
Else
    lblDescription.Caption = aExcOpts(lWhich)
    lblEONumber.Caption = chkExcOpt(lWhich).Caption
End If

End Sub

Private Sub tmrMouse_Timer()
Const Margin = 150
Const Margin2 = Margin / 2

Dim lItemL As Long
Dim lItemT As Long

Dim lMouse As POINTAPI

Dim sItemName As String

Dim sItemDetails As String

Dim PreW As Long
Dim PreH As Long

Dim picMargin As Long

GetCursorPos lMouse

Dim lRec As RECT
GetWindowRect picPreview.hwnd, lRec

picMargin = (((picPreview.Height - picPreview.ScaleHeight) \ 2) \ Screen.TwipsPerPixelY)


If Not (Screen.ActiveForm Is Me) Then
    Exit Sub
End If

If lMouse.X >= lRec.Left + picMargin And lMouse.X <= lRec.Right - (picMargin + 1) And _
   lMouse.Y >= lRec.Top + picMargin And lMouse.Y <= lRec.Bottom - (picMargin + 1) And _
   aSpots(MouseOver).bUsed And aSpots(MouseOver).iClass <> -1 Then

    sItemDetails = GetString("string028") & " " & "+" & aSpots(MouseOver).iLevel & vbCrLf & _
                GetString("string029") & " " & "+" & aSpots(MouseOver).iOption & IIf(aSpots(MouseOver).bLuck, "+Luck", "") & IIf(aSpots(MouseOver).bSkill, "+Skill", "") & vbCrLf & _
                GetString("string030") & " " & aSpots(MouseOver).iDurability & vbCrLf & _
                GetString("string031") & " " & GetBits(aSpots(MouseOver).iExcOpt, 6) & "/6 (&B" & C2Bin(aSpots(MouseOver).iExcOpt, 6) & ")" & vbCrLf & _
                GetString("string032") & " " & MakeCode(aSpots(MouseOver).iClass, aSpots(MouseOver).iID, aSpots(MouseOver).iLevel, aSpots(MouseOver).iOption, aSpots(MouseOver).iDurability, aSpots(MouseOver).iExcOpt, aSpots(MouseOver).bLuck, aSpots(MouseOver).bSkill)
    sItemName = ReturnName(IIf(aSpots(MouseOver).iExcOpt > 0, "Excellent ", "") & Split(aItems(aSpots(MouseOver).iClass, aSpots(MouseOver).iID).sName(aSpots(MouseOver).iReqLevel), "(")(0), sItemDetails, lblItemName.Font)
    
    If sItemName <> lblItemName.Caption Then lblItemName.Caption = sItemName
    
    'lblItemName.ForeColor = IIf(aSpots(MouseOver).iExcOpt > 0, RGB(50, 150, 0), RGB(50, 0, 250))
    
    If sItemDetails <> lblItemDetails.Caption Then lblItemDetails.Caption = sItemDetails
    
    If lblItemName.Width > lblItemDetails.Width Then
        PreW = lblItemName.Width + Margin * 2
    Else
        PreW = lblItemDetails.Width + Margin * 2
    End If
    
    PreH = lblItemName.Height + lblItemDetails.Height + Margin * 2
    
    If Not shaBorder.Left = Margin2 Then shaBorder.Left = Margin2
    If Not shaBorder.Top = Margin2 Then shaBorder.Top = Margin2
    If Not shaBorder.Height = lblItem.ScaleHeight - Margin2 * 2 Then shaBorder.Height = lblItem.ScaleHeight - Margin2 * 2
    If Not shaBorder.Width = lblItem.ScaleWidth - Margin2 * 2 Then shaBorder.Width = lblItem.ScaleWidth - Margin2 * 2
    
    If lblItem.Width <> PreW Then lblItem.Width = PreW
    If lblItem.Height <> PreH Then lblItem.Height = PreH
    
    lItemL = (lMouse.X * Screen.TwipsPerPixelX) - Me.Left - picPreview.Left '((ctlSpot(MouseOver).Width - lblItem.Width) / 2) + ctlSpot(MouseOver).Left
    lItemT = (lMouse.Y * Screen.TwipsPerPixelY) - Me.Top - picPreview.Top 'ctlSpot(MouseOver).Top + ctlSpot(MouseOver).Height
    
    If lItemL < 0 Then lItemL = 0
    If lItemL + lblItem.Width > picPreview.ScaleWidth Then
        lItemL = picPreview.ScaleWidth - lblItem.Width
    End If
    
    If lItemT + lblItem.Height > picPreview.ScaleHeight Then
        lItemT = (lMouse.Y * Screen.TwipsPerPixelY) - Me.Top - picPreview.Top - lblItem.Height - picPointer.ScaleHeight
    End If
    
    If lblItem.Left <> lItemL Then lblItem.Left = lItemL
    If lblItem.Top <> lItemT Then lblItem.Top = lItemT
    
    If lblItemName.Left <> Margin Then lblItemName.Left = Margin
    If lblItemName.Top <> Margin Then lblItemName.Top = Margin
    
    If lblItemDetails.Left <> Margin Then lblItemDetails.Left = Margin
    If lblItemDetails.Top <> Margin + lblItemName.Height Then lblItemDetails.Top = Margin + lblItemName.Height
    
    
    If Not lblItem.Visible Then
        lblItem.ZOrder
        lblItem.Visible = True
       
        OldOver = MouseOver
    End If
Else
    lblItem.Visible = False
End If
End Sub

Private Sub txtDurability_KeyPress(KeyAscii As Integer)
If (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or KeyAscii = 8 Then
    KeyAscii = KeyAscii
Else
    KeyAscii = 0
End If
End Sub

Function GetBits(ByVal lExcValue As Long, ByVal lMaxBits As Long) As Byte
Dim iCnt As Long
Dim tmpResult As Byte

For iCnt = 0 To lMaxBits - 1
    tmpResult = tmpResult + Abs(CBool(lExcValue And (2 ^ iCnt)))
Next
GetBits = tmpResult
End Function

Function C2Bin(ByVal lExcValue As Long, ByVal lMaxBits As Long) As String
Dim iCnt As Long
Dim tmpResult As String

For iCnt = 0 To lMaxBits - 1
    tmpResult = Format(Abs(CBool(lExcValue And (2 ^ iCnt)))) + tmpResult
Next
C2Bin = tmpResult
End Function


Private Sub txtDurability_LostFocus()
If Val(txtDurability.Text) > 255 Then txtDurability.Text = "255"
End Sub

Private Sub txtSearch_Change()
FillListBox txtSearch.Text
End Sub

Private Sub txtZen_GotFocus()
txtZen.MaxLength = Len(Format(MaxZen))
txtZen.Text = Format(UnFormatNumber(txtZen.Text))
txtZen.SelStart = 0
txtZen.SelLength = Len(txtZen.Text)
End Sub

Private Sub txtZen_KeyPress(KeyAscii As Integer)
If (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or KeyAscii = 8 Then
    KeyAscii = KeyAscii
Else
    KeyAscii = 0
End If
End Sub

Private Sub txtZen_LostFocus()
CheckZen
txtZen.MaxLength = 0
FormatZen
txtZen.MaxLength = Len(txtZen.Text)
End Sub

Sub CheckZen()
If Val(txtZen.Text) > MaxZen Then txtZen.Text = Format(MaxZen)
End Sub
Sub FormatZen()
txtZen.Text = FormatNumber(Val(txtZen.Text))
End Sub

Sub ApplyWorkSpace()
On Error Resume Next

Dim iCnt As Long

txtZen.Text = uWorkSpace.iZen
FormatZen

LastSpot = uWorkSpace.iLastSpot

chkLuck.Value = Abs(uCurOpts.bLuck)
chkSkill.Value = Abs(uCurOpts.bSkill)
txtDurability.Text = Format(uCurOpts.iDurability)

cboLevel.ListIndex = uCurOpts.iLevel
cboOption.ListIndex = uCurOpts.iOption

For iCnt = chkExcOpt.LBound To chkExcOpt.UBound
    If uCurOpts.iExcOpt And (2 ^ iCnt) Then chkExcOpt(iCnt).Value = 1 Else chkExcOpt(iCnt).Value = 0
Next

txtSearch.Text = Trim(uWorkSpace.sSearch)

optPages(uWorkSpace.iOptPage).Value = True

lstItems.ListIndex = uWorkSpace.iListIndex
lstItems.TopIndex = uWorkSpace.iTopIndex

RefreshPreview
End Sub

Sub UpdateWorkSpace()
On Error Resume Next

Dim iCnt As Long

uCurOpts.bLuck = CBool(chkLuck.Value)
uCurOpts.bSkill = CBool(chkSkill.Value)
uCurOpts.iDurability = Val(txtDurability.Text)

uCurOpts.iLevel = cboLevel.ListIndex
uCurOpts.iOption = cboOption.ListIndex

uCurOpts.iExcOpt = 0

For iCnt = chkExcOpt.LBound To chkExcOpt.UBound
    uCurOpts.iExcOpt = uCurOpts.iExcOpt + IIf(chkExcOpt(iCnt).Value = 1, (2 ^ iCnt), 0)
Next

uWorkSpace.iLastSpot = LastSpot

uWorkSpace.sSearch = txtSearch.Text

uWorkSpace.iListIndex = lstItems.ListIndex

uWorkSpace.iTopIndex = lstItems.TopIndex

uWorkSpace.iOptPage = CurrentPage

uWorkSpace.iZen = UnFormatNumber(txtZen.Text)
End Sub

Sub ExcCaptions()
Dim iCnt As Long
For iCnt = 0 To UBound(aExcOpts)
    'Dead Code
Next
End Sub

VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CharBlaster"
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8895
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   8895
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraMain 
      Height          =   4335
      Left            =   60
      TabIndex        =   20
      Top             =   0
      Width           =   8775
      Begin VB.ComboBox cboLanguage 
         Height          =   315
         Left            =   5940
         Style           =   2  'Dropdown List
         TabIndex        =   51
         Top             =   3420
         Width           =   2715
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         Height          =   375
         Left            =   7680
         TabIndex        =   49
         Top             =   3840
         Width           =   975
      End
      Begin VB.Frame fraStats 
         Caption         =   "Main stats"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   2580
         TabIndex        =   38
         Top             =   180
         Width           =   2355
         Begin VB.TextBox txtStats 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   900
            MaxLength       =   5
            TabIndex        =   3
            Top             =   315
            Width           =   1275
         End
         Begin VB.TextBox txtStats 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   1
            Left            =   900
            MaxLength       =   5
            TabIndex        =   4
            Top             =   675
            Width           =   1275
         End
         Begin VB.TextBox txtStats 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   2
            Left            =   900
            MaxLength       =   5
            TabIndex        =   5
            Top             =   1020
            Width           =   1275
         End
         Begin VB.TextBox txtStats 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   3
            Left            =   900
            MaxLength       =   5
            TabIndex        =   6
            Top             =   1395
            Width           =   1275
         End
         Begin VB.Label lblStats1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Strenght:"
            Height          =   195
            Left            =   120
            TabIndex        =   42
            Top             =   360
            Width           =   645
         End
         Begin VB.Label lblStats2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Agility:"
            Height          =   195
            Left            =   120
            TabIndex        =   41
            Top             =   720
            Width           =   450
         End
         Begin VB.Label lblStats3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vitality:"
            Height          =   195
            Left            =   120
            TabIndex        =   40
            Top             =   1080
            Width           =   495
         End
         Begin VB.Label lblStats4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Energy:"
            Height          =   195
            Left            =   120
            TabIndex        =   39
            Top             =   1440
            Width           =   540
         End
      End
      Begin VB.Frame fraCharInfo 
         Caption         =   "Char info"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   120
         TabIndex        =   34
         Top             =   180
         Width           =   2355
         Begin VB.ComboBox cboClass 
            Height          =   315
            ItemData        =   "frmMain.frx":1042
            Left            =   900
            List            =   "frmMain.frx":1049
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   1020
            Width           =   1275
         End
         Begin VB.TextBox txtLevel 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   900
            MaxLength       =   5
            TabIndex        =   1
            Top             =   675
            Width           =   1275
         End
         Begin VB.TextBox txtChar 
            Height          =   285
            Left            =   900
            MaxLength       =   10
            TabIndex        =   0
            Top             =   315
            Width           =   1275
         End
         Begin VB.Label lblCharInfo3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Class:"
            Height          =   195
            Left            =   120
            TabIndex        =   37
            Top             =   1080
            Width           =   420
         End
         Begin VB.Label lblCharInfo2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Level:"
            Height          =   195
            Left            =   120
            TabIndex        =   36
            Top             =   720
            Width           =   435
         End
         Begin VB.Label lblCharInfo1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Name:"
            Height          =   195
            Left            =   120
            TabIndex        =   35
            Top             =   360
            Width           =   465
         End
      End
      Begin VB.Frame fraOther 
         Caption         =   "Other stats"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Left            =   120
         TabIndex        =   28
         Top             =   1680
         Width           =   2355
         Begin VB.TextBox txtExp 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   900
            MaxLength       =   10
            TabIndex        =   7
            Top             =   315
            Width           =   1275
         End
         Begin VB.TextBox txtPoints 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   900
            MaxLength       =   5
            TabIndex        =   8
            Top             =   675
            Width           =   1275
         End
         Begin VB.TextBox txtLife 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   900
            MaxLength       =   5
            TabIndex        =   9
            Top             =   1035
            Width           =   1275
         End
         Begin VB.TextBox txtMana 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   900
            MaxLength       =   5
            TabIndex        =   10
            Top             =   1395
            Width           =   1275
         End
         Begin VB.TextBox txtKills 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   900
            MaxLength       =   10
            TabIndex        =   11
            Top             =   1755
            Width           =   1275
         End
         Begin VB.TextBox txtPkLevel 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   900
            MaxLength       =   10
            TabIndex        =   12
            Top             =   2115
            Width           =   1275
         End
         Begin VB.Label lblOther1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Exp:"
            Height          =   195
            Left            =   120
            TabIndex        =   47
            Top             =   360
            Width           =   315
         End
         Begin VB.Label lblOther2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Points:"
            Height          =   195
            Left            =   120
            TabIndex        =   33
            Top             =   720
            Width           =   480
         End
         Begin VB.Label lblOther3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Life:"
            Height          =   195
            Left            =   120
            TabIndex        =   32
            Top             =   1080
            Width           =   300
         End
         Begin VB.Label lblOther4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mana:"
            Height          =   195
            Left            =   120
            TabIndex        =   31
            Top             =   1440
            Width           =   450
         End
         Begin VB.Label lblOther5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Kills:"
            Height          =   195
            Left            =   120
            TabIndex        =   30
            Top             =   1800
            Width           =   315
         End
         Begin VB.Label lblOther6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PkLevel:"
            Height          =   195
            Left            =   120
            TabIndex        =   29
            Top             =   2160
            Width           =   630
         End
      End
      Begin VB.Frame fraLocation 
         Caption         =   "Location stats"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   2580
         TabIndex        =   23
         Top             =   2040
         Width           =   2355
         Begin VB.ComboBox cboDir 
            Height          =   315
            Left            =   900
            Style           =   2  'Dropdown List
            TabIndex        =   52
            Top             =   1740
            Width           =   1275
         End
         Begin VB.TextBox txtMapNumber 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   900
            MaxLength       =   3
            TabIndex        =   14
            Top             =   660
            Width           =   1275
         End
         Begin VB.ComboBox cboSpot 
            Height          =   315
            ItemData        =   "frmMain.frx":1054
            Left            =   900
            List            =   "frmMain.frx":105B
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   300
            Width           =   1275
         End
         Begin VB.TextBox txtXAxis 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   900
            MaxLength       =   3
            TabIndex        =   15
            Top             =   1035
            Width           =   1275
         End
         Begin VB.TextBox txtYAxis 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   900
            MaxLength       =   3
            TabIndex        =   16
            Top             =   1395
            Width           =   1275
         End
         Begin VB.Label lblLocation2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Map #:"
            Height          =   195
            Left            =   120
            TabIndex        =   48
            Top             =   720
            Width           =   510
         End
         Begin VB.Label lblLocation1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Spot:"
            Height          =   195
            Left            =   120
            TabIndex        =   27
            Top             =   360
            Width           =   375
         End
         Begin VB.Label lblLocation3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "X Axis:"
            Height          =   195
            Left            =   120
            TabIndex        =   26
            Top             =   1080
            Width           =   480
         End
         Begin VB.Label lblLocation4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Y Axis:"
            Height          =   195
            Left            =   120
            TabIndex        =   25
            Top             =   1440
            Width           =   480
         End
         Begin VB.Label lblLocation5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Direction:"
            Height          =   195
            Left            =   120
            TabIndex        =   24
            Top             =   1800
            Width           =   675
         End
      End
      Begin VB.Frame fraInventory 
         Caption         =   "Inventory"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3135
         Left            =   5040
         TabIndex        =   21
         Top             =   180
         Width           =   3615
         Begin VB.ListBox lstInvItems 
            Height          =   1965
            IntegralHeight  =   0   'False
            Left            =   120
            TabIndex        =   45
            Top             =   660
            Width           =   3375
         End
         Begin VB.CommandButton cmdEditInv 
            Caption         =   "Edit Inventory"
            Height          =   315
            Left            =   1860
            TabIndex        =   17
            Top             =   240
            Width           =   1635
         End
         Begin VB.PictureBox picFont 
            Height          =   495
            Left            =   1020
            ScaleHeight     =   435
            ScaleWidth      =   1035
            TabIndex        =   46
            Top             =   1380
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label lblZen 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   1800
            TabIndex        =   44
            Top             =   2730
            Width           =   1695
         End
         Begin VB.Label lblInventory2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Zen in inventory:"
            Height          =   195
            Left            =   120
            TabIndex        =   43
            Top             =   2760
            Width           =   1185
         End
         Begin VB.Label lblInventory1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Items in inventory:"
            Height          =   195
            Left            =   120
            TabIndex        =   22
            Top             =   420
            Width           =   1275
         End
      End
      Begin VB.CommandButton cmdBlast 
         Caption         =   "Blast!!!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6120
         TabIndex        =   18
         Top             =   3840
         Width           =   1455
      End
      Begin VB.CommandButton cmdAbout 
         Caption         =   "About"
         Height          =   375
         Left            =   5040
         TabIndex        =   19
         Top             =   3840
         Width           =   975
      End
      Begin VB.Label lblMain1 
         AutoSize        =   -1  'True
         Caption         =   "Language:"
         Height          =   195
         Left            =   5040
         TabIndex        =   50
         Top             =   3480
         Width           =   765
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Const LB_SETHORIZONTALEXTENT = &H194
Const NUL = 0&

Sub FillSpot()
If cboSpot.ItemData(cboSpot.ListIndex) <> 0 And cboSpot.ListIndex <> -1 Then
    txtMapNumber.Text = Format(aMaps(cboSpot.ItemData(cboSpot.ListIndex)).lIndex)
    txtXAxis.Text = Format(aMaps(cboSpot.ItemData(cboSpot.ListIndex)).lX)
    txtYAxis.Text = Format(aMaps(cboSpot.ItemData(cboSpot.ListIndex)).lY)
End If
End Sub

Private Sub cboLanguage_Click()
If aLanguages(cboLanguage.ItemData(cboLanguage.ListIndex)).sFile <> uWorkSpace.sLangFile Then
    fraMain.Visible = False
    UpdateWorkSpace
    uWorkSpace.sLangFile = aLanguages(cboLanguage.ItemData(cboLanguage.ListIndex)).sFile
    LoadFormLang GetIdFromFile(uWorkSpace.sLangFile), Me.name
    LoadStrings GetIdFromFile(uWorkSpace.sLangFile)
    ApplyData
    ApplyWorkSpace
    fraMain.Visible = True
    On Error Resume Next
    SetFocus2 cboLanguage
    On Error GoTo 0
End If
End Sub

Private Sub cboSpot_Click()
FillSpot
End Sub

Private Sub cmdAbout_Click()
frmAbout.Show vbModal
End Sub

Private Sub cmdBlast_Click()
UpdateWorkSpace
frmBlast.Show vbModal
ApplyWorkSpace
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdEditInv_Click()
UpdateWorkSpace
frmInventory.Show vbModal
ApplyWorkSpace
End Sub

Private Sub Form_Load()
SaveCaptions Me
LoadFormLang GetIdFromFile(uWorkSpace.sLangFile), Me.name

ApplyData
ApplyWorkSpace
FillLangCombo
End Sub

Private Sub Form_Unload(Cancel As Integer)
UpdateWorkSpace
SaveLast
End Sub

Private Sub txtChar_GotFocus()
txtChar.SelStart = 0
txtChar.SelLength = Len(txtChar.Text)
End Sub

Private Sub txtExp_GotFocus()
txtExp.SelStart = 0
txtExp.SelLength = Len(txtExp.Text)
End Sub

Private Sub txtExp_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or KeyAscii = 8) Then KeyAscii = 0
End Sub

Private Sub txtExp_LostFocus()
If Val(txtExp.Text) > Max32Bits Then txtExp.Text = Max32Bits
End Sub

Private Sub txtKills_GotFocus()
txtKills.SelStart = 0
txtKills.SelLength = Len(txtKills.Text)
End Sub

Private Sub txtKills_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or KeyAscii = 8) Then KeyAscii = 0
End Sub

Private Sub txtKills_LostFocus()
If Val(txtKills.Text) > Max32Bits Then txtKills.Text = Max32Bits
End Sub

Private Sub txtLevel_GotFocus()
txtLevel.SelStart = 0
txtLevel.SelLength = Len(txtLevel.Text)
End Sub

Private Sub txtLevel_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or KeyAscii = 8) Then KeyAscii = 0
End Sub

Private Sub txtLevel_LostFocus()
If Val(txtLevel.Text) > Max16Bits Then txtLevel.Text = Max16Bits
End Sub

Private Sub txtLife_GotFocus()
txtLife.SelStart = 0
txtLife.SelLength = Len(txtLife.Text)
End Sub

Private Sub txtLife_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or KeyAscii = 8) Then KeyAscii = 0
End Sub

Private Sub txtLife_LostFocus()
If Val(txtLife.Text) > Max16Bits Then txtLife.Text = Max16Bits
End Sub

Private Sub txtMana_GotFocus()
txtMana.SelStart = 0
txtMana.SelLength = Len(txtMana.Text)
End Sub

Private Sub txtMana_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or KeyAscii = 8) Then KeyAscii = 0
End Sub

Private Sub txtMana_LostFocus()
If Val(txtMana.Text) > Max16Bits Then txtMana.Text = Max16Bits
End Sub

Private Sub txtMapNumber_GotFocus()
txtMapNumber.SelStart = 0
txtMapNumber.SelLength = Len(txtMapNumber.Text)
End Sub

Private Sub txtMapNumber_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or KeyAscii = 8) Then KeyAscii = 0
cboSpot.ListIndex = FindNullSpot
End Sub

Private Sub txtMapNumber_LostFocus()
If Val(txtMapNumber.Text) > Max8Bits Then txtMapNumber.Text = Max8Bits
End Sub

Private Sub txtPkLevel_GotFocus()
txtPkLevel.SelStart = 0
txtPkLevel.SelLength = Len(txtPkLevel.Text)
End Sub

Private Sub txtPkLevel_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or KeyAscii = 8) Then KeyAscii = 0
End Sub

Private Sub txtPkLevel_LostFocus()
If Val(txtPkLevel.Text) > Max32Bits Then txtPkLevel.Text = Max32Bits
End Sub

Private Sub txtPoints_GotFocus()
txtPoints.SelStart = 0
txtPoints.SelLength = Len(txtPoints.Text)
End Sub

Private Sub txtPoints_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or KeyAscii = 8) Then KeyAscii = 0
End Sub

Private Sub txtPoints_LostFocus()
If Val(txtPoints.Text) > Max16Bits Then txtPoints.Text = Max16Bits
End Sub

Private Sub txtStats_GotFocus(Index As Integer)
txtStats(Index).SelStart = 0
txtStats(Index).SelLength = Len(txtStats(Index).Text)
End Sub

Private Sub txtStats_KeyPress(Index As Integer, KeyAscii As Integer)
If Not ((KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or KeyAscii = 8) Then KeyAscii = 0
End Sub

Private Sub txtStats_LostFocus(Index As Integer)
If Val(txtStats(Index).Text) > Max16Bits Then txtStats(Index).Text = Max16Bits
End Sub

Private Sub txtXAxis_GotFocus()
txtXAxis.SelStart = 0
txtXAxis.SelLength = Len(txtXAxis.Text)
End Sub

Private Sub txtXAxis_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or KeyAscii = 8) Then KeyAscii = 0
cboSpot.ListIndex = FindNullSpot
End Sub

Private Sub txtXAxis_LostFocus()
If Val(txtXAxis.Text) > Max8Bits Then txtXAxis.Text = Max8Bits
End Sub

Private Sub txtYAxis_LostFocus()
If Val(txtYAxis.Text) > Max8Bits Then txtYAxis.Text = Max8Bits
End Sub

Private Sub txtYAxis_GotFocus()
txtYAxis.SelStart = 0
txtYAxis.SelLength = Len(txtYAxis.Text)
End Sub

Private Sub txtYAxis_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or KeyAscii = 8) Then KeyAscii = 0
cboSpot.ListIndex = FindNullSpot
End Sub

Function FindNullSpot() As Long
Dim iCnt As Long

For iCnt = 0 To cboSpot.ListCount - 1
    If cboSpot.ItemData(iCnt) = 0 Then FindNullSpot = iCnt
Next
End Function

Sub ApplyWorkSpace()
'    iZen As Long
'
'    sChar As String
'    iLevel As Long
'    iClass As Long
'    iStr As Long
'    iAgi As Long
'    iVit As Long
'    iEnrg As Long
'    iLife As Long
'    iMana As Long
'    iPoints As Long
'    iKills As Long
'    iPkLevel As Long
'
'    iMap As Long
'    iXAxis As Long
'    iYAxis As Long
'    iDir As Long

txtChar.Text = uWorkSpace.sChar
txtLevel.Text = Format(uWorkSpace.iLevel)
cboClass.ListIndex = Format(uWorkSpace.iClass)
txtStats(0).Text = Format(uWorkSpace.iStr)
txtStats(1).Text = Format(uWorkSpace.iAgi)
txtStats(2).Text = Format(uWorkSpace.iVit)
txtStats(3).Text = Format(uWorkSpace.iEnrg)
txtLife.Text = Format(uWorkSpace.iLife)
txtMana.Text = Format(uWorkSpace.iMana)
txtExp.Text = Format(uWorkSpace.iExp)
txtPoints.Text = Format(uWorkSpace.iPoints)
txtKills.Text = Format(uWorkSpace.iKills)
txtPkLevel.Text = Format(uWorkSpace.iPkLevel)

cboSpot.ListIndex = Format(uWorkSpace.iSpot)
txtMapNumber.Text = Format(uWorkSpace.iMap)
txtXAxis.Text = Format(uWorkSpace.iXAxis)
txtYAxis.Text = Format(uWorkSpace.iYAxis)
cboDir.ListIndex = FindDir(uWorkSpace.iDir)

lblZen.Caption = FormatNumber(uWorkSpace.iZen)

lstInvItems.Visible = False
InvCount

If lstInvItems.ListCount > uWorkSpace.iInvShowT Then
    lstInvItems.TopIndex = uWorkSpace.iInvShowT
Else
    lstInvItems.TopIndex = lstInvItems.ListCount - 1
End If

lstInvItems.Visible = True
End Sub

Sub ApplyData()
Dim iCnt As Long

cboSpot.Clear
cboSpot.AddItem GetString("string001")
cboSpot.ItemData(0) = 0
For iCnt = 1 To UBound(aMaps)
    cboSpot.AddItem aMaps(iCnt).sName
    cboSpot.ItemData(cboSpot.NewIndex) = iCnt
Next

cboClass.Clear
For iCnt = 1 To UBound(aClasses)
    cboClass.AddItem aClasses(iCnt).sName
    cboClass.ItemData(cboClass.NewIndex) = iCnt
Next

cboDir.Clear
For iCnt = 0 To UBound(aDirs)
    cboDir.AddItem ReturnDirName(aDirs(iCnt))
    cboDir.ItemData(cboDir.NewIndex) = iCnt
Next
End Sub


Sub UpdateWorkSpace()
'    iZen As Long
'
'    sChar As String
'    iLevel As Long
'    iClass As Long
'    iStr As Long
'    iAgi As Long
'    iVit As Long
'    iEnrg As Long
'    iLife As Long
'    iMana As Long
'    iPoints As Long
'    iKills As Long
'    iPkLevel As Long
'
'    iMap As Long
'    iXAxis As Long
'    iYAxis As Long
'    iDir As Long

uWorkSpace.sChar = txtChar.Text
uWorkSpace.iLevel = CheckValue(Val(txtLevel.Text), Max16Bits)
uWorkSpace.iClass = cboClass.ListIndex
uWorkSpace.iStr = CheckValue(Val(txtStats(0).Text), Max16Bits)
uWorkSpace.iAgi = CheckValue(Val(txtStats(1).Text), Max16Bits)
uWorkSpace.iVit = CheckValue(Val(txtStats(2).Text), Max16Bits)
uWorkSpace.iEnrg = CheckValue(Val(txtStats(3).Text), Max16Bits)
uWorkSpace.iLife = CheckValue(Val(txtLife.Text), Max16Bits)
uWorkSpace.iMana = CheckValue(Val(txtMana.Text), Max16Bits)
uWorkSpace.iPoints = CheckValue(Val(txtPoints.Text), Max16Bits)
uWorkSpace.iExp = CheckValue(Val(txtExp.Text), Max32Bits)
uWorkSpace.iKills = CheckValue(Val(txtKills.Text), Max32Bits)
uWorkSpace.iPkLevel = CheckValue(Val(txtPkLevel.Text), Max32Bits)

uWorkSpace.iSpot = cboSpot.ListIndex
uWorkSpace.iMap = CheckValue(Val(txtMapNumber.Text), Max8Bits)
uWorkSpace.iXAxis = CheckValue(Val(txtXAxis.Text), Max8Bits)
uWorkSpace.iYAxis = CheckValue(Val(txtYAxis.Text), Max8Bits)
uWorkSpace.iDir = cboDir.ItemData(cboDir.ListIndex)

uWorkSpace.iInvShowT = lstInvItems.TopIndex
End Sub

Function FindDir(ByVal lIData As Long) As Long
Dim iCnt As Long

For iCnt = 0 To cboDir.ListCount - 1
    If cboDir.ItemData(iCnt) = lIData Then
        FindDir = iCnt
        Exit Function
    End If
Next
FindDir = -1
End Function

Function CheckValue(ByVal lValue As Long, ByVal lMaxValue As Long) As Long
If lValue > lMaxValue Then
    CheckValue = lMaxValue
Else
    CheckValue = lValue
End If
End Function

Sub InvCount()
Dim iCnt As Long
Dim sName As String
Dim clNames As New Collection
Dim clCont() As Long
Dim lIndex As Long

For iCnt = 0 To UBound(aSpots)
    If aSpots(iCnt).bUsed And aSpots(iCnt).iClass <> -1 Then
        sName = Empty
        sName = sName + IIf(aSpots(iCnt).iExcOpt > 0, "Exc ", "")
        sName = sName + aItems(aSpots(iCnt).iClass, aSpots(iCnt).iID).sName(aSpots(iCnt).iReqLevel)
        If Not aItems(aSpots(iCnt).iClass, aSpots(iCnt).iID).bReqLevel Then
            sName = sName + IIf(aSpots(iCnt).iLevel <> 0 Or aSpots(iCnt).iOption <> 0, "+" + Format(aSpots(iCnt).iLevel), "")
            sName = sName + IIf(aSpots(iCnt).iOption <> 0, "+" + Format(aSpots(iCnt).iOption), "")
            sName = sName + IIf(aSpots(iCnt).bLuck, "+L", "")
            sName = sName + IIf(aSpots(iCnt).bSkill, "+S", "")
        End If
        lIndex = FindInCol(clNames, sName)
        If lIndex <> -1 Then
            clCont(lIndex) = clCont(lIndex) + 1
        Else
            clNames.Add sName
            ReDim Preserve clCont(1 To clNames.Count) As Long
            clCont(clNames.Count) = 1
        End If
    End If
Next

Dim sItem As String
Dim lMaxW As Long

lstInvItems.Clear

Set picFont.Font = lstInvItems.Font

For iCnt = 1 To clNames.Count
    sItem = clCont(iCnt) & " " & clNames(iCnt)
    If picFont.TextWidth(sItem) > lMaxW Then lMaxW = picFont.TextWidth(sItem)
    lstInvItems.AddItem sItem
Next

If lstInvItems.ListCount = 0 Then
    lstInvItems.AddItem GetString("string002")
    lstInvItems.Enabled = False
Else
    lstInvItems.Enabled = True
End If

lMaxW = lMaxW + (picFont.Width - picFont.ScaleWidth)
SendMessage lstInvItems.hwnd, LB_SETHORIZONTALEXTENT, lMaxW / Screen.TwipsPerPixelX, NUL
End Sub

Function FindInCol(ByRef inColl As Collection, ByVal sName As String) As Long
Dim iCnt As Long

FindInCol = -1
For iCnt = 1 To inColl.Count
    If inColl(iCnt) = sName Then FindInCol = iCnt
Next
End Function

Sub FillLangCombo()
Dim iCnt As Long

cboLanguage.AddItem OriginalName
cboLanguage.ItemData(cboLanguage.NewIndex) = 0

For iCnt = 1 To UBound(aLanguages)
    cboLanguage.AddItem aLanguages(iCnt).sName & " (" & aLanguages(iCnt).sFile & ")"
    cboLanguage.ItemData(cboLanguage.NewIndex) = iCnt
Next

For iCnt = 0 To cboLanguage.ListCount - 1
    If cboLanguage.ItemData(iCnt) = GetIdFromFile(uWorkSpace.sLangFile) Then cboLanguage.ListIndex = iCnt
Next

End Sub

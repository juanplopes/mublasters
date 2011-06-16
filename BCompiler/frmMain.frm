VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   Caption         =   "BCompiler output"
   ClientHeight    =   4740
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8385
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4740
   ScaleWidth      =   8385
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrCaret 
      Interval        =   10
      Left            =   2580
      Top             =   1680
   End
   Begin VB.TextBox txtConsole 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   3195
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "frmMain.frx":08CA
      Top             =   0
      Width           =   6555
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ShowCaret Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function HideCaret Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long


Public Sub AddItem(sItem As String)
LockWindowUpdate txtConsole.hwnd
txtConsole.Text = txtConsole.Text + sItem + vbCrLf
LockWindowUpdate 0
DoEvents
End Sub

Public Sub FeedLine()
AddItem ""
End Sub

Public Sub Clear()
txtConsole.Text = Empty
End Sub

Private Sub Resize()
txtConsole.Left = 0
txtConsole.Top = 0
txtConsole.Width = Me.ScaleWidth
txtConsole.Height = Me.ScaleHeight
End Sub

Private Sub SaveSizes()
SaveSetting App.Title, "Window", "Left", Me.Left
SaveSetting App.Title, "Window", "Top", Me.Top
SaveSetting App.Title, "Window", "Width", Me.Width
SaveSetting App.Title, "Window", "Height", Me.Height
End Sub

Private Sub LoadSizes()
Me.Left = GetSetting(App.Title, "Window", "Left", Me.Left)
Me.Top = GetSetting(App.Title, "Window", "Top", Me.Top)
Me.Width = GetSetting(App.Title, "Window", "Width", Me.Width)
Me.Height = GetSetting(App.Title, "Window", "Height", Me.Height)
End Sub

Private Sub Form_Load()
LoadSizes
Form_Resize
Me.Tag = Me.Caption
Clear
End Sub

Sub Finish()
Me.Caption = Me.Tag + " [finished]"
End Sub

Private Sub Form_Resize()
txtConsole.Visible = False
Resize
txtConsole.Visible = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
SaveSizes
End Sub

Private Sub tmrCaret_Timer()
HideCaret txtConsole.hwnd
End Sub

Private Sub txtConsole_Change()
txtConsole.SelStart = Len(txtConsole.Text)
txtConsole.Refresh
End Sub

VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5070
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   5070
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1898
      TabIndex        =   7
      Top             =   4320
      Width           =   1275
   End
   Begin VB.Frame fraAbout 
      Height          =   4215
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   4935
      Begin VB.Frame fraHelp 
         Height          =   1275
         Left            =   120
         TabIndex        =   9
         Top             =   2820
         Width           =   4695
         Begin VB.Label lblHelp1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   $"frmAbout.frx":1042
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   120
            TabIndex        =   10
            Top             =   180
            Width           =   4425
            WordWrap        =   -1  'True
         End
      End
      Begin VB.PictureBox picIcon 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   180
         Picture         =   "frmAbout.frx":1124
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   1
         Top             =   240
         Width           =   480
      End
      Begin VB.Label lblAbout1 
         AutoSize        =   -1  'True
         Caption         =   "Developed by"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   5
         Top             =   840
         Width           =   1185
      End
      Begin VB.Label lblUndetec 
         AutoSize        =   -1  'True
         Caption         =   "ZeroValue (zerovalue@gmail.com)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   999
         Left            =   180
         TabIndex        =   8
         Top             =   1080
         Width           =   2910
      End
      Begin VB.Label lblUndetec 
         AutoSize        =   -1  'True
         Caption         =   "CharBlaster"
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
         Index           =   998
         Left            =   780
         TabIndex        =   6
         Top             =   240
         Width           =   2025
      End
      Begin VB.Label lblAbout2 
         AutoSize        =   -1  'True
         Caption         =   "ItemCode Algorithm is based (but not copied) in the ItemProject source code. I don't know who codes it, but thanks a lot man!"
         Height          =   390
         Left            =   180
         TabIndex        =   4
         Top             =   1380
         Width           =   4695
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblAbout4 
         AutoSize        =   -1  'True
         Caption         =   "This program has only educational purposes. Any other use of it is your entire responsibility."
         Height          =   390
         Left            =   180
         TabIndex        =   3
         Top             =   2340
         Width           =   4665
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblAbout3 
         AutoSize        =   -1  'True
         Caption         =   "Thanks Doidin, Poke and everyone who helped me to develop this program (testing or having his server hacked =D)."
         Height          =   390
         Left            =   180
         TabIndex        =   2
         Top             =   1860
         Width           =   4590
         WordWrap        =   -1  'True
      End
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOk_Click()
Unload Me
End Sub

Private Sub Form_Load()
SaveCaptions Me
LoadFormLang GetIdFromFile(uWorkSpace.sLangFile), Me.name
End Sub


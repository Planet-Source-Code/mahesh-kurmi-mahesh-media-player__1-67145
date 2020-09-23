VERSION 5.00
Object = "{86135EDC-6265-45AA-8A47-6C463280490B}#1.0#0"; "AudioControls2.ocx"
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      BackColor       =   &H00404040&
      Caption         =   "PLAYLIST EDITOR"
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1335
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   735
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   3840
         TabIndex        =   9
         Text            =   "Text2"
         Top             =   840
         Width           =   615
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H80000009&
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   2040
         TabIndex        =   6
         Top             =   600
         Visible         =   0   'False
         Width           =   735
         Begin VB.Label Label4 
            BackColor       =   &H00800000&
            BackStyle       =   0  'Transparent
            Caption         =   " delete"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   0
            TabIndex        =   8
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label3 
            BackColor       =   &H00800000&
            BackStyle       =   0  'Transparent
            Caption         =   " play"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   0
            TabIndex        =   7
            Top             =   0
            Width           =   810
         End
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   1335
         Left            =   1920
         Max             =   8
         TabIndex        =   5
         Top             =   120
         Width           =   135
      End
      Begin VB.Label track 
         BackStyle       =   0  'Transparent
         Caption         =   "Label3"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Visible         =   0   'False
         Width           =   3600
      End
      Begin VB.Label track 
         BackStyle       =   0  'Transparent
         Caption         =   "Label3"
         Height          =   195
         Index           =   8
         Left            =   120
         TabIndex        =   17
         Top             =   2160
         Visible         =   0   'False
         Width           =   3600
      End
      Begin VB.Label track 
         BackStyle       =   0  'Transparent
         Caption         =   "Label3"
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   16
         Top             =   1920
         Visible         =   0   'False
         Width           =   3600
      End
      Begin VB.Label track 
         BackStyle       =   0  'Transparent
         Caption         =   "Label3"
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   15
         Top             =   1680
         Visible         =   0   'False
         Width           =   3600
      End
      Begin VB.Label track 
         BackStyle       =   0  'Transparent
         Caption         =   "Label3"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   14
         Top             =   1440
         Visible         =   0   'False
         Width           =   3600
      End
      Begin VB.Label track 
         BackStyle       =   0  'Transparent
         Caption         =   "Label3"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   13
         Top             =   1200
         Visible         =   0   'False
         Width           =   3600
      End
      Begin VB.Label track 
         BackStyle       =   0  'Transparent
         Caption         =   "Label3"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Visible         =   0   'False
         Width           =   3600
      End
      Begin VB.Label track 
         BackStyle       =   0  'Transparent
         Caption         =   "Label3"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Visible         =   0   'False
         Width           =   3600
      End
      Begin VB.Label track 
         BackStyle       =   0  'Transparent
         Caption         =   "Label3"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Visible         =   0   'False
         Width           =   3600
      End
   End
   Begin AUDIOCONTROLS2Lib.Knob Knob1 
      Height          =   1335
      Left            =   360
      TabIndex        =   3
      Top             =   1320
      Width           =   1215
      _Version        =   65536
      _ExtentX        =   2143
      _ExtentY        =   2355
      _StockProps     =   0
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   3360
      TabIndex        =   2
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   360
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   1815
      ItemData        =   "Form2.frx":0000
      Left            =   1800
      List            =   "Form2.frx":0016
      TabIndex        =   0
      Top             =   480
      Width           =   975
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'List1.AddItem("a", 2) = Text1.Text
'LED.YellowZone = 123
End Sub

Private Sub List1_Click()
'Text1.Text = List1.List
End Sub

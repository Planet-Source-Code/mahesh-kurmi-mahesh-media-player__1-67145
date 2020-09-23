VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00808080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " About Mahesh's MediaPlayer"
   ClientHeight    =   3495
   ClientLeft      =   1905
   ClientTop       =   2520
   ClientWidth     =   5205
   Icon            =   "about me.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   5205
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      BackColor       =   &H00808080&
      Cancel          =   -1  'True
      Caption         =   "OK"
      Height          =   330
      Left            =   4170
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3030
      Width           =   825
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00808080&
      Height          =   675
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Text            =   "about me.frx":0442
      Top             =   1140
      Width           =   4815
   End
   Begin VB.Label compname 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Computer name"
      Height          =   330
      Left            =   270
      TabIndex        =   8
      Top             =   2520
      Width           =   3165
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   180
      X2              =   5040
      Y1              =   2940
      Y2              =   2940
   End
   Begin VB.Label user 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "User name"
      Height          =   330
      Left            =   270
      TabIndex        =   6
      Top             =   2130
      Width           =   3165
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "This Product is licenced to:"
      Height          =   195
      Left            =   270
      TabIndex        =   5
      Top             =   1920
      Width           =   1920
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "For SGSITS Corp."
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   3150
      TabIndex        =   3
      Top             =   480
      Width           =   1275
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "All rights reserved"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   300
      TabIndex        =   2
      Top             =   750
      Width           =   1245
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright(C) 2005-2006 SGSITS Corp."
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   270
      TabIndex        =   1
      Top             =   480
      Width           =   2715
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mahesh's Mediaplayer  Version 2.1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   300
      TabIndex        =   0
      Top             =   210
      Width           =   2460
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim FSys As New Shell32.ShellFolderView
Public Sub Command1_Click()
Unload Me
End Sub

Public Sub Form_Load()
Dim k

Dim buffer As String
buffer = Space(60)
k = GetUserName(buffer, 60)
user.Caption = " " + buffer
k = GetComputerName(buffer, 60)
compname.Caption = " " + buffer
k = SetWindowPos(Me.hwnd, -1, Me.Left / 12000 * 800, Me.Top / 9000 * 600, Me.width / 12000 * 800, Me.height / 9000 * 600, &H40)
'Form1.Show
End Sub




Private Sub Timer2_Timer()

End Sub

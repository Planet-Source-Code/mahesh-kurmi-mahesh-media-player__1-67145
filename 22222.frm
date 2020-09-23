VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3720
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5205
   LinkTopic       =   "Form1"
   ScaleHeight     =   3720
   ScaleWidth      =   5205
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Text            =   "Text3"
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   3120
      TabIndex        =   2
      Top             =   1440
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   2535
      Left            =   840
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "22222.frx":0000
      Top             =   360
      Width           =   2055
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   2535
      Left            =   4440
      TabIndex        =   0
      Top             =   480
      Width           =   255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text1.Text = 2
Text1.Text = 3
Text1.Text = 4
Text1.Text = 5
Text1.Text = Chr(26)
End Sub

Private Sub Text2_Change()
'Dim k As
'k = keycode(Text2.Text)
'Text1.Text = k
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
Text1.Text = KeyCode
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
Text1.Text = KeyAscii
End Sub

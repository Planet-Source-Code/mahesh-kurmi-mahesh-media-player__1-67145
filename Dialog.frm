VERSION 5.00
Begin VB.Form Dialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dialog Caption"
   ClientHeight    =   5250
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   8355
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   8355
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "OK"
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   3960
      Width           =   1215
   End
   Begin VB.FileListBox File2 
      Height          =   2820
      Left            =   3840
      MultiSelect     =   2  'Extended
      TabIndex        =   2
      Top             =   1200
      Width           =   3735
   End
   Begin VB.DirListBox Dir2 
      Height          =   2565
      Left            =   960
      TabIndex        =   1
      Top             =   1200
      Width           =   2655
   End
   Begin VB.DriveListBox Drive2 
      Height          =   315
      Left            =   1680
      TabIndex        =   0
      Top             =   240
      Width           =   4095
   End
   Begin VB.Label Label1 
      Caption         =   "Look in"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub OKButton_Click()

End Sub


VERSION 5.00
Begin VB.Form frmMirror 
   BorderStyle     =   1  'Fixed Single
   Caption         =   ". Mahesh Mediaplayer ."
   ClientHeight    =   870
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   2370
   ClipControls    =   0   'False
   Icon            =   "frmMirror.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   870
   ScaleWidth      =   2370
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmMirror"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    musicsystem.Show
   
End Sub
Private Sub Form_Paint()
  If Me.WindowState = vbNormal Then
       ' musicsystem.WindowState = vbNormal
        musicsystem.Show
        If mini <> True Then plst.Visible = True
        musicsystem.Image22.Picture = musicsystem.Image21.Picture
        'musicsystem.Timer7.Enabled = False
  ElseIf Me.WindowState = vbMinimized Then
       ' musicsystem.WindowState = vbMinimized
        musicsystem.Hide
             If attach3 = False Then
             If plst.Visible = True Then mini = False
             plst.Visible = False
             End If
             
             
  End If
    'musicsystem.width = plst.width
   ' musicsystem.Refresh
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Me.Visible = False
    musicsystem.Visible = False
    plst.Visible = False
    Form1.Visible = False
    Set frmMirror = Nothing
    'musicsystem.mp3.PlayStop
    DoEvents
    DoEvents
       Unload Form2
       Unload plst
       Unload Form1
       Unload musicsystem
        
    End
End Sub

Private Sub Form_Resize()
Call Form_Paint
   
End Sub


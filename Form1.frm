VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "msdxm.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   4035
   ClientLeft      =   6150
   ClientTop       =   1200
   ClientWidth     =   8460
   ControlBox      =   0   'False
   FillColor       =   &H00404040&
   FillStyle       =   0  'Solid
   FontTransparent =   0   'False
   ForeColor       =   &H00404040&
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4035
   ScaleWidth      =   8460
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   5040
      Top             =   900
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   5040
      Top             =   480
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   245
      Left            =   2070
      TabIndex        =   4
      Top             =   3660
      Width           =   1595
      Begin VB.Image downtime 
         Height          =   120
         Left            =   1330
         Picture         =   "Form1.frx":0442
         Top             =   110
         Width           =   225
      End
      Begin VB.Image uptime 
         Height          =   120
         Left            =   1330
         Picture         =   "Form1.frx":0604
         Top             =   10
         Width           =   225
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00404040&
         X1              =   1320
         X2              =   1320
         Y1              =   0
         Y2              =   240
      End
      Begin VB.Label trackno 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   " T 1"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   150
         Left            =   30
         TabIndex        =   6
         Top             =   40
         Width           =   180
      End
      Begin VB.Label trackpos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00:00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   210
         Left            =   880
         TabIndex        =   5
         ToolTipText     =   "Time elapsed"
         Top             =   20
         Width           =   405
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00404040&
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   240
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00404040&
         X1              =   0
         X2              =   1590
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00404040&
         X1              =   1580
         X2              =   1580
         Y1              =   0
         Y2              =   240
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00404040&
         X1              =   0
         X2              =   1590
         Y1              =   225
         Y2              =   225
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   265
      Left            =   240
      TabIndex        =   3
      Top             =   3630
      Width           =   1755
      Begin VB.Image cbutton 
         Height          =   270
         Index           =   0
         Left            =   0
         Picture         =   "Form1.frx":07C6
         ToolTipText     =   "Previous track(B)"
         Top             =   0
         Width           =   360
      End
      Begin VB.Image cbutton 
         Height          =   270
         Index           =   1
         Left            =   1410
         Picture         =   "Form1.frx":0D18
         ToolTipText     =   "Next track(N)"
         Top             =   0
         Width           =   360
      End
      Begin VB.Image cbutton 
         Height          =   270
         Index           =   6
         Left            =   1050
         Picture         =   "Form1.frx":126A
         ToolTipText     =   "Stop(S)"
         Top             =   0
         Width           =   360
      End
      Begin VB.Image cbutton 
         Height          =   270
         Index           =   3
         Left            =   690
         Picture         =   "Form1.frx":17BC
         ToolTipText     =   "Pause(P)"
         Top             =   0
         Width           =   360
      End
      Begin VB.Image cbutton 
         Height          =   270
         Index           =   2
         Left            =   330
         Picture         =   "Form1.frx":1D0E
         ToolTipText     =   "Play(P)"
         Top             =   0
         Width           =   360
      End
      Begin VB.Image Image4 
         Height          =   375
         Left            =   -2640
         MousePointer    =   7  'Size N S
         Picture         =   "Form1.frx":2260
         Top             =   -60
         Width           =   2220
      End
   End
   Begin MSComctlLib.Slider text1 
      Height          =   0
      Left            =   2190
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   270
      Width           =   15
      _ExtentX        =   26
      _ExtentY        =   0
      _Version        =   393216
      LargeChange     =   1
      Max             =   100
      SelStart        =   14
      Value           =   14
   End
   Begin VB.Image Image5 
      Height          =   120
      Left            =   5160
      Picture         =   "Form1.frx":4DFE
      Top             =   1320
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image downdown 
      Height          =   120
      Left            =   5010
      Picture         =   "Form1.frx":4FC0
      Top             =   2040
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image updown 
      Height          =   120
      Left            =   5070
      Picture         =   "Form1.frx":5182
      Top             =   1800
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image downup 
      Height          =   120
      Left            =   5040
      Picture         =   "Form1.frx":5344
      Top             =   2670
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image upup 
      Height          =   120
      Left            =   5070
      Picture         =   "Form1.frx":5506
      Top             =   2460
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image Imagetop2 
      Height          =   480
      Left            =   2040
      MousePointer    =   7  'Size N S
      Picture         =   "Form1.frx":56C8
      Top             =   3540
      Width           =   2265
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   12000
   End
   Begin VB.Image ibottom 
      Height          =   480
      Left            =   0
      MousePointer    =   7  'Size N S
      Picture         =   "Form1.frx":900A
      Top             =   3540
      Width           =   10875
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "MAHESH'S VIDEO"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   150
      Left            =   1620
      TabIndex        =   2
      Top             =   60
      Width           =   1020
   End
   Begin VB.Image Image1 
      Height          =   2835
      Left            =   0
      Picture         =   "Form1.frx":1A04C
      Top             =   6540
      Width           =   165
   End
   Begin VB.Image Ileft 
      Enabled         =   0   'False
      Height          =   6615
      Left            =   0
      MousePointer    =   9  'Size W E
      Picture         =   "Form1.frx":1BB22
      Top             =   0
      Width           =   165
   End
   Begin VB.Image Image2 
      Height          =   3420
      Left            =   5820
      Picture         =   "Form1.frx":1F968
      Top             =   6360
      Width           =   180
   End
   Begin VB.Image min 
      Height          =   135
      Left            =   3900
      Picture         =   "Form1.frx":219BA
      ToolTipText     =   "Minimise"
      Top             =   0
      Width           =   165
   End
   Begin VB.Image meclose 
      Height          =   135
      Left            =   4080
      Picture         =   "Form1.frx":21B40
      ToolTipText     =   "Close"
      Top             =   30
      Width           =   135
   End
   Begin VB.Image closedown 
      Height          =   150
      Left            =   6960
      Picture         =   "Form1.frx":21C7E
      Top             =   720
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Image closeup 
      Height          =   135
      Left            =   7170
      Picture         =   "Form1.frx":21E00
      Top             =   690
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Image minup 
      Height          =   135
      Left            =   6930
      Picture         =   "Form1.frx":21F3E
      Top             =   2220
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Image mindown 
      Height          =   120
      Left            =   7140
      Picture         =   "Form1.frx":220C4
      Top             =   2220
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Image iright 
      Height          =   6555
      Left            =   4140
      MousePointer    =   9  'Size W E
      Picture         =   "Form1.frx":221E6
      Top             =   -30
      Width           =   180
   End
   Begin VB.Image itop 
      Height          =   255
      Left            =   -30
      Picture         =   "Form1.frx":25F9C
      Top             =   0
      Width           =   10410
   End
   Begin VB.Image Image3 
      Height          =   255
      Left            =   6750
      Picture         =   "Form1.frx":2EA42
      Top             =   0
      Width           =   8310
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "MAHESH'S VIDEO"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   120
      Left            =   150
      TabIndex        =   0
      Top             =   -15
      Width           =   870
   End
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer1 
      CausesValidation=   0   'False
      DragIcon        =   "Form1.frx":35904
      Height          =   3405
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   4740
      AudioStream     =   0
      AutoSize        =   0   'False
      AutoStart       =   0   'False
      AnimationAtStart=   0   'False
      AllowScan       =   0   'False
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   0   'False
      CursorType      =   0
      CurrentPosition =   1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   255
      DisplayForeColor=   16711680
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   0   'False
      EnablePositionControls=   0   'False
      EnableFullScreenControls=   0   'False
      EnableTracker   =   0   'False
      Filename        =   ""
      InvokeURLs      =   0   'False
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   0
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   0   'False
      SendWarningEvents=   0   'False
      SendErrorEvents =   0   'False
      SendKeyboardEvents=   -1  'True
      SendMouseClickEvents=   -1  'True
      SendMouseMoveEvents=   -1  'True
      SendPlayStateChangeEvents=   0   'False
      ShowCaptioning  =   0   'False
      ShowControls    =   0   'False
      ShowAudioControls=   0   'False
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   0   'False
      ShowStatusBar   =   0   'False
      ShowTracker     =   0   'False
      TransparentAtStart=   -1  'True
      VideoBorderWidth=   0
      VideoBorderColor=   255
      VideoBorder3D   =   0   'False
      Volume          =   -10
      WindowlessVideo =   0   'False
   End
   Begin VB.Menu property 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu wtop 
         Caption         =   " Always on top"
         Checked         =   -1  'True
      End
      Begin VB.Menu bar1 
         Caption         =   "-"
      End
      Begin VB.Menu fullscreen 
         Caption         =   "Switch Fullscreen           Alt+Enter"
      End
      Begin VB.Menu zoom 
         Caption         =   "Zoom"
         Begin VB.Menu zoom50 
            Caption         =   "50%"
         End
         Begin VB.Menu zoom100 
            Caption         =   "100%"
         End
         Begin VB.Menu zoom200 
            Caption         =   "200%"
         End
      End
      Begin VB.Menu bar4 
         Caption         =   "-"
      End
      Begin VB.Menu ratio43 
         Caption         =   "Force Ratio 4:3"
      End
      Begin VB.Menu volume 
         Caption         =   "Volume"
         Begin VB.Menu volup 
            Caption         =   "Up"
         End
         Begin VB.Menu voldown 
            Caption         =   "Down"
         End
         Begin VB.Menu volmute 
            Caption         =   "mute"
         End
      End
      Begin VB.Menu playpause 
         Caption         =   "play/pause"
      End
      Begin VB.Menu bar5 
         Caption         =   "-"
      End
      Begin VB.Menu about 
         Caption         =   "About MediaPlayer"
      End
      Begin VB.Menu exitt 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim max As Boolean
Public Sub Command1_Click()
'MediaPlayer1.ImageSourceHeight = MediaPlayer1.ImageSourceHeight + 100
End Sub

Public Sub about_Click()
On Error Resume Next
Form2.Show
End Sub

Public Sub cbutton_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
cbutton(Index).Picture = musicsystem.cbuttonE(Index).Picture
End Sub




Public Sub cbutton_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 cbutton(Index).Picture = musicsystem.cbuttonD(Index).Picture

End Sub

Private Sub cbutton_Click(Index As Integer)
Call musicsystem.cbutton_Click(Index)
text1.SetFocus
End Sub

Private Sub downtime_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
downtime.Picture = downdown.Picture
Timer3.Enabled = True
End Sub

Private Sub downtime_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
downtime.Picture = downup.Picture
Timer3.Enabled = False
End Sub

Public Sub exitt_Click()
Unload Me
End Sub

Private Sub Form_Activate()
'k = SetWindowPos(Me.hwnd, -1, Me.Left / 12000 * 800, Me.Top / 9000 * 600, Me.width / 12000 * 800, Me.height / 9000 * 600, &H40)
On Error GoTo 1
cmo = True

If plst.List1.ListIndex <> -1 Then plst.List1.Selected(plst.List1.ListIndex) = False
plst.scroll.Top = 90
'form1.scroll.top

Call FormResize
'MediaPlayer1.SetFocus
1:
End Sub

Private Sub Form_DblClick()
If MediaPlayer1.DisplaySize <> mpFullScreen Then
If Me.WindowState <> 2 Then
Me.WindowState = 2
Else
Me.WindowState = 0
End If
Call FormResize
End If
'If Me.WindowState <> 1 Then Me.WindowState = 0
1:
End Sub

Private Sub Form_GotFocus()
'k = SetWindowPos(Me.hwnd, -1, Me.Left / 12000 * 800, Me.Top / 9000 * 600, Me.width / 12000 * 800, Me.height / 9000 * 600, &H40)
On Error GoTo 1
cmo = True

plst.List1.Selected(plst.List1.ListIndex) = False
plst.scroll.Top = 90
'form1.scroll.top

Call FormResize
MediaPlayer1.SetFocus
1:
End Sub

Private Sub Form_KeyDown(Keycode As Integer, Shift As Integer)
If Keycode = 112 Then
'AppActivate App.Title
'SendKeys "{F1}", True

SendKeys "%+{F4}", True
End If
End Sub

Public Sub Form_Load()
'k = SetWindowPos(Me.hwnd, -1, 360, 100, 420, 340, &H40)
'Me.Hide

Me.Top = GetSetting(App.EXEName, "Startup", "form1Top", musicsystem.Left + musicsystem.width)
Me.Left = GetSetting(App.EXEName, "Startup", "form1Left", musicsystem.Top)
Me.height = GetSetting(App.EXEName, "Startup", "form1height", musicsystem.height + plst.height)
Me.width = GetSetting(App.EXEName, "Startup", "form1width", musicsystem.width)
'Me.Top = 400
'Me.Left = 400
If Me.Top + Me.height < 10 Or Me.Top > 12000 Then Me.Top = 400
If Me.Left + Me.Left < 10 Or Me.Left > 12000 Then Me.Top = 400
If Me.height <= 1755 Then Me.height = 4125
'If Me.height < 1770 Then Me.height = 1770
'If Me.width < 4335 Then Me.height = 4125
If Me.Top < -Me.height + 20 Then Me.Top = musicsystem.Top
If Me.Left < -Me.height + 20 Then Me.Left = musicsystem.Left + musicsystem.width

'Call EnableFileDrops(Me)
'MediaPlayer1.height = (Me.height - 550)
'End If
' / 100105 *
'MediaPlayer1.width = Me.width - 550 ' 55
MediaPlayer1.Top = 160
MediaPlayer1.Left = 160
itop.Top = 0
Ileft.Left = 0
iright.Left = Me.width - 240
ibottom.Top = Me.height - 345
Imagetop2.Top = ibottom.Top 'ibottom.height - 10
Imagetop2.Left = Me.width - Imagetop2.width
ratio = False
 k = SetWindowPos(Form1.hwnd, -1, Form1.Left / 12000 * 800, Form1.Top / 9000 * 600, Form1.width / 12000 * 800, Form1.height / 9000 * 600, &H40)
text1.max = musicsystem.Slider3.max
text1.min = musicsystem.Slider3.min
text1.Value = musicsystem.Slider3.Value

Me.Hide
'Call zoom50_Click
End Sub

Private Sub Form_LostFocus()
If attach1 <> True Then

If plst.Left >= Me.Left + Me.width - 20 And plst.Left <= Me.Left + Me.width + 20 Then
 plst.Left = Me.Left + Me.width
 attach3 = True
 ElseIf Me.Left >= plst.Left + plst.width - 20 And Me.Left <= plst.Left + plst.width + 20 Then
 Me.Left = plst.Left + plst.width
 attach3 = True
ElseIf plst.Top >= Me.Top + Me.height - 20 And plst.Top <= Me.Top + Me.height + 20 Then
 plst.Top = Me.Top + Me.height
 attach3 = True
 ElseIf Me.Top >= plst.Top + plst.height - 20 And Me.Top <= plst.Top + plst.height + 20 Then
 Me.Top = plst.Top + plst.height
 attach3 = True

'plst.Left = Me.Left + Me.width
'attach1 = True

Else
attach3 = False
End If
End If
End Sub

Public Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
musicsystem.Timer5.Enabled = False

End Sub

Public Sub FormResize()
'DoEvents
On Error GoTo 1
DoEvents
itop.Top = -10
Ileft.Left = -10
iright.Left = Me.width - 165
Image2.Left = iright.Left
Imagetop2.Top = Me.height - 165
Image1.Left = Ileft.Left
ibottom.Top = Me.height - ibottom.height 'ibottom.height - 10
Imagetop2.Top = ibottom.Top 'ibottom.height - 10
Imagetop2.Left = Me.width - Imagetop2.width - 10
Image3.Top = itop.Top
Frame1.Top = ibottom.Top + 100
Frame2.Top = ibottom.Top + 130
'uptime.Top = ibottom.Top + 120
'downtime.Top = uptime.Top + 110

'If Form1.Caption <> "" Then
'MediaPlayer1.height = (Me.height - 400)
'Else
MediaPlayer1.height = (Me.height - itop.height - ibottom.height)
MediaPlayer1.Top = itop.height
'End If
' / 100105 *
'MediaPlayer1.Left = 250
MediaPlayer1.width = Me.width - 360 ' 55
'MediaPlayer1.Top = 130
'MediaPlayer1.Left = 130
MediaPlayer1.ShowControls = False

'MediaPlayer1.DisplaySize = mpFullScreen
MediaPlayer1.EnableFullScreenControls = False
'MediaPlayer1.ShowControls = False
' / 100 102 *
'MediaPlayer1.Left = Me.Left + 120
'Text1.Text = MediaPlayer1.Left
'If MediaPlayer1.Height <= 3600 Then
'MediaPlayer1.Top = -240 * (MediaPlayer1.Height) / 3600 '-1000 * Me.Width / 12000
'Else
'MediaPlayer1.Top = -360 '-1000 * Me.Width / 12000
'End If
meclose.Left = Me.width - 200
min.Left = Me.width - 360
Label1.Left = (Me.width - Label1.width) / 2
1:
If Form1.height <= musicsystem.height + 20 And Form1.height >= musicsystem.height - 20 Then
 Form1.height = musicsystem.height
ElseIf Form1.height <= plst.height + 20 And Form1.height >= plst.height - 20 Then
 Form1.height = plst.height
ElseIf Form1.height <= musicsystem.height + plst.height + 20 And Form1.height >= musicsystem.height + plst.height - 20 Then
 Form1.height = musicsystem.height + plst.height + 10
End If

If Form1.width <= musicsystem.width + 20 And Form1.width >= musicsystem.width - 20 Then Form1.width = musicsystem.width

End Sub

Private Sub Form_Resize()

If attach3 = True Then plst.Show
End Sub

Public Sub fullscreen_Click()
'Call MediaPlayer1_DblClick(1, 0, 1, 1)
'MediaPlayer1.SetFocus
'app
If MediaPlayer1.DisplaySize <> mpFullScreen Then
Me.WindowState = 2

MediaPlayer1.DisplaySize = mpFullScreen
 
Else
Me.WindowState = 0
MediaPlayer1.DisplaySize = mpFitToSize
FormResize
End If
FormResize

End Sub


Private Sub ibottom_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim k
Dim m As POINTAPI
k = GetCursorPos(m)
initx = m.X
inity = m.Y
End Sub

Private Sub Ibottom_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim k
Dim m As POINTAPI
On Error GoTo 1
If Button = 1 Then
k = GetCursorPos(m)
'[DoEvents
'k = MoveWindow(Me.hwnd, Me.Left * 800 / 12000, Me.Top * 800 / 12000, (Me.width) / 12000 * 800, (Me.height + 210) / 12000 * 800, 1)
If m.Y >= inity + 2 Then
If Me.height >= 8000 Then
Me.height = 8000
If ratio = True Then Me.width = 4# / 3# * Me.height
Call FormResize
Exit Sub
End If
Me.height = Me.height + (m.Y - inity) * 12000 / 800
inity = m.Y
If ratio = True Then Me.width = 4# / 3# * Me.height

Call FormResize
ElseIf m.Y <= inity - 2 Then
If Me.height <= 1770 Then
Me.height = 1770
If ratio = True Then Me.width = 4# / 3# * Me.height

Call FormResize
Exit Sub
End If
Me.height = Me.height + (m.Y - inity) * 12000 / 800
inity = m.Y
If ratio = True Then Me.width = 4# / 3# * Me.height

Call FormResize
End If
'Me.width = Me.width + (m.X - initx) * 12000 / 800

End If
1:
End Sub

Private Sub Imagetop2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim k
Dim m As POINTAPI
k = GetCursorPos(m)
initx = m.X
inity = m.Y
End Sub

Private Sub Imagetop2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim k
Dim m As POINTAPI
On Error GoTo 1
If Button = 1 Then
k = GetCursorPos(m)
'[DoEvents
'k = MoveWindow(Me.hwnd, Me.Left * 800 / 12000, Me.Top * 800 / 12000, (Me.width) / 12000 * 800, (Me.height + 210) / 12000 * 800, 1)
If m.Y >= inity + 2 Then
If Me.height >= 8000 Then
Me.height = 8000
If ratio = True Then Me.width = 4# / 3# * Me.height
Call FormResize
Exit Sub
End If
Me.height = Me.height + (m.Y - inity) * 12000 / 800
inity = m.Y
If ratio = True Then Me.width = 4# / 3# * Me.height

Call FormResize
ElseIf m.Y <= inity - 2 Then
If Me.height <= 1770 Then
Me.height = 1770
If ratio = True Then Me.width = 4# / 3# * Me.height

Call FormResize
Exit Sub
End If
Me.height = Me.height + (m.Y - inity) * 12000 / 800
inity = m.Y
If ratio = True Then Me.width = 4# / 3# * Me.height

Call FormResize
End If
'Me.width = Me.width + (m.X - initx) * 12000 / 800

End If
1:
End Sub


Private Sub Itop_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim m As POINTAPI
Dim k
moovevideo = True
k = GetCursorPos(m)
If attach1 <> True Then
If plst.Left >= Me.Left + Me.width - 40 And plst.Left <= Me.Left + Me.width + 40 Then
 plst.Left = Me.Left + Me.width
 attach3 = True
 ElseIf Me.Left >= plst.Left + plst.width - 40 And Me.Left <= plst.Left + plst.width + 40 Then
 Me.Left = plst.Left + plst.width
 attach3 = True
ElseIf plst.Top >= Me.Top + Me.height - 40 And plst.Top <= Me.Top + Me.height + 40 Then
 plst.Top = Me.Top + Me.height
 attach3 = True
 ElseIf Me.Top >= plst.Top + plst.height - 40 And Me.Top <= plst.Top + plst.height + 40 Then
 Me.Top = plst.Top + plst.height
 attach3 = True

Else
attach3 = False
End If
End If
k = GetCursorPos(m)
initialypos = m.Y - Me.Top / 12000 * 800
initialxpos = m.X - Me.Left / 12000 * 800
initialpx = m.X - plst.Left / 12000 * 800
initialpy = m.Y - plst.Top / 12000 * 800

End Sub

Public Sub Itop_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim k
Dim m As POINTAPI
If moovevideo = True Then
k = GetCursorPos(m)
'm.X = (X - Me.Left) * 277 / 4110
If Button = 1 Then
m.X = m.X - initialxpos
m.Y = m.Y - initialypos

k = MoveWindow(Me.hwnd, m.X, m.Y, Me.width / 12000 * 800, Me.height / 12000 * 800, 1)
If attach3 = True Then k = MoveWindow(plst.hwnd, m.X - initialpx + initialxpos, m.Y - initialpy + initialypos, plst.width / 12000 * 800, plst.height / 12000 * 800, 1)
End If
End If
End Sub


Private Sub Label15_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Itop_MouseDown(Button, Shift, Label15.Left + X, Y)

End Sub

Private Sub Label15_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Call Itop_MouseMove(Button, Shift, Label15.Left + X, Y)
End Sub



Private Sub itop_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
moovevideo = False
If attach1 <> True Then

If plst.Left >= Me.Left + Me.width - 20 And plst.Left <= Me.Left + Me.width + 20 Then
 plst.Left = Me.Left + Me.width
 attach3 = True
 ElseIf Me.Left >= plst.Left + plst.width - 20 And Me.Left <= plst.Left + plst.width + 20 Then
 Me.Left = plst.Left + plst.width
 attach3 = True
ElseIf plst.Top >= Me.Top + Me.height - 20 And plst.Top <= Me.Top + Me.height + 20 Then
 plst.Top = Me.Top + Me.height
 attach3 = True
 ElseIf Me.Top >= plst.Top + plst.height - 20 And Me.Top <= plst.Top + plst.height + 20 Then
 Me.Top = plst.Top + plst.height
 attach3 = True

'plst.Left = Me.Left + Me.width
'attach1 = True

Else
attach3 = False
End If
End If

End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Itop_MouseDown(Button, Shift, Label1.Left + X, Y)

End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Itop_MouseMove(Button, Shift, Label1.Left + X, Y)

End Sub

Private Sub MediaPlayer1_Click(Button As Integer, ShiftState As Integer, X As Single, Y As Single)
On Error Resume Next
'musicsystem.Show
text1.SetFocus
End Sub

Public Sub MediaPlayer1_DblClick(Button As Integer, ShiftState As Integer, X As Single, Y As Single)
On Error GoTo 1
'If MediaPlayer1.DisplaySize <> mpFullScreen Then
'Form1.WindowState = 2
'MediaPlayer1.width = 14000
'MediaPlayer1.height = 12000
'MediaPlayer1.ShowControls = True
'MediaPlayer1.DisplaySize = mpFullScreen
'MediaPlayer1.EnableFullScreenControls = True
'moving.Enabled = False
'minimise.Enabled = False
'Else
'moving.Enabled = True
'minimise.Enabled = True

'MediaPlayer1.ShowControls = False
'Form1.WindowState = 0
'MediaPlayer1.Width =
'MediaPlayer1.Height = 12000
'MediaPlayer1.DisplaySize = mpFitToSize
'MediaPlayer1.ShowControls = False
'MediaPlayer1.EnableFullScreenControls = False
'Call FormResize
'form1.WindowState = 2
' mpFullScreen
'End If
If MediaPlayer1.DisplaySize <> mpFullScreen Then
If Me.WindowState <> 2 Then
Me.WindowState = 2
Else
Me.WindowState = 0
End If
Call FormResize
End If
'Call MediaPlayer1_Click
'If Me.WindowState <> 1 Then Me.WindowState = 0
1:
End Sub


Private Sub MediaPlayer1_Error()
Exit Sub
End Sub

Public Sub MediaPlayer1_KeyDown(Keycode As Integer, ShiftState As Integer)
On Error GoTo 9

If Keycode = 112 Then
'AppActivate App.Title
'SendKeys "{F1}", True

SendKeys "%+{F4}", True
End If
If Keycode = 70 Then Call fullscreen_Click
'MediaPlayer1.SetFocus
text1.SetFocus
9:
End Sub

Private Sub MediaPlayer1_KeyPress(Keycode As Integer)
If Keycode = 112 Then
'AppActivate App.Title
'SendKeys "{F1}", True

SendKeys "%+{F4}", True
End If
End Sub

Public Sub MediaPlayer1_MouseDown(Button As Integer, ShiftState As Integer, X As Single, Y As Single)

Me.CurrentX = 1000
Me.CurrentY = 2000
Me.Print "hero heeralal"
On Error GoTo 1
moovevideo = True
If MediaPlayer1.DisplaySize <> mpFullScreen Then
MediaPlayer1.EnableFullScreenControls = False
MediaPlayer1.ShowControls = False
End If
Dim m As POINTAPI
Dim k
k = GetCursorPos(m)
attach3 = False

If attach1 <> True Then

If plst.Left >= Me.Left + Me.width - 20 And plst.Left <= Me.Left + Me.width + 20 Then
 plst.Left = Me.Left + Me.width
 attach3 = True
 ElseIf Me.Left >= plst.Left + plst.width - 20 And Me.Left <= plst.Left + plst.width + 20 Then
 Me.Left = plst.Left + plst.width
 attach3 = True
ElseIf plst.Top >= Me.Top + Me.height - 20 And plst.Top <= Me.Top + Me.height + 20 Then
 plst.Top = Me.Top + Me.height
 attach3 = True
 ElseIf Me.Top >= plst.Top + plst.height - 20 And Me.Top <= plst.Top + plst.height + 20 Then
 Me.Top = plst.Top + plst.height
 attach3 = True

'plst.Left = Me.Left + Me.width
'attach1 = True

Else
attach3 = False
End If
End If
k = GetCursorPos(m)
initialypos = m.Y - Me.Top / 12000 * 800
initialxpos = m.X - Me.Left / 12000 * 800
initialpx = m.X - plst.Left / 12000 * 800
initialpy = m.Y - plst.Top / 12000 * 800

'initialypos = m.Y - Me.Top / 12000 * 800

'initialxpos = m.X - Me.Left / 12000 * 800

If Me.WindowState = 2 And Button = 1 And MediaPlayer1.DisplaySize <> mpFullScreen Then
If ckk = True Then
k = SetWindowPos(musicsystem.hwnd, -1, musicsystem.Left / 12000 * 800, musicsystem.Top / 9000 * 600, musicsystem.width / 12000 * 800, musicsystem.height / 9000 * 600, &H40)
'k = SetWindowPos(Me.hwnd, -2, Me.Left / 12000 * 800, Me.Top / 9000 * 600, Me.width / 12000 * 800, Me.height / 9000 * 600, &H40)

ckk = False
Else
'k = SetWindowPos(musicsystem.hwnd, -2, musicsystem.Left / 12000 * 800, musicsystem.Top / 9000 * 600, musicsystem.width / 12000 * 800, musicsystem.height / 9000 * 600, &H40)
k = SetWindowPos(Me.hwnd, -1, Me.Left / 12000 * 800, Me.Top / 9000 * 600, Me.width / 12000 * 800, Me.height / 9000 * 600, &H40)
ckk = True
text1.SetFocus
End If
End If
'

1:
End Sub

Public Sub MediaPlayer1_MouseMove(Button As Integer, ShiftState As Integer, X As Single, Y As Single)
Dim k
Dim m As POINTAPI
If moovevideo = True Then
k = GetCursorPos(m)
'm.X = (X - Me.Left) * 277 / 4110
m.X = m.X - initialxpos
m.Y = m.Y - initialypos

If Button = 1 Then
k = MoveWindow(Me.hwnd, m.X, m.Y, Me.width / 12000 * 800, Me.height / 12000 * 800, 1)
If attach3 = True Then k = MoveWindow(plst.hwnd, m.X - initialpx + initialxpos, m.Y - initialpy + initialypos, plst.width / 12000 * 800, plst.height / 12000 * 800, 1)
End If
'If Form1.WindowState = 2 Then Timer1.Enabled = True
End If
text1.SetFocus
End Sub

Public Sub MediaPlayer1_MouseUp(Button As Integer, ShiftState As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
'Form1.Height = Form1.Height - 330
'Form1.Top = Form1.Top + 320
'MediaPlayer1.Height = MediaPlayer1.Height - 330


'MediaPlayer1.Height = MediaPlayer1.Height - 360
PopupMenu property
ElseIf Button = vbLeftButton Then
 If attach1 <> True Then

If plst.Left >= Me.Left + Me.width - 20 And plst.Left <= Me.Left + Me.width + 20 Then
 plst.Left = Me.Left + Me.width
 attach3 = True
 ElseIf Me.Left >= plst.Left + plst.width - 20 And Me.Left <= plst.Left + plst.width + 20 Then
 Me.Left = plst.Left + plst.width
 attach3 = True
ElseIf plst.Top >= Me.Top + Me.height - 20 And plst.Top <= Me.Top + Me.height + 20 Then
 plst.Top = Me.Top + Me.height
 attach3 = True
 ElseIf Me.Top >= plst.Top + plst.height - 20 And Me.Top <= plst.Top + plst.height + 20 Then
 Me.Top = plst.Top + plst.height
 attach3 = True

'plst.Left = Me.Left + Me.width
'attach1 = True

Else
attach3 = False
End If
End If
End If
moovevideo = False
End Sub

Public Sub minimise_Click()
Form1.WindowState = 1
End Sub

Public Sub moving_Click()
On Error GoTo 1

If Me.Caption = "" Then
Me.height = Me.height - 600
Me.Top = Me.Top + 590
Me.Caption = "Mahesh's video"

'MediaPlayer1.Height = MediaPlayer1.Height - 330
End If
1:
End Sub

Private Sub MediaPlayer1_Warning(ByVal WarningType As Long, ByVal Param As Long, ByVal Description As String)
Exit Sub
End Sub

Public Sub playpause_Click()
If Form1.MediaPlayer1.PlayState = mpStopped Then
'musicsystem.cbutton(2).Enabled = True
Else
If Form1.MediaPlayer1.PlayState = mpPlaying Then
Form1.MediaPlayer1.Pause
'musicsystem.cbutton(2).Picture = musicsystem.cbuttonE(2).Picture
'musicsystem.cbutton(2).Enabled = True
Else
If Form1.MediaPlayer1.OpenState Then
Form1.MediaPlayer1.Play
'musicsystem.cbutton(2).Picture = musicsystem.cbuttonD(2).Picture
'musicsystem.cbutton(2).Enabled = False

End If
End If
End If
End Sub






Private Sub ratio43_Click()
If MediaPlayer1.DisplaySize = mpFullScreen Then Exit Sub

If ratio43.Checked = True Then
ratio43.Checked = False
ratio = False
Else
ratio43.Checked = True
ratio = True
End If
End Sub


Private Sub text1_GotFocus()
text1.Value = musicsystem.Slider3.Value
End Sub

Private Sub Text1_KeyDown(Keycode As Integer, Shift As Integer)

On Error Resume Next
'On Error GoTo 9
If Keycode = 112 Then
SendKeys "%+{F4}", True
End If
 If Keycode = 39 Then
   If text1.Value >= text1.min + 1 Then text1.Value = text1.Value - 1
 Call text1_Scroll
 If Form1.MediaPlayer1.CurrentPosition < Form1.MediaPlayer1.Duration - 20 Then
 Form1.MediaPlayer1.CurrentPosition = Form1.MediaPlayer1.CurrentPosition + 20
 Else
 'Form1.MediaPlayer1.Stop
 Form1.MediaPlayer1.CurrentPosition = Form1.MediaPlayer1.Duration - 0.5
 End If
  musicsystem.Imgbar.Left = musicsystem.Image8(1).Left + Form1.MediaPlayer1.CurrentPosition / Form1.MediaPlayer1.Duration * 3270
ElseIf Keycode = 37 Then

If text1.Value <= text1.max - 1 Then text1.Value = text1.Value + 1
  Call text1_Scroll
 If Form1.MediaPlayer1.CurrentPosition > 20 Then
   Form1.MediaPlayer1.CurrentPosition = Form1.MediaPlayer1.CurrentPosition - 20
 Else
   Form1.MediaPlayer1.CurrentPosition = 0
 End If
  musicsystem.Imgbar.Left = musicsystem.Image8(1).Left + Form1.MediaPlayer1.CurrentPosition / Form1.MediaPlayer1.Duration * 3270
End If
'If KeyCode = 66 And musicsystem.cbutton(0).Enabled = True Then Call musicsystem.cbutton_Click(0)
'If KeyCode = 78 And musicsystem.cbutton(1).Enabled = True Then Call musicsystem.cbutton_Click(1)
'MediaPlayer1.SetFocus
If Keycode = 78 Then
Call musicsystem.cbutton_Click(1)
ElseIf Keycode = 66 Then
Call musicsystem.cbutton_Click(0)
ElseIf Keycode = 79 Then
Call musicsystem.cbutton_Click(7)
ElseIf Keycode = 83 Then
Call musicsystem.cbutton_Click(6)
ElseIf Keycode = 80 Then
If Form1.MediaPlayer1.PlayState <> mpPlaying Then
Call musicsystem.cbutton_Click(2)
Else
Call musicsystem.cbutton_Click(3)
End If
End If

'**** the following event is done to reverse the action of left-right keys on slider
'*** slider default action takes place after keydown event so max variable is used

If Keycode = 38 Then
If text1.Value <= text1.max - 2 Then
max = False
 text1.Value = text1.Value + 2
 'Call text1_Scroll
Else
max = True
End If
'Call musicsystem.Image10_MouseDown(1, 0, musicsystem.volumebar.Left - 100 - musicsystem.Image10.Left, musicsystem.Image10.Top + 40)
'text1.SetFocus
End If

If Keycode = 40 Then
max = False

If text1.Value >= text1.min + 2 Then
 text1.Value = text1.Value - 2
'Call text1_Scroll
Else
text1.Value = text1.min
End If
End If

If Keycode = 70 Then
Call fullscreen_Click
End If
'Form1.Show
text1.SetFocus

9:
End Sub

Private Sub text1_KeyPress(Keycode As Integer)

If Keycode = 13 Then
If MediaPlayer1.FileName <> "" Then MediaPlayer1.Play
End If
End Sub

Private Sub text1_Scroll()
If max = True Then text1.Value = text1.max
musicsystem.Slider3.Value = text1.Value
'Label2.Caption = text1.Value
Call musicsystem.Slider3_Scroll
End Sub



Private Sub Timer2_Timer()
On Error GoTo 1
If musicsystem.Slider2.Value >= musicsystem.Slider2.max - 2 Then
Timer2.Enabled = False
Exit Sub
Else
musicsystem.Slider2.Value = musicsystem.Slider2.Value + 2
Form1.MediaPlayer1.CurrentPosition = musicsystem.Slider2.Value * Form1.MediaPlayer1.Duration / musicsystem.Slider2.max
musicsystem.Imgbar.Left = musicsystem.Image8(1).Left + Form1.MediaPlayer1.CurrentPosition / Form1.MediaPlayer1.Duration * 3270 - 30
End If
1:
End Sub

Private Sub Timer3_Timer()
On Error GoTo 1
If musicsystem.Slider2.Value <= 2 Then
Timer2.Enabled = False
Exit Sub
Else
musicsystem.Slider2.Value = musicsystem.Slider2.Value - 2
Form1.MediaPlayer1.CurrentPosition = musicsystem.Slider2.Value * Form1.MediaPlayer1.Duration / musicsystem.Slider2.max
musicsystem.Imgbar.Left = musicsystem.Image8(1).Left + Form1.MediaPlayer1.CurrentPosition / Form1.MediaPlayer1.Duration * 3270 - 30
End If
1:

End Sub

Private Sub uptime_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
uptime.Picture = updown.Picture
Timer2.Enabled = True
End Sub

Private Sub uptime_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
uptime.Picture = upup.Picture
Timer2.Enabled = False
End Sub

Private Sub volmute_Click()
If MediaPlayer1.Mute = False Then
MediaPlayer1.Mute = True
Else
MediaPlayer1.Mute = False
End If
End Sub

Private Sub volup_Click()

If musicsystem.Slider3.Value < 21 Then
musicsystem.Slider3.Value = musicsystem.Slider3.Value + 1
volumebar.Left = 1550 + Slider3.Value * 40
musicsystem.volinf.Caption = "Volume" + str((musicsystemSlider3.Value - 1) * 5) + " %" '+ Form1.MediaPlayer1.volume) / 25))) + " %"
musicsystem.vol.WaveLevel = (musicsystem.Slider3.Value - 1) * 3275
End If
End Sub
Private Sub voldown_Click()
If musicsystem.Slider3.Value > 1 Then
musicsystem.Slider3.Value = musicsystem.Slider3.Value - 1
volumebar.Left = 1550 + Slider3.Value * 40
musicsystem.volinf.Caption = "Volume" + str((musicsystemSlider3.Value - 1) * 5) + " %" '+ Form1.MediaPlayer1.volume) / 25))) + " %"
musicsystem.vol.WaveLevel = (musicsystem.Slider3.Value - 1) * 3275
End If
End Sub

Public Sub wtop_Click()
Dim k
On Error GoTo 1
If MediaPlayer1.DisplaySize = mpFullScreen Then Exit Sub
If wtop.Checked = False Then
k = SetWindowPos(Me.hwnd, -1, Me.Left / 12000 * 800, Me.Top / 9000 * 600, Me.width / 12000 * 800, Me.height / 9000 * 600, &H40)
wtop.Checked = True
Else
k = SetWindowPos(Me.hwnd, -2, Me.Left / 12000 * 800, Me.Top / 9000 * 600, Me.width / 12000 * 800, Me.height / 9000 * 600, &H40)
wtop.Checked = False
End If
1:
End Sub

Private Sub zoom_Click()
If MediaPlayer1.DisplaySize = mpFullScreen Then Exit Sub

End Sub

Public Sub zoom100_Click()
On Error GoTo 1
Form1.height = 5500
Form1.width = 6285
Call FormResize
1:
End Sub

Public Sub zoom200_Click()
On Error GoTo 1
Form1.height = 8000
Form1.width = 10000
Call FormResize
1:
End Sub

Public Sub zoom50_Click()
On Error GoTo 1
Form1.height = 3730
Form1.width = 4120
Call FormResize
1:
End Sub

Private Sub iright_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim k
Dim m As POINTAPI
k = GetCursorPos(m)
initx = m.X
inity = m.Y
End Sub

Private Sub Iright_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim k
Dim m As POINTAPI
On Error GoTo 1
If Button = 1 Then
k = GetCursorPos(m)
'DoEvents
'k = MoveWindow(Me.hwnd, Me.Left * 800 / 12000, Me.Top * 800 / 12000, (Me.width) / 12000 * 800, (Me.width + 210) / 12000 * 800, 1)
If m.X >= initx + 2 Then

Me.width = Me.width + (m.X - initx) * 12000 / 800
If ratio = True Then Me.height = 3# / 4# * Me.width

initx = m.X
If Me.width >= 10300 Then
Me.width = 10300
If ratio = True Then Me.height = 3# / 4# * Me.width

Call FormResize
Exit Sub
End If
Call FormResize

ElseIf m.X <= initx - 2 Then
If Me.width <= 4125 Then
Me.width = 4125
If ratio = True Then Me.height = 3# / 4# * Me.width
Call FormResize
Exit Sub
End If
Me.width = Me.width + (m.X - initx) * 12000 / 800

If ratio = True Then Me.height = 3# / 4# * Me.width
initx = m.X
Call FormResize
End If
'Me.width = Me.width + (m.X - initx) * 12000 / 800
'Call FormResize

End If
1:

End Sub
Public Sub meclose_Click()
'Unload Form1
'Unload Form2
Unload Me
'End
End Sub

Public Sub meclose_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
meclose.Picture = closedown.Picture

End Sub

Public Sub meclose_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
meclose.Picture = closeup.Picture

End Sub
Public Sub min_Click()
'Me.Caption = ""
Me.WindowState = 1
If attach3 = True Then plst.Hide
End Sub

Public Sub min_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
min.Picture = mindown.Picture
End Sub

Public Sub min_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
min.Picture = minup.Picture
End Sub

Private Sub ileft_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim k
Dim m As POINTAPI
k = GetCursorPos(m)
initx = m.X
inity = m.Y
End Sub

Private Sub Ileft_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim k
Dim m As POINTAPI
On Error GoTo 1
If Button = 1 Then
k = GetCursorPos(m)
'DoEvents
'k = MoveWindow(Me.hwnd, Me.Left * 800 / 12000, Me.Top * 800 / 12000, (Me.width) / 12000 * 800, (Me.width + 210) / 12000 * 800, 1)
If m.X <= initx - 5 Then
Me.Left = Me.Left + (m.X - initx) * 12000 / 800

Me.width = Me.width - (m.X - initx) * 12000 / 800
If ratio = True Then Me.height = 3# / 4# * Me.width

initx = m.X
If Me.width >= 10300 Then
Me.width = 10300
If ratio = True Then Me.height = 3# / 4# * Me.width

Call FormResize
Exit Sub
End If
Call FormResize

ElseIf m.X >= initx + 5 Then
If Me.width <= 800 Then
Me.width = 800
If ratio = True Then Me.height = 3# / 4# * Me.width

Call FormResize
Exit Sub
End If
Me.Left = Me.Left + (m.X - initx) * 12000 / 800

Me.width = Me.width - (m.X - initx) * 12000 / 800
If ratio = True Then Me.height = 3# / 4# * Me.width
initx = m.X
Call FormResize
End If
'Me.width = Me.width + (m.X - initx) * 12000 / 800
'Call FormResize

End If
1:

End Sub

VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form plst 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   1770
   ClientLeft      =   2010
   ClientTop       =   4845
   ClientWidth     =   4125
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   ScaleHeight     =   1770
   ScaleWidth      =   4125
   ShowInTaskbar   =   0   'False
   Begin VB.FileListBox File3 
      Height          =   1260
      Left            =   4230
      Pattern         =   " *.DAT;*.mpg;*.avi ;*.jpg;*.mpeg;*.mp3;*.dat"
      TabIndex        =   20
      Top             =   2100
      Width           =   2595
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   540
      Index           =   1
      Left            =   180
      TabIndex        =   19
      Top             =   990
      Visible         =   0   'False
      Width           =   375
      Begin VB.Image addtracks 
         Height          =   270
         Left            =   30
         Picture         =   "playlist.frx":0000
         ToolTipText     =   "Add files to playlist"
         Top             =   270
         Width           =   345
      End
      Begin VB.Image ddown 
         Height          =   270
         Left            =   1200
         Picture         =   "playlist.frx":0552
         Top             =   180
         Width           =   345
      End
      Begin VB.Image dup 
         Height          =   270
         Left            =   720
         Picture         =   "playlist.frx":0AA4
         Top             =   240
         Width           =   345
      End
      Begin VB.Image fdown 
         Height          =   270
         Left            =   1590
         Picture         =   "playlist.frx":0FF6
         Top             =   180
         Width           =   345
      End
      Begin VB.Image fup 
         Height          =   270
         Left            =   720
         Picture         =   "playlist.frx":1548
         Top             =   30
         Width           =   345
      End
      Begin VB.Image adddir 
         Height          =   270
         Left            =   30
         Picture         =   "playlist.frx":1A9A
         Top             =   0
         Width           =   345
      End
      Begin VB.Image Image28 
         Height          =   540
         Left            =   0
         Picture         =   "playlist.frx":1FEC
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.FileListBox File2 
      Height          =   675
      Left            =   4920
      Pattern         =   " *.DAT;*.mpg;*.avi ;*.jpg;*.mpeg;*.mp3;*.dat"
      TabIndex        =   18
      Top             =   240
      Width           =   2595
   End
   Begin VB.DirListBox Dir1 
      Height          =   315
      Left            =   6210
      TabIndex        =   15
      Top             =   390
      Width           =   585
   End
   Begin VB.FileListBox File1 
      Height          =   480
      Left            =   4590
      Pattern         =   "*.mmpl"
      TabIndex        =   14
      Top             =   1200
      Width           =   1875
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   815
      Index           =   0
      Left            =   3430
      TabIndex        =   5
      Top             =   720
      Visible         =   0   'False
      Width           =   380
      Begin VB.Image lpldown 
         Height          =   270
         Left            =   1320
         Picture         =   "playlist.frx":2ADE
         Top             =   720
         Width           =   375
      End
      Begin VB.Image spldown 
         Height          =   270
         Left            =   1380
         Picture         =   "playlist.frx":3078
         Top             =   390
         Width           =   375
      End
      Begin VB.Image npldown 
         Height          =   270
         Left            =   1290
         Picture         =   "playlist.frx":3612
         Top             =   60
         Width           =   375
      End
      Begin VB.Image splup 
         Height          =   270
         Left            =   930
         Picture         =   "playlist.frx":3BAC
         Top             =   660
         Width           =   375
      End
      Begin VB.Image lplup 
         Height          =   270
         Left            =   900
         Picture         =   "playlist.frx":4146
         Top             =   390
         Width           =   375
      End
      Begin VB.Image nplup 
         Height          =   270
         Left            =   900
         Picture         =   "playlist.frx":46E0
         Top             =   60
         Width           =   375
      End
      Begin VB.Image savepl 
         Height          =   270
         Left            =   0
         Picture         =   "playlist.frx":4C7A
         Top             =   270
         Width           =   375
      End
      Begin VB.Image loadpl 
         Height          =   270
         Left            =   0
         Picture         =   "playlist.frx":5214
         Top             =   540
         Width           =   375
      End
      Begin VB.Image Newpl 
         Height          =   270
         Left            =   0
         Picture         =   "playlist.frx":57AE
         Top             =   0
         Width           =   375
      End
      Begin VB.Image Image14 
         Height          =   810
         Left            =   0
         Picture         =   "playlist.frx":5D48
         Top             =   30
         Width           =   375
      End
      Begin VB.Label savepl1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "   SAVE        LIST"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   5.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   315
         Left            =   930
         TabIndex        =   8
         Top             =   750
         Width           =   345
      End
      Begin VB.Label loadpl1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "   LOAD           LIST"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   5.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   255
         Left            =   840
         TabIndex        =   7
         Top             =   450
         Width           =   405
      End
      Begin VB.Label newpl1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "   NEW            LIST"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   5.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   315
         Left            =   840
         TabIndex        =   6
         Top             =   90
         Width           =   375
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   960
      Left            =   3840
      TabIndex        =   4
      Top             =   210
      Width           =   645
      Begin VB.Image scroll 
         DragIcon        =   "playlist.frx":6D92
         Height          =   360
         Left            =   40
         Picture         =   "playlist.frx":709C
         Top             =   30
         Width           =   120
      End
      Begin VB.Image Image24 
         Height          =   6555
         Left            =   0
         Picture         =   "playlist.frx":731E
         Top             =   0
         Width           =   195
      End
      Begin VB.Image Image23 
         Height          =   5100
         Left            =   90
         Picture         =   "playlist.frx":B7A8
         Top             =   -390
         Width           =   75
      End
      Begin VB.Image Image2 
         Enabled         =   0   'False
         Height          =   6555
         Left            =   180
         Picture         =   "playlist.frx":CD2A
         Top             =   0
         Width           =   105
      End
   End
   Begin VB.Frame frmpl 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   765
      Left            =   0
      TabIndex        =   0
      Top             =   1170
      Width           =   4215
      Begin VB.Image showaddmenu 
         Height          =   270
         Left            =   210
         Picture         =   "playlist.frx":F664
         ToolTipText     =   "Add files to playlist"
         Top             =   90
         Width           =   330
      End
      Begin VB.Image showplmenu 
         Height          =   285
         Left            =   3480
         Picture         =   "playlist.frx":FB6E
         Top             =   90
         Width           =   330
      End
      Begin VB.Label image5 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Height          =   105
         Left            =   0
         MousePointer    =   7  'Size N S
         TabIndex        =   17
         Top             =   450
         Width           =   4155
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "TRACK"
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
         Left            =   1950
         TabIndex        =   16
         ToolTipText     =   "Position of playing track in playlist"
         Top             =   75
         Width           =   420
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   " 1 / 15"
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
         Left            =   2400
         TabIndex        =   10
         ToolTipText     =   "Position of playing track in playlist"
         Top             =   75
         Width           =   285
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "00:00"
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
         Height          =   135
         Left            =   3000
         TabIndex        =   9
         ToolTipText     =   "Time elapsed"
         Top             =   75
         Width           =   375
      End
      Begin VB.Image Image10 
         Height          =   135
         Left            =   1950
         Picture         =   "playlist.frx":100BC
         Stretch         =   -1  'True
         Top             =   90
         Width           =   1005
      End
      Begin VB.Label Label17 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "00:00"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   165
         Left            =   2910
         TabIndex        =   13
         ToolTipText     =   "Position of playing track in playlist"
         Top             =   270
         Width           =   405
      End
      Begin VB.Image Image8 
         Height          =   90
         Left            =   3060
         Picture         =   "playlist.frx":1039E
         Top             =   300
         Width           =   105
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00404040&
         BorderWidth     =   2
         X1              =   4200
         X2              =   4200
         Y1              =   60
         Y2              =   600
      End
      Begin VB.Image Image21 
         Height          =   795
         Left            =   4200
         Picture         =   "playlist.frx":10470
         Top             =   420
         Width           =   75
      End
      Begin VB.Image deletetrack 
         Height          =   270
         Left            =   660
         Picture         =   "playlist.frx":10802
         ToolTipText     =   "Delete files from playlist"
         Top             =   90
         Width           =   315
      End
      Begin VB.Image uptrack 
         Height          =   270
         Left            =   1080
         Picture         =   "playlist.frx":10CC4
         ToolTipText     =   "Move up the track"
         Top             =   90
         Width           =   330
      End
      Begin VB.Image downtrack 
         Height          =   270
         Left            =   1500
         Picture         =   "playlist.frx":111CE
         ToolTipText     =   "Move down the track"
         Top             =   90
         Width           =   390
      End
      Begin VB.Image Image11 
         Height          =   120
         Left            =   1950
         Picture         =   "playlist.frx":117B0
         Top             =   300
         Width           =   105
      End
      Begin VB.Image Image12 
         Height          =   120
         Left            =   2100
         Picture         =   "playlist.frx":118B2
         Top             =   300
         Width           =   90
      End
      Begin VB.Image Image13 
         Height          =   120
         Left            =   2250
         Picture         =   "playlist.frx":11994
         Top             =   300
         Width           =   135
      End
      Begin VB.Image Image15 
         Height          =   105
         Left            =   2400
         Picture         =   "playlist.frx":11AB6
         Top             =   300
         Width           =   105
      End
      Begin VB.Image Image16 
         Height          =   105
         Left            =   2520
         Picture         =   "playlist.frx":11BA0
         Top             =   300
         Width           =   105
      End
      Begin VB.Image Image17 
         Height          =   120
         Left            =   2640
         Picture         =   "playlist.frx":11C8A
         Top             =   300
         Width           =   135
      End
      Begin VB.Line Line23 
         BorderColor     =   &H00404040&
         X1              =   4140
         X2              =   4140
         Y1              =   2190
         Y2              =   1530
      End
      Begin VB.Image Image1 
         Height          =   570
         Left            =   0
         Picture         =   "playlist.frx":11DAC
         Top             =   -30
         Width           =   4125
      End
      Begin VB.Image Image7 
         Height          =   360
         Left            =   -60
         Picture         =   "playlist.frx":198D6
         Top             =   -60
         Width           =   105
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   990
      Left            =   4070
      TabIndex        =   12
      Top             =   1170
      Width           =   375
      Begin VB.Line Line6 
         BorderColor     =   &H00C0C0C0&
         X1              =   60
         X2              =   60
         Y1              =   0
         Y2              =   990
      End
      Begin VB.Image Image9 
         Height          =   6690
         Left            =   0
         Picture         =   "playlist.frx":19B58
         Top             =   -390
         Width           =   120
      End
   End
   Begin VB.Frame lblplmenu 
      BackColor       =   &H00808080&
      Caption         =   "Frame6"
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   6000
      TabIndex        =   2
      Top             =   960
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   2115
      Left            =   8490
      TabIndex        =   1
      Top             =   810
      Width           =   510
      _ExtentX        =   900
      _ExtentY        =   3731
      _Version        =   393216
      Orientation     =   1
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   570
      Top             =   2220
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "*.dat"
      DialogTitle     =   "Open file"
      Filter          =   $"playlist.frx":1C56A
      FilterIndex     =   4
      InitDir         =   "e:\"
      MaxFileSize     =   22000
   End
   Begin MSComDlg.CommonDialog cd2 
      Left            =   1320
      Top             =   2190
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      FileName        =   "maheshpl"
      Filter          =   "Playlistfiles(*.mmpl)|*.mmpl"
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   990
      IntegralHeight  =   0   'False
      ItemData        =   "playlist.frx":1C672
      Left            =   180
      List            =   "playlist.frx":1C674
      TabIndex        =   3
      Top             =   300
      Width           =   4245
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   0
      X2              =   4190
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Image closedown 
      Height          =   150
      Left            =   7230
      Picture         =   "playlist.frx":1C676
      Top             =   3210
      Width           =   150
   End
   Begin VB.Image closeup 
      Height          =   135
      Left            =   7440
      Picture         =   "playlist.frx":1C7F8
      Top             =   3180
      Width           =   135
   End
   Begin VB.Image meclose 
      Height          =   135
      Left            =   3950
      Picture         =   "playlist.frx":1C936
      ToolTipText     =   "Close"
      Top             =   30
      Width           =   135
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      X1              =   4500
      X2              =   6510
      Y1              =   2310
      Y2              =   2850
   End
   Begin VB.Image Image4 
      Height          =   105
      Left            =   840
      Top             =   1110
      Width           =   225
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00000000&
      BorderWidth     =   3
      X1              =   0
      X2              =   4120
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Image Image3 
      Height          =   180
      Left            =   5000
      Picture         =   "playlist.frx":1CA74
      Stretch         =   -1  'True
      Top             =   0
      Width           =   360
   End
   Begin VB.Image cbuttonE 
      Height          =   285
      Index           =   7
      Left            =   6360
      Picture         =   "playlist.frx":1CF6E
      Top             =   2220
      Width           =   330
   End
   Begin VB.Line Line4 
      X1              =   4140
      X2              =   4140
      Y1              =   0
      Y2              =   600
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "PLAYLIST"
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
      Left            =   1770
      TabIndex        =   11
      Top             =   30
      Width           =   555
   End
   Begin VB.Shape Shape8 
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   4560
      Top             =   2580
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.Image cbuttonD 
      Height          =   270
      Index           =   7
      Left            =   7680
      Picture         =   "playlist.frx":1D4BC
      Top             =   1830
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgaddtracksup 
      Height          =   270
      Left            =   3540
      Picture         =   "playlist.frx":1DA0E
      Top             =   2970
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Image imgaddtracksdown 
      Height          =   270
      Left            =   3540
      Picture         =   "playlist.frx":1DF18
      Top             =   3300
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Image imgdeletetrackup 
      Height          =   270
      Left            =   3900
      Picture         =   "playlist.frx":1E422
      Top             =   2970
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image imgdeletetrackdown 
      Height          =   270
      Left            =   3900
      Picture         =   "playlist.frx":1E8E4
      Top             =   3300
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Image imguptrackup 
      Height          =   270
      Left            =   4620
      Picture         =   "playlist.frx":1EDEE
      Top             =   3090
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Image imguptrackdown 
      Height          =   270
      Left            =   5880
      Picture         =   "playlist.frx":1F2F8
      Top             =   2940
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Image imgdowntrackdown 
      Height          =   270
      Left            =   2550
      Picture         =   "playlist.frx":1F802
      Top             =   2670
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Image imgdowntrackup 
      Height          =   270
      Left            =   3030
      Picture         =   "playlist.frx":1FDE4
      Top             =   2040
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Image scrollup 
      DragIcon        =   "playlist.frx":203C6
      DragMode        =   1  'Automatic
      Height          =   255
      Left            =   5070
      Picture         =   "playlist.frx":206D0
      Top             =   480
      Width           =   105
   End
   Begin VB.Image scrolldown 
      DragIcon        =   "playlist.frx":208AA
      DragMode        =   1  'Automatic
      Height          =   255
      Left            =   4590
      Picture         =   "playlist.frx":20BB4
      Top             =   0
      Width           =   105
   End
   Begin VB.Image Image18 
      Height          =   285
      Left            =   0
      Picture         =   "playlist.frx":20D8E
      Top             =   0
      Width           =   4140
   End
   Begin VB.Image Image6 
      Enabled         =   0   'False
      Height          =   6450
      Left            =   0
      Picture         =   "playlist.frx":24B44
      Top             =   150
      Width           =   180
   End
End
Attribute VB_Name = "plst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpbrowseinfo As browseinfo) As Long
Private Declare Function SHget Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal folder As String) As Long
Private Declare Sub DragAcceptFiles Lib "shell32.dll" _
        (ByVal hwnd As Long, _
         ByVal fAccept As Long)
         

Private Type browseinfo
 howner As Long
 pidlroot As Long
 pszdisplayname As String
 lpsztitle As String
 ulflags As Long
 lpfn As Long
 lParam As Long
 iimage As Long
End Type



' Define variants to hold keyword parsing arrays
' *** Establish constants for the registry functions

Private Sub ResultsImportFile(sFileName As String, media As Boolean)

Dim listname As String
Dim i As Integer
If media = True Then
   Dim temptrack As file
  ' currenttrack = 1 ' = sFileName
   'plst.newpl_Click
    'Open App.path + "\mplayerlist.mmpl" For Random As #1 Len = 155
   totaltracks = totaltracks + 1
      File1.path = ""
     File1.FileName = sFileName

    temptrack.path = sFileName
     listname = str(totaltracks) + ". " + LCase(File1.List(0))
            If Len(listname) >= 34 Then listname = Left(listname, 34) + "..."
     'listname = LCase(listname)

        plst.List1.AddItem (listname)
      temptrack.name = listname 'Str(totaltracks) + ". " + Left(cd1.FileTitle, 30) + Space(7) + tracklength
     i = InStr(listname, ".")
      temptrack.name = LCase(Mid(listname, i + 1, Len(listname)))
      'Put #1, totaltracks, track
      'Close #1
      track(totaltracks) = temptrack
      If totaltracks < 12 Then musicsystem.Label6(totaltracks - 1).Enabled = True
    'Form1.MediaPlayer1.FileName = sFileName
       'Form1.MediaPlayer1.Play
Else
    Open sFileName For Random As #1 Len = 155
    i = 0
    'currentplaylist = App.path + "\mplayerlist.mmpl"
    Do While (1)
        Get #1, i + 1, track(totaltracks + i + 1)
        'On Error GoTo 3
        If track(totaltracks + i + 1).name = "" Then Exit Do
        plst.List1.AddItem str(totaltracks + i + 1) + "." + track(totaltracks + i + 1).name
        i = i + 1
    Loop
3:
       totaltracks = totaltracks + i
       For i = 1 To totaltracks Step 1
        If i <= 12 Then
          musicsystem.Label6(i - 1).Enabled = True
          musicsystem.Label6(i - 1).ForeColor = musicsystem.Label6(15).ForeColor
        End If
       Next
     Close #1

End If
List1.ListIndex = totaltracks - 1
End Sub












Public Sub DroppedFiles(vFileList As Variant)

Dim nLoopCtr            As Integer
Dim sFileName           As String

' *** Loop through the variant array to process
' each dropped file
For nLoopCtr = 0 To UBound(vFileList)
    ' Get the current file name from the array
    sFileName = vFileList(nLoopCtr)
    ' Make sure it is a SAD format file
    If UCase$(Right$(sFileName, 4)) = ".MP3" Or UCase$(Right$(sFileName, 4)) = ".MPG" Or UCase$(Right$(sFileName, 4)) = ".DAT" Or UCase$(Right$(sFileName, 4)) = ".WAV" Or UCase$(Right$(sFileName, 5)) = ".MPEG" Or UCase$(Right$(sFileName, 4)) = ".AVI" Then
      'Form1.MediaPlayer1.FileName = sFileName
        ' It is the correct file type so import it
        Call ResultsImportFile(sFileName, True)
     ElseIf UCase$(Right$(sFileName, 5)) = ".MMPL" Then
        Call ResultsImportFile(sFileName, False)
    End If
Next nLoopCtr

End Sub






 '.FileSystemObject










Public Sub updateplaylist()
Dim i, k As Integer
On Error GoTo 9

k = plst.List1.ListIndex
plst.List1.Clear
'Dim track As file
'Open App.path + "\mplayerlist.mmpl" For Random As 2 Len = 155
'Put #2, plst.list1.ListIndex + 1, removedtrack
For i = 1 To totaltracks Step 1
'Get #2, i, track
plst.List1.AddItem str(i) + "." + track(i).name
Next
'Close #2
On Error GoTo 1
If k < List1.ListCount Then
 List1.ListIndex = k
Else
 List1.ListIndex = List1.ListCount - 1
End If
1:
If totaltracks < 12 Then

For i = totaltracks To 11 Step 1
musicsystem.Label6(i).Enabled = False

'If i <=12Then
'musicsystem.label6(8 - i).ForeColor = RGB(255, 20, 105)

'musicsystem.label6(8 - i).Enabled = False

Next
End If

9:
End Sub





Private Sub adddir_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Frame1(1).Visible = False
Dim b As browseinfo
'b.ulflags = &H1000
On Error GoTo 1
totaltracks = List1.ListCount
Dim k
Dim folder As String
Dim pidl As Long
folder = Space(300)
With b
.howner = hwnd
.pidlroot = 0
.lpsztitle = "mahesh"
.lpsztitle = "Select Media folder for maheshmp Playlist"
End With

pidl = SHBrowseForFolder(b)
k = SHget(ByVal pidl, ByVal folder)
folder = Left(folder, InStr(folder, Chr$(0)) - 1)
File3.path = folder
'Open App.path + "\mplayerlist.mmpl" For Random As #1 Len = 155
Dim temptrack As file
For i = 0 To File3.ListCount - 1 Step 1
temptrack.name = LCase(str(totaltracks + 1) + "." + Left(File3.List(i), 34))
List1.AddItem temptrack.name
temptrack.name = LCase(Left(File3.List(i), 34))
temptrack.path = folder + "\" + File3.List(i)
totaltracks = totaltracks + 1
track(totaltracks) = temptrack
'Put #1, totaltracks, track
Next
'Close #1
'totaltracks = totaltracks + File3.ListCount
For i = 1 To totaltracks Step 1
If i <= 12 Then musicsystem.Label6(i - 1).Enabled = True
Next

scrollunit = sizescroll / totaltracks
If List1.ListIndex <> -1 Then ListIndex = totaltracks - 1
'List1.ListIndex = totaltracks - 1
scroll.Top = scrollunit * totaltracks
Label18.Caption = str(currenttrack) + "/" + str(totaltracks)
Close #1
 Slider1.SetFocus

If totaltracks > 1 Then
 For i = 0 To totaltracks - 1 Step 1
 If i < 12 Then
 musicsystem.Label6(i).Enabled = True
 musicsystem.Label6(i).ForeColor = musicsystem.Label6(15).ForeColor
 End If
 Next
End If
Label18.Caption = str(currenttrack) + "/" + str(totaltracks)
 
1:
'Close #1
musicsystem.Enabled = True
Me.Enabled = True
 Me.Enabled = True
 totaltracks = List1.ListCount
End Sub

Private Sub adddir_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
addtracks.Picture = fup.Picture
adddir.Picture = ddown.Picture
End Sub

Private Sub addtracks_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
addtracks.Picture = fdown.Picture
adddir.Picture = dup.Picture
End Sub

Private Sub closeup_Click()
Me.Hide
End Sub






Private Sub Form_Activate()
'plst.List1.Selected(plst.List1.ListIndex) = False
mini = False
On Error GoTo 1
cmo = False

plst.List1.Selected(currenttrack - 1) = True
1:
End Sub

Private Sub Form_GotFocus()
'Label1.ForeColor = &H80FF80    'RGB(200, 0, 200)
'On Error GoTo 1
'plst.List1.Selected(currenttrack - 1) = True

'1:
End Sub

Private Sub Form_Load()

cd1.Flags = cdlOFNFileMustExist
cd2.Flags = cdlOFNFileMustExist
Dim k
cd2.Flags = &H2 'cdlOFNOverwritePrompt &H2
  cd1.Flags = cdlOFNAllowMultiselect Or cdlOFNExplorer Or cdlOFNNoDereferenceLinks
height_incr = plst.height
'fet top & left& height settings from rgistry

plst.Top = GetSetting(App.EXEName, "Startup", "plstTop", musicsystem.Top + musicsystem.height)
plst.Left = GetSetting(App.EXEName, "Startup", "plstLeft", musicsystem.Left)
plst.height = GetSetting(App.EXEName, "Startup", "plstheight", 1755)
height_incr = plst.height - height_incr
plst.List1.height = plst.List1.height + height_incr
plst.Frame8.height = plst.Frame8.height + height_incr
plst.frmpl.Top = plst.frmpl.Top + height_incr '# * 12000 / 800)
'Image24.height = Image24.height = 210
plst.Frame1(0).Top = plst.Frame1(0).Top + height_incr
plst.Frame1(1).Top = plst.Frame1(1).Top + height_incr
sizescroll = sizescroll + height_incr

mhHookWindow2 = Me.hwnd
k = SetWindowPos(Me.hwnd, -1, Me.Left / 12000 * 800, Me.Top / 9000 * 600, Me.width / 12000 * 800, Me.height / 9000 * 600, &H40)
plst.List1.Visible = True

Call EnableFileDrops(Me)

Dim mlPrevWndProc
' Set the subclassing window message hook
'///mlPrevWndProc = SetWindowLong(mhHookWindow2, _
                -4, _
                AddressOf HookCallback)

' Tell the OS that the specified window accepts
' dropped files
'//Call DragAcceptFiles(mhHookWindow2, True)


End Sub

Private Sub Form_LostFocus()
'Label1.ForeColor = RGB(0, 0, 0)
Frame1(0).Visible = False
Frame1(1).Visible = False
End Sub



Private Sub frmpl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Frame1(0).Visible = False
Frame1(1).Visible = False

End Sub

Private Sub Image1_Click()
Frame1(0).Visible = False
Frame1(1).Visible = False

End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim m As POINTAPI
Dim k
Frame1(0).Visible = False
Frame1(1).Visible = False

k = GetCursorPos(m)
initialypos = m.Y - Me.Top / 12000 * 800
initialxpos = m.X - Me.Left / 12000 * 800
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim k
Dim m As POINTAPI
k = GetCursorPos(m)
'm.X = (X - Me.Left) * 277 / 4110
If Frame1(0).Visible Or Frame1(1).Visible Then
addtracks.Picture = fup.Picture
adddir.Picture = dup.Picture
loadpl.Picture = lplup.Picture
savepl.Picture = splup.Picture
Newpl.Picture = nplup.Picture
End If

If Button = 1 Then
m.X = m.X - initialxpos
m.Y = m.Y - initialypos

k = MoveWindow(Me.hwnd, m.X, m.Y, Me.width / 12000 * 800, Me.height / 12000 * 800, 1)
End If
End Sub

Private Sub Image18_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim m As POINTAPI
Dim k
mooveplst = True
updown = True
k = GetCursorPos(m)
initialypos = m.Y - Me.Top / 12000 * 800
initialxpos = m.X - Me.Left / 12000 * 800
If plst.Left >= Me.Left + Me.width - 40 And plst.Left <= Me.Left + Me.width + 40 Then
 plst.Left = Me.Left + Me.width
 attach1 = True
 ElseIf Me.Left >= plst.Left + plst.width - 40 And Me.Left <= plst.Left + plst.width + 40 Then
  Me.Left = plst.Left + plst.width
 attach1 = True
ElseIf plst.Top >= Me.Top + Me.height - 40 And plst.Top <= Me.Top + Me.height + 40 Then
 plst.Top = Me.Top + Me.height
 attach1 = True
 ElseIf Me.Top >= plst.Top + plst.height - 40 And Me.Top <= plst.Top + plst.height + 40 Then
 Me.Top = plst.Top + plst.height
 attach1 = True

'plst.Left = Me.Left + Me.width
'attach1 = True

Else
attach1 = False
End If



If attach1 = False Then
If plst.Left >= Form1.Left + Form1.width - 40 And plst.Left <= Form1.Left + Form1.width + 40 Then
 plst.Left = Form1.Left + Form1.width
 attach3 = True
 ElseIf Form1.Left >= plst.Left + plst.width - 40 And Form1.Left <= plst.Left + plst.width + 40 Then
 Form1.Left = plst.Left + plst.width
 attach3 = True
ElseIf plst.Top >= Form1.Top + Form1.height - 40 And plst.Top <= Form1.Top + Form1.height + 40 Then
 plst.Top = Form1.Top + Form1.height
 attach3 = True
 ElseIf Form1.Top >= plst.Top + plst.height - 40 And Form1.Top <= plst.Top + plst.height + 40 Then
 Form1.Top = plst.Top + plst.height
 attach3 = True
'plst.Left = form1.Left + form1.width
'attach1 = True
Else
attach3 = False
End If
If attach3 = True Then attach1 = False
End If
End Sub

Public Sub Image18_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim k
Dim m As POINTAPI
k = GetCursorPos(m)
'm.X = (X - Me.Left) * 277 / 4110
If mooveplst = True Then
If Button = 1 Then
m.X = m.X - initialxpos
m.Y = m.Y - initialypos

k = MoveWindow(Me.hwnd, m.X, m.Y, Me.width / 12000 * 800, Me.height / 12000 * 800, 1)
End If
ElseIf updown = True Then
If m.Y <= initialypos - 2 Then
Call uptrack_Click
initialypos = m.Y
End If
If m.Y <= ((plst.Top * 800 / 12000)) Then
k = SetCursorPos(m.X, (((plst.Top) * 800 / 12000) + 3))
initialypos = (((plst.Top) * 800 / 12000) + 3)
End If
End If
End Sub

Private Sub Image18_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
mooveplst = False
updown = False
If plst.Left >= Me.Left + Me.width - 40 And plst.Left <= Me.Left + Me.width + 40 Then
 plst.Left = Me.Left + Me.width
 attach1 = True
 ElseIf Me.Left >= plst.Left + plst.width - 40 And Me.Left <= plst.Left + plst.width + 40 Then
  Me.Left = plst.Left + plst.width
 attach1 = True
ElseIf plst.Top >= Me.Top + Me.height - 40 And plst.Top <= Me.Top + Me.height + 40 Then
 plst.Top = Me.Top + Me.height
 attach1 = True
 ElseIf Me.Top >= plst.Top + plst.height - 40 And Me.Top <= plst.Top + plst.height + 40 Then
 Me.Top = plst.Top + plst.height
 attach1 = True

'plst.Left = Me.Left + Me.width
'attach1 = True

Else
attach1 = False
End If



If attach1 = False Then
If plst.Left >= Form1.Left + Form1.width - 40 And plst.Left <= Form1.Left + Form1.width + 40 Then
 plst.Left = Form1.Left + Form1.width
 attach3 = True
 ElseIf Form1.Left >= plst.Left + plst.width - 40 And Form1.Left <= plst.Left + plst.width + 40 Then
 Form1.Left = plst.Left + plst.width
 attach3 = True
ElseIf plst.Top >= Form1.Top + Form1.height - 40 And plst.Top <= Form1.Top + Form1.height + 40 Then
 plst.Top = Form1.Top + Form1.height
 attach3 = True
 ElseIf Form1.Top >= plst.Top + plst.height - 40 And Form1.Top <= plst.Top + plst.height + 40 Then
 Form1.Top = plst.Top + plst.height
 attach3 = True
'plst.Left = form1.Left + form1.width
'attach1 = True
Else
attach3 = False
End If
If attach3 = True Then attach1 = False

End If
End Sub

Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim m As POINTAPI
Dim k
k = GetCursorPos(m)
initialypos = m.Y - Me.Top / 12000 * 800
initialxpos = m.X - Me.Left / 12000 * 800
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim k
Dim m As POINTAPI
k = GetCursorPos(m)
'm.X = (X - Me.Left) * 277 / 4110

If Button = 1 Then
m.X = m.X - initialxpos
m.Y = m.Y - initialypos

k = MoveWindow(Me.hwnd, m.X, m.Y, Me.width / 12000 * 800, Me.height / 12000 * 800, 1)
End If
End Sub

Private Sub Image2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
mooveplst = False
updown = False
If plst.Left >= Me.Left + Me.width - 40 And plst.Left <= Me.Left + Me.width + 40 Then
 plst.Left = Me.Left + Me.width
 attach1 = True
 ElseIf Me.Left >= plst.Left + plst.width - 40 And Me.Left <= plst.Left + plst.width + 40 Then
  Me.Left = plst.Left + plst.width
 attach1 = True
ElseIf plst.Top >= Me.Top + Me.height - 40 And plst.Top <= Me.Top + Me.height + 40 Then
 plst.Top = Me.Top + Me.height
 attach1 = True
 ElseIf Me.Top >= plst.Top + plst.height - 40 And Me.Top <= plst.Top + plst.height + 40 Then
 Me.Top = plst.Top + plst.height
 attach1 = True

'plst.Left = Me.Left + Me.width
'attach1 = True

Else
attach1 = False
End If



If attach1 = False Then
If plst.Left >= Form1.Left + Form1.width - 40 And plst.Left <= Form1.Left + Form1.width + 40 Then
 plst.Left = Form1.Left + Form1.width
 attach3 = True
 ElseIf Form1.Left >= plst.Left + plst.width - 40 And Form1.Left <= plst.Left + plst.width + 40 Then
 Form1.Left = plst.Left + plst.width
 attach3 = True
ElseIf plst.Top >= Form1.Top + Form1.height - 40 And plst.Top <= Form1.Top + Form1.height + 40 Then
 plst.Top = Form1.Top + Form1.height
 attach3 = True
 ElseIf Form1.Top >= plst.Top + plst.height - 40 And Form1.Top <= plst.Top + plst.height + 40 Then
 Form1.Top = plst.Top + plst.height
 attach3 = True
'plst.Left = form1.Left + form1.width
'attach1 = True
Else
attach3 = False
End If
If attach3 = True Then attach1 = False

End If
End Sub

Public Sub Image24_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
Dim i As Integer
Dim k As Integer
On Error GoTo 1 'If ll = False Then

If Y > scroll.Top Then

i = ((Y - scroll.Top) / scrollunit)

If i >= 1 Then
For k = 0 To i Step 1
scroll.Top = scroll.Top + i '* scrollunit
List1.ListIndex = List1.ListIndex + 1
Next
End If
ElseIf Y < scroll.Top Then
i = Int((-(Y - scroll.Top) / scrollunit))
'For k = 0 To i Step 1
'If i >= 1 Then
scrolltop = Y + 5
scrolltop = scroll.Top
If i >= 1 Then
For k = 0 To i Step 1
scroll.Top = scroll.Top - scrollunit
List1.ListIndex = List1.ListIndex - 1
Next
End If
End If

If scroll.Top <= 70 Then scroll.Top = 70
If scroll.Top >= sizescroll Then scroll.Top = sizescroll + 70
1:
If List1.ListIndex = -1 Then List1.ListIndex = 0

End Sub

Public Sub Image24_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Slider1.Max = (totaltracks)
On Error GoTo 1
scrollunit = sizescroll / (totaltracks)
Call Image24_DragOver(scroll, X + 60, Y, 2)
'Slider1.SetFocus
1:
End Sub

Public Sub Image24_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo 9

Slider1.Value = List1.ListIndex
Slider1.max = (totaltracks)
scrollunit = sizescroll / (totaltracks)
Slider1.SetFocus
Frame1(0).Visible = False
Frame1(1).Visible = False

9:
End Sub

Public Sub Image24_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'll = True
End Sub

Private Sub Image27_Click()
End Sub

Private Sub Image5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim k
resize = True

Dim m As POINTAPI
k = GetCursorPos(m)
initx = m.X
inity = m.Y
Frame1(0).Visible = False
Frame1(1).Visible = False

End Sub

Private Sub Image5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim k
Dim m As POINTAPI
'Frame1(0).Visible = False
'Frame1(1).Visible = False
If Frame1(0).Visible Or Frame1(1).Visible Then
addtracks.Picture = fup.Picture
adddir.Picture = dup.Picture
loadpl.Picture = lplup.Picture
savepl.Picture = splup.Picture
Newpl.Picture = nplup.Picture
End If

'On Error GoTo 12
If resize = True Then
If Button = 1 Then
k = GetCursorPos(m)
If m.Y > inity + (210# * 800 / 12000) Then

If plst.height >= 7000 Then Exit Sub
'k = MoveWindow(Me.hwnd, Me.Left * 800 / 12000, Me.Top * 800 / 12000, (Me.width) / 12000 * 800, (Me.height + 210) / 12000 * 800, 1)

plst.height = plst.height + (210)

List1.height = List1.height + 210
Frame8.height = Frame8.height + 210
sizescroll = sizescroll + 210

frmpl.Top = frmpl.Top + 210 '# * 12000 / 800)
'Image24.height = Image24.height = 210
Frame1(0).Top = Frame1(0).Top + 210
Frame1(1).Top = Frame1(1).Top + 210

'Line1.Y2 = Line1.Y2 + 210
'Line7.Y1 = Line7.Y1 + 210
'Line7.Y2 = Line7.Y2 + 210

inity = m.Y
ElseIf m.Y < inity - (210# * 800 / 12000) Then
If plst.height < 1200 Then Exit Sub
'k = MoveWindow(Me.hwnd, Me.Left * 800 / 12000, Me.Top * 800 / 12000, (Me.width) / 12000 * 800, (Me.height - 210) / 12000 * 800, 1)
plst.height = plst.height - (210)
List1.height = List1.height - 210
Frame8.height = Frame8.height - 210
sizescroll = sizescroll - 210
frmpl.Top = frmpl.Top - 210 '# * 12000 / 800)
'Image24.height = Image24.height = 210
Frame1(0).Top = Frame1(0).Top - 210
Frame1(1).Top = Frame1(1).Top - 210

Frame2.Top = Frame2.Top - 210
'Line1.Y2 = Line1.Y2 - 210
'Line7.Y1 = Line7.Y1 - 210
'Line7.Y2 = Line7.Y2 - 210

inity = m.Y
End If
End If
End If
Image6.Top = 0
12:
If plst.height < musicsystem.height + 50 And plst.height > musicsystem.height - 50 Then plst.height = musicsystem.height
End Sub

Private Sub image5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
resize = False
End Sub

Private Sub Image6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim m As POINTAPI
Dim k
k = GetCursorPos(m)
initialypos = m.Y - Me.Top / 12000 * 800
initialxpos = m.X - Me.Left / 12000 * 800
End Sub

Private Sub Image6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim k
Dim m As POINTAPI
k = GetCursorPos(m)
'm.X = (X - Me.Left) * 277 / 4110

If Button = 1 Then
m.X = m.X - initialxpos
m.Y = m.Y - initialypos

k = MoveWindow(Me.hwnd, m.X, m.Y, Me.width / 12000 * 800, Me.height / 12000 * 800, 1)
End If
End Sub

Private Sub Image6_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
mooveplst = False
updown = False
If plst.Left >= Me.Left + Me.width - 40 And plst.Left <= Me.Left + Me.width + 40 Then
 plst.Left = Me.Left + Me.width
 attach1 = True
 ElseIf Me.Left >= plst.Left + plst.width - 40 And Me.Left <= plst.Left + plst.width + 40 Then
  Me.Left = plst.Left + plst.width
 attach1 = True
ElseIf plst.Top >= Me.Top + Me.height - 40 And plst.Top <= Me.Top + Me.height + 40 Then
 plst.Top = Me.Top + Me.height
 attach1 = True
 ElseIf Me.Top >= plst.Top + plst.height - 40 And Me.Top <= plst.Top + plst.height + 40 Then
 Me.Top = plst.Top + plst.height
 attach1 = True

'plst.Left = Me.Left + Me.width
'attach1 = True

Else
attach1 = False
End If



If attach1 = False Then
If plst.Left >= Form1.Left + Form1.width - 40 And plst.Left <= Form1.Left + Form1.width + 40 Then
 plst.Left = Form1.Left + Form1.width
 attach3 = True
 ElseIf Form1.Left >= plst.Left + plst.width - 40 And Form1.Left <= plst.Left + plst.width + 40 Then
 Form1.Left = plst.Left + plst.width
 attach3 = True
ElseIf plst.Top >= Form1.Top + Form1.height - 40 And plst.Top <= Form1.Top + Form1.height + 40 Then
 plst.Top = Form1.Top + Form1.height
 attach3 = True
 ElseIf Form1.Top >= plst.Top + plst.height - 40 And Form1.Top <= plst.Top + plst.height + 40 Then
 Form1.Top = plst.Top + plst.height
 attach3 = True
'plst.Left = form1.Left + form1.width
'attach1 = True
Else
attach3 = False
End If
If attach3 = True Then attach1 = False
End If

End Sub


Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Image18_MouseDown(Button, Shift, Label1.Left + X, Y)

End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Call Image18_MouseMove(Button, Shift, Label1.Left + X, Y)
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
mooveplst = False
moovemain = False
updown = True
resize = False
Dim k
Dim m As POINTAPI
k = GetCursorPos(m)
initialypos = m.Y '- Me.Top / 12000 * 800
initialxpos = m.X '- Me.Left / 12000 * 800
Frame1(0).Visible = False
Frame1(1).Visible = False
Frame1(1).Visible = False

End Sub

Private Sub List1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
updown = False
End Sub

Private Sub meclose_Click()
Me.Hide
mini = True
End Sub

Private Sub scroll_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim m As POINTAPI
Dim k
On Error GoTo 1
k = GetCursorPos(m)
initialypos = m.Y '- Me.Top / 12000 * 800

initialxpos = m.X '- Me.Left / 12000 * 800

scrolltop = scroll.Top

Slider1.Value = List1.ListIndex
Slider1.max = (totaltracks)
scrollunit = sizescroll / (totaltracks)
Slider1.SetFocus
cmo = True
1:
End Sub

Public Sub scroll_MouseMove(Button As Integer, Shift As Integer, X As Single, Y1 As Single)
On Error GoTo 9
Frame1(0).Visible = False
Frame1(1).Visible = False

'Slider1.SetFocus
Dim m As POINTAPI
Dim k
Dim i As Integer
'On Error GoTo 1
If Button = 1 Then
k = GetCursorPos(m)
If Abs(m.Y - initialypos) >= 2 Then
scroll.Top = scroll.Top + ((m.Y - initialypos) * 12000 / 800#)  '+ 20
Y = m.Y * 12000 / 800 - Frame8.Top
initialypos = m.Y
i = Int((scroll.Top) / (scrollunit))
If i >= 1 Then List1.ListIndex = i  ' - 1

If scroll.Top <= 10 Then List1.ListIndex = 0

End If
'DoEvents

If scroll.Top <= 10 Then scroll.Top = 10
If scroll.Top >= sizescroll + 60 Then
scroll.Top = sizescroll + 70
List1.ListIndex = totaltracks - 1
End If
End If

9:
End Sub

Public Sub List1_Click()
On Error GoTo 9

scrollunit = sizescroll / (totaltracks)
If List1.ListIndex > 1 Then
If cmo = False Then scroll.Top = (List1.ListIndex + 1) * scrollunit + 90
Else
If cmo = False Then scroll.Top = (List1.ListIndex) * scrollunit + 90
End If
'Call scroll_Click '
9:
End Sub

Public Sub List1_GotFocus()
Slider1.SetFocus
End Sub

Public Sub List1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim k
Dim m As POINTAPI
If Frame1(0).Visible = True Then Frame1(0).Visible = False
If Frame1(1).Visible = True Then Frame1(1).Visible = False

Slider1.SetFocus
If updown = True Then
If Button = 1 Then

k = GetCursorPos(m)
If m.Y - initialypos > 11 Then
Call downtrack_Click
initialypos = m.Y
ElseIf initialypos - m.Y > 11 Then
Call uptrack_Click
initialypos = m.Y
End If
k = GetCursorPos(m)

'End If
If m.Y <= (((plst.Top + 240) * 800 / 12000) + 10) And m.Y >= (((plst.Top + 240) * 800 / 12000) - 10) And m.Y <= initialypos Then
If m.Y <= initialypos - 2 Then
Call uptrack_Click
initialypos = m.Y
End If
End If
If m.Y >= (((plst.Top + List1.height + 1) * 800 / 12000)) Then
If m.Y >= initialypos + 1 Then
Call downtrack_Click
initialypos = m.Y
k = SetCursorPos(m.X, (((plst.Top + List1.height + 1) * 800 / 12000) + 3))
initialypos = (((plst.Top + List1.height + 1) * 800 / 12000) + 3)
End If
'If m.y >= (((plst.Top + List1.height + 2) * 800 / 12000)) Then
End If

End If
End If

End Sub



Public Sub loadpl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
savepl.Picture = splup.Picture
Newpl.Picture = nplup.Picture
loadpl.Picture = lpldown.Picture
End Sub
Public Sub addtracks_Click()
Dim temptrack As file
Dim k
Dim listname As String
Dim dummylist As String
Dim tracklength As String
Dim kl As String
Dim spposition, filenames, pathname, trackno, i
On Error GoTo 2
'DoEvents
cd1.DialogTitle = "Add Files"
Dim a, b As Long
'a = musicsystem.Top
'b = musicsystem.Left
'c = Me.Top
'd = Me.Left
Me.Enabled = False
musicsystem.Enabled = False
cd1.Action = 1
'Open App.path + "\mplayerlist.mmpl" For Random As #1 Len = 155
filenames = cd1.FileName
'Text1.Text = cd1.filename
spposition = InStr(filenames, Chr(0))
If (spposition = 0) Then
totaltracks = totaltracks + 1
temptrack.path = cd1.FileName
listname = str(totaltracks) + ". " + cd1.FileTitle
If Len(listname) >= 34 Then listname = Left(listname, 34) + "..."
listname = LCase(listname)
'track.name = LCase(cd1.FileTitle)

'musicsystem.MediaPlayer2.FileName = track.path
'dummylist = listname
'listname = listname + " (" + tracklength + ")"

List1.AddItem (listname)
temptrack.name = listname 'Str(totaltracks) + ". " + Left(cd1.FileTitle, 30) + Space(7) + tracklength
i = InStr(listname, ".")
temptrack.name = LCase(Mid(listname, i + 1, Len(listname)))
track(totaltracks) = temptrack
GoTo 2
End If
    pathname = Left(filenames, spposition - 1)
    If Right(pathname, 1) <> "\" Then pathname = pathname + "\"
      filenames = Mid(filenames, spposition + 1)
' then extract each space delimited file name
    If Len(filenames) = 0 Then
       ' List1.AddItem "No files selected"
        Exit Sub
    Else
        spposition = InStr(filenames, Chr(0))
        While spposition > 0
        totaltracks = totaltracks + 1
            temptrack.path = pathname + Left(filenames, spposition - 1)
            temptrack.name = LCase(Left(filenames, spposition - 1))
            'musicsystem.MediaPlayer2.FileName = track.path
            'tracklength = Format((Int(musicsystem.MediaPlayer2.Duration / 60)), "00") + ":" + Format(Int(Int(musicsystem.MediaPlayer2.Duration Mod 60)), "00")
            listname = str(totaltracks) + ". " + temptrack.name
            'If TextWidth(listname) > TextWidth(Text1.Text) - 817 Then
             If Len(listname) >= 34 Then listname = LCase(str(totaltracks) + ". " + Left(temptrack.name, 34)) + "..."
 
         'listname = listname + " (" + tracklength + ")"
             List1.AddItem listname
              'DoEvents

              i = InStr(listname, ".")
             temptrack.name = LCase(Mid(listname, i + 1, Len(listname)))
              
              track(totaltracks) = temptrack
              
            filenames = Mid(filenames, spposition + 1)
            spposition = InStr(filenames, Chr(0))
        Wend
' Add the last file's name to the list
' (the last file name isn't followed by a space)
        totaltracks = totaltracks + 1
           temptrack.name = LCase(filenames)
            temptrack.path = pathname + (filenames)
            'musicsystem.MediaPlayer2.FileName = track.path
            'tracklength = Format((Int(musicsystem.MediaPlayer2.Duration / 60)), "00") + ":" + Format(Int(Int(musicsystem.MediaPlayer2.Duration Mod 60)), "00")
             
             
             listname = LCase(str(totaltracks) + ". " + LCase(filenames))
             'If TextWidth(listname) > TextWidth(Text1.Text) - 817 Then
              If Len(listname) >= 34 Then listname = LCase(str(totaltracks) + ". " + Left(filenames, 34) + "...")



             'listname = listname + " (" + tracklength + ")"
             List1.AddItem LCase(listname)
               
                     
             i = InStr(listname, ".")
             temptrack.name = LCase(Mid(listname, i + 1, Len(listname)))
              track(totaltracks) = temptrack
  End If
For i = 1 To totaltracks Step 1
If i <= 12 Then musicsystem.Label6(i - 1).Enabled = True
Next

scrollunit = sizescroll / totaltracks
If List1.ListIndex <> -1 Then ListIndex = totaltracks - 1
List1.ListIndex = totaltracks - 1
scroll.Top = scrollunit * totaltracks
Label18.Caption = str(currenttrack) + "/" + str(totaltracks)
Close #1
 Slider1.SetFocus
2:
If totaltracks > 1 Then
 For i = 0 To totaltracks - 1 Step 1
 If i < 12 Then
 musicsystem.Label6(i).Enabled = True
 musicsystem.Label6(i).ForeColor = musicsystem.Label6(15).ForeColor
 End If
 Next
End If
Label18.Caption = str(currenttrack) + "/" + str(totaltracks)
 DoEvents
musicsystem.Enabled = True
Me.Enabled = True
 Me.Enabled = True
musicsystem.Enabled = True
List1.ListIndex = List1.ListCount - 1

End Sub
Public Sub addtracks_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'addtracks.Picture = imgaddtracksdown.Picture
Frame1(0).Visible = False

Call addtracks_Click
Frame1(1).Visible = False
Frame1(0).Visible = False
End Sub

Public Sub addtracks_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'addtracks.Picture = imgaddtracksup.Picture

End Sub

Public Sub deletetrack_Click()
Dim removedtrack As file
'Dim track As file
Dim i As Integer
'On Error GoTo 9

If List1.ListIndex < 0 Then Exit Sub
removedtrack.path = ""
removedtrack.name = ""
totaltracks = totaltracks - 1
If currenttrack >= List1.ListIndex + 1 Then currenttrack = currenttrack - 1

'Open App.path + "\mplayerlist.mmpl" For Random As 2 Len = 155
'Put #2, List1.ListIndex + 1, removedtrack
For i = List1.ListIndex + 1 To totaltracks Step 1
track(i) = track(i + 1)
Next
'Close #2

'If totaltracks > 1 Then
'For i = 0 To 1 Step 1
'musicsystem.cbutton(i).Picture = musicsystem.cbuttonE(i).Picture
'musicsystem.cbutton(i).Enabled = True
'Next
'Else
'For i = 0 To 1 Step 1
'musicsystem.cbutton(i).Picture = musicsystem.cbuttonD(i).Picture
'musicsystem.cbutton(i).Enabled = False
'Next
'End If

updateplaylist
List1.SetFocus
9:
End Sub

Public Sub downtrack_Click()

Dim temptrack As file
Dim track2 As file
Dim i As Integer
'On Error GoTo 9

i = List1.ListIndex
If i >= totaltracks - 1 Then Exit Sub


'swap track with lower one
temptrack = track(i + 2)
track(i + 2) = track(i + 1)
track(i + 1) = temptrack
'assign new names to list
 List1.List(i) = str(i + 1) + "." + track(i + 1).name
 List1.List(i + 1) = str(i + 2) + "." + track(i + 2).name

List1.ListIndex = i + 1
If currenttrack = i + 1 Then
If i < 12 Then musicsystem.Label6(currenttrack - 1).ForeColor = musicsystem.Label6(9).ForeColor
currenttrack = i
'tracklength = Format((Int(Form1.MediaPlayer1.Duration / 60)), "00") + ":" + Format(Int(Int(musicsystem.MediaPlayer2.Duration Mod 60)), "00")

musicsystem.playtrack.Caption = " *** " + (str(i + 2) + "." + track(i + 2).name) ' +' tracklength

'playtrack.Caption = (Str(i) + "." + track2.name)
End If
List1.SetFocus
9:



End Sub

Private Sub scroll_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmo = False
End Sub

Private Sub showaddmenu_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Frame1(1).Visible = True
addtracks.Picture = fdown.Picture
Frame1(0).Visible = False
End Sub

Public Sub uptrack_Click()
Dim temptrack As file
Dim track2 As file
Dim i As Integer
'On Error GoTo 9

i = List1.ListIndex
If i <= 0 Then Exit Sub


'swap track with upper one
temptrack = track(i)
track(i) = track(i + 1)
track(i + 1) = temptrack
'assign new names to list
 List1.List(i) = str(i + 1) + "." + track(i + 1).name
 List1.List(i - 1) = str(i) + "." + track(i).name

List1.ListIndex = i - 1
If currenttrack = i + 1 Then
If i <= 12 Then musicsystem.Label6(currenttrack - 1).ForeColor = musicsystem.Label6(9).ForeColor
currenttrack = i
'tracklength = Format((Int(Form1.MediaPlayer1.Duration / 60)), "00") + ":" + Format(Int(Int(musicsystem.MediaPlayer2.Duration Mod 60)), "00")

musicsystem.playtrack.Caption = "   " + (str(i) + "." + track(i).name) ' +' tracklength

'playtrack.Caption = (Str(i) + "." + track2.name)
End If
List1.SetFocus
9:
End Sub

Public Sub uptrack_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
uptrack.Picture = imguptrackdown.Picture
End Sub

Public Sub uptrack_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
uptrack.Picture = imguptrackup.Picture

End Sub
Public Sub downtrack_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
downtrack.Picture = imgdowntrackdown.Picture
End Sub

Public Sub downtrack_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
downtrack.Picture = imgdowntrackup.Picture

End Sub
Public Sub deletetrack_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
deletetrack.Picture = imgdeletetrackdown.Picture
End Sub

Public Sub deletetrack_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
deletetrack.Picture = imgdeletetrackup.Picture

End Sub
Public Sub Image11_Click()
Call musicsystem.cbutton_Click(0)
End Sub

Public Sub Image12_Click()
Call musicsystem.cbutton_Click(2)

End Sub

Public Sub Image13_Click()
Call musicsystem.cbutton_Click(3)

End Sub

Public Sub Image15_Click()
Call musicsystem.cbutton_Click(6)

End Sub

Public Sub Image16_Click()
Call musicsystem.cbutton_Click(1)

End Sub

Public Sub Image17_Click()
Call musicsystem.cbutton_Click(7)

End Sub

Public Sub Label17_Click()
If showtime = 0 Then
showtime = 1
Label17.ToolTipText = "Time remaining"
Else
showtime = 0
Label17.ToolTipText = "Time elapsed"
End If
End Sub



Public Sub loadpl_Click()
Dim track As file
Dim i As Integer
Frame1(0).Visible = False
Frame1(1).Visible = False

On Error GoTo 1
currentplaylist = cd2.FileName
cd2.DialogTitle = "Open Playlist for Maheshmp"
cd2.Action = 1

Open cd2.FileName For Random As #1 Len = 155
i = 0
List1.Clear
currenttrack = 0
totaltracks = 0
Do While (1)
Get #1, i + 1, track
'On Error GoTo 3
If track.name = "" Then Exit Do
List1.AddItem str(i + 1) + ". " + track.name
i = i + 1
Loop
3:
totaltracks = i
Open App.path + "\mplayerlist.mmpl" For Random As 2 Len = 155
'defaultplaylist = cd2.FileName
If totaltracks >= 1 Then
For i = 1 To totaltracks Step 1
Get #1, i, track
Put #2, i, track
Next
Close #2
End If
Close #2

Close #1
For i = 1 To totaltracks Step 1
If i <= 12 Then musicsystem.Label6(i - 1).Enabled = True
Next

If totaltracks > 1 Then
 For i = 0 To 1 Step 1
' musicsystem.cbutton(i).Picture = musicsystem.cbuttonE(i).Picture
' musicsystem.cbutton(i).Enabled = True
 Next
Else
 For i = 0 To 1 Step 1
  'musicsystem.cbutton(i).Picture = musicsystem.cbuttonD(i).Picture
  'musicsystem.cbutton(i).Enabled = False
 Next
End If
If totaltracks > 0 Then
 For i = 2 To 6 Step 1
 'musicsystem.cbutton(i).Picture = musicsystem.cbuttonE(i).Picture
 'musicsystem.cbutton(i).Enabled = True
 Next
End If

If Form1.MediaPlayer1.PlayState = mpPlaying Then
 'musicsystem.cbutton(3).Picture = musicsystem.cbuttonE(3).Picture
' musicsystem.cbutton(3).Enabled = True
 'musicsystem.cbutton(2).Picture = musicsystem.cbuttonD(2).Picture
' musicsystem.cbutton(2).Enabled = False
Else
 'musicsystem.cbutton(2).Picture = musicsystem.cbuttonE(2).Picture
 'musicsystem.cbutton(2).Enabled = True
 'musicsystem.cbutton(3).Picture = musicsystem.cbuttonD(3).Picture
 'musicsystem.cbutton(3).Enabled = False
End If
If List1.ListIndex = -1 Then List1.ListIndex = 0
Label18.Caption = str(currenttrack) + "/" + str(totaltracks)

1:
End Sub

Public Sub newpl_Click()
Dim i As Integer
'On Error GoTo 9

List1.Clear
Dim k
k = DeleteFile(App.path + "\mplayerlist.mmpl")
Frame1(0).Visible = False
Frame1(1).Visible = False

totaltracks = 0
currenttrack = 1
For i = 0 To 1
'musicsystem.cbutton(i).Picture = musicsystem.cbuttonD(i).Picture
'musicsystem.cbutton(i).Enabled = False
Next
If Form1.MediaPlayer1.FileName = "" Then
For i = 2 To 5
'musicsystem.cbutton(i).Picture = musicsystem.cbuttonD(i).Picture
'musicsystem.cbutton(i).Enabled = False
Next
ElseIf Form1.MediaPlayer1.PlayState = mpPlaying Then
'Call musicsystem.cbutton_Click(3)
End If
currenttrack = 1
totaltracks = 0
For i = 0 To 11 Step 1
musicsystem.Label6(i).Enabled = False
Next
9: 'Frame1(0).Visible = False
Label18.Caption = str(0) + "/" + str(totaltracks)

End Sub
Public Sub savepl_Click()
Dim i As Integer
Dim temptrack As file
Frame1(0).Visible = False
Frame1(1).Visible = False

On Error GoTo 1
cd2.DialogTitle = "Save Playlist for Maheshmp as"

cd2.Action = 2
Open cd2.FileName For Random As #1 Len = 155
'Open App.path + "\mplayerlist.mmpl" For Random As 2 Len = 155
For i = 1 To totaltracks
'Get #2, i, track
Put #1, i, track(i)
Next
temptrack.path = ""
temptrack.name = ""
Put #1, totaltracks + 1, temptrack
'Close #2
Close #1
1:
'Frame1(0).Visible = False
End Sub
Public Sub savepl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
savepl.Picture = spldown.Picture
Newpl.Picture = nplup.Picture
loadpl.Picture = lplup.Picture
End Sub

Public Sub newpl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
savepl.Picture = splup.Picture
Newpl.Picture = npldown.Picture
loadpl.Picture = lplup.Picture
End Sub

Public Sub plclose_Click()


End Sub

Public Sub plclose_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
plclose.Picture = closedown.Picture
End Sub

Public Sub plclose_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
plclose.Picture = closeup.Picture

End Sub
Public Sub showplmenu_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Frame1(0).Visible = True
Dim k
Call loadpl_MouseMove(0, 0, X, Y)
Frame1(1).Visible = False

End Sub

Public Sub slider1_GotFocus()
On Error GoTo 1
scroll.Picture = scrolldown.Picture
plst.Slider1.max = totaltracks
plst.Slider1.Value = plst.List1.ListIndex + 1
scrollunit = sizescroll / totaltracks * 1#
1:
End Sub

Public Sub slider1_KeyDown(Keycode As Integer, Shift As Integer)
On Error Resume Next
If Keycode = 46 Then Call deletetrack_Click
If Keycode = 13 Then Call plst.list1_DblClick
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
'ElseIf KeyCode = 39 Then
'Call musicsystem.cbutton_Click(3)

'Call musicsystem.cbutton_Click(4)
End If
If Keycode = 39 Then
'List1.Selected(List1.ListIndex - 1) = True
Slider1.Value = Slider1.Value - 1
'SendKeys "{up}", True
'List1.ListIndex = List1.ListIndex - 1
If Form1.MediaPlayer1.CurrentPosition < Form1.MediaPlayer1.Duration - 20 Then
Form1.MediaPlayer1.CurrentPosition = Form1.MediaPlayer1.CurrentPosition + 20
Else
'Form1.MediaPlayer1.Stop
Form1.MediaPlayer1.CurrentPosition = Form1.MediaPlayer1.Duration - 0.5
End If

'Call musicsystem.cbutton_Click(5)
'If bar.Width <= Label3.Width - 100 Then bar.Width = bar.Width + 120
'Form1.MediaPlayer1.CurrentPosition = Form1.MediaPlayer1.Duration * bar.Width / Label13.Width
ElseIf Keycode = 37 Then
Slider1.Value = Slider1.Value + 1
'SendKeys "{down}", True
If Form1.MediaPlayer1.CurrentPosition > 20 Then
Form1.MediaPlayer1.CurrentPosition = Form1.MediaPlayer1.CurrentPosition - 20
Else
Form1.MediaPlayer1.CurrentPosition = 0
End If
'List1.ListIndex = List1.ListIndex + 1

'SendKeys "{up}", True
End If
musicsystem.Imgbar.Left = musicsystem.Image8(1).Left + Form1.MediaPlayer1.CurrentPosition / Form1.MediaPlayer1.Duration * 3270 - 30
Slider1.SetFocus
End Sub

Public Sub slider1_LostFocus()
scroll.Picture = scrollup.Picture

End Sub

Public Sub slider1_Scroll()
On Error GoTo 1
'If plst.list1.ListIndex = -1 Then ' Then
'plst.list1.ListIndex = 0
'scroll.Top = 10

'Call plst.list1_Click
'End If
'scroll.Top = plst.Slider1.Value * scrollunit + 20
If plst.Slider1.Value >= 1 Then plst.List1.ListIndex = plst.Slider1.Value - 1
1:
End Sub
Public Sub list1_DblClick()
Dim i, m As Integer
Dim dummy As String
Dim tracklength As String
'Dim track As file
Dim k
'For i = 1 To totaltracks Step 1
'If plst.list1.Selected(i - 1) = True Then
On Error GoTo 1
'Call reset_Click
i = plst.List1.ListIndex
currenttrack = i + 1
i = InStr(plst.List1.List(currenttrack - 1), "   ")
If i = 0 Then i = Len(plst.List1.List(currenttrack - 1))
'playtrack.Caption = (Str(i + 2) + "." + track2.name) + tracklength

musicsystem.playtrack.Caption = " *** " + UCase(Left(plst.List1.List(currenttrack - 1), i + 1))
musicsystem.trackno.Caption = " Track" + str(currenttrack) 'plst.list1.List(i - 1)
'On Error Resume Next
k = FileLen(track(currenttrack).path)
'On Error Resume Next

Form1.MediaPlayer1.FileName = track(currenttrack).path
tracklength = format((Int(Form1.MediaPlayer1.Duration / 60)), "00") + ":" + format(Int(Int(Form1.MediaPlayer1.Duration Mod 60)), "00")
musicsystem.Label17.Caption = tracklength

'tracklength = Format((Int(Form1.MediaPlayer1.Duration / 60)), "00") + ":" + Format(Int(Int(musicsystem.MediaPlayer2.Duration Mod 60)), "00")
Label2.Caption = tracklength

'List1.List(currenttrack - 1) = List1.List(currenttrack - 1) + tracklength
'Else 'If Form1.MediaPlayer1.FileName = "" Then
'musicsystem.Width = 4890 - 120
'k = MsgBox("Mahesh's mediaplayer can't play this file.This is not valid file for Mahesh's Mediaplayer", vbCritical, "PLayback error")
'Call musicsystem.nexttrackplay
'Form1.MediaPlayer1.Stop
'musicsystem.Text1.SetFocus
'Close #1
'Exit Sub
'End If
'If form1.MediaPlayer1.Duration = 0 And form1.MediaPlayer1.ImageSourceHeight = 0 Then
'End If
Form1.MediaPlayer1.Play
'musicsystem.cbutton(3).Picture = musicsystem.cbuttonE(3).Picture
'musicsystem.cbutton(2).Picture = musicsystem.cbuttonD(2).Picture
'Call musicsystem.cbutton_Click(2)

'totalreplayedtracks = 0
If checkback = False Then
totalreplayedtracks = 0
If totalplayedtracks >= 50 Then totalplayedtracks = 1
totalplayedtracks = totalplayedtracks + 1
previoustrack(totalplayedtracks) = currenttrack
End If
If Form1.MediaPlayer1.ImageSourceHeight = 0 Then

'If musicsystem.WindowState <> 1 Then musicsystem.Width = 4890 - 120
musicsystem.mediatype.Caption = "(MP3)"
  Form1.Hide
  Form1.MediaPlayer1.Visible = False
' musicsystem.Label3.Enabled = True
Else
  Form1.MediaPlayer1.Visible = True
  'If musicsystem.WindowState <> 1 Then musicsystem.Width = 10395 + 120
40
 Form1.trackno.Caption = "T" + str(currenttrack) + " [" + tracklength + "]"

  If Form1.MediaPlayer1.Duration = 0 Then
   musicsystem.mediatype.Caption = "(IMAGE)"
    musicsystem.Timer2.Enabled = False
   For i = 0 To 8 Step 1
    musicsystem.sp(i).Top = 720 + 285
    musicsystem.sp(i).height = 35 '360 + (620)
   Next
 '  If allowslide = True Then Timer4.Enabled = True
   Form1.Show

   Form1.MediaPlayer1.Pause
   For i = 2 To 6 Step 1
   'musicsystem.cbutton(i).Picture = musicsystem.cbuttonD(i).Picture
   'musicsystem.cbutton(i).Enabled = False
   Next
  ' Label3.Enabled = False
  Else
   musicsystem.mediatype.Caption = "(VIDEO)"
   'musicsystem.Label3.Enabled = True
 'k = SetWindowPos(Form1.hwnd, -1, Form1.Left / 12000 * 800, Form1.Top / 9000 * 600, Form1.width / 12000 * 800, Form1.height / 9000 * 600, &H40)
 Form1.Show

  End If
  Form1.Show
  If ckk = False Then
 ' musicsystem.Show
  End If
End If

Close #1
Close #2

musicsystem.Timer2.Enabled = True
For i = 0 To 11 Step 1
musicsystem.Label6(i).ForeColor = musicsystem.Label6(15).ForeColor
Next
'plst.Slider1.SetFocus

1:
Close #1
'plst.list1.SetFocus
'plst.Slider1.SetFocus

End Sub


Public Sub list1_KeyDown(Keycode As Integer, Shift As Integer)
On Error GoTo 9

scrollunit = sizescroll / (totaltracks)

If Keycode = 39 Then
'AppActivate musicsystem.Caption
'plst.list1.SetFocus
SendKeys "{up}", True
Text1.SetFocus
End If
If Keycode = 37 Then
'AppActivate musicsystem.Caption
'plst.list1.SetFocus
SendKeys "{down}", True
Text1.SetFocus
'scroll.Top = scroll.Top - scrollunit

End If
If Keycode = 38 Then scroll.Top = scroll.Top - scrollunit + 60
If Keycode = 40 Then scroll.Top = scroll.Top + scrollunit + 60
If scroll.Top >= sizescroll Then scroll.Top = sizescroll + 60
If scroll.Top <= 10 Then scroll.Top = 90

9: 'If KeyCode = 40 Or KeyCode = 38 Then Text1.SetFocus
End Sub

Public Sub list1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Call plst.list1_DblClick
End Sub




Public Sub next_Click()
On Error GoTo 1
If musicsystem.cbutton(1).Enabled = True Then musicsystem.cbutton_Click (1)
1:
End Sub
Public Sub meclose_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
meclose.Picture = closedown.Picture

End Sub

Public Sub meclose_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
meclose.Picture = closeup.Picture

End Sub

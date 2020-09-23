VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form musicsystem 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   5775
   ClientLeft      =   2025
   ClientTop       =   1200
   ClientWidth     =   7725
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H00404040&
   ForeColor       =   &H00404040&
   Icon            =   "01.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   5775
   ScaleWidth      =   7725
   ShowInTaskbar   =   0   'False
   Begin MCI.MMControl MMControl2 
      Height          =   375
      Left            =   6600
      TabIndex        =   29
      Top             =   4470
      Width           =   840
      _ExtentX        =   1482
      _ExtentY        =   661
      _Version        =   393216
      PauseEnabled    =   -1  'True
      RecordEnabled   =   -1  'True
      PrevVisible     =   0   'False
      NextVisible     =   0   'False
      PlayVisible     =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      StopVisible     =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   "waveaudio"
      FileName        =   ""
      MouseIcon       =   "01.frx":0442
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   4950
      TabIndex        =   37
      Text            =   "Text5"
      Top             =   2820
      Width           =   2655
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   4500
      TabIndex        =   36
      Text            =   "Text4"
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Timer Timer6 
      Interval        =   200
      Left            =   5550
      Top             =   3450
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1935
      Left            =   4320
      TabIndex        =   34
      Top             =   3060
      Width           =   3375
      Begin VB.Image recplay 
         Height          =   270
         Left            =   2520
         Picture         =   "01.frx":0894
         Top             =   960
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Image cbuttone 
         Height          =   270
         Index           =   7
         Left            =   120
         Picture         =   "01.frx":0DE6
         Top             =   240
         Width           =   360
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "12"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   5.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   195
         Index           =   15
         Left            =   150
         TabIndex        =   35
         ToolTipText     =   "Track 12"
         Top             =   3090
         Width           =   195
      End
      Begin VB.Image colorbar 
         Height          =   150
         Left            =   180
         Picture         =   "01.frx":1338
         Top             =   2430
         Width           =   975
      End
      Begin VB.Image volumebarup 
         DragIcon        =   "01.frx":1B22
         DragMode        =   1  'Automatic
         Height          =   75
         Left            =   180
         Picture         =   "01.frx":1E2C
         ToolTipText     =   "Volume bar"
         Top             =   2790
         Width           =   210
      End
      Begin VB.Image volumebardown 
         DragIcon        =   "01.frx":1F4A
         DragMode        =   1  'Automatic
         Height          =   75
         Left            =   450
         Picture         =   "01.frx":2254
         ToolTipText     =   "Volume bar"
         Top             =   2790
         Width           =   210
      End
      Begin VB.Image imgbardown 
         Height          =   90
         Left            =   180
         Picture         =   "01.frx":2372
         Top             =   2640
         Width           =   435
      End
      Begin VB.Image imgbarup 
         Height          =   90
         Left            =   180
         Picture         =   "01.frx":25C4
         Top             =   2940
         Width           =   435
      End
      Begin VB.Image recpause 
         Height          =   270
         Left            =   2610
         Picture         =   "01.frx":2816
         Top             =   240
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Image cbuttonD 
         Height          =   270
         Index           =   7
         Left            =   120
         Picture         =   "01.frx":2D68
         Top             =   570
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Image cbuttone 
         Height          =   270
         Index           =   6
         Left            =   0
         Picture         =   "01.frx":32BA
         Top             =   4350
         Width           =   360
      End
      Begin VB.Image cbuttone 
         Height          =   270
         Index           =   3
         Left            =   630
         Picture         =   "01.frx":380C
         Top             =   4620
         Width           =   360
      End
      Begin VB.Image cbuttone 
         Height          =   270
         Index           =   0
         Left            =   540
         Picture         =   "01.frx":3D5E
         Top             =   240
         Width           =   360
      End
      Begin VB.Image cbuttonD 
         Height          =   270
         Index           =   6
         Left            =   1740
         Picture         =   "01.frx":42B0
         Top             =   570
         Width           =   360
      End
      Begin VB.Image cbuttonD 
         Height          =   270
         Index           =   3
         Left            =   2190
         Picture         =   "01.frx":4802
         Top             =   570
         Width           =   360
      End
      Begin VB.Image cbuttonD 
         Height          =   270
         Index           =   2
         Left            =   930
         Picture         =   "01.frx":4D54
         Top             =   570
         Width           =   360
      End
      Begin VB.Image cbuttonD 
         Height          =   270
         Index           =   1
         Left            =   1320
         Picture         =   "01.frx":52A6
         Top             =   570
         Width           =   360
      End
      Begin VB.Image cbuttone 
         Height          =   270
         Index           =   2
         Left            =   930
         Picture         =   "01.frx":57F8
         Top             =   240
         Width           =   360
      End
      Begin VB.Image cbuttone 
         Height          =   270
         Index           =   1
         Left            =   1320
         Picture         =   "01.frx":5D4A
         Top             =   240
         Width           =   360
      End
      Begin VB.Image cbuttonD 
         Height          =   270
         Index           =   0
         Left            =   540
         Picture         =   "01.frx":629C
         Top             =   570
         Width           =   360
      End
      Begin VB.Image autooffdown 
         Height          =   195
         Left            =   120
         Picture         =   "01.frx":67EE
         Top             =   1140
         Width           =   495
      End
      Begin VB.Image autooffup 
         Height          =   195
         Left            =   120
         Picture         =   "01.frx":6D44
         Top             =   900
         Width           =   495
      End
      Begin VB.Image p 
         Height          =   255
         Index           =   3
         Left            =   120
         Picture         =   "01.frx":729A
         Top             =   1350
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.Image p 
         Height          =   240
         Index           =   4
         Left            =   120
         Picture         =   "01.frx":78B4
         Top             =   1620
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.Image p 
         Height          =   240
         Index           =   1
         Left            =   600
         Picture         =   "01.frx":7E76
         Top             =   1620
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.Image p 
         Height          =   240
         Index           =   2
         Left            =   600
         Picture         =   "01.frx":8438
         Top             =   1350
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.Image autoondown 
         Height          =   195
         Left            =   660
         Picture         =   "01.frx":89FA
         Top             =   1140
         Width           =   495
      End
      Begin VB.Image autoonup 
         Height          =   195
         Left            =   660
         Picture         =   "01.frx":8F50
         Top             =   900
         Width           =   495
      End
      Begin VB.Image presetdown 
         Height          =   210
         Left            =   120
         Picture         =   "01.frx":94A6
         Top             =   1920
         Width           =   690
      End
      Begin VB.Image eqonup 
         Height          =   195
         Left            =   1800
         Picture         =   "01.frx":9C90
         Top             =   900
         Width           =   390
      End
      Begin VB.Image eqondown 
         Height          =   195
         Left            =   1800
         Picture         =   "01.frx":A0E2
         Top             =   1140
         Width           =   390
      End
      Begin VB.Image eqoffdown 
         Height          =   195
         Left            =   1320
         Picture         =   "01.frx":A534
         Top             =   1140
         Width           =   390
      End
      Begin VB.Image eqoffup 
         Height          =   195
         Left            =   1320
         Picture         =   "01.frx":A986
         Top             =   900
         Width           =   390
      End
      Begin VB.Image presetup 
         Height          =   210
         Left            =   120
         Picture         =   "01.frx":ADD8
         Top             =   2160
         Width           =   690
      End
      Begin VB.Image plup 
         Height          =   135
         Left            =   1110
         Picture         =   "01.frx":B5C2
         Top             =   1410
         Width           =   270
      End
      Begin VB.Image equp 
         Height          =   135
         Left            =   1440
         Picture         =   "01.frx":B7FC
         Top             =   1410
         Width           =   255
      End
      Begin VB.Image pldown 
         Height          =   135
         Left            =   1110
         Picture         =   "01.frx":BA12
         Top             =   1590
         Width           =   255
      End
      Begin VB.Image eqdown 
         Height          =   150
         Left            =   1440
         Picture         =   "01.frx":BC28
         Top             =   1590
         Width           =   270
      End
      Begin VB.Image Infdown 
         Height          =   165
         Left            =   1950
         Picture         =   "01.frx":BE9A
         Top             =   1590
         Width           =   150
      End
      Begin VB.Image infup 
         Height          =   165
         Left            =   1950
         Picture         =   "01.frx":C03C
         Top             =   1410
         Width           =   165
      End
      Begin VB.Image mindown 
         Height          =   120
         Left            =   2190
         Picture         =   "01.frx":C20A
         Top             =   1590
         Width           =   135
      End
      Begin VB.Image minup 
         Height          =   135
         Left            =   2160
         Picture         =   "01.frx":C32C
         Top             =   1410
         Width           =   165
      End
      Begin VB.Image closeup 
         Height          =   135
         Left            =   1770
         Picture         =   "01.frx":C4B2
         Top             =   1410
         Width           =   135
      End
      Begin VB.Image closedown 
         Height          =   150
         Left            =   1770
         Picture         =   "01.frx":C5F0
         Top             =   1590
         Width           =   150
      End
   End
   Begin VB.FileListBox File1 
      Height          =   480
      Left            =   4440
      Pattern         =   "*.wav"
      TabIndex        =   33
      Top             =   2760
      Width           =   3255
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Use control panel recording format"
      Height          =   255
      Left            =   4920
      TabIndex        =   32
      Top             =   4470
      Value           =   1  'Checked
      Width           =   2895
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   4680
      TabIndex        =   31
      Text            =   "<select soundcard>"
      Top             =   4260
      Width           =   4215
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   4560
      TabIndex        =   30
      Text            =   "<select soundcard>"
      Top             =   4500
      Width           =   4215
   End
   Begin VB.Timer rec_timer 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5430
      Top             =   3840
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   5430
      Locked          =   -1  'True
      TabIndex        =   22
      TabStop         =   0   'False
      Text            =   "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaa"
      Top             =   4260
      Width           =   2865
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   4770
      Locked          =   -1  'True
      TabIndex        =   21
      TabStop         =   0   'False
      Text            =   "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaa"
      Top             =   4080
      Width           =   2835
   End
   Begin MSComctlLib.Slider Slider4 
      Height          =   165
      Left            =   4710
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   4530
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   291
      _Version        =   393216
      SelStart        =   5
      Value           =   5
   End
   Begin MSComctlLib.Slider Slider3 
      Height          =   225
      Left            =   5250
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   4050
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   397
      _Version        =   393216
      LargeChange     =   1
      Max             =   100
      SelStart        =   14
      Value           =   14
   End
   Begin MSComctlLib.Slider Slider2 
      Height          =   315
      Left            =   5790
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3900
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   556
      _Version        =   393216
      Max             =   70
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Left            =   5310
      Top             =   4440
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   700
      Left            =   5730
      Top             =   3960
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   6000
      Top             =   4140
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Left            =   5100
      Top             =   4290
   End
   Begin VB.TextBox Text1 
      Height          =   345
      Left            =   6570
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa"
      Top             =   3300
      Width           =   4545
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   6510
      Top             =   4080
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   8460
      TabIndex        =   4
      Top             =   4800
      Visible         =   0   'False
      Width           =   1605
      Begin VB.Image Image3 
         Height          =   240
         Index           =   1
         Left            =   2700
         Picture         =   "01.frx":C772
         Top             =   300
         Width           =   420
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   3
         Left            =   840
         Picture         =   "01.frx":CCF4
         Top             =   450
         Width           =   420
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   2
         Left            =   2010
         Picture         =   "01.frx":D276
         Top             =   390
         Width           =   420
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   4
         Left            =   1470
         Picture         =   "01.frx":D7F8
         Top             =   390
         Width           =   420
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00FFFFFF&
         X1              =   5670
         X2              =   30
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00FFFFFF&
         X1              =   5790
         X2              =   -810
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00FFFFFF&
         X1              =   5520
         X2              =   5520
         Y1              =   4560
         Y2              =   120
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   5580
      Top             =   3690
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "wav"
      DialogTitle     =   "Save recoeded file as"
      FileName        =   "*.wav"
      Filter          =   "*.wav"
   End
   Begin VB.Image Image21 
      Height          =   5430
      Left            =   4410
      Picture         =   "01.frx":DD7A
      Top             =   2490
      Width           =   4155
   End
   Begin VB.Image Image24 
      Height          =   585
      Left            =   0
      Picture         =   "01.frx":5763C
      Top             =   390
      Width           =   240
   End
   Begin VB.Image eqfup 
      Height          =   165
      Left            =   4650
      Picture         =   "01.frx":57DCE
      Top             =   810
      Width           =   120
   End
   Begin VB.Image eqfdown 
      Height          =   165
      Left            =   4680
      Picture         =   "01.frx":57F18
      Top             =   1260
      Width           =   105
   End
   Begin VB.Image eqf 
      Height          =   165
      Index           =   10
      Left            =   360
      Picture         =   "01.frx":58062
      Top             =   2660
      Width           =   120
   End
   Begin VB.Label rec_count 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00:02:34"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   150
      Left            =   645
      TabIndex        =   28
      Top             =   540
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Label reclabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "REC"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   120
      Left            =   320
      TabIndex        =   27
      Top             =   570
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image Image1 
      Height          =   345
      Left            =   -30
      Picture         =   "01.frx":581AC
      Stretch         =   -1  'True
      Top             =   3210
      Width           =   4110
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "EQUALIZER"
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
      Left            =   4470
      TabIndex        =   26
      Top             =   2370
      Width           =   675
   End
   Begin VB.Image eqclose 
      Height          =   135
      Left            =   3930
      Picture         =   "01.frx":5BF16
      ToolTipText     =   "Close"
      Top             =   1770
      Width           =   135
   End
   Begin VB.Line Line15 
      BorderColor     =   &H00004080&
      BorderStyle     =   3  'Dot
      DrawMode        =   9  'Not Mask Pen
      X1              =   1300
      X2              =   2970
      Y1              =   2150
      Y2              =   2150
   End
   Begin VB.Image auto 
      Height          =   195
      Left            =   600
      Picture         =   "01.frx":5C054
      Top             =   2010
      Width           =   495
   End
   Begin VB.Image preset 
      Height          =   210
      Left            =   3240
      Picture         =   "01.frx":5C5AA
      Top             =   1970
      Width           =   690
   End
   Begin VB.Image eqon 
      Height          =   195
      Left            =   210
      Picture         =   "01.frx":5CD94
      Top             =   2010
      Width           =   390
   End
   Begin VB.Image eqf 
      Height          =   165
      Index           =   9
      Left            =   3650
      Picture         =   "01.frx":5D1E6
      Top             =   2550
      Width           =   120
   End
   Begin VB.Image eqf 
      Height          =   165
      Index           =   8
      Left            =   3380
      Picture         =   "01.frx":5D330
      Top             =   2550
      Width           =   120
   End
   Begin VB.Image eqf 
      Height          =   165
      Index           =   7
      Left            =   3110
      Picture         =   "01.frx":5D47A
      Top             =   2700
      Width           =   120
   End
   Begin VB.Image eqf 
      Height          =   165
      Index           =   6
      Left            =   2840
      Picture         =   "01.frx":5D5C4
      Top             =   2700
      Width           =   120
   End
   Begin VB.Image eqf 
      Height          =   165
      Index           =   5
      Left            =   2570
      Picture         =   "01.frx":5D70E
      Top             =   2700
      Width           =   120
   End
   Begin VB.Image eqf 
      Height          =   165
      Index           =   4
      Left            =   2300
      Picture         =   "01.frx":5D858
      Top             =   2730
      Width           =   120
   End
   Begin VB.Image eqf 
      Height          =   165
      Index           =   3
      Left            =   2020
      Picture         =   "01.frx":5D9A2
      Top             =   2700
      Width           =   120
   End
   Begin VB.Image eqf 
      Height          =   165
      Index           =   2
      Left            =   1750
      Picture         =   "01.frx":5DAEC
      Top             =   2670
      Width           =   120
   End
   Begin VB.Image eqf 
      Height          =   165
      Index           =   1
      Left            =   1480
      Picture         =   "01.frx":5DC36
      Top             =   2640
      Width           =   120
   End
   Begin VB.Image eqf 
      Height          =   165
      Index           =   0
      Left            =   1200
      Picture         =   "01.frx":5DD80
      Top             =   2400
      Width           =   120
   End
   Begin VB.Image balancebar 
      DragIcon        =   "01.frx":5DECA
      Height          =   75
      Left            =   2880
      Picture         =   "01.frx":5E1D4
      ToolTipText     =   "Balance bar"
      Top             =   930
      Width           =   210
   End
   Begin VB.Image imgbalancemeter 
      Height          =   150
      Left            =   2640
      Picture         =   "01.frx":5E2F2
      Top             =   900
      Width           =   660
   End
   Begin VB.Image Image20 
      Height          =   165
      Left            =   3570
      Picture         =   "01.frx":5E85C
      Top             =   0
      Width           =   225
   End
   Begin VB.Image Image18 
      Height          =   300
      Left            =   3630
      Picture         =   "01.frx":5EAAE
      Top             =   1290
      Width           =   315
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
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
      ForeColor       =   &H000040C0&
      Height          =   165
      Left            =   2310
      TabIndex        =   24
      ToolTipText     =   "Current Media Duration"
      Top             =   600
      Width           =   330
   End
   Begin VB.Shape sp 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   7
      Left            =   1170
      Top             =   750
      Width           =   90
   End
   Begin VB.Shape sp 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00C0FFC0&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   5
      Left            =   930
      Top             =   750
      Width           =   90
   End
   Begin VB.Shape sp 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   0
      Left            =   330
      Top             =   750
      Width           =   90
   End
   Begin VB.Shape sp 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   1
      Left            =   450
      Top             =   750
      Width           =   90
   End
   Begin VB.Shape sp 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0080C0FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   2
      Left            =   570
      Top             =   750
      Width           =   90
   End
   Begin VB.Shape sp 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00C0FFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   3
      Left            =   690
      Top             =   750
      Width           =   90
   End
   Begin VB.Shape sp 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   4
      Left            =   810
      Top             =   750
      Width           =   90
   End
   Begin VB.Shape sp 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0080FF80&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   6
      Left            =   1050
      Top             =   750
      Width           =   90
   End
   Begin VB.Shape sp 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00008000&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   8
      Left            =   1290
      Top             =   750
      Width           =   90
   End
   Begin VB.Label trackno 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   " Track"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   165
      Left            =   1680
      TabIndex        =   17
      Top             =   600
      Width           =   390
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
      Left            =   315
      TabIndex        =   18
      ToolTipText     =   "Time elapsed"
      Top             =   395
      Width           =   405
   End
   Begin VB.Image Image16 
      Height          =   750
      Left            =   240
      Picture         =   "01.frx":5EFF0
      Top             =   300
      Width           =   1350
   End
   Begin VB.Image EQ 
      Height          =   135
      Left            =   3450
      Picture         =   "01.frx":62552
      ToolTipText     =   "Toggle Graphic EQualisers"
      Top             =   900
      Width           =   255
   End
   Begin VB.Image PL 
      Height          =   135
      Left            =   3720
      Picture         =   "01.frx":62768
      ToolTipText     =   "Toggle PlaylistEditor"
      Top             =   900
      Width           =   270
   End
   Begin VB.Image Image2 
      Height          =   225
      Left            =   3210
      Picture         =   "01.frx":629A2
      Stretch         =   -1  'True
      Top             =   870
      Width           =   735
   End
   Begin VB.Line Line21 
      BorderColor     =   &H00404000&
      X1              =   3960
      X2              =   4020
      Y1              =   300
      Y2              =   300
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "MAHESH'S MPLAYER"
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
      Left            =   4560
      TabIndex        =   23
      Top             =   3360
      Width           =   1245
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "12"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   120
      Index           =   11
      Left            =   3810
      TabIndex        =   16
      ToolTipText     =   "Track 12"
      Top             =   660
      Width           =   105
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "9"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   120
      Index           =   8
      Left            =   3690
      TabIndex        =   13
      ToolTipText     =   "Track 9"
      Top             =   660
      Width           =   60
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   120
      Index           =   5
      Left            =   3570
      TabIndex        =   12
      ToolTipText     =   "Track 6"
      Top             =   660
      Width           =   60
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "11"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   120
      Index           =   10
      Left            =   3810
      TabIndex        =   15
      ToolTipText     =   "Track 11"
      Top             =   510
      Width           =   90
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "8"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   120
      Index           =   7
      Left            =   3690
      TabIndex        =   11
      ToolTipText     =   "Track 8"
      Top             =   510
      Width           =   60
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   120
      Index           =   4
      Left            =   3570
      TabIndex        =   9
      ToolTipText     =   "Track 5"
      Top             =   510
      Width           =   60
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   120
      Index           =   9
      Left            =   3810
      TabIndex        =   14
      ToolTipText     =   "Track 10"
      Top             =   390
      Width           =   105
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "7"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   120
      Index           =   6
      Left            =   3690
      TabIndex        =   10
      ToolTipText     =   "Track 7"
      Top             =   390
      Width           =   60
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   120
      Index           =   3
      Left            =   3570
      TabIndex        =   6
      ToolTipText     =   "Track 4"
      Top             =   390
      Width           =   60
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   120
      Index           =   2
      Left            =   3450
      TabIndex        =   8
      ToolTipText     =   "Track 3"
      Top             =   660
      Width           =   60
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   120
      Index           =   1
      Left            =   3450
      TabIndex        =   7
      ToolTipText     =   "Track 2"
      Top             =   510
      Width           =   60
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   120
      Index           =   0
      Left            =   3450
      TabIndex        =   5
      ToolTipText     =   "Track 1"
      Top             =   390
      Width           =   45
   End
   Begin VB.Image volumebar 
      DragIcon        =   "01.frx":631A0
      Height          =   75
      Left            =   2210
      Picture         =   "01.frx":634AA
      ToolTipText     =   "Volume bar"
      Top             =   930
      Width           =   210
   End
   Begin VB.Image meclose 
      Height          =   135
      Left            =   3930
      Picture         =   "01.frx":635C8
      ToolTipText     =   "Close"
      Top             =   30
      Width           =   135
   End
   Begin VB.Image Inf 
      Height          =   165
      Left            =   60
      Picture         =   "01.frx":63706
      ToolTipText     =   "Configuration"
      Top             =   0
      Width           =   165
   End
   Begin VB.Line Line16 
      X1              =   3960
      X2              =   3960
      Y1              =   330
      Y2              =   810
   End
   Begin VB.Image Image10 
      Height          =   150
      Left            =   1640
      Picture         =   "01.frx":638D4
      Top             =   900
      Width           =   975
   End
   Begin VB.Image min 
      Height          =   135
      Left            =   3780
      Picture         =   "01.frx":640BE
      ToolTipText     =   "Minimise"
      Top             =   30
      Width           =   165
   End
   Begin VB.Label mediatype 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(MEDIA)"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   120
      Left            =   2700
      TabIndex        =   19
      Top             =   610
      Width           =   420
   End
   Begin VB.Image Imgbar 
      DragIcon        =   "01.frx":64244
      Height          =   90
      Left            =   240
      Picture         =   "01.frx":6454E
      ToolTipText     =   "Seeking Bar"
      Top             =   1110
      Width           =   435
   End
   Begin VB.Image Image3 
      Height          =   240
      Index           =   0
      Left            =   3150
      Picture         =   "01.frx":647A0
      ToolTipText     =   "Toggle repeat"
      Top             =   1350
      Width           =   420
   End
   Begin VB.Image p 
      Height          =   240
      Index           =   0
      Left            =   2730
      Picture         =   "01.frx":64D22
      ToolTipText     =   "Toggle shuffle"
      Top             =   1350
      Width           =   435
   End
   Begin VB.Image cbutton 
      Height          =   270
      Index           =   7
      Left            =   2010
      Picture         =   "01.frx":652E4
      ToolTipText     =   "Record/Record Pause"
      Top             =   1320
      Width           =   360
   End
   Begin VB.Image cbutton 
      Height          =   270
      Index           =   2
      Left            =   540
      Picture         =   "01.frx":65836
      ToolTipText     =   "Play(P)"
      Top             =   1320
      Width           =   360
   End
   Begin VB.Image cbutton 
      Height          =   270
      Index           =   3
      Left            =   900
      Picture         =   "01.frx":65D88
      ToolTipText     =   "Pause(P)"
      Top             =   1320
      Width           =   360
   End
   Begin VB.Image cbutton 
      Height          =   270
      Index           =   6
      Left            =   1260
      Picture         =   "01.frx":662DA
      ToolTipText     =   "Stop(S)"
      Top             =   1320
      Width           =   360
   End
   Begin VB.Image cbutton 
      Height          =   270
      Index           =   1
      Left            =   1620
      Picture         =   "01.frx":6682C
      ToolTipText     =   "Next track(N)"
      Top             =   1320
      Width           =   360
   End
   Begin VB.Image cbutton 
      Height          =   270
      Index           =   0
      Left            =   210
      Picture         =   "01.frx":66D7E
      ToolTipText     =   "Previous track(B)"
      Top             =   1320
      Width           =   360
   End
   Begin VB.Image Image8 
      Height          =   135
      Index           =   1
      Left            =   220
      Picture         =   "01.frx":672D0
      Stretch         =   -1  'True
      ToolTipText     =   "Seeking Bar"
      Top             =   1100
      Width           =   3705
   End
   Begin VB.Image Image23 
      Height          =   450
      Left            =   3360
      Picture         =   "01.frx":6909A
      Stretch         =   -1  'True
      Top             =   330
      Width           =   585
   End
   Begin VB.Label playtrack 
      BackStyle       =   0  'Transparent
      Caption         =   ". Mahesh Mediaplayer ."
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
      Left            =   1680
      TabIndex        =   20
      Top             =   390
      Width           =   1770
   End
   Begin VB.Label volinf 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Left            =   1720
      TabIndex        =   25
      Top             =   390
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Image Image15 
      Height          =   450
      Left            =   1620
      Picture         =   "01.frx":6B35C
      Stretch         =   -1  'True
      Top             =   330
      Width           =   2385
   End
   Begin VB.Image Image17 
      Height          =   360
      Left            =   3920
      Picture         =   "01.frx":6D61E
      Stretch         =   -1  'True
      Top             =   300
      Width           =   150
   End
   Begin VB.Image Image22 
      Height          =   3630
      Left            =   -30
      Picture         =   "01.frx":6D8A0
      Top             =   -15
      Width           =   4125
   End
   Begin VB.Menu file 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu about 
         Caption         =   "About Mahesh'sMplayer..."
      End
      Begin VB.Menu atop 
         Caption         =   "Always on top"
      End
      Begin VB.Menu ple 
         Caption         =   "Playlist Editor"
         Checked         =   -1  'True
      End
      Begin VB.Menu eqe 
         Caption         =   "Graphic Equaliser"
         Checked         =   -1  'True
      End
      Begin VB.Menu explored 
         Caption         =   "Explore mmpWorkarea Folder"
      End
      Begin VB.Menu plb 
         Caption         =   "playback"
      End
      Begin VB.Menu exe 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "musicsystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l As Long
Dim m As Integer
Public vol As New clsVolume


' Define variants to hold keyword parsing arrays

' *** Establish constants for the registry functions
Const HKEY_CLASSES_ROOT = &H80000000
Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002
Const HKEY_USERS = &H80000003
Const HKEY_DYN_DATA = &H80000004

Const REG_SZ = 1

' Registry error constants
Const ERROR_SUCCESS = 0&
Const ERROR_BADKEY = 1010&

' Registry API prototypes
Private Declare Function RegCreateKey Lib "advapi32.dll" _
  Alias "RegCreateKeyA" _
    (ByVal hkey As Long, _
     ByVal lpSubKey As String, _
     phkResult As Long) As Long
     
Private Declare Function RegSetValueEx Lib "advapi32.dll" _
  Alias "RegSetValueExA" _
    (ByVal hkey As Long, _
     ByVal lpValueName As String, _
     ByVal Reserved As Long, _
     ByVal dwType As Long, _
     lpData As Any, _
     ByVal cbData As Long) As Long

Private Declare Function RegDeleteKey Lib "advapi32.dll" _
  Alias "RegDeleteKeyA" _
    (ByVal hkey As Long, _
     ByVal lpSubKey As String) As Long
     












Private Sub ResultsImportFile(sFileName As String, media As Boolean)

Dim listname As String
Dim i As Integer
If media = True Then
Dim temptrack As file
currenttrack = 1 ' = sFileName
plst.newpl_Click
'Open App.path + "\mplayerlist.mmpl" For Random As #1 Len = 155
totaltracks = 1
temptrack.path = sFileName
File1.path = ""
File1.FileName = sFileName
listname = File1.List(0)
listname = str(totaltracks) + ". " + listname
If Len(listname) >= 34 Then listname = Left(listname, 34) + "..."
listname = LCase(listname)
plst.List1.AddItem (listname)
temptrack.name = listname 'Str(totaltracks) + ". " + Left(cd1.FileTitle, 30) + Space(7) + tracklength
i = InStr(listname, ".")
temptrack.name = LCase(Mid(listname, i + 1, Len(listname)))
'Put #1, totaltracks, track
'Close #1
currenttrack = 1
plst.List1.ListIndex = 0
track(totaltracks) = temptrack
Call plst.list1_DblClick

'playtrack.Caption = temptrack.name
Label6(0).Enabled = True
Else
End If

End Sub
Private Sub GetAppSettings()

Dim nLeft           As Integer
Dim nTop            As Integer

' Get the previous left position of the form
nLeft = GetSetting(App.Title, "General", "Left", 0)

' Get the previous top position of the form
nTop = GetSetting(App.Title, "General", "Top", 0)

' Now reposition the form
Me.Move nLeft, nTop

End Sub

Private Sub SaveAppSettings()

' Save the current form position and dimensions
SaveSetting App.Title, "General", "Left", Me.Left
SaveSetting App.Title, "General", "Top", Me.Top

End Sub

Public Sub DroppedFiles(vFileList As Variant)

Dim nLoopCtr            As Integer
Dim sFileName           As String

' *** Loop through the variant array to process
' each dropped file
'For nLoopCtr = 0 To UBound(vFileList)
    ' Get the current file name from the array
    sFileName = vFileList(nLoopCtr)
    
    ' Make sure it is a SAD format file
    If UCase$(Right$(sFileName, 4)) = ".MP3" Or UCase$(Right$(sFileName, 4)) = ".DAT" Or UCase$(Right$(sFileName, 4)) = ".WAV" Or UCase$(Right$(sFileName, 4)) = ".MPG" Or UCase$(Right$(sFileName, 5)) = ".MPEG" Or UCase$(Right$(sFileName, 4)) = ".AVI" Then
      'Form1.MediaPlayer1.FileName = sFileName
        ' It is the correct file type so import it
        Call ResultsImportFile(sFileName, True)
    End If
'Next nLoopCtr

End Sub

Private Sub SelfRegister(sCommandLine As String)

Dim lResult                 As Long
    On Error Resume Next

 DeleteSetting App.Title, "General"
    DeleteSetting App.Title
' Check to see if the code needs to establish registry settings
If InStr(1, UCase$(sCommandLine), "/REGSERVER") <> 0 Then
    ' Create the file type entries
    lResult = DeleteRegKey(HKEY_CLASSES_ROOT, _
                ".DAT", _
                "")
  ' lResult = DeleteRegKey(HKEY_CLASSES_ROOT, _
                "winamp.file", _
                "")
  ' lResult = DeleteRegKey(HKEY_CLASSES_ROOT, _
                "MaheshMp3file\Shell\open", _
                "")
   ' lResult = DeleteRegKey(HKEY_CLASSES_ROOT, _
                "MaheshMp31", _
                "")
    ' Register the file extension'("")value for default registry
    ' don't write" " for ""
    lResult = SetRegValue(HKEY_CLASSES_ROOT, _
                ".DAT", _
                "", _
                "MaheshDATfile")
                
    ' Register the extension shell handling
    lResult = SetRegValue(HKEY_CLASSES_ROOT, _
                "MaheshDATfile", _
                "", _
                "Maheshmp media file")
     'assigns type of file in detailed view
    lResult = SetRegValue(HKEY_CLASSES_ROOT, _
                "MaheshDATfile\DefaultIcon", _
                "", _
                "d:\maheshmp.exe" & ",0")
      'assigns exe icon to filename type
    lResult = SetRegValue(HKEY_CLASSES_ROOT, _
                "MaheshDATfile\Shell\play with mahesh\Command", _
                "", _
                 "d:\maheshmp.exe" & " " & Chr$(34) & "%1" & Chr$(34))
     ' "%1" is replaced by the correspomding filename selected & henc ew thic particular
     'combination gives commmand as <Exe filepath filename>
     'and filename is sent as commandline argument
      lResult = SetRegValue(HKEY_CLASSES_ROOT, _
                "MaheshDATfile\Shell\open", _
                "", _
                "open with Mahesh Mplayer")
     'Instructs to display "open with Mplayer " in place of "open" menu command in right click
      lResult = SetRegValue(HKEY_CLASSES_ROOT, _
                "MaheshDATfile\Shell\play with mahesh", _
                "", _
                "play with Mahesh Mplayer")
                
     lResult = SetRegValue(HKEY_CLASSES_ROOT, _
                "MaheshDATfile\Shell\open\Command", _
                "", _
                Chr$(34) & "d:\maheshmp.exe" & Chr$(34) & " " & Chr$(34) & "%1" & Chr$(34))
     'sends command prompt the instruction <exepath "filename">
     'replacing  "%1"by filename string
  
    End

End If

' Check to see if the code needs to remove registry settings
If InStr(1, UCase$(sCommandLine), "/UNREGSERVER") <> 0 Then
    ' Delete the form and application settings
    
    ' The error trapping is in case the application
    ' specific entries do not exist
    On Error Resume Next
    DeleteSetting App.Title, "General"
    DeleteSetting App.Title
    On Error GoTo 0
    
    ' Delete the file type entries
    lResult = DeleteRegKey(HKEY_CLASSES_ROOT, _
                ".mp3", _
                "")
    lResult = DeleteRegKey(HKEY_CLASSES_ROOT, _
                "MaheshMp3", _
                "")
                
    ' Exit from the program
    End

End If

' If the program gets here there were no self-registration
' tasks to perform, so go back to form load

End Sub

Private Function SetRegValue _
            (lKeyRoot As Long, _
             tRegistryKey As String, _
             tSubKey As String, _
             tKeyValue As String) As Long

Dim lKeyId              As Long
Dim lResult             As Long

SetRegValue = 0                ' Assume succcess

If Len(tRegistryKey) = 0 Then
    ' The key parameter is not set
    SetRegValue = ERROR_BADKEY
    Exit Function
End If

' Open the key by attempting to create it. If it
' already exists we get back an ID.
lResult = RegCreateKey(lKeyRoot, _
                tRegistryKey, _
                lKeyId)

If lResult <> 0 Then
    ' Call failed, can't open the key so exit
    SetRegValue = lResult
    Exit Function
End If

If Len(tKeyValue) = 0 Then
    ' No key value, so clear any existing entry
    SetRegValue = RegSetValueEx(lKeyId, _
                tSubKey, _
                0&, _
                REG_SZ, _
                0&, _
                0&)
  Else
    ' Set the registry entry to the value
    SetRegValue = RegSetValueEx(lKeyId, _
                tSubKey, _
                0&, _
                REG_SZ, _
                ByVal tKeyValue, _
                Len(tKeyValue) + 1)
End If

End Function

Private Function DeleteRegKey _
            (lKeyRoot As Long, _
             tRegistryKey As String, _
             tSubKey As String) As Long

Dim lKeyId          As Long
Dim lResult         As Long

DeleteRegKey = 0           ' Assume succcess

If Len(tRegistryKey) = 0 Then
    ' The key parameter is not set
    DeleteRegKey = ERROR_BADKEY
    Exit Function
End If

' Open the key by attempting to create it. If it
' already exists we get back an ID.
lResult = RegCreateKey(lKeyRoot, tRegistryKey, lKeyId)
If lResult = 0 Then
    ' We got a key ID so we can delete the entry
    DeleteRegKey = RegDeleteKey(lKeyId, ByVal tSubKey)
End If

End Function


 '.FileSystemObject



Public Sub bar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'bar.Move X, Y

text1.Text = str(X) + "   " + str(Y)
End Sub

Public Sub button_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)


End Sub

Public Sub button_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

End Sub

Public Sub about_Click()
Form2.Show
End Sub


Public Sub addtracks_Click()

End Sub

Public Sub atop_Click()
Dim k
On Error GoTo 1

If atop.Checked = False Then
k = SetWindowPos(Me.hwnd, -1, Me.Left / 12000 * 800, Me.Top / 9000 * 600, Me.width / 12000 * 800, Me.height / 9000 * 600, &H40)
k = SetWindowPos(plst.hwnd, -1, plst.Left / 12000 * 800, plst.Top / 9000 * 600, Me.width / 12000 * 800, plst.height / 9000 * 600, &H40)

atop.Checked = True
Else
k = SetWindowPos(Me.hwnd, -2, Me.Left / 12000 * 800, Me.Top / 9000 * 600, Me.width / 12000 * 800, Me.height / 9000 * 600, &H40)
k = SetWindowPos(plst.hwnd, -2, plst.Left / 12000 * 800, plst.Top / 9000 * 600, Me.width / 12000 * 800, plst.height / 9000 * 600, &H40)

atop.Checked = False
End If
1:
End Sub

Public Sub back_Click()
On Error GoTo 1
If musicsystem.cbutton(0).Enabled = True Then musicsystem.cbutton_Click (0)
1:
End Sub

Private Sub balancebar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim m As POINTAPI
Dim k
Dim m1 As Integer
k = GetCursorPos(m)
initialypos = m.Y '- Me.Top / 12000 * 800

initialxpos = m.X '- Me.Left / 12000 * 800
volinf.Visible = True
playtrack.Visible = False
If balancebar.Left > 2868 Then
m1 = Int((balancebar.Left - 2868) / 2.5)
If m1 > 99 Then m1 = 100
If m1 < 2 Then m1 = 0
volinf.Caption = "Balance= " + str(m1) + "% Right"
ElseIf balancebar.Left < 2870 Then
m1 = Int(-(balancebar.Left - 2880) / 2.5)
If m1 > 99 Then m1 = 100
If m1 < 2 Then m1 = 0
volinf.Caption = "Balance= " + str(m1) + "% Left"
End If
If balancebar.Left < 2880 And balancebar.Left > 2850 Then
volinf.Caption = "Balance= Centre" '+ 'Str(Int(-(balancebar.Left - 2870) / 2.5)) + "% Left"
Form1.MediaPlayer1.Balance = 0
End If

End Sub

Public Sub balancebar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim m As POINTAPI
Dim k
If Button = 1 Then
k = GetCursorPos(m)
If Abs(m.X - initialxpos) >= 1 Then
balancebar.Left = balancebar.Left + ((m.X - initialxpos) * 12000 / 800)
initialxpos = m.X
End If

Dim m1 As Integer

If balancebar.Left <= 2600 Then balancebar.Left = 2620
If (balancebar.Left >= 3120) Then balancebar.Left = 3120

Form1.MediaPlayer1.Balance = (balancebar.Left - 2850) * 9

Slider4.Value = (balancebar.Left - 2610) / 72
'Slider3.SetFocus
playtrack.Visible = False

If balancebar.Left > 2868 Then
m1 = Int((balancebar.Left - 2868) / 2.5)
If m1 > 99 Then m1 = 100
If m1 < 2 Then m1 = 0
volinf.Caption = "Balance= " + str(m1) + "% Right"
ElseIf balancebar.Left < 2870 Then
m1 = Int(-(balancebar.Left - 2880) / 2.5)
If m1 > 99 Then m1 = 100
If m1 < 2 Then m1 = 0
volinf.Caption = "Balance= " + str(m1) + "% Left"
End If
If balancebar.Left < 2880 And balancebar.Left > 2850 Then
volinf.Caption = "Balance= Centre" '+ 'Str(Int(-(balancebar.Left - 2870) / 2.5)) + "% Left"
Form1.MediaPlayer1.Balance = 0
End If

End If

'Slider4.SetFocus
Slider4.SetFocus

End Sub

Private Sub balancebar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
volinf.Visible = False
playtrack.Visible = True
End Sub

Public Sub cbutton_Click(Index As Integer)
On Error GoTo 1
If plst.List1.ListCount >= 1 Then plst.List1.ListIndex = currenttrack - 1
Select Case Index

Case 0 'back
totalreplayedtracks = totalreplayedtracks + 1

'totalplayedtracks = totalplayedtracks + 1
'previoustrack(totalplayedtracks) = previoustrack(totalplayedtracks - 1)
checkback = True
If totalreplayedtracks < totalplayedtracks Then
plst.List1.ListIndex = previoustrack(totalplayedtracks - totalreplayedtracks) - 1
If plst.List1.ListIndex <= 0 Then plst.List1.ListIndex = 0
Call plst.list1_DblClick
Else
plst.List1.ListIndex = 0
End If
checkback = False
Case 1 'next

Call nexttrackplay

Case 3 'pause
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
Case 2 'play
If Form1.MediaPlayer1.PlayState = mpPaused Then
Form1.MediaPlayer1.Play
ElseIf Form1.MediaPlayer1.PlayState = mpStopped Then
Form1.MediaPlayer1.CurrentPosition = 0
Form1.MediaPlayer1.Play
Imgbar.Left = Image8(1).width
ElseIf Form1.MediaPlayer1.FileName = "" Then
Call plst.list1_DblClick
End If
'musicsystem.cbutton(4).Picture = musicsystem.cbuttonE(4).Picture
'musicsystem.cbutton(4).Enabled = True
'musicsystem.cbutton(5).Picture = musicsystem.cbuttonE(5).Picture
'musicsystem.cbutton(5).Enabled = True

'musicsystem.cbutton(2).Picture = musicsystem.cbuttonD(2).Picture
'musicsystem.cbutton(3).Picture = musicsystem.cbuttonE(3).Picture
'musicsystem.cbutton(2).Enabled = False
'musicsystem.cbutton(6).Picture = musicsystem.cbuttonE(6).Picture
'musicsystem.cbutton(6).Enabled = True
'musicsystem.cbutton(3).Enabled = True
Case 4

If Form1.MediaPlayer1.CurrentPosition > 20 Then
Form1.MediaPlayer1.CurrentPosition = Form1.MediaPlayer1.CurrentPosition - 20
Else
Form1.MediaPlayer1.CurrentPosition = 0
End If
'Call musicsystem.cbutton_Click(2)
If Form1.MediaPlayer1.Duration >= 1 Then Imgbar.Left = Image8(1).Left + Form1.MediaPlayer1.CurrentPosition / Form1.MediaPlayer1.Duration * (3270)

Case 5
If Form1.MediaPlayer1.CurrentPosition < Form1.MediaPlayer1.Duration - 20 Then
Form1.MediaPlayer1.CurrentPosition = Form1.MediaPlayer1.CurrentPosition + 20
Else
'Form1.MediaPlayer1.Stop
'musicsystem.cbutton(2).Enabled = True
'musicsystem.cbutton(2).Picture = musicsystem.cbuttonE(2).Picture
'bar.Width = Label3.Width
Form1.MediaPlayer1.CurrentPosition = Form1.MediaPlayer1.Duration - 0.5
'musicsystem.cbutton(3).Picture = musicsystem.cbuttonD(3).Picture
'musicsystem.cbutton(3).Enabled = False
End If
'Call musicsystem.cbutton_Click(2)
If Form1.MediaPlayer1.Duration > 0 Then Imgbar.Left = Image8(1).Left + Form1.MediaPlayer1.CurrentPosition / Form1.MediaPlayer1.Duration * 3270

Case 6
If rec = True Then
 cbutton(7).Picture = cbuttonE(7).Picture
 rectime = 0
 musicsystem.rec_timer.Enabled = False
 musicsystem.reclabel.Caption = "FILE"
 musicsystem.reclabel.ForeColor = musicsystem.trackpos.ForeColor
' MMControl2.Command = "stop"
   'Call stop_record
   'CommonDialog2.ShowSave
   Call getSaveFile
    MMControl2.FileName = SaveAudiofile
MMControl2.Command = "save"
MMControl2.FileName = ""
  MMControl2.Command = "close"
  rec_count.Left = reclabel.Left + reclabel.width + 40

rec = False
rec_pause = False
'Form1.MediaPlayer1.FileName = SaveAudiofile
'Form1.Hide
'Form1.MediaPlayer1.Play
'Form1.Hide

Else
Imgbar.Left = Image8(1).Left

'Form1.MediaPlayer1.Stop
Form1.MediaPlayer1.CurrentPosition = 0.5
Form1.MediaPlayer1.Stop 'Pause
trackpos.Caption = "00:00"
'If Form1.Visible = True Then
Form1.trackpos.Caption = trackpos.Caption

playtrack.Left = 1700
'musicsystem.cbutton(2).Enabled = True
'musicsystem.cbutton(2).Picture = musicsystem.cbuttonE(2).Picture
'musicsystem.cbutton(3).Picture = musicsystem.cbuttonD(3).Picture
'musicsystem.cbutton(3).Enabled = False
'musicsystem.cbutton(6).Picture = musicsystem.cbuttonD(6).Picture
'musicsystem.cbutton(6).Enabled = False
'form1.MediaPlayer1.CurrentPosition = form1.MediaPlayer1.Duration - 1
'musicsystem.cbutton(4).Picture = musicsystem.cbuttonD(4).Picture
'musicsystem.cbutton(4).Enabled = False
'musicsystem.cbutton(5).Picture = musicsystem.cbuttonD(5).Picture
'musicsystem.cbutton(5).Enabled = False
End If
Case 7 'open
If rec = False Then
   rec = True
   cbutton(7).Picture = cbuttonE(7).Picture
   MMControl2.Command = "close"    ' close previously open file
   MMControl2.FileName = "new"
   MMControl2.Command = "open"
   MMControl2.Notify = True
   MMControl2_RecordClick (0)
   MMControl2.Command = "record"
   rectime = 0
   rec_pause = False
   rec_timer.Enabled = True
   musicsystem.reclabel.Caption = "REC"
   musicsystem.rec_count.Left = 630
   musicsystem.rec_count.Visible = True
   musicsystem.reclabel.Visible = True
   musicsystem.reclabel.ForeColor = musicsystem.rec_count.ForeColor
ElseIf rec_pause = False Then
   rec = True
   rec_pause = True
   cbutton(7).Picture = recpause.Picture
   rec_timer.Enabled = False
     ' MMControl2.Notify = True
   MMControl2.Command = "pause"
Else
   rec = True
   rec_pause = False
   cbutton(7).Picture = cbuttonE(7).Picture
   rec_timer.Enabled = True
   'MMControl2.Notify = True
   
   MMControl2.Command = "pause"
End If

rec_count.Left = reclabel.Left + reclabel.width + 40


End Select
'MMC.Command = "close"
1:
Dim k
'k = SetWindowPos(Me.hwnd, -1, Me.Left / 12000 * 800, Me.Top / 9000 * 600, Me.width / 12000 * 800, Me.height / 9000 * 600, &H40)
'rectime = 520
'Text1.SetFocus
End Sub

Public Sub cbutton_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
cbutton(Index).Picture = musicsystem.cbuttonD(Index).Picture
End Sub

Public Sub cbutton_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'Dim i As Integer
'For i = 0 To 7 Step 1
'musicsystem.cbutton(i).Picture = musicsystem.cbuttonE(i).Picture
'Next
'musicsystem.cbutton(Index).Picture = musicsystem.cbuttonD(Index).Picture

On Error GoTo 9
text1.SetFocus
9:
End Sub

Public Sub cbutton_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
cbutton(Index).Picture = musicsystem.cbuttonE(Index).Picture
End Sub
























Public Sub EQ_Click()
Dim k As Integer
 If plst.Top >= Me.Top + Me.height - 40 And plst.Top <= Me.Top + Me.height + 40 Then
  plst.Top = Me.Top + Me.height
  k = 1
  attach1 = True
 ElseIf Me.Top >= plst.Top + plst.height - 40 And Me.Top <= plst.Top + plst.height + 40 Then
  Me.Top = plst.Top + plst.height
   attach1 = True
  k = 2
 Else
  attach1 = False
 End If
If Me.height > 2500 Then
Me.height = 1765
Else
Me.height = 3610
If attach1 = True Then plst.Top = plst.Top - plst.height - 10
End If

If attach1 = True And k = 1 Then
plst.Top = Me.Top + Me.height - 10

plst.Show

End If

End Sub

Public Sub EQ_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
EQ.Picture = eqdown.Picture
End Sub

Public Sub EQ_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
EQ.Picture = equp.Picture

End Sub

Public Sub eqclose_Click()

Call EQ_Click
End Sub

Public Sub eqclose_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
eqclose.Picture = closedown.Picture

End Sub

Public Sub eqclose_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
eqclose.Picture = closeup.Picture

End Sub

Public Sub eqe_Click()
Call EQ_Click
'If frmeq.Visible = True Then eqe.Checked = True
'If frmeq.Visible = False Then eqe.Checked = False

End Sub


Private Sub eqf_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim m As POINTAPI
Dim k
If Button = 1 Then
k = GetCursorPos(m)
initialypos = m.Y
End If
eqf(Index).Picture = eqfdown.Picture

End Sub

Private Sub eqf_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim m As POINTAPI
Dim k
k = GetCursorPos(m)
If Button = 1 Then
If Abs(m.Y - initialypos) >= 1 Then
eqf(Index).Top = eqf(Index).Top + (m.Y - initialypos) * 12000 / 800
initialypos = m.Y
End If
If eqf(Index).Top >= 3060 Then eqf(Index).Top = 3060
If eqf(Index).Top <= 2310 Then eqf(Index).Top = 2310
'Image12(Index).height = -(590 - eqf(Index).Top) + 90
End If
If Index = 10 And Button = 1 Then
 If playspeed <> (eqf(10).Top - 2450) / 700# - 0.4 Then
  playspeed = (eqf(10).Top - 2450) / 700# - 0.4
  Form1.MediaPlayer1.Rate = 1 - playspeed
 End If
End If
'Dim i As Integer
'Dim l1, l2 As Integer
'Dim ind As Integer
'Line (800, 2000)-(3700, 2200), RGB(0, 0, 0), BF
'Line8.Visible = False
'Line8.Visible = True

'For i = 1300 To 2970 Step 20

'k = k + 1

'l1 = 200 * (eqf(ind).Top - 570) / 700
'l2 = 200 * (eqf(ind + 1).Top - 570) / 700
'If i > eqf(ind).Left Then
'Line (i, 2000 + l1)-(i + 270, 2000 + l2), RGB((200 - l1) + 50, l1, 0)

'ind = ind + 1
'End If
'Next

End Sub

Private Sub eqf_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
eqf(Index).Picture = eqfup.Picture

End Sub

Private Sub eqon_Click()
If eqon.Picture = eqonup.Picture Then
allowrate = False
eqf(10).Enabled = False
Form1.MediaPlayer1.Rate = 1
eqon.Picture = eqoffup.Picture
Else
eqon.Picture = eqonup.Picture
eqf(10).Enabled = True
Form1.MediaPlayer1.Rate = 1 - playspeed
allowrate = True
End If
End Sub

Private Sub eqon_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If eqon.Picture = eqonup.Picture Then
eqon.Picture = eqondown.Picture
Else
eqon.Picture = eqoffdown.Picture
End If
End Sub


Private Sub eqon_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If eqon.Picture = eqondown.Picture Then
eqon.Picture = eqonup.Picture
Else
eqon.Picture = eqoffup.Picture
End If

End Sub

Private Sub auto_Click()
If auto.Picture = autoonup.Picture Then
auto.Picture = autooffup.Picture
Else
auto.Picture = autoonup.Picture
End If
Form1.MediaPlayer1.Rate = 1
playspeed = 0
eqf(10).Top = 2660
End Sub

Private Sub auto_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If auto.Picture = autoonup.Picture Then
auto.Picture = autoondown.Picture
Else
auto.Picture = autooffdown.Picture
End If
End Sub


Private Sub auto_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If auto.Picture = autoondown.Picture Then
auto.Picture = autoonup.Picture
Else
auto.Picture = autooffup.Picture
End If
End Sub


Private Sub Image12_Click(Index As Integer)
End Sub

Private Sub Form_Paint()
Static k1, k2, a, b
Me.Refresh
'DoEvents
If musicsystem.WindowState = 1 Then
'Me.Caption = "Mahesh's Mediaplayer"
If attach3 = False Then plst.Hide
'
Else 'If 'Me.WindowState = 0 Then

If mini = True Then plst.Show 'eql.WindowState = 2
'Call Image22_MouseDown(1, 0, 50, 59)
End If
End Sub

Private Sub Image18_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
moovemain = False
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
End Sub

Private Sub Image22_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
moovemain = False
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
End Sub

Private Sub Image24_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim m As POINTAPI
Dim k
moovemain = True
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
If Form1.Left >= Me.Left + Me.width - 40 And Form1.Left <= Me.Left + Me.width + 40 Then
 Form1.Left = Me.Left + Me.width
 attach2 = True
 ElseIf Me.Left >= Form1.Left + Form1.width - 40 And Me.Left <= Form1.Left + Form1.width + 40 Then
  Me.Left = Form1.Left + Form1.width
 attach2 = True
ElseIf Form1.Top >= Me.Top + Me.height - 40 And Form1.Top <= Me.Top + Me.height + 40 Then
 Form1.Top = Me.Top + Me.height
 attach2 = True
 ElseIf Me.Top >= Form1.Top + Form1.height - 40 And Me.Top <= Form1.Top + Form1.height + 40 Then
 Me.Top = Form1.Top + Form1.height
 attach2 = True

'form1.Left = Me.Left + Me.width
'attach2 = True

Else
attach2 = False
End If
initialpx = m.X - plst.Left / 12000 * 800
initialpy = m.Y - plst.Top / 12000 * 800
initialfx = m.X - Form1.Left / 12000 * 800
initialfy = m.Y - Form1.Top / 12000 * 800
If attach3 = True And attach2 = True Then
attach1 = True
attach3 = False
End If
End Sub

Private Sub Image24_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim k
Dim m As POINTAPI
Dim width As Long
Dim height As Long
text1.SetFocus
k = GetCursorPos(m)
'm.X = (X - Me.Left) * 277 / 4110
'm.x = m.x - initialxpos
'm.y = m.y - initialypos
If moovemain = True Then
width = Me.width / 12000 * 800
height = Me.height / 12000 * 800
If Button = 1 Then

If (((attach1 = False) Or (plst.Visible = False)) And ((attach2 = False) Or (Form1.Visible = False))) Then
Call Drag(Me)

Else
k = MoveWindow(Me.hwnd, m.X - initialxpos, m.Y - initialypos, width, height, 1)
If attach2 = True Then k = MoveWindow(Form1.hwnd, m.X - initialfx, m.Y - initialfy, Form1.width / 12000 * 800, Form1.height / 12000 * 800, 1)
If attach1 = True Then k = MoveWindow(plst.hwnd, m.X - initialpx, m.Y - initialpy, plst.width / 12000 * 800, plst.height / 12000 * 800, 1)
End If
End If
End If
End Sub

Private Sub Image24_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
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
End Sub

Private Sub MMControl2_RecordClick(Cancel As Integer)
Dim parms As MCI_WAVE_SET_PARMS
   Dim rc As Long
   Dim Msg As String * 300
   
   ' Set the record/playback device for mmcontrol2
   parms.wInput = Combo2.ListIndex
   parms.wOutput = Combo2.ListIndex
   rc = mciSendCommand(MMControl2.DEVICEID, _
                        MCI_SET, _
                        MCI_WAVE_INPUT Or MCI_WAVE_OUTPUT, _
                        parms)
                        
   If (rc <> NO_ERROR) Then
      mciGetErrorString rc, Msg, Len(Msg)
     ' MsgBox msg
   End If
   
   ' Use control panel record format if option is checked
   If (Check1.Value = 1) Then
      Dim wf As WAVEFORMAT
      If (GetDefaultWaveFormat(wf) = False) Then
       '/  MsgBox "Couldn't get the default format"
      Else
         parms.nAvgBytesPerSec = wf.nAvgBytesPerSec
         parms.nBlockAlign = wf.nBlockAlign
         parms.nChannels = wf.nChannels
         parms.nSamplesPerSec = wf.nSamplesPerSec
         parms.wBitsPerSample = wf.wBitsPerSample
         parms.wFormatTag = wf.wFormatTag
         rc = mciSendCommand(MMControl2.DEVICEID, _
                              MCI_SET, _
                              MCI_WAVE_SET_SAMPLESPERSEC Or _
                              MCI_WAVE_SET_AVGBYTESPERSEC Or _
                              MCI_WAVE_SET_BITSPERSAMPLE Or _
                              MCI_WAVE_SET_BLOCKALIGN Or _
                              MCI_WAVE_SET_CHANNELS Or _
                              MCI_WAVE_SET_FORMATTAG, _
                              parms)
                              
         If (rc <> NO_ERROR) Then
            mciGetErrorString rc, Msg, Len(Msg)
            MsgBox Msg
         End If
      End If
   End If
End Sub


Private Sub preset_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
preset.Picture = presetdown.Picture
End Sub


Private Sub preset_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
preset.Picture = presetup.Picture

End Sub

Public Sub exe_Click()
Call meclose_Click
End Sub

Public Sub fast_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 text1.SetFocus

End Sub



Private Sub Form_Activate()
On Error GoTo 1
cmo = True
plst.List1.Selected(plst.List1.ListIndex) = False
plst.Scroll.Top = 90
plst.Frame1(0).Visible = False
plst.Frame1(1).Visible = False
Slider3.Value = vol.WaveLevel / 3275 + 1
volumebar.Left = 1550 + Slider3.Value * 40
volinf.Caption = "Volume" + str((Slider3.Value - 1) * 5) + " %" '+ Form1.MediaPlayer1.volume) / 25))) + " %"

'form1.scroll.top
1:
End Sub
Public Sub SetWindowStyle(ByVal hwnd As Long, ByVal extended_style As Boolean, ByVal style_value As Long, ByVal new_value As Boolean, ByVal brefresh As Boolean)
   
   Dim style_type As Long
   Dim style As Long
   
   If extended_style Then
      style_type = GWL_EXSTYLE
   Else
      style_type = GWL_STYLE
   End If
   
   ' Get the current style.
   style = GetWindowLong(hwnd, style_type)
   
   ' Add or remove the indicated value.
   If new_value Then
      style = style Or style_value
   Else
      style = style And Not style_value
   End If
   
   ' Hide Window if Changing ShowInTaskBar
   If brefresh Then
      ShowWindow hwnd, 5
   End If
   
   ' Set the style.
   SetWindowLong hwnd, style_type, style
   
   ' Show Window if Changing ShowInTaskBar
   If brefresh Then
      ShowWindow hwnd, 0
   End If
   
   ' Make the window redraw.
   
End Sub
Private Sub Timer6_Timer()
Dim str, str1 As String
If playtrack.Caption = "" Then playtrack.Caption = ". Mahesh's Mediaplayer . "
   ' str = Mid$(frmMirror.Caption, 2, Len(frmMirror.Caption))
   ' frmMirror.Caption = str & Mid$(frmMirror.Caption, 1, 1)
    str1 = Mid$(playtrack.Caption, 2, Len(playtrack))
    playtrack.Caption = str1 & Mid$(playtrack.Caption, 1, 1)
    frmMirror.Caption = playtrack.Caption
End Sub
Private Function GetDefaultWaveFormat(format As WAVEFORMAT) As Boolean
'////////////////////////////////////////////////////////////////////////////////////
' This user-defined function retrieves the default wave format from the registry.
'////////////////////////////////////////////////////////////////////////////////////
   Dim rc As Long
   Dim key1 As Long
   Dim key2 As Long
   Dim formatName As String * 50
   Dim length As Long
   
   ' Initialize return code
   GetDefaultWaveFormat = False
    
   rc = RegOpenKeyEx(HKEY_CURRENT_USER, _
                     "Software\Microsoft\Multimedia\Audio", _
                     0, _
                     KEY_READ, _
                     key1)
   If (rc <> 0) Then
      Exit Function
   End If
   
   length = Len(formatName)
   rc = RegQueryValueString(key1, "DefaultFormat", 0, 0, formatName, length)
   
   If (NO_ERROR = rc) Then
      rc = RegOpenKeyEx(HKEY_CURRENT_USER, _
                        "Software\Microsoft\Multimedia\Audio\WaveFormats", _
                        0, _
                        KEY_READ, _
                        key2)
                        
      If (NO_ERROR = rc) Then
         length = Len(format)
         rc = RegQueryValueEx(key2, _
                              formatName, _
                              0, _
                              0, _
                              format, _
                              length)
         RegCloseKey key2
         
         If (NO_ERROR = rc) Then
            GetDefaultWaveFormat = True
         End If
      End If
   End If
   RegCloseKey key1
End Function

Public Sub Form_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
'Source.Left = X
End Sub

Private Sub Form_GotFocus()
On Error GoTo 9
'Slider1.SetFocus
'Label15.ForeColor = &H80FF80    'RGB(200, 0, 200)
'plst.Label1.ForeColor = RGB(0, 0, 0)
On Error GoTo 1
plst.List1.Selected(plst.List1.ListIndex) = False
1:
9:
End Sub

Public Sub Form_LostFocus()
Label15.ForeColor = RGB(0, 0, 0)
'Me.Caption = ""
'Timer5.Enabled = False
End Sub

Public Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer
On Error GoTo 9
'Me.Caption = ""
'Me.BorderStyle = 0
Imgbar.Picture = imgbarup.Picture
For i = 0 To 7 Step 1
musicsystem.cbutton(i).Picture = musicsystem.cbuttonE(i).Picture
Next
'If Frame1.Visible = True Then Frame1.Visible = False
text1.SetFocus
9:
End Sub

Public Sub Form_Resize()
Form_Paint
End Sub

Public Sub Form_Unload(Cancel As Integer)

End Sub

Public Sub Image10_DragOver(Source As Control, X As Single, Y As Single, State As Integer)

End Sub

Public Sub Image10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Call Image10_DragOver(volumebar, X, Y, 2)
volumebar.Left = Image10.Left + X - 100
If volumebar.Left > 2400 Then
  volumebar.Left = 2400
  Slider3.Value = 21
ElseIf (volumebar.Left < 1670) Then
   volumebar.Left = Image10.Left
   Slider3.Value = 1
Else
  Slider3.Value = Int(((volumebar.Left - Image10.Left - 15) / 37)) + 1
End If
'Form1.MediaPlayer1.volume = -(6500 - (volumebar.Left - Image10.Left) / 835 * 6500) + 300
vol.WaveLevel = (Slider3.Value - 1) * 3275
playtrack.Visible = False

'playtrack.Left = 1620
'Text1.SetFocus
volinf.Visible = True

volinf.Caption = "Volume" + str((Slider3.Value - 1) * 5) + " %" '+ Form1.MediaPlayer1.volume) / 25))) + " %"

End Sub

Public Sub Image10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Slider3.SetFocus
End Sub


Public Sub Image17_Click()

End Sub



Public Sub Image19_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Me.Caption = ""

End Sub

Public Sub Image20_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Me.Caption = "Mahesh's MediaPlayer"
Me.BorderStyle = 4
End Sub

Public Sub Image22_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim m As POINTAPI
Dim k
moovemain = True
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
If Form1.Left >= Me.Left + Me.width - 40 And Form1.Left <= Me.Left + Me.width + 40 Then
 Form1.Left = Me.Left + Me.width
 attach2 = True
 ElseIf Me.Left >= Form1.Left + Form1.width - 40 And Me.Left <= Form1.Left + Form1.width + 40 Then
  Me.Left = Form1.Left + Form1.width
 attach2 = True
ElseIf Form1.Top >= Me.Top + Me.height - 40 And Form1.Top <= Me.Top + Me.height + 40 Then
 Form1.Top = Me.Top + Me.height
 attach2 = True
 ElseIf Me.Top >= Form1.Top + Form1.height - 40 And Me.Top <= Form1.Top + Form1.height + 40 Then
 Me.Top = Form1.Top + Form1.height
 attach2 = True

'form1.Left = Me.Left + Me.width
'attach2 = True

Else
attach2 = False
End If
initialpx = m.X - plst.Left / 12000 * 800
initialpy = m.Y - plst.Top / 12000 * 800
initialfx = m.X - Form1.Left / 12000 * 800
initialfy = m.Y - Form1.Top / 12000 * 800
If attach3 = True And attach2 = True Then
attach1 = True
attach3 = False
End If
 musicsystem.Image22.Picture = musicsystem.Image21.Picture

End Sub

Public Sub Image22_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim k
Dim m As POINTAPI
Dim width As Long
Dim height As Long

text1.SetFocus
k = GetCursorPos(m)
'm.X = (X - Me.Left) * 277 / 4110
'm.x = m.x - initialxpos
'm.y = m.y - initialypos
If moovemain = True Then
width = Me.width / 12000 * 800
height = Me.height / 12000 * 800
If Button = 1 Then
If (((attach1 = False) Or (plst.Visible = False)) And ((attach2 = False) Or (Form1.Visible = False))) Then
Call Drag(Me)
Else
k = MoveWindow(Me.hwnd, m.X - initialxpos, m.Y - initialypos, width, height, 1)
If attach2 = True Then k = MoveWindow(Form1.hwnd, m.X - initialfx, m.Y - initialfy, Form1.width / 12000 * 800, Form1.height / 12000 * 800, 1)
If attach1 = True Then k = MoveWindow(plst.hwnd, m.X - initialpx, m.Y - initialpy, plst.width / 12000 * 800, plst.height / 12000 * 800, 1)

End If
End If
End If





End Sub


Public Sub Image3_Click(Index As Integer)
If repeat = False Then
Image3(0).Picture = Image3(1).Picture
repeat = True
ElseIf repeat = True Then
Image3(0).Picture = Image3(2).Picture
repeat = False
End If
' Text1.SetFocus

End Sub

Public Sub Image3_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Image3(0).Picture = Image3(1).Picture Then Image3(0).Picture = Image3(4).Picture
If Image3(0).Picture = Image3(2).Picture Then Image3(0).Picture = Image3(3).Picture

End Sub

Public Sub Image3_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Image3(0).Picture = Image3(4).Picture Then Image3(0).Picture = Image3(1).Picture
If Image3(0).Picture = Image3(3).Picture Then Image3(0).Picture = Image3(2).Picture

End Sub

Public Sub Image7_Click(Index As Integer)
Form2.Show
End Sub



Public Sub Image8_DragOver(Index As Integer, Source As Control, X As Single, Y As Single, State As Integer)

End Sub

Public Sub Image8_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo 1
Imgbar.Left = Image8(1).Left + X - 160
If Imgbar.Left >= 3480 Then
Imgbar.Left = 3510
Form1.MediaPlayer1.CurrentPosition = Form1.MediaPlayer1.Duration - 1
ElseIf Imgbar.Left <= Image8(1).Left + 20 Then
Imgbar.Left = Image8(1).Left + 10
Form1.MediaPlayer1.CurrentPosition = 2
Else
Form1.MediaPlayer1.CurrentPosition = Form1.MediaPlayer1.Duration * ((Imgbar.Left - Image8(1).Left) / (Image8(1).width - 435))
End If
Slider2.Value = (Form1.MediaPlayer1.CurrentPosition / Form1.MediaPlayer1.Duration) * 70
Slider2.SetFocus
1:
End Sub

Public Sub Image8_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'Imgbar.Picture = imgbarup.Picture
Slider2.SetFocus
End Sub

Public Sub imgbalancemeter_DragOver(Source As Control, X As Single, Y As Single, State As Integer)

End Sub

Public Sub imgbalancemeter_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo 1

balancebar.Left = imgbalancemeter.Left + X
If balancebar.Left <= 2600 Then balancebar.Left = 2600
If (balancebar.Left >= 3120) Then balancebar.Left = 3120

Form1.MediaPlayer1.Balance = (balancebar.Left - 2850) * 9

Slider4.Value = (balancebar.Left - 2610) / 72
'Slider3.SetFocus




'Slider4.SetFocus
Slider4.SetFocus
1:
End Sub

Public Sub Imgbar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim m As POINTAPI
Dim k
k = GetCursorPos(m)
initialypos = m.Y '- Me.Top / 12000 * 800

initialxpos = m.X '- Me.Left / 12000 * 800



Slider2.SetFocus

End Sub

Public Sub Imgbar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim m As POINTAPI
Dim k
On Error GoTo 1
If Button = 1 Then
k = GetCursorPos(m)
If Abs(m.X - initialxpos) >= 2 Then
Imgbar.Left = Imgbar.Left + ((m.X - initialxpos) * 12000 / 800)
initialxpos = m.X
End If


'If X <= (Image8(1).width - 60) Then Imgbar.Left = Image8(1).Left + X
If Imgbar.Left >= 3480 Then Imgbar.Left = 3510
If Imgbar.Left <= Image8(1).Left + 20 Then Imgbar.Left = Image8(1).Left
Form1.MediaPlayer1.CurrentPosition = Form1.MediaPlayer1.Duration * ((Imgbar.Left - Image8(1).Left) / (Image8(1).width - 435))
Slider2.Value = (Form1.MediaPlayer1.CurrentPosition / Form1.MediaPlayer1.Duration) * 70
Slider2.SetFocus
End If

'Slider2.SetFocus
1:
Slider2.SetFocus
End Sub


Public Sub Imgbar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Imgbar.Picture = imgbarup.Picture

End Sub

Public Sub Inf_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Inf.Picture = Infdown.Picture
End Sub

Public Sub Inf_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Inf.Picture = infup.Picture
PopupMenu file
End Sub





Public Sub label18_Click()

End Sub

Public Sub loadpl_Click()

End Sub



Public Sub meclose_Click()
On Error GoTo 9
Dim i As Integer
Dim temptrack As file
temptrack.path = ""
temptrack.name = ""
Open App.path + "\mplayerlist.mmpl" For Random As #1 Len = 155
For i = 1 To totaltracks Step 1
Put #1, i, track(i)
Next
Put #1, totaltracks + 1, temptrack
Close #1
' *** Clear the file drop windows hook
SaveSetting App.EXEName, "Startup", "Top", Me.Top
SaveSetting App.EXEName, "Startup", "Left", Me.Left
SaveSetting App.EXEName, "Startup", "height", Me.height
SaveSetting App.EXEName, "Startup", "plstheight", plst.height
SaveSetting App.EXEName, "Startup", "plsttop", plst.Top
SaveSetting App.EXEName, "Startup", "plstleft", plst.Left
SaveSetting App.EXEName, "Startup", "form1top", Form1.Top
SaveSetting App.EXEName, "Startup", "form1left", Form1.Left
SaveSetting App.EXEName, "Startup", "form1height", Form1.height
SaveSetting App.EXEName, "Startup", "form1width", Form1.width
'Call DisableFileDrops
9:
Unload frmMirror



'Unload aboutme
'Unload Formrec
'mp3.PlayStop

End Sub

Public Sub meclose_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
meclose.Picture = closedown.Picture

End Sub

Public Sub meclose_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
meclose.Picture = closeup.Picture

End Sub

Public Sub min_Click()
'Me.Caption = ""

'Me.Visible = False
'plst.Visible = False
frmMirror.WindowState = vbMinimized
End Sub

Public Sub min_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
min.Picture = mindown.Picture
End Sub

Public Sub min_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
min.Picture = minup.Picture
End Sub


Public Sub PL_Click()
If plst.Visible = True Then
plst.Hide 'Show
mini = True
Else
plst.Show 'Show
mini = False
End If
End Sub

Public Sub PL_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
PL.Picture = pldown.Picture

End Sub

Public Sub PL_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
PL.Picture = plup.Picture

End Sub



Public Sub ple_Click()

Call PL_Click
'If frmpl.Visible = True Then ple.Checked = True
'If frmpl.Visible = False Then ple.Checked = False

End Sub



Public Sub savepl_Click()

End Sub


Public Sub showplmenu_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Frame1.Visible = False

End Sub



Private Sub rec_timer_Timer()
On Error Resume Next
Dim k As Long
rectime = (rectime) + 1
rec_count.Caption = format$(Int((rectime) / 600), "00") + ":" + format$(Int((((rectime) Mod 600) - 1) / 10), "00") + ":" + format$((rectime Mod 10) * 10, "00")
End Sub

Private Sub reclabel_Click()
On Error GoTo 1
If SaveAudiofile <> "" Then
Form1.MediaPlayer1.FileName = SaveAudiofile
'Form1.Hide
Form1.MediaPlayer1.Play
Form1.Hide
'tracklength.Caption = ""
'trackpos.Caption = ""
musicsystem.trackno.Caption = "Rec File"
Label17.Caption = rec_count.Caption
playtrack.Caption = "Recording File"
mediatype.Caption = ""
rec_count.Left = reclabel.Left + reclabel.width + 40
End If
1:
End Sub

Public Sub Slider2_GotFocus()

On Error GoTo 1
Imgbar.Picture = imgbardown.Picture
Slider2.max = 70
Slider2.Value = Form1.MediaPlayer1.CurrentPosition / Form1.MediaPlayer1.Duration * 70

1:
End Sub

Private Sub Slider2_KeyDown(Keycode As Integer, Shift As Integer)
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
End Sub

Public Sub Slider2_LostFocus()
Imgbar.Picture = imgbarup.Picture

End Sub

Public Sub Slider2_Scroll()
On Error GoTo 1
Form1.MediaPlayer1.CurrentPosition = Slider2.Value * Form1.MediaPlayer1.Duration / Slider2.max
Imgbar.Left = Image8(1).Left + Form1.MediaPlayer1.CurrentPosition / Form1.MediaPlayer1.Duration * 3270 - 30
1:
End Sub



Public Sub Slider3_GotFocus()
'volumebar.Picture = volumebardown.Picture
volinf.Visible = True
playtrack.Visible = False

End Sub

Private Sub Slider3_KeyDown(Keycode As Integer, Shift As Integer)
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
End Sub

Public Sub Slider3_LostFocus()
volumebar.Picture = volumebarup.Picture
volinf.Visible = False
playtrack.Visible = True
End Sub

Public Sub Slider3_Scroll()
'On Error Resume Next
'Static k As Boolean
If Slider3.Value <= 1 Then Slider3.Value = 1
volumebar.Left = 1550 + Slider3.Value * 40
If volumebar.Left > 2400 Then
volumebar.Left = 2400
ElseIf (volumebar.Left < 1660) Then volumebar.Left = Image10.Left ' + 10
End If

 'If Slider3.Value <= 1 Then k = True
' Slider3.Value = 0
vol.WaveLevel = (Slider3.Value - 1) * 3275
volinf.Caption = "Volume" + str((Slider3.Value - 1) * 5) + " %" '+ Form1.MediaPlayer1.volume) / 25))) + " %"
'Slider3.Value = ((volumebar.Left - Image10.Left) / 38.25)

'Call Image10_DragOver(volumebar, Slider3.Value * 90, Image10.Top - 40, 2)
End Sub

Public Sub Slider4_Change()
On Error GoTo 1
1:
End Sub

Public Sub Slider4_GotFocus()
balancebar.Picture = volumebardown.Picture
End Sub

Private Sub Slider4_KeyDown(Keycode As Integer, Shift As Integer)
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
End Sub

Public Sub Slider4_LostFocus()
balancebar.Picture = volumebarup.Picture
End Sub

Public Sub Slider4_Scroll()

balancebar.Left = imgbalancemeter.Left - 40 + Slider4.Value * 72

If balancebar.Left <= 2600 Then balancebar.Left = 2600
If (balancebar.Left >= 3120) Then balancebar.Left = 3120

Form1.MediaPlayer1.Balance = (balancebar.Left - 2850) * 9

'Slider4.Value = (balancebar.Left - 2610) / 72
'Slider3.SetFocus




'Slider4.SetFocus
Slider4.SetFocus
End Sub

Public Sub Command9_Click()
'MMC.FileName = form1.MediaPlayer1.FileName
'MMC.Command = "open"
End Sub



Public Sub Form_Load()
Dim oldCaption As String
Dim otherWnd As Long
Dim retValue As Long
Dim k
'On Error GoTo 44
'dim wavctrl as vol
'ok = GetVolumeControl(hmixer, _
        MIXERLINE_COMPONENTTYPE_SRC_WAVEOUT, _
        MIXERCONTROL_CONTROLTYPE_VOLUME, _
        wavctrl)
'trackno.Caption = "Mahesh the hero"




Dim sCommandLine            As String
mini = False
' *** Make a local copy of the command line
sCommandLine = Command()
' *** Check for registration commands
If Len(sCommandLine) <> 0 And _
   InStr(1, sCommandLine, "/") <> 0 Then
    ' There are command line arguments with switches
    ' indicating the possibility of registry actions
    Call SelfRegister(sCommandLine)
End If
'100,400 are default values if no entry is obtained
Me.Top = GetSetting(App.EXEName, "Startup", "Top", 100)
Me.Left = GetSetting(App.EXEName, "Startup", "Left", 400)
Me.height = GetSetting(App.EXEName, "Startup", "height", 3660)
'SaveSetting App.EXEName, "Startup", "plstheight", plst.height
'SaveSetting App.EXEName, "Startup", "plsttop", plst.Top
'SaveSetting App.EXEName, "Startup", "plstleft", plst.Left



    
Dim i As Integer





Slider3.min = vol.WaveMin / 3275
Slider3.max = vol.WaveMax / 3275 + 1
Slider3.Value = vol.WaveLevel / 3275 + 1
volumebar.Left = 1550 + Slider3.Value * 40
volinf.Caption = "Volume" + str((Slider3.Value - 1) * 5) + " %" '+ Form1.MediaPlayer1.volume) / 25))) + " %"

'Dim track As file


    
    ckk = True
    attach3 = False
    attach1 = True
    attach2 = True
    
eqon.Picture = eqonup.Picture
    
'DoEvents

sizescroll = 630
'On Error GoTo 1
DoEvents
rec_pause = False
rec = False
'colorbar.width = 345
attach1 = False
'DoEvents
'plst.Show


checkback = False
'
Image3(0).Picture = Image3(1).Picture


'On Error Resume Next
p(0).Picture = p(2).Picture
shuffle = False
i = 0
'
' *** Check for a non-registration command line with a
' file to import

Dim fill As String
Dim listname As String
Me.Visible = False
plst.Visible = True
Call EnableFileDrops(Me)
'Dim mm As Form
'Set mm = Me
'mhHookWindow1 = mm.hWnd
'If mhHookWindow1 <> 0 Then Call DisableFileDrops

'Dim mlPrevWndProc
' Set the subclassing window message hook
'mlPrevWndProc = SetWindowLong(mhHookWindow1, _
                -4, _
                AddressOf HookCallback)

' Tell the OS that the specified window accepts
' dropped files
'Call DragAcceptFiles(mhHookWindow1, True)






If Len(sCommandLine) <> 0 Then
'plst.newpl_Click
    ' Make sure it is a SAD format file
    
    'If (Right$(sCommandLine, 1) = """") Then sCommandLine = Left(sCommandLine, Len(sCommandLine) - 1)
    'If (Left$(sCommandLine, 1) = """") Then sCommandLine = Right(sCommandLine, Len(sCommandLine) - 1)
   
   If (Right$(sCommandLine, 1) = """") Then sCommandLine = Left(sCommandLine, (Len(sCommandLine) - 1))
    If (Left$(sCommandLine, 1) = """") Then sCommandLine = Right(sCommandLine, (Len(sCommandLine) - 1))
     Text4.Text = sCommandLine
     'sCommandLine = Text4.Text
    If UCase$(Right$(sCommandLine, 4)) = ".MP3" Or UCase$(Right$(sCommandLine, 4)) = ".DAT" Or UCase$(Right$(sCommandLine, 4)) = ".WAV" Or UCase$(Right$(sCommandLine, 4)) = ".MPG" Or UCase$(Right$(sCommandLine, 5)) = ".MPEG" Or UCase$(Right$(sCommandLine, 4)) = ".AVI" Then
        ' It is the correct file type so import it
        'Open App.path + "\mplayerlist.mmpl" For Random As #1 Len = 155
         totaltracks = 1
         track(1).path = sCommandLine
         File1.path = ""
         File1.FileName = sCommandLine
         listname = str(totaltracks) + ". " + LCase(File1.FileName)
         If Len(listname) >= 34 Then listname = Left(listname, 34) + "..."
         plst.List1.AddItem listname
         i = InStr(listname, ".")
         track(1).name = LCase(Mid(listname, i + 1, Len(listname)))
         ' Form1.MediaPlayer1.FileName = sCommandLine
         ' Form1.MediaPlayer1.Play
       ' Put #1, 1, track
        ' Close #1
    ElseIf UCase$(Right$(sCommandLine, 5)) = ".mmpl" Then
      i = 0
       fill = sCommandLine
       Open fill For Random As #2 Len = 155
      ' Open App.path + "\mplayerlist.mmpl" For Random As #1 Len = 155
       Do While (1)
          Get #2, i + 1, track(i + 1)
          'Put #1, i + 1, track
           If track(i + 1).name = "" Then Exit Do
           i = i + 1
       Loop
       'Close #1
       Close #2
       'fill = App.path + "\mplayerlist.mmpl"
      ' k = FileCopy(sCommandLine, App.path + "\mplayerlist.mmpl")
    End If
Else

i = 0
Open App.path + "\mplayerlist.mmpl" For Random As #1 Len = 155
currentplaylist = App.path + "\mplayerlist.mmpl"
Do While (1)
Get #1, i + 1, track(i + 1)
'On Error GoTo 3
If track(i + 1).name = "" Then Exit Do
plst.List1.AddItem str(i + 1) + "." + track(i + 1).name
i = i + 1
Loop
3:
totaltracks = i
For i = 1 To totaltracks Step 1
If i <= 12 Then
musicsystem.Label6(i - 1).Enabled = True
Label6(i - 1).ForeColor = Label6(15).ForeColor
End If
Next
Close #1
Close #2
End If
'......if player is being opened from supp. files then copy their paths to default playlist......

Me.width = plst.width '= 4125




If plst.List1.List(0) <> "" Then plst.List1.ListIndex = 0
'Get #1, 1, track
'form1.MediaPlayer1.FileName = track.path
currenttrack = 1
totalplayedtracks = 1




repeat = True
allowslide = True
If totaltracks > 0 Then
scrollunit = sizescroll / totaltracks
plst.Slider1.max = totaltracks
End If
1:
Close #1
If plst.List1.ListIndex < 0 Then totaltracks = 0

plst.Label18.Caption = str(currenttrack) + "/" + str(totaltracks)
'Text1.SetFocus
'plst.slider1.SetFocus
'Call EQ_Click
'If totaltracks = 1 Then Call plst.list1_DblClick

attach3 = False
    attach1 = True
    attach2 = True
    
For i = 1 To totaltracks Step 1
If i <= 12 Then musicsystem.Label6(i - 1).Enabled = True
Next
Dim numDevs As Long
   Dim idx As Long
   Dim outcaps As WAVEOUTCAPS
   Dim incaps As WAVEINCAPS
   
   ' Get the available wave output devices
   numDevs = waveOutGetNumDevs
   
   For idx = 0 To numDevs - 1
      waveOutGetDevCaps idx, outcaps, Len(outcaps)
      Combo1.AddItem outcaps.szPname
   Next
   
   ' Get the available wave input devices
   numDevs = waveInGetNumDevs
   
   For idx = 0 To numDevs - 1
      waveInGetDevCaps idx, incaps, Len(incaps)
      Combo2.AddItem incaps.szPname
   Next
'Formrec.Combo2.ListIndex = 0
'Formrec.Combo1.ListIndex = 0
rec = False
rec_pause = False
'If Right(App.path, 1) <> "\" Then
'File1.path = App.path + "\mmp Workarea"
'Else
'File1.path = App.path + "mmp Workarea"
'End If
'Dim mm As SECURITY_ATTRIBUTES
'fs=v
On Error Resume Next
MkDir (App.path + "\mmp Workarea")
File1.path = App.path + "\mmp Workarea"
'File1.path = "c:\mmp Workarea"
44:
k = SetWindowPos(Me.hwnd, -1, Me.Left / 12000 * 800, Me.Top / 9000 * 600, Me.width / 12000 * 800, Me.height / 9000 * 600, &H40)
frmMirror.Left = 99999
    frmMirror.Top = 99999
    DoEvents
    DoEvents
    SetWindowStyle hwnd, False, WS_CAPTION, False, False
    DoEvents
    DoEvents
frmMirror.Show
mini = False
If plst.List1.ListCount = 1 Then
    currenttrack = 1
    plst.List1.ListIndex = 0
    Call plst.list1_DblClick
    plst.List1.ListIndex = -1
    totaltracks = 1
    End If
'DoEvents
Me.Visible = True
End Sub


Public Sub Frame3_Click()
'Text2.SetFocus
'form1.MediaPlayer1.SetFocus
'SendKeys ("%{enter}"), True

End Sub

Public Sub Frame2_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
text1.Text = "frame"
End Sub

Public Sub Frame2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'bar.DragMode = 0

End Sub




Public Sub Frame7_DragDrop(Source As Control, X As Single, Y As Single)
On Error GoTo 9

If X >= 2760 And X <= 4200 Then
Source.Left = X
Form1.MediaPlayer1.Rate = 1 + (X - 34800) / 120 * 0.14
End If

9:
End Sub

Public Sub Frame7_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
On Error GoTo 9

If X >= 2760 And X <= 4200 Then
Source.Left = X
Form1.MediaPlayer1.Rate = 1 + (X - 3480) / 120 * 0.14
End If
9:
End Sub

Public Sub Image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Image2.BorderStyle = 1
End Sub

Public Sub Image2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Image2.BorderStyle = 0

End Sub







Public Sub fullscr_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 text1.SetFocus

End Sub











Public Sub Label16_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'
'musicsystem.Caption = " Mahesh's MediaPlayer"

End Sub




Public Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Text1.Text = Str(X) + ".." + Str(Y)
'bar.DragMode = 1
'bar.Enabled = True

End Sub



Public Sub label6_Click(Index As Integer)
Dim i As Integer
On Error GoTo 1
For i = 0 To 11 Step 1
 musicsystem.Label6(i).ForeColor = musicsystem.Label6(15).ForeColor
Next
musicsystem.Label6(Index).ForeColor = trackno.ForeColor

  plst.List1.ListIndex = Index
  Call plst.list1_DblClick
text1.SetFocus
1:

End Sub


Public Sub label6_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
musicsystem.Label6(Index).Appearance = 1
'musicsystem.Label6(Index).BackColor = Frame6.BackColor

End Sub











Public Sub repeattracks_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 text1.SetFocus

End Sub






Public Sub SLOW_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 text1.SetFocus

End Sub

Public Sub Text1_KeyDown(Keycode As Integer, Shift As Integer)
On Error GoTo 1
text1.Text = Keycode
If Keycode = 39 Then
Call musicsystem.cbutton_Click(5)
'If bar.Width <= Label3.Width - 100 Then bar.Width = bar.Width + 120
'Form1.MediaPlayer1.CurrentPosition = Form1.MediaPlayer1.Duration * bar.Width / Label13.Width
ElseIf Keycode = 37 Then
Call musicsystem.cbutton_Click(4)

'If bar.Width >= 120 Then bar.Width = bar.Width - 120
'Form1.MediaPlayer1.CurrentPosition = Form1.MediaPlayer1.Duration * bar.Width / Label13.Width
End If
If Keycode = 40 Then


Call Image10_MouseDown(1, 0, volumebar.Left - 100 - Image10.Left, Image10.Top + 40)
text1.SetFocus
End If

If Keycode = 38 Then

Call Image10_MouseDown(1, 0, volumebar.Left - Image10.Left + volumebar.width - 100, Image10.Top + 40)
text1.SetFocus

End If
text1.SetFocus

If Keycode = 13 Then
Call plst.list1_DblClick
End If

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
1:
End Sub

Public Sub text1_KeyPress(KeyAscii As Integer)
'Text2.Text = KeyAscii
End Sub





Public Sub Timer1_Timer()
'HScroll1.Value = (MMC.Position / 5000)
Dim i
On Error GoTo 3

If Form1.MediaPlayer1.Duration = 0 Then Exit Sub
  If Form1.MediaPlayer1.PlayState = mpPlaying Then
 If Form1.MediaPlayer1.CurrentPosition >= Form1.MediaPlayer1.Duration - 1 Then
 
Call nexttrackplay
End If
'Sleep(500)


    
Form1.MediaPlayer1.DisplayMode = mpTime
  trackpos.Caption = format((Int(Form1.MediaPlayer1.CurrentPosition / 60)), "00") + ":" + format(Int(Int(Form1.MediaPlayer1.CurrentPosition Mod 60)), "00")
  'If Form1.Visible = True Then Form1.trackpos.Caption = trackpos.Caption

  If showtime = 0 Then plst.Label17.Caption = trackpos.Caption
  If showtime = 1 Then plst.Label17.Caption = "-" + format((Int((Form1.MediaPlayer1.Duration - Form1.MediaPlayer1.CurrentPosition) / 60)), "00") + ":" + format(Int(Int((Form1.MediaPlayer1.Duration - Form1.MediaPlayer1.CurrentPosition) Mod 60)), "00") ' Form1.MediaPlayer1.Duration - MediaPlayer1.CurrentPosition)
   trackpos.Caption = plst.Label17.Caption
  If Form1.Visible = True Then Form1.trackpos.Caption = trackpos.Caption

  Imgbar.Left = Image8(1).Left + Form1.MediaPlayer1.CurrentPosition / Form1.MediaPlayer1.Duration * 3270

    If currenttrack <= 12 Then
If musicsystem.Label6(currenttrack - 1).ForeColor = musicsystem.Label6(9).ForeColor Then
musicsystem.Label6(currenttrack - 1).ForeColor = RGB(255, 200, 105)
Else
musicsystem.Label6(currenttrack - 1).ForeColor = musicsystem.Label6(9).ForeColor
End If
End If
End If
GoTo 4
3:
Timer2.Enabled = False
Timer3.Enabled = False
 For i = 0 To 8 Step 1
    sp(i).Top = 720
    sp(i).height = 285 '- 285 + (sp(0).Height)
   Next
4:
plst.Label18.Caption = str(currenttrack) + "/" + str(totaltracks)
End Sub

Public Sub Timer2_Timer()
Dim i As Integer
'If MMC.n = False Then
On Error GoTo 9

  If Form1.MediaPlayer1.PlayState = mpPlaying Then
  Timer3.Enabled = True
   For i = 0 To 8 Step 1
    
    sp(i).height = Rnd() * (300)
    sp(i).Top = (720 + 285) - sp(i).height
   Next
   'bar.Left = Label3.Left
 
  
 Else
 Timer3.Enabled = False
   For i = 0 To 8 Step 1
    sp(i).Top = 720 + 285
    sp(i).height = 35 '360 + (620)
   Next
   If Form1.MediaPlayer1.Duration = 0 Then Imgbar.Left = Image8(1).Left
 End If
9:
End Sub

Public Sub nexttrackplay()
Dim fullscr As Boolean
Dim i
On Error GoTo 5
'Form1.MediaPlayer1.Stop
DoEvents
'Me.Hide
'Me.Enabled = True
'SendKeys "%{enter}", True
'musicsystem.SetFocus
'plst.list1.SetFocus
'SendKeys "%{enter}", True
fullscr = False
Form1.MediaPlayer1.Stop
If Form1.MediaPlayer1.DisplaySize = mpFullScreen Then
Form1.MediaPlayer1.DisplaySize = mpFitToSize
fullscr = True
End If

'plst.list1.SetFocus

plst.List1.ListIndex = currenttrack - 1
If totaltracks = 1 And repeat <> True Then Exit Sub
If repeat = True And currenttrack = totaltracks And shuffle = False Then
plst.List1.ListIndex = 0
currenttrack = 1
GoTo 4
ElseIf repeat = False And currenttrack = totaltracks And shuffle = False Then
'form1.MediaPlayer1.FileName = "F:\Documents and Settings\Brijendra Tiwari\Desktop\video.bmp"
'form1.MediaPlayer1.Play
'Call musicsystem.cbutton_Click(6)
 'Label3.Width

Form1.MediaPlayer1.Stop
Form1.MediaPlayer1.CurrentPosition = 0
'form1.MediaPlayer1.CurrentPosition = 0.5 'Pause
'musicsystem.cbutton(2).Enabled = True
'musicsystem.cbutton(2).Picture = musicsystem.cbuttonE(2).Picture
'musicsystem.cbutton(3).Picture = musicsystem.cbuttonD(3).Picture
'musicsystem.cbutton(3).Enabled = False
'musicsystem.cbutton(6).Picture = musicsystem.cbuttonD(6).Picture
'musicsystem.cbutton(6).Enabled = False
'form1.MediaPlayer1.CurrentPosition = form1.MediaPlayer1.Duration - 1
'musicsystem.cbutton(4).Picture = musicsystem.cbuttonD(4).Picture
'musicsystem.cbutton(4).Enabled = False
'musicsystem.cbutton(5).Picture = musicsystem.cbuttonD(5).Picture
'musicsystem.cbutton(5).Enabled = False
'bar.width = 30
'form1.MediaPlayer1.Stop
GoTo 5
End If


totalreplayedtracks = 0

If shuffle = True Then
1:
 i = Int(Rnd() * totaltracks)
 If (i = currenttrack Or i = 0) And totaltracks > 1 Then
  GoTo 1:
 Else
  plst.List1.ListIndex = i - 1
 
  currenttrack = i
 End If
 ElseIf currenttrack <> totaltracks Then
 plst.List1.ListIndex = currenttrack
 currenttrack = currenttrack + 1
  
End If
'form1.MediaPlayer1.SetFocus

4:

plst.list1_DblClick
If fullscr = True Then Form1.MediaPlayer1.DisplaySize = mpFullScreen
fullscr = False
5:
End Sub

Public Sub Timer3_Timer()

End Sub





Public Sub Timer4_Timer()
Static twice As Integer
On Error GoTo 9

twice = twice + 1
'form1.MediaPlayer1.FileName = "E:\efysoftware\vcd\video1.jpg"
'form1.MediaPlayer1.Play
Form1.MediaPlayer1.Pause
If twice Mod 2 = 0 Then
Timer4.Enabled = False
'musicsystem.Width = 4890 - 190
Call nexttrackplay
End If

9:
End Sub

Public Sub Timer5_Timer()
'DoEvents
On Error GoTo 9

'If musicsystem.WindowState = 1 Then
'Form1.WindowState = 1
'Else
'Form1.WindowState = 0
'End If
'If Form1.Left >= musicsystem.Left + musicsystem.width - 1000 And Form1.Left <= musicsystem.Left + musicsystem.width + 1200 Then
'Form1.Left = musicsystem.Left + musicsystem.width
'Form1.Top = musicsystem.Top
'End If
9:
End Sub






Public Sub trackpos_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call plst.Label17_Click
trackpos.ToolTipText = plst.Label17.ToolTipText
text1.SetFocus
End Sub






Public Sub volumebar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim m As POINTAPI
Dim k
k = GetCursorPos(m)
initialypos = m.Y '- Me.Top / 12000 * 800

initialxpos = m.X '- Me.Left / 12000 * 800

Form1.MediaPlayer1.Mute = False
playtrack.Visible = False

'playtrack.Left = 1620
volinf.Caption = "Volume" + str(Abs(Int((volumebar.Left - Image10.Left) / 7.65))) + " %" '+ Form1.MediaPlayer1.volume) / 25))) + " %"
'Text1.SetFocus
volinf.Visible = True

End Sub

Public Sub volumebar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim m As POINTAPI
Dim k
On Error GoTo 1

If Button = 1 Then
k = GetCursorPos(m)
If Abs(m.X - initialxpos) >= 2 Then
volumebar.Left = volumebar.Left + ((m.X - initialxpos) * 12000 / 800)
initialxpos = m.X
End If

If volumebar.Left > 2400 Then
  volumebar.Left = 2400
  Slider3.Value = 21
ElseIf (volumebar.Left < 1670) Then
   volumebar.Left = Image10.Left
   Slider3.Value = 1
Else
  Slider3.Value = Int(((volumebar.Left - Image10.Left - 15) / 37)) + 1
End If
'Form1.MediaPlayer1.volume = -(6500 - (volumebar.Left - Image10.Left) / 835 * 6500) + 300
vol.WaveLevel = (Slider3.Value - 1) * 3275

volinf.Caption = "Volume" + str((Slider3.Value - 1) * 5) + " %" '+ Form1.MediaPlayer1.volume) / 25))) + " %"
'Slider3.Value = ((volumebar.Left - Image10.Left) / 38.25)

playtrack.Visible = False

'Slider3.Value = (volumebar.Left - Image10.Left) / 90
End If
Slider3.SetFocus
1:
End Sub

Private Sub Label15_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Image22_MouseDown(Button, Shift, Label15.Left + X, Y)

End Sub

Private Sub Label15_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Call Image22_MouseMove(Button, Shift, Label15.Left + X, Y)
End Sub

Private Sub volumebar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
playtrack.Visible = True
volinf.Visible = False
End Sub


Public Sub Image18_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim m As POINTAPI
Dim k
moovemain = True
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
If Form1.Left >= Me.Left + Me.width - 40 And Form1.Left <= Me.Left + Me.width + 40 Then
 Form1.Left = Me.Left + Me.width
 attach2 = True
 ElseIf Me.Left >= Form1.Left + Form1.width - 40 And Me.Left <= Form1.Left + Form1.width + 40 Then
  Me.Left = Form1.Left + Form1.width
 attach2 = True
ElseIf Form1.Top >= Me.Top + Me.height - 40 And Form1.Top <= Me.Top + Me.height + 40 Then
 Form1.Top = Me.Top + Me.height
 attach2 = True
 ElseIf Me.Top >= Form1.Top + Form1.height - 40 And Me.Top <= Form1.Top + Form1.height + 40 Then
 Me.Top = Form1.Top + Form1.height
 attach2 = True

'form1.Left = Me.Left + Me.width
'attach2 = True

Else
attach2 = False
End If
initialpx = m.X - plst.Left / 12000 * 800
initialpy = m.Y - plst.Top / 12000 * 800
initialfx = m.X - Form1.Left / 12000 * 800
initialfy = m.Y - Form1.Top / 12000 * 800
If attach3 = True And attach2 = True Then
attach1 = True
attach3 = False
End If
End Sub

Public Sub Image18_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim k
Dim m As POINTAPI
Dim width As Long
Dim height As Long
text1.SetFocus
k = GetCursorPos(m)
'm.X = (X - Me.Left) * 277 / 4110
'm.x = m.x - initialxpos
'm.y = m.y - initialypos
If moovemain = True Then
width = Me.width / 12000 * 800
height = Me.height / 12000 * 800
If Button = 1 Then
If (((attach1 = False) Or (plst.Visible = False)) And ((attach2 = False) Or (Form1.Visible = False))) Then
Call Drag(Me)
        musicsystem.Image22.Picture = musicsystem.Image21.Picture

Else
k = MoveWindow(Me.hwnd, m.X - initialxpos, m.Y - initialypos, width, height, 1)
If attach2 = True Then k = MoveWindow(Form1.hwnd, m.X - initialfx, m.Y - initialfy, Form1.width / 12000 * 800, Form1.height / 12000 * 800, 1)
If attach1 = True Then k = MoveWindow(plst.hwnd, m.X - initialpx, m.Y - initialpy, plst.width / 12000 * 800, plst.height / 12000 * 800, 1)
End If
End If
End If

End Sub

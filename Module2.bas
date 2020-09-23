Attribute VB_Name = "Module2"

Declare Function LocalHandle Lib "kernel32" (wMem As Any) As Long

Public allowrate As Boolean
Public playingtrack As String
'Public playingtrack As String
'Public selectedtrack As String
Public currenttrack As Integer
Public selectedtrack As Integer
Public previoustrack(1 To 50) As Integer
Public totalplayedtracks As Integer
Public totalreplayedtracks As Integer
Public SaveAudiofile As String
Public totaltracks As Integer
Public rec As Boolean
Public rec_pause As Boolean

'Public trackpos As Integer
Public shuffle As Boolean
Public repeat As Boolean
'Public track As String
'Declare Function mciSendCommand Lib "winmm.dll" Alias "mciSendCommandA" (ByVal wDeviceID As Long, ByVal uMessage As Long, ByVal dwParam1 As Long, ByVal dwParam2 As Any) As Long
Public min As Boolean
Public mooveplst As Boolean
Public moovemain As Boolean
Public moovevideo As Boolean
Public playspeed As Double
Public allowslide As Boolean

Type file
 path As String
 name As String
End Type
Public track(1 To 220) As file
Public currentplaylist As String
Public defaultplaylist As String

Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Type POINTAPI
        X As Long
        Y As Long
End Type
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public fx, fy As Integer
Public checkback As Boolean
Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public showtime As Integer
Public ll As Boolean
Public scrolltop As Long
Public scrollunit As Double
Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long

Public updown As Boolean
Public resize As Boolean

Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
    (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, _
    ByVal nCmdShow As Long) As Long

Public Const SW_SHOWNORMAL = 1

'public Sub Form_Load()
    
'End Sub


Public Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long

Public initialpx As Long

Public initialpy As Long
Public initialfx As Long

Public initialfy As Long

    
Public initialypos As Long
Public attach2 As Boolean
Public attach3 As Boolean

Public initialxpos As Long
Public attach1 As Boolean
Public sizescroll As Long
Public initx As Long
Public inity As Long
Public ratio As Boolean

'Public mini As Boolean
Public rectime As Long
Public ckk As Boolean
Public cmo As Boolean

Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long

Public Sub start_record()
'musicsystem.rec_timer.Enabled = True

musicsystem.reclabel.Caption = "REC"
musicsystem.rec_count.Visible = True
musicsystem.reclabel.Visible = True
musicsystem.reclabel.ForeColor = musicsystem.rec_count.ForeColor
End Sub

Public Sub stop_record()
rectime = 0
musicsystem.rec_timer.Enabled = False
musicsystem.reclabel.Caption = "FILE"
musicsystem.reclabel.ForeColor = musicsystem.trackpos.ForeColor
' MMControl2.Command = "stop"
   CommonDialog2.ShowSave
   MMControl2.FileName = CommonDialog2.FileName
MMControl2.Command = "save"
  MMControl2.Command = "close"
'musicsystem.rec_count.Visible = False
'musicsystem.reclabel.Visible = False
End Sub



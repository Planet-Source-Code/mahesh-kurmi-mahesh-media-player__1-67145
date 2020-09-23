Attribute VB_Name = "modDragDrop"
Option Explicit
Public mhHookWindow1 As Long
Public mhHookWindow2 As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public height_incr As Integer

Private Declare Function CallWindowProc Lib "user32" _
    Alias "CallWindowProcA" _
        (ByVal lpPrevWndFunc As Long, _
         ByVal hwnd As Long, _
         ByVal Msg As Long, _
         ByVal wParam As Long, _
         ByVal lParam As Long) As Long

Public Declare Function SetWindowLong Lib "user32" _
    Alias "SetWindowLongA" _
        (ByVal hwnd As Long, _
         ByVal nIndex As Long, _
         ByVal dwNewLong As Long) As Long

Private mlPrevWndProc       As Long
Private mhHookWindow        As Long
Private moDropForm          As Object

Public Declare Sub DragAcceptFiles Lib "shell32.dll" _
        (ByVal hwnd As Long, _
         ByVal fAccept As Long)
         
Private Declare Sub DragFinish Lib "shell32.dll" _
        (ByVal hDrop As Long)
        
Private Declare Function DragQueryFile Lib "shell32.dll" _
    Alias "DragQueryFileA" _
        (ByVal hDrop As Long, _
         ByVal UINT As Long, _
         ByVal lpStr As String, _
         ByVal ch As Long) As Long

Private Const GWL_WNDPROC = -4
Private Const WM_DROPFILES = &H233

Public Sub EnableFileDrops(oDropTarget As Form)

' Be sure we are not arleady subclassing drops
'If mhHookWindow <> 0 Then Call DisableFileDrops

' Save the handle and object reference
' of the calling window
Set moDropForm = oDropTarget
mhHookWindow = moDropForm.hwnd

' Set the subclassing window message hook
mlPrevWndProc = SetWindowLong(mhHookWindow, _
                GWL_WNDPROC, _
                AddressOf HookCallback)

' Tell the OS that the specified window accepts
' dropped files
Call DragAcceptFiles(mhHookWindow, True)

End Sub

Public Sub DisableFileDrops()

Dim lReturn         As Long

' Check to be sure that there is a hook active
If mhHookWindow = 0 Then Exit Sub
If IsEmpty(mhHookWindow) = True Then Exit Sub
If IsNull(mhHookWindow) = True Then Exit Sub

' Tell the OS that the specified window no longer
' accepts dropped files
Call DragAcceptFiles(mhHookWindow, False)

' Remove our hook from the system
lReturn = SetWindowLong(mhHookWindow, _
                GWL_WNDPROC, _
                mlPrevWndProc)

' Clear the window handles and references
mhHookWindow = 0
mlPrevWndProc = 0
Set moDropForm = Nothing

End Sub

Function HookCallback(ByVal hwnd As Long, _
            ByVal lMsg As Long, _
            ByVal wParam As Long, _
            ByVal lParam As Long) As Long
            
Select Case hwnd
    Case mhHookWindow
        ' The message is for the window we subclassed so ...
        ' See if this is a file drop message
        If lMsg = WM_DROPFILES Then
        Set moDropForm = musicsystem

            ' Get the list of dropped files from the OS
            Call GetDropFileList(wParam, lParam)
        End If
     Case plst.hwnd
           If lMsg = WM_DROPFILES Then
            ' Get the list of dropped files from the OS
            Set moDropForm = plst
            Call GetDropFileList(wParam, lParam)
        End If
      
        
        
    Case Else
        ' The message is for some other window
        
End Select

' Pass the message through to the next message processor
HookCallback = CallWindowProc(mlPrevWndProc, _
                              hwnd, _
                              lMsg, _
                              wParam, _
                              lParam)

End Function

Public Sub GetDropFileList(wParam As Long, _
                       lParam As Long)
                
Dim nDropCount          As Integer
Dim nLoopCtr            As Integer
Dim lReturn             As Long
Dim hDrop               As Long
Dim sFileName           As String
Dim vFileNames          As Variant

' Save the drop structure handle
hDrop = wParam

' Allocate space for the return value
sFileName = Space$(255)

' Get the number of file names dropped
nDropCount = DragQueryFile(hDrop, -1, sFileName, 254)

' Allocate variant array elements to store the
' dropped file names
vFileNames = Array(" ")
ReDim vFileNames(nDropCount - 1) As String

' Loop to get each dropped file name and
' add it to the variant array
For nLoopCtr = 0 To nDropCount - 1
    ' Allocate space for the return value
    sFileName = Space$(255)
    ' Get a dropped file name
    lReturn = DragQueryFile(hDrop, nLoopCtr, sFileName, 254)
    vFileNames(nLoopCtr) = Left$(sFileName, lReturn)
Next nLoopCtr

' Release the drop structure from memory
Call DragFinish(hDrop)

' Call the form method to pass the list of dropped files
Call moDropForm.DroppedFiles(vFileNames)

End Sub


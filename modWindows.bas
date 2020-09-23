Attribute VB_Name = "modWindows"
Option Explicit

Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Const SPI_SCREENSAVERRUNNING = 97
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_COMMAND = &H111
Private Const MIN_ALL = 419
Private Const MIN_ALL_UNDO = 416
Private Const LB_SETHORIZONTALEXTENT = &H194
Private Declare Sub ClipCursor Lib "user32" (lpRect As Any)
Private Declare Sub GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT)
Private Declare Sub ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINT)
Private Declare Sub OffsetRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal y As Long)
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type POINT
    X As Long
    y As Long
End Type

'Other APIs
Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal y As Long) As Long
Private Declare Function BlockInput Lib "user32" (ByVal fBlock As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' Minimize all the windows on the desktop (and optionally restore them)
' This has the same effect as pressing the Windows+M key combination

Private Sub MinWindows(Optional Restore As Boolean)
    Dim hWnd As Long
    ' get the handle of the taskbar
    hWnd = FindWindow("Shell_TrayWnd", vbNullString)
    ' Minimize or restore all windows
    If Restore Then
        SendMessage hWnd, WM_COMMAND, MIN_ALL_UNDO, ByVal 0&
    Else
        SendMessage hWnd, WM_COMMAND, MIN_ALL, ByVal 0&
    End If
End Sub


Public Function HideWindows()
    'System Bar
    Dim Handle As Long
    Handle& = FindWindow("Shell_TrayWnd", vbNullString)
    ShowWindow Handle&, 0
    'Ctrl-Alt-Delete
    SystemParametersInfo SPI_SCREENSAVERRUNNING, True, vbNullString, 0
    'Icons
    Dim hWnd As Long
    hWnd = FindWindowEx(0&, 0&, "Progman", vbNullString)
    ShowWindow hWnd, 0
    'Minimize All Windows
    MinWindows False
End Function

Public Function ShowWindows()
    'System Bar
    Dim Handle As Long
    Handle& = FindWindow("Shell_TrayWnd", vbNullString)
    ShowWindow Handle&, 1
    'Ctrl-Alt-Delete
    SystemParametersInfo SPI_SCREENSAVERRUNNING, False, vbNullString, 0
    'Icons
    Dim hWnd As Long
    hWnd = FindWindowEx(0&, 0&, "Progman", vbNullString)
    ShowWindow hWnd, 5
    'Normalize All Windows
    MinWindows True
End Function

Private Function DisButtons()
    SystemParametersInfo SPI_SCREENSAVERRUNNING, True, vbNullString, 0
End Function

Public Property Let IsWinCursorVisible(Visible As Boolean)
On Error Resume Next
    If Visible Then
        Do
            If ShowCursor(1) >= 0 Then Exit Do
        Loop
    Else
        Do
            If ShowCursor(0) < 0 Then Exit Do
        Loop
    End If
End Property


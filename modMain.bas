Attribute VB_Name = "modMain"
Option Explicit

' a picture to draw on before showing it on screen
Private BackBuffer As New cPicture24

' the number of stars used
Private NUM_STARS As Long
Private Const DEFAULT_STARS = 10000
Private Const MIN_STARS = 100
Private Const MAX_STARS = 100000

' the info needed for a star
Private Type STARINFO
    X As Integer
    y As Integer
End Type ' only 4 bytes per star
' each star also has a z position, but it is implicit because it comes from it's position in the buffer
Private Const LEN_STAR = 4

' an array of stars
Private Star() As STARINFO

' the index of the newest star
Private newStar As Long

' the spread of stars
Private Const STAR_WIDTH = 32676
Private Const STAR_HEIGHT = 32676
Private Const STAR_DEPTH = 200

' the rate the stars are recycled (also speed)
Private STAR_SPEED As Long
Private Const DEFAULT_SPEED = 100
Private Const MIN_SPEED = 1
Private Const MAX_SPEED = 10000

' monitor settings
Private Const SCREEN_WIDTH = 800
Private Const SCREEN_HEIGHT = 600
Private Const SCREEN_BPP = 24

' for memory copying/moving/zeroing/filling
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlCopyMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)
Private Declare Sub FillMemory Lib "kernel32.dll" Alias "RtlFillMemory" (dest As Any, ByVal numBytes As Long, Fill As Byte)

' for getting a pointer to a variable
Private Declare Function VarPtr Lib "msvbvm50.dll" (Ptr As Any) As Long
Private Declare Function VarPtrArray Lib "msvbvm50.dll" Alias "VarPtr" (Ptr() As Any) As Long

' for the time in milliseconds
Private Declare Function timeGetTime Lib "winmm.dll" () As Long

' for reading keyboard input
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Const KEY_IS_DOWN = -32767



' the program starts here
Sub Main()
    If Init Then MainLoop
    CleanUp
End Sub



' changes screen res, hides windows, loads buffers
Function Init() As Boolean
Dim strResult As String
Dim l As Long
' x and y values for stars
Dim X As Integer
Dim y As Integer
On Error Resume Next
    Randomize Timer
    ' intro
    MsgBox "This program is a fast 3D starfield demo written in Visual Basic by Simon Price. Win32 API is used, but NO DirectX or OpenGL or other 3rd party DLL's are used whatsoever!", vbInformation, "3D Starfield"
    MsgBox "You must have a video card capable of 800 x 600 resolution and 24 bit color.", vbInformation, "Requirements"
    MsgBox "This program has a high frame rate, press the ""F"" key to show/hide the FPS counter", vbInformation, "Controls"
    MsgBox "You can choose a few variables which affect the starfield, but please leave them at the defaults for the first time, then try experimenting!", vbInformation, "Starfield Variables"
    ' get the number of stars
EnterNumStars:
    NUM_STARS = Val(InputBox("Enter the number of stars in the starfield (" & MIN_STARS & " to " & MAX_STARS & ")", "# of stars", DEFAULT_STARS))
    Select Case NUM_STARS
        Case MIN_STARS To MAX_STARS
        Case Else
            MsgBox "Invalid #"
            GoTo EnterNumStars
    End Select
    ReDim Star(NUM_STARS)
    ' get the speed of stars
EnterStarSpeed:
    STAR_SPEED = Val(InputBox("Enter the speed of flight (" & MIN_SPEED & " to " & MAX_SPEED & ")", "speed", DEFAULT_SPEED))
    Select Case STAR_SPEED
        Case MIN_SPEED To MAX_SPEED
        Case Else
            MsgBox "Invalid #"
            GoTo EnterStarSpeed
    End Select
    ' change screen res
    If MsgBox("Changing screen resolution to " & SCREEN_WIDTH & " x " & SCREEN_HEIGHT & "  x " & SCREEN_BPP & ", continue?", vbQuestion Or vbYesNo) = vbNo Then Exit Function
    strResult = modScreen.SetDisplayMode(SCREEN_WIDTH, SCREEN_HEIGHT, SCREEN_BPP)
    If strResult = "OK" Then
        MsgBox "Screen resolution changed successfully"
    Else
        modScreen.RestoreDisplayMode
        If MsgBox("Error - could not change screen resolution! Continue anyway?", vbYesNo Or vbExclamation) = vbNo Then Exit Function
    End If
    ' hide all windows
    modWindows.HideWindows
    ' hide cursor
    modWindows.IsWinCursorVisible = False
    For l = 1 To 10
        DoEvents ' the windows are a bit slow at hiding, this gives them time
    Next
    ' init backbuffer
    If Not BackBuffer.Init(SCREEN_WIDTH, SCREEN_HEIGHT) Then Exit Function
    ' create some random stars
    newStar = NUM_STARS
    For l = 1 To NUM_STARS ' fill every star
        ' random x value
        X = Rnd * STAR_WIDTH * 2 - STAR_WIDTH
        Star(l).X = X
        ' random y value
        y = Rnd * STAR_HEIGHT * 2 - STAR_HEIGHT
        Star(l).y = y
    Next
    ' init was successful
    Init = True
End Function



' the main program loop
Sub MainLoop()
' pointer to the current pixel
Dim pPixel As Long
' untransformed z coord of star
Dim z As Byte
' transformed x, y coords of star
Dim X2 As Long
Dim Y2 As Long
' for looping through the stars
Dim iStar As Long
' for making stars gradually closer
Dim fz As Single
Dim DZ As Single
DZ = STAR_DEPTH / NUM_STARS
' used for 3d to 2d transformation
Const ZOOM = 0.9
Const HALF_WIDTH = SCREEN_WIDTH / 2
Const HALF_HEIGHT = SCREEN_HEIGHT / 2
' a temp value to store a result, saving us a divide during transformation
Dim LensDivZ As Single
' times to calculate the fps
Dim timeLast As Long
Dim timeNow As Long
Dim timeElapsed As Long
' the fps and whether the fps is displayed
Dim FPS As Single
Dim ShowFPS As Boolean
On Error Resume Next
    timeLast = timeGetTime
    Do
        DoEvents ' allow events to be processed
        BackBuffer.ClearToBlack ' clear picture
        ' render stars
        fz = STAR_DEPTH ' start from back (painters algo)
        'pStar = StarBuffer.pBits ' reset pointer to stars
        iStar = newStar
        Do ' loop through each star
            iStar = iStar + 1
            If iStar > NUM_STARS Then iStar = 1
            ' transform 3d coord to 2d (this is actually upside down, but in a starfield who would notice?)
            LensDivZ = ZOOM / fz
            X2 = Star(iStar).X * LensDivZ + HALF_WIDTH
            Y2 = Star(iStar).y * LensDivZ + HALF_HEIGHT
            ' makes sure the 2d coords are in bounds so we dont write to invalid memory
            If X2 < 0 Then GoTo Skip
            If Y2 < 0 Then GoTo Skip
            If X2 > SCREEN_WIDTH Then GoTo Skip
            If Y2 > SCREEN_HEIGHT Then GoTo Skip
            'X2 = X2 And SCREEN_WIDTH ' can anyone get these 2 lines working?
            'Y2 = Y2 And SCREEN_HEIGHT
            ' calculate star brightness and size
            z = CByte(255 - fz)
            ' draw the star
            pPixel = BackBuffer.pBits + (Y2 * SCREEN_WIDTH + X2) * 3
            FillMemory ByVal pPixel, (z \ 50 + 1) * 3, ByVal z
            ' move foward
            fz = fz - DZ
Skip:
        Loop Until iStar = newStar
        ' animate the stars
        newStar = newStar - STAR_SPEED
        If newStar <= 0 Then newStar = NUM_STARS
        Star(newStar).X = Rnd * STAR_WIDTH * 2 - STAR_WIDTH
        Star(newStar).y = Rnd * STAR_HEIGHT * 2 - STAR_HEIGHT
        ' check for f key input
        If GetAsyncKeyState(vbKeyF) Then ShowFPS = Not ShowFPS
        ' calculate and display fps
        If ShowFPS Then
            timeNow = timeGetTime
            timeElapsed = timeNow - timeLast
            timeLast = timeNow
            FPS = 1000 / timeElapsed
            BackBuffer.PrintText "FPS = " & FPS
        End If
        ' show picture in backbuffer on the screen
        BackBuffer.Show
    Loop Until GetAsyncKeyState(vbKeyEscape) = KEY_IS_DOWN ' end when escape is pressed
End Sub



' deallocate resources
Sub CleanUp()
On Error Resume Next
    ' delete backbuffer
    Set BackBuffer = Nothing
    ' reset screen res
    modScreen.RestoreDisplayMode
    ' restore windows
    modWindows.ShowWindows
    ' show cursor
    modWindows.IsWinCursorVisible = True
    MsgBox "Thankyou, if you liked this starfield, please visit my site www.VBgames.co.uk and vote for my source code on www.planet-source-code.com", vbInformation, "Thankyou!"
End Sub

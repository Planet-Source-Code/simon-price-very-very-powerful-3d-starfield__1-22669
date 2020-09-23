VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPicture24"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public hDC As Long
Public hBitmap As Long
Public pBits As Long
Public lSize As Long
Public lWidth As Long
Public lHeight As Long

' types
Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type BITMAPFILEHEADER
   bfType As Integer
   bfSize As Long
   bfReserved1 As Integer
   bfReserved2 As Integer
   bfOffBits As Long
End Type
Private Type BITMAPINFOHEADER
   biSize As Long
   biWidth As Long
   biHeight As Long
   biPlanes As Integer
   biBitCount As Integer
   biCompression As Long
   biSizeImage As Long
   biXPelsPerMeter As Long
   biYPelsPerMeter As Long
   biClrUsed As Long
   biClrImportant As Long
End Type
Private Type BITMAPINFO_24
   bmiHeader As BITMAPINFOHEADER
End Type
Private Type RGBQUAD
        rgbBlue As Byte
        rgbGreen As Byte
        rgbRed As Byte
        rgbReserved As Byte
End Type
Private Type BITMAPINFO
        bmiHeader As BITMAPINFOHEADER
        bmiColors As RGBQUAD
End Type

' constants
Private Const DIB_RGB_COLORS = 0

' win32 api function declarations
Private Declare Function VarPtr Lib "msvbvm50.dll" (Ptr As Any) As Long
Private Declare Function VarPtrArray Lib "msvbvm50.dll" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlCopyMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)
Private Declare Sub FillMemory Lib "kernel32.dll" Alias "RtlFillMemory" (dest As Any, ByVal numBytes As Long, Fill As Byte)
Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SetDIBits_24 Lib "gdi32" Alias "SetDIBits" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpbi As BITMAPINFO_24, ByVal wUsage As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function GetBkColor Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hDC As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, ByRef lplpVoid As Long, ByVal Handle As Long, ByVal dw As Long) As Long
Private Declare Function SaveDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function RestoreDC Lib "gdi32" (ByVal hDC As Long, ByVal nSavedDC As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function SetTextAlign Lib "gdi32" (ByVal hDC As Long, ByVal wFlags As Long) As Long
Private Declare Function GetTextAlign Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function GetTextColor Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDC As Long, ByVal X As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long



' init to a blank bitmap of specified size
Function Init(Optional Width As Long, Optional Height As Long) As Boolean
On Error Resume Next
Dim Info As BITMAPINFO
Dim DisplayDC As Long
    CleanUp
    ' fill bitmap info structure
    With Info.bmiHeader
        .biBitCount = 24
        .biWidth = Width
        .biHeight = Height
        .biPlanes = 1
        .biSize = Len(Info.bmiHeader)
    End With
    ' get a handle to the display
    DisplayDC = GetDC(0)
    ' create a dc compatible with the display
    hDC = CreateCompatibleDC(DisplayDC)
    SaveDC hDC
    If hDC = 0 Then GoTo InitFailed ' check we received a valid handle
    ' create a device independant bitmap
    hBitmap = CreateDIBSection(hDC, Info, DIB_RGB_COLORS, pBits, 0, 0)
    If hBitmap = 0 Then GoTo InitFailed ' check we received a valid handle
    ' use the bitmap for the dc
    SelectObject hDC, hBitmap
    ' clean up the display dc
    ReleaseDC 0, DisplayDC
    ' compute size
    lWidth = Width
    lHeight = Height
    lSize = lWidth * lHeight * 3
    Init = True
    Exit Function
InitFailed:
End Function



' init from a bitmap file
Function InitFromFile(FileName As String) As Boolean
Dim Info As BITMAPINFO
Dim f As Integer
Dim DisplayDC As Long
Dim FileHead As BITMAPFILEHEADER
Dim bmpInfo_24 As BITMAPINFO_24
Dim bmpInfoHeader As BITMAPINFOHEADER
Dim picBytes() As Byte
Dim picOffset As Long, picLen As Long
On Error Resume Next
    ' open the file and read the header info
    picOffset = 1
    picLen = FileLen(FileName)
    f = FreeFile
    Open FileName For Binary Access Read As #f
    Get #f, picOffset, FileHead
    Get #f, picOffset + 14, bmpInfoHeader
    ' check the bpp (bits per pixel)
    Select Case bmpInfoHeader.biBitCount
        Case 24
            Get #f, picOffset + 14, bmpInfo_24
        Case Else
            GoTo InitFailed ' other formats unsupported
    End Select
    ' get each byte of the picture
    ReDim picBytes(0 To picLen - 44)
    Get #f, , picBytes
    ' close file
    Close #f
    ' get a handle to the display
    DisplayDC = GetDC(0)
    ' create a new dc that is compatible with the display
    hDC = CreateCompatibleDC(DisplayDC)
    ' create a dc that is compatible with the display, and the size of the bitmap
'    hBitmap = CreateCompatibleBitmap(DisplayDC, bmpInfoHeader.biWidth, bmpInfoHeader.biHeight)
    ' fill bitmap info structure
    With Info.bmiHeader
        .biBitCount = bmpInfoHeader.biBitCount
        .biWidth = bmpInfoHeader.biWidth
        .biHeight = bmpInfoHeader.biHeight
        .biPlanes = bmpInfoHeader.biPlanes
        .biSize = Len(Info.bmiHeader)
    End With
    ' create a device independant bitmap
    hBitmap = CreateDIBSection(hDC, Info, DIB_RGB_COLORS, pBits, 0, 0)
    '
    SelectObject hDC, hBitmap
    ' set the bits of the bitmap to the temporary array
    SetDIBits_24 hDC, hBitmap, 0, bmpInfoHeader.biHeight, picBytes(0), bmpInfo_24, 0
    ' clean up the display dc
    ReleaseDC 0, DisplayDC
    ' compute size
    lWidth = bmpInfoHeader.biWidth
    lHeight = bmpInfoHeader.biHeight
    lSize = lWidth * lHeight * 3
    InitFromFile = True
    Exit Function
InitFailed:
End Function



' sets an individual pixel to the specified red, green and blue values
Function SetPixel(X As Long, y As Long, r As Byte, g As Byte, b As Byte)
On Error Resume Next
Dim p As Long
    p = pBits + (y * lWidth + X) * 3
    CopyMemory ByVal p, ByVal VarPtr(b), 1
    p = p + 1
    CopyMemory ByVal p, ByVal VarPtr(g), 1
    p = p + 1
    CopyMemory ByVal p, ByVal VarPtr(r), 1
End Function
' gets the red, green and blue of an individual pixel
Sub GetPixel(X As Long, y As Long, r As Byte, g As Byte, b As Byte)
On Error Resume Next
Dim p As Long
    p = pBits + (y * lWidth + X) * 3
    CopyMemory ByVal VarPtr(b), ByVal p, 1
    p = p + 1
    CopyMemory ByVal VarPtr(g), ByVal p, 1
    p = p + 1
    CopyMemory ByVal VarPtr(r), ByVal p, 1
End Sub
' gets an individual pixel in long format
Function GetPixelLong(X As Long, y As Long) As Long
On Error Resume Next
Dim p As Long
Dim r As Byte
Dim g As Byte
Dim b As Byte
    p = pBits + (y * lWidth + X) * 3
    CopyMemory ByVal VarPtr(b), ByVal p, 1
    p = p + 1
    CopyMemory ByVal VarPtr(g), ByVal p, 1
    p = p + 1
    CopyMemory ByVal VarPtr(r), ByVal p, 1
    GetPixelLong = RGB(r, g, b)
End Function



' clearing functions
Sub ClearToBlack()
    ZeroMemory ByVal pBits, lSize
End Sub
Sub ClearToGrey(Shade As Byte)
    FillMemory ByVal pBits, lSize, ByVal Shade
End Sub
Sub Clear(r As Byte, g As Byte, b As Byte)
Dim p As Long
Dim p2 As Long
Dim pr As Long
Dim pg As Long
Dim pb As Long
    p = pBits
    p2 = pBits + lSize
    pr = VarPtr(r)
    pg = VarPtr(g)
    pb = VarPtr(b)
    Do
        CopyMemory ByVal p, ByVal pb, 1
        p = p + 1
        CopyMemory ByVal p, ByVal pg, 1
        p = p + 1
        CopyMemory ByVal p, ByVal pr, 1
        p = p + 1
    Loop Until p > p2
End Sub



' prints some text onto the picture
Sub PrintText(str As String, Optional X As Long = 0, Optional y As Long = 0)
    TextOut hDC, X, y, str, Len(str)
End Sub



' background color
Property Get Backcolor() As Long
    Backcolor = GetBkColor(hDC)
End Property
Property Let Backcolor(Color As Long)
    SetBkColor hDC, Color
End Property



' text color
Property Get TextColor() As Long
    TextColor = GetTextColor(hDC)
End Property
Property Let TextColor(Color As Long)
    SetTextColor hDC, Color
End Property



' text alignment
Property Get TextAlign() As Long
    TextAlign = GetTextAlign(hDC)
End Property
Property Let TextAlign(Align As Long)
    SetTextAlign hDC, Align
End Property



' shows the picture on the screen
Sub Show()
    BitBlt GetDC(0), 0, 0, lWidth, lHeight, hDC, 0, 0, vbSrcCopy
End Sub



' deallocate resources
Sub CleanUp()
On Error Resume Next
    ' delete bitmap
    DeleteObject hBitmap
    ' delete dc
    RestoreDC hDC, -1
    DeleteDC hDC
End Sub
Private Sub Class_Terminate()
    CleanUp
End Sub


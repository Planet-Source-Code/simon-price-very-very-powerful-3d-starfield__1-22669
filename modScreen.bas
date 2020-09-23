Attribute VB_Name = "modScreen"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'       modScreen.bas - Provides monitor setting functions
'
'                        By Simon Price
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit


' constants for ChangeDisplaySettings API function
Private Const CCDEVICENAME = 32
Private Const CCFORMNAME = 32
Private Const DISP_CHANGE_SUCCESSFUL = 0
Private Const DISP_CHANGE_RESTART = 1
Private Const DISP_CHANGE_FAILED = -1
Private Const DISP_CHANGE_BADMODE = -2
Private Const DISP_CHANGE_NOTUPDATED = -3
Private Const DISP_CHANGE_BADFLAGS = -4
Private Const DISP_CHANGE_BADPARAM = -5
Private Const CDS_UPDATEREGISTRY = &H1
Private Const CDS_TEST = &H2
Private Const DM_BITSPERPEL = &H40000
Private Const DM_PELSWIDTH = &H80000
Private Const DM_PELSHEIGHT = &H100000


' this type contains info on all devices for Windows
Private Type DEVMODE
  dmDeviceName As String * CCDEVICENAME
  dmSpecVersion As Integer
  dmDriverVersion As Integer
  dmSize As Integer
  dmDriverExtra As Integer
  dmFields As Long
  dmOrientation As Integer
  dmPaperSize As Integer
  dmPaperLength As Integer
  dmPaperWidth As Integer
  dmScale As Integer
  dmCopies As Integer
  dmDefaultSource As Integer
  dmPrintQuality As Integer
  dmColor As Integer
  dmDuplex As Integer
  dmYResolution As Integer
  dmTTOption As Integer
  dmCollate As Integer
  dmFormName As String * CCFORMNAME
  dmUnusedPadding As Integer
  dmBitsPerPel As Integer
  dmPelsWidth As Long
  dmPelsHeight As Long
  dmDisplayFlags As Long
  dmDisplayFrequency As Long
End Type


' API function declarations
Private Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lpDevMode As Any) As Boolean
Private Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lpDevMode As Any, ByVal dwflags As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long


' variables
Private lastWidth As Integer
Private lastHeight As Integer
Private lastColors As Integer



' this function changes the screen settings and returns either
' "OK" to show success or another string stating the error

Public Function SetDisplayMode(ByVal Width As Integer, ByVal Height As Integer, ByVal Colors As Integer) As String
Dim DeviceMode As DEVMODE, lastDeviceMode As DEVMODE, lTemp As Long, lIndex As Long

' remember current screen resolution
lIndex = GetSystemMetrics(0&)
lTemp = EnumDisplaySettings(0&, lIndex, lastDeviceMode)
lastWidth = lastDeviceMode.dmPelsWidth
lastHeight = lastDeviceMode.dmPelsHeight
lastColors = lastDeviceMode.dmBitsPerPel
' try again, I don't know why but 1 in 2 times this routine fails
If lastWidth = 0 Then SetDisplayMode Width, Height, Colors

lIndex = 0
' loop through all settings
Do
    lTemp = EnumDisplaySettings(0&, lIndex, DeviceMode)
    If lTemp = 0 Then Exit Do
    lIndex = lIndex + 1
    With DeviceMode
        ' check if the current setting is the one we want
        If .dmPelsWidth = Width And .dmPelsHeight = Height And .dmBitsPerPel = Colors Then
            ' yes, we can change the screen settings and stop looking
            lTemp = ChangeDisplaySettings(DeviceMode, CDS_UPDATEREGISTRY)
            Exit Do
        End If
    End With
Loop
' check for errors
Select Case lTemp
    Case DISP_CHANGE_SUCCESSFUL
        ' report OK message and remember previous display settings
        SetDisplayMode = "OK"
    Case DISP_CHANGE_RESTART
        SetDisplayMode = "The computer must be restarted in order for the graphics mode to work"
    Case DISP_CHANGE_FAILED
        SetDisplayMode = "The display driver failed the specified graphics mode"
    Case DISP_CHANGE_BADMODE
        SetDisplayMode = "The graphics mode is not supported"
    Case DISP_CHANGE_NOTUPDATED
        SetDisplayMode = "Unable to write settings to the registry"
    Case DISP_CHANGE_BADFLAGS
        SetDisplayMode = "An invalid set of flags was passed in"
End Select
End Function



' this function restores the screen to it's previous settings

Public Function RestoreDisplayMode() As String
    RestoreDisplayMode = SetDisplayMode(lastWidth, lastHeight, lastColors)
End Function


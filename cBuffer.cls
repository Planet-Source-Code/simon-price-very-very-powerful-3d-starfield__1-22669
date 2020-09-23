VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function LocalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal wBytes As Long) As Long
Private Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlCopyMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)
Private Declare Sub FillMemory Lib "kernel32.dll" Alias "RtlFillMemory" (dest As Any, ByVal numBytes As Long, Fill As Byte)
Private Declare Function VarPtr Lib "msvbvm50.dll" (Ptr As Any) As Long
Private Declare Function VarPtrArray Lib "msvbvm50.dll" Alias "VarPtr" (Ptr() As Any) As Long

Public pBits As Long
Public lSize As Long



' reserves a fixed memory block
Function Init(Size As Long) As Boolean
On Error Resume Next
    pBits = LocalAlloc(0, Size)
    Init = CBool(pBits)
    If Init Then lSize = Size
End Function



' resets to zero
Sub Clear()
    ZeroMemory ByVal pBits, ByVal lSize
End Sub



' fills the memory with a specific byte value
Sub Fill(FillByte As Byte)
    FillMemory pBits, lSize, FillByte
End Sub



' copies memory from one place to another
Sub Copy(pDst As Long, pSrc As Long, ByteLen As Long)
    CopyMemory ByVal pDst, ByVal pSrc, ByVal ByteLen
End Sub



' same as copying, except this works even when the 2 blocks of memory overlap
Sub Move(pDst As Long, pSrc As Long, ByteLen As Long)
    MoveMemory ByVal pDst, ByVal pSrc, ByVal ByteLen
End Sub



' gets a variable from the buffer
Sub GetVar(pGet As Long, myVar)
    CopyMemory ByVal VarPtr(myVar), ByVal pGet, ByVal Len(myVar)
End Sub



' free memory
Private Sub Class_Terminate()
On Error Resume Next
    LocalFree pBits
End Sub

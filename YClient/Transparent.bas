Attribute VB_Name = "Transparent_Mod"
Option Explicit
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWnsenddInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Declare Function UpdateLayeredWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal hdcDst As Long, pptDst As Any, psize As Any, ByVal hdcSrc As Long, pptSrc As Any, crKey As Long, ByVal pblend As Long, ByVal dwFlags As Long) As Long
Public Type POINTAPI
    X As Long
    Y As Long
    End Type
Public Type SIZE
    cx As Long
    cy As Long
    End Type
Public Type BLENDFUNCTION
    BlendOp As Byte
    BlendFlags As Byte
    SourceConstantAlpha As Byte
    AlphaFormat As Byte
    End Type
    Public Const GWL_STYLE = (-16)
    Public Const GWL_EXSTYLE = (-20)
    Public Const WS_EX_LAYERED = &H80000
    Public Const ULW_COLORKEY = &H1
    Public Const ULW_ALPHA = &H2
    Public Const ULW_OPAQUE = &H4
    Public Const AC_SRC_OVER = &H0
    Public Const AC_SRC_ALPHA = &H1
    Public Const AC_SRC_NO_PREMULT_ALPHA = &H1
    Public Const AC_SRC_NO_ALPHA = &H2
    Public Const AC_DST_NO_PREMULT_ALPHA = &H10
    Public Const AC_DST_NO_ALPHA = &H20
    Public Const LWA_COLORKEY = &H1
    Public Const LWA_ALPHA = &H2
    Public Const HWND_TOPMOST = -1
    Public Const SWP_NOMOVE = &H2
    Public Const SWP_NOSIZE = &H1
    Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Public Sub StayOnTop(frm As Form)
Call SetWindowPos(frm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
End Sub
Public Sub Trans(hwnd As Long, Rate As Byte)
    On Error Resume Next
    Dim WinInfo As Long
    WinInfo = GetWindowLong(hwnd, GWL_EXSTYLE)
    WinInfo = WinInfo Or WS_EX_LAYERED
    SetWindowLong hwnd, GWL_EXSTYLE, WinInfo
    SetLayeredWindowAttributes hwnd, 0, Rate, LWA_ALPHA
End Sub


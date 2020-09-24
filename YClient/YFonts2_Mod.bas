Attribute VB_Name = "YFonts2_Mod"
Option Explicit
Public Const LF_FACESIZE = 32
Public Const WM_USER = &H400
Public Const EM_SETCHARFORMAT = (WM_USER + 68)
Public Const CFM_BACKCOLOR = &H4000000
Public Const SCF_SELECTION = &H1

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Type CHARFORMAT2
    cbSize As Integer    '2
    wPad1 As Integer    '4
    dwMask As Long    '8
    dwEffects As Long    '12
    yHeight As Long    '16
    yOffset As Long    '20
    crTextColor As Long    '24
    bCharSet As Byte    '25
    bPitchAndFamily As Byte    '26
    szFaceName(0 To LF_FACESIZE - 1) As Byte    ' 58
    wPad2 As Integer    ' 60

    wWeight As Integer
    sSpacing As Integer
    crBackColor As Long
    lLCID As Long
    dwReserved As Long
    sStyle As Integer
    wKerning As Integer
    bUnderlineType As Byte
    bAnimation As Byte
    bRevAuthor As Byte
    bReserved1 As Byte
End Type



Private Type Size
    cx As Long
    cy As Long
End Type

Private Type POINTAPI
    x As Long
    Y As Long
End Type

Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type


Private Declare Function CreateMetaFile Lib "gdi32" Alias "CreateMetaFileA" (ByVal lpString As String) As Long
Private Declare Function CloseMetaFile Lib "gdi32" (ByVal hDCMF As Long) As Long
Private Declare Function DeleteMetaFile Lib "gdi32" (ByVal hMF As Long) As Long

Private Declare Function SetMapMode Lib "gdi32" (ByVal hdc As Long, ByVal nMapMode As Long) As Long
Private Declare Function SetWindowExtEx Lib "gdi32" (ByVal hdc As Long, ByVal nX As Long, ByVal nY As Long, lpSize As Size) As Long
Private Declare Function SetWindowOrgEx Lib "gdi32" (ByVal hdc As Long, ByVal nX As Long, ByVal nY As Long, lpPoint As POINTAPI) As Long
Private Declare Function SaveDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function RestoreDC Lib "gdi32" (ByVal hdc As Long, ByVal nSavedDC As Long) As Long

Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long


Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Private Const MM_ANISOTROPIC = 8 ' Map mode anisotropic

Function getTempName(Optional anExt As String = "tmp") As String
    Dim tempPath    As String
    Dim fileName    As String
    Dim I           As Long
    
    Const validChars As String = "123567890qwertyuiopasdfghjklzxcvbnm"
    
    tempPath = String$(255, " ")
    GetTempPath 255, tempPath
    tempPath = Left$(tempPath, InStr(tempPath, Chr$(0)) - 1)
    fileName = Space(12)
    Mid$(fileName, 1, 1) = "T"
    Mid$(fileName, Len(fileName) - Len(anExt), 1) = "."
    Mid$(fileName, Len(fileName) - Len(anExt) + 1, Len(anExt)) = anExt
    Randomize
    For I = 2 To Len(fileName) - 4
        Mid$(fileName, I, 1) = Mid$(validChars, CLng(Rnd() * (Len(validChars)) + 1), 1)
    Next I
    tempPath = tempPath & fileName
    getTempName = tempPath
    
End Function
Function Pic(aStdPic As StdPicture) As String
    Dim hMetaDC     As Long
    Dim hMeta       As Long
    Dim hPicDC      As Long
    Dim hOldBmp     As Long
    Dim aBMP        As BITMAP
    Dim aSize       As Size
    Dim aPt         As POINTAPI
    Dim fileName    As String
    Dim screenDC    As Long
    Dim headerStr   As String
    Dim retStr      As String
    Dim byteStr     As String
    Dim bytes()     As Byte
    Dim filenum     As Integer
    Dim numBytes    As Long
    Dim I           As Long
    

    fileName = getTempName("WMF")
    hMetaDC = CreateMetaFile(fileName)
    SetMapMode hMetaDC, MM_ANISOTROPIC
    SetWindowOrgEx hMetaDC, 0, 0, aPt
    GetObject aStdPic.Handle, Len(aBMP), aBMP
    SetWindowExtEx hMetaDC, aBMP.bmWidth, aBMP.bmHeight, aSize
    SaveDC hMetaDC
    screenDC = GetDC(0)
    hPicDC = CreateCompatibleDC(screenDC)
    ReleaseDC 0, screenDC
    hOldBmp = SelectObject(hPicDC, aStdPic.Handle)
    BitBlt hMetaDC, 0, 0, aBMP.bmWidth, aBMP.bmHeight, hPicDC, 0, 0, vbSrcCopy
    SelectObject hPicDC, hOldBmp
    DeleteDC hPicDC
    DeleteObject hOldBmp
    RestoreDC hMetaDC, True
    hMeta = CloseMetaFile(hMetaDC)
    DeleteMetaFile hMeta
    headerStr = "{\rtf1\ansi"
    headerStr = headerStr & _
                "{\pict\picscalex100\picscaley100" & _
                "\picw" & aStdPic.Width & "\pich" & aStdPic.Height & _
                "\picwgoal" & aBMP.bmWidth * Screen.TwipsPerPixelX & _
                "\pichgoal" & aBMP.bmHeight * Screen.TwipsPerPixelY & _
                "\wmetafile8"
    
    numBytes = FileLen(fileName)
    ReDim bytes(1 To numBytes)
    filenum = FreeFile()
    Open fileName For Binary Access Read As #filenum
    Get #filenum, , bytes
    Close #filenum
    byteStr = String(numBytes * 2, "0")
    For I = LBound(bytes) To UBound(bytes)
        If bytes(I) > &HF Then
            Mid$(byteStr, 1 + (I - 1) * 2, 2) = Hex$(bytes(I))
        Else
            Mid$(byteStr, 2 + (I - 1) * 2, 1) = Hex$(bytes(I))
        End If
    Next I
    retStr = headerStr & " " & byteStr & "}"
    retStr = retStr & "}"
        
    Pic = retStr
    On Local Error Resume Next
    If Dir(fileName) <> "" Then Kill fileName
End Function

Sub ColourBG(Col As ColorConstants, Rch As RichTextBox)
Dim BGCol As CHARFORMAT2
 With BGCol
    .cbSize = Len(BGCol)
    .dwMask = CFM_BACKCOLOR
    .crBackColor = Col
 End With
SendMessage Rch.hwnd, EM_SETCHARFORMAT, SCF_SELECTION, BGCol
End Sub


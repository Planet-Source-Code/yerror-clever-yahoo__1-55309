Attribute VB_Name = "YFonts_Mod"
Sub ProcessDing(Rch As RichTextBox, Move As Form)
Rch.SelFontSize = 10
Rch.SelFontName = "Arial"
Rch.SelColor = vbRed
Rch.SelUnderline = False
Rch.SelBold = True
Move.Left = Move.Left + 200
Move.Top = Move.Top + 200
Move.Left = Move.Left - 200
Move.Top = Move.Top - 200
Rch.SelItalic = False
If frmConfig.Check12.Value = 1 Then
PlayWav frmConfig.Text3
End If
Rch.SelText = "BUZZ!!!" & vbCrLf
Rch.SelBold = False
Rch.SelColor = vbBlack
End Sub
Sub ProcessText(ByVal TXT As String, Rch As RichTextBox)
On Error Resume Next
Dim N As Integer, T As String
Rch.SelStart = Len(Rch.Text)
Dim I As Integer
For I = 1 To Len(TXT)
    N = 2
    If LCase(Left(TXT, 3)) = "<b>" Then
        Rch.SelBold = True
        N = 4
    ElseIf LCase(Left(TXT, 5)) = "[38m" Then
        Rch.SelColor = vbRed
        N = 6
    ElseIf LCase(Left(TXT, 5)) = "[30m" Then
        Rch.SelColor = vbBlack
        N = 6
    ElseIf LCase(Left(TXT, 5)) = "[34m" Then
        Rch.SelColor = vbGreen
        N = 6
    ElseIf LCase(Left(TXT, 5)) = "[39m" Then
        Rch.SelColor = &H8080&
        N = 6
    ElseIf LCase(Left(TXT, 5)) = "[31m" Then
        Rch.SelColor = vbBlue
        N = 6
    ElseIf LCase(Left(TXT, 5)) = "[36m" Then
        Rch.SelColor = &H800080
        N = 6
    ElseIf LCase(Left(TXT, 5)) = "[32m" Then
        Rch.SelColor = &H404000
        N = 6
    ElseIf LCase(Left(TXT, 5)) = "[35m" Then
        Rch.SelColor = &HFF80FF
        N = 6
    ElseIf LCase(Left(TXT, 5)) = "[33m" Then
        Rch.SelColor = &H8000000C
        N = 6
    ElseIf LCase(Left(TXT, 5)) = "[37m" Then
        Rch.SelColor = &H80FF&
        N = 6
    ElseIf LCase(Left(TXT, 3)) = "[#" Then
        Rch.SelColor = Hex2RGB(Mid(TXT, 4, 6))
        N = 11
    ElseIf LCase(Left(TXT, 2)) = "<#" Then
        Rch.SelColor = Hex2RGB(Mid(TXT, 3, 6))
        N = InStr(TXT, ">") + 1
    ElseIf LCase(Left(TXT, 4)) = "[1m" Then
        Rch.SelBold = True
        N = 5
    ElseIf LCase(Left(TXT, 4)) = "[2m" Then
        Rch.SelItalic = True
        N = 5
    ElseIf LCase(Left(TXT, 4)) = "[4m" Then
        Rch.SelUnderline = True
        N = 5
    ElseIf LCase(Left(TXT, 4)) = "<alt" Then
        N = InStr(TXT, ">") + 1
    ElseIf LCase(Left(TXT, 5)) = "<fade" Then
        N = InStr(TXT, ">") + 1
    ElseIf LCase(Left(TXT, 7)) = "</fade>" Then
        N = 8
    ElseIf LCase(Left(TXT, 6)) = "</alt>" Then
        N = 7
    ElseIf LCase(Left(TXT, 3)) = "<i>" Then
        Rch.SelItalic = True
        N = 4
    ElseIf LCase(Left(TXT, 3)) = "<u>" Then
        Rch.SelUnderline = True
        N = 4
    ElseIf LCase(Left(TXT, 5)) = "<red>" Then
        Rch.SelColor = vbRed
        N = 6
    ElseIf LCase(Left(TXT, 7)) = "<green>" Then
        Rch.SelColor = vbGreen
        N = 8
    ElseIf LCase(Left(TXT, 6)) = "<blue>" Then
        Rch.SelColor = vbBlue
        N = 7
    ElseIf LCase(Left(TXT, 7)) = "<black>" Then
        Rch.SelColor = vbBlack
        N = 8
    ElseIf LCase(Left(TXT, 8)) = "<yellow>" Then
        Rch.SelColor = vbYellow
        N = 9
    ElseIf LCase(Left(TXT, 6)) = "<gray>" Then
        Rch.SelColor = &H808080
        N = 7
    ElseIf LCase(Left(TXT, 7)) = "</gray>" Then
        Rch.SelColor = vbBlack
        N = 8
    ElseIf LCase(Left(TXT, 9)) = "</yellow>" Then
        Rch.SelColor = vbYellow
        N = 10
    ElseIf LCase(Left(TXT, 7)) = "</blue>" Then
        Rch.SelColor = vbBlack
        N = 8
    ElseIf LCase(Left(TXT, 8)) = "</green>" Then
        Rch.SelColor = vbBlack
        N = 9
    ElseIf LCase(Left(TXT, 6)) = "</red>" Then
        Rch.SelColor = vbRed
        N = 7
    ElseIf LCase(Left(TXT, 4)) = "</b>" Then
        Rch.SelBold = False
        N = 5
    ElseIf LCase(Left(TXT, 4)) = "</i>" Then
        Rch.SelItalic = False
        N = 5
    ElseIf LCase(Left(TXT, 4)) = "</u>" Then
        Rch.SelUnderline = False
        N = 5
    ElseIf LCase(Left(TXT, 7)) = "http://" Or LCase(Left(TXT, 4)) = "www." Then
    
        If Not InStr(TXT, " ") Then TXT = TXT & " "
        N = InStr(TXT, " ")
        TXT = Left(TXT, N - 1)
        Dim c As ColorConstants, U As Boolean
        With Rch
        U = .SelUnderline
        c = .SelColor
        .SelColor = vbBlue
        .SelUnderline = True
        .SelText = TXT
        .SelUnderline = U
        .SelColor = c
        End With
        
    ElseIf LCase(Left(TXT, 5)) = "<font" Then
        On Error Resume Next
        Dim FontS As String
        Dim FontC As String
        Dim FontF As String
        Dim AllFont As String
        Dim FontI As Integer
        FontS = Split(TXT, "size=" & Code)(1)
        FontS = Split(FontS, Code)(0)
        FontC = Split(TXT, "color=" & Code)(1)
        FontC = Split(FontC, Code)(0)
        FontF = Split(TXT, "face=" & Code)(1)
        FontF = Split(FontF, Code)(0)
        Rch.SelFontName = FontF
        Rch.SelColor = FontC
        Rch.SelFontSize = FontS
        N = InStr(TXT, ">") + 1
    ElseIf LCase(Left(TXT, 6)) = "</font" Then
        Rch.SelFontName = "Arial"
        Rch.SelFontSize = "10"
        Rch.SelColor = vbBlack
        N = 8
    ElseIf LCase(Left(TXT, 2)) = ":))" Then
        Rch.SelRTF = Pic(frmMenus.Image1)
        N = 3
    ElseIf LCase(Left(TXT, 3)) = ":-))" Then
        Rch.SelRTF = Pic(frmMenus.Image1)
        N = 4
    ElseIf LCase(Left(TXT, 2)) = ";)" Then
        Rch.SelRTF = Pic(frmMenus.Image2)
        N = 3
    ElseIf LCase(Left(TXT, 3)) = ";-)" Then
        Rch.SelRTF = Pic(frmMenus.Image2)
        N = 4
    ElseIf LCase(Left(TXT, 2)) = ":d" Then
        Rch.SelRTF = Pic(frmMenus.Image3)
        N = 3
    ElseIf LCase(Left(TXT, 3)) = ":-d" Then
        Rch.SelRTF = Pic(frmMenus.Image3)
        N = 4
    ElseIf LCase(Left(TXT, 2)) = ":x" Then
        Rch.SelRTF = Pic(frmMenus.Image4)
        N = 3
    ElseIf LCase(Left(TXT, 3)) = ":-x" Then
        Rch.SelRTF = Pic(frmMenus.Image4)
        N = 4
    ElseIf LCase(Left(TXT, 2)) = "x(" Then
        Rch.SelRTF = Pic(frmMenus.Image5)
        N = 3
    ElseIf LCase(Left(TXT, 3)) = "x-(" Then
        Rch.SelRTF = Pic(frmMenus.Image5)
        N = 4
    ElseIf LCase(Left(TXT, 2)) = ":p" Then
        Rch.SelRTF = Pic(frmMenus.Image6)
        N = 3
    ElseIf LCase(Left(TXT, 3)) = ":-p" Then
        Rch.SelRTF = Pic(frmMenus.Image6)
        N = 4
    ElseIf LCase(Left(TXT, 2)) = ":o" Then
        Rch.SelRTF = Pic(frmMenus.Image7)
        N = 3
    ElseIf LCase(Left(TXT, 3)) = ":-o" Then
        Rch.SelRTF = Pic(frmMenus.Image7)
        N = 4
    ElseIf LCase(Left(TXT, 2)) = ":""" Then
        Rch.SelRTF = Pic(frmMenus.Image8)
        N = 3
    ElseIf LCase(Left(TXT, 3)) = ":-""" Then
        Rch.SelRTF = Pic(frmMenus.Image8)
        N = 4
    ElseIf LCase(Left(TXT, 3)) = "/:)" Then
        Rch.SelRTF = Pic(frmMenus.Image9)
        N = 4
    ElseIf LCase(Left(TXT, 3)) = ":)" Then
        Rch.SelRTF = Pic(frmMenus.Image10)
        N = 4
    ElseIf LCase(Left(TXT, 4)) = ":-)" Then
        Rch.SelRTF = Pic(frmMenus.Image10)
        N = 5
    ElseIf LCase(Left(TXT, 3)) = ":((" Then
        Rch.SelRTF = Pic(frmMenus.Image11)
        N = 4
    ElseIf LCase(Left(TXT, 4)) = ":-((" Then
        Rch.SelRTF = Pic(frmMenus.Image11)
        N = 5
    ElseIf LCase(Left(TXT, 4)) = ":-((" Then
        Rch.SelRTF = Pic(frmMenus.Image11)
        N = 5
    ElseIf LCase(Left(TXT, 3)) = ":-/" Then
        Rch.SelRTF = Pic(frmMenus.Image12)
        N = 4
    ElseIf LCase(Left(TXT, 2)) = ":/" Then
        Rch.SelRTF = Pic(frmMenus.Image12)
        N = 3
    ElseIf LCase(Left(TXT, 3)) = ":-|" Then
        Rch.SelRTF = Pic(frmMenus.Image13)
        N = 4
    ElseIf LCase(Left(TXT, 2)) = ":|" Then
        Rch.SelRTF = Pic(frmMenus.Image13)
        N = 3
    ElseIf LCase(Left(TXT, 2)) = "=;" Then
        Rch.SelRTF = Pic(frmMenus.Image14)
        N = 3
    ElseIf LCase(Left(TXT, 2)) = ":>" Then
        Rch.SelRTF = Pic(frmMenus.Image15)
        N = 3
    ElseIf LCase(Left(TXT, 3)) = ":->" Then
        Rch.SelRTF = Pic(frmMenus.Image15)
        N = 4
    ElseIf LCase(Left(TXT, 2)) = ":s" Then
        Rch.SelRTF = Pic(frmMenus.Image16)
        N = 3
    ElseIf LCase(Left(TXT, 3)) = ":-s" Then
        Rch.SelRTF = Pic(frmMenus.Image16)
        N = 4
    ElseIf LCase(Left(TXT, 2)) = "b)" Then
        Rch.SelRTF = Pic(frmMenus.Image17)
        N = 3
    ElseIf LCase(Left(TXT, 3)) = "b-)" Then
        Rch.SelRTF = Pic(frmMenus.Image17)
        N = 4
    ElseIf LCase(Left(TXT, 2)) = ":(" Then
        Rch.SelRTF = Pic(frmMenus.Image18)
        N = 3
    ElseIf LCase(Left(TXT, 3)) = ":-(" Then
        Rch.SelRTF = Pic(frmMenus.Image18)
        N = 4
    ElseIf LCase(Left(TXT, 3)) = ":"">" Then
        Rch.SelRTF = Pic(frmMenus.Image19)
        N = 4
    ElseIf LCase(Left(TXT, 3)) = ">:)" Then
        Rch.SelRTF = Pic(frmMenus.Image20)
        N = 4
    ElseIf LCase(Left(TXT, 3)) = ";;)" Then
        Rch.SelRTF = Pic(frmPMSmileys.Image20)
        N = 4
    ElseIf LCase(Left(TXT, 2)) = ":*" Then
        Rch.SelRTF = Pic(frmPMSmileys.Image22)
        N = 3
    ElseIf LCase(Left(TXT, 3)) = ":-*" Then
        Rch.SelRTF = Pic(frmPMSmileys.Image22)
        N = 4
    ElseIf LCase(Left(TXT, 3)) = "=d>" Then
        Rch.SelRTF = Pic(frmPMSmileys.Image23)
        N = 4
    ElseIf LCase(Left(TXT, 3)) = "=p~" Then
        Rch.SelRTF = Pic(frmPMSmileys.Image24)
        N = 4
    ElseIf LCase(Left(TXT, 3)) = ":-/" Then
        Rch.SelRTF = Pic(frmPMSmileys.Image25)
        N = 4
    ElseIf LCase(Left(TXT, 2)) = ":/" Then
        Rch.SelRTF = Pic(frmPMSmileys.Image25)
        N = 3
    ElseIf LCase(Left(TXT, 3)) = "i-)" Then
        Rch.SelRTF = Pic(frmPMSmileys.Image26)
        N = 4
    ElseIf LCase(Left(TXT, 4)) = "o:-)" Then
        Rch.SelRTF = Pic(frmPMSmileys.Image27)
        N = 5
    ElseIf LCase(Left(TXT, 3)) = "o:)" Then
        Rch.SelRTF = Pic(frmPMSmileys.Image27)
        N = 4
    ElseIf LCase(Left(TXT, 3)) = ":-&" Then
        Rch.SelRTF = Pic(frmPMSmileys.Image28)
        N = 4
    ElseIf LCase(Left(TXT, 3)) = ":-?" Then
        Rch.SelRTF = Pic(frmPMSmileys.Image30)
        N = 4
    ElseIf LCase(Left(TXT, 3)) = ":-b" Then
        Rch.SelRTF = Pic(frmPMSmileys.Image31)
        N = 4
    ElseIf LCase(Left(TXT, 3)) = "[-(" Then
        Rch.SelRTF = Pic(frmPMSmileys.Image32)
        N = 4
    Else
        If frmConfig.Check6.Value = 1 Then Rch.SelFontSize = 10: Rch.SelFontName = "Arial"
        If frmConfig.Check4.Value = 1 Then Rch.SelColor = vbBlack
        T = Left(TXT, 1)
        Rch.SelText = T
    End If
    TXT = Mid(TXT, N)
    Rch.SelStart = Len(Rch.Text)
    If TXT = "" Then Exit Sub
Next I
End Sub
Function Hex2RGB(ByVal Hex As String, Optional Default As ColorConstants = vbBlack) As ColorConstants
On Error GoTo Error
Hex = Replace(Trim(Hex), ",", "")
If Left(Hex, 1) = "#" Then Hex = Mid(Hex, 2)
If Not Len(Hex) = 6 Then GoTo Error
Dim r As Integer, G As Integer, B As Integer
r = Val("&h" & Mid(Hex, 1, 2))
G = Val("&h" & Mid(Hex, 3, 2))
B = Val("&h" & Mid(Hex, 5, 2))
Hex2RGB = RGB(r, G, B)
Exit Function
Error:
Hex2RGB = Default
End Function
Sub CheckUrl(Rch As RichTextBox)

End Sub
Function TextClicked(Rch As RichTextBox, Optional sStr As String = " ") As String
On Error GoTo Error
Dim l As Long, S As String
l = Rch.SelStart
l = InStrRev(Rch.Text, sStr, l)
S = Mid(Rch.Text, l + 1)
S = S & " "
l = InStr(S, " ")
S = Trim(Left(S, l))
TextClicked = S
Exit Function
Error:
TextClicked = ""
End Function
Function ClickedURL(Rch As RichTextBox) As Boolean
Dim Str As String
Str = TextClicked(Rch)
If LCase(Left(Str, 7)) = "http://" Or LCase(Left(Str, 4)) = "www." Then
    URL = Str
    URL = Split(URL, Chr(&HA))(0)
    If URL = "" Then URL = Str
    If LCase(Left(Str, 4)) = "www." Then
    Shell "Explorer http://" & URL
    Else
    Shell "Explorer " & URL
    End If
    ClickedURL = True
Else
    ClickedURL = False
End If
End Function
Public Function RefreshPics(Rch As RichTextBox)
Dim lFoundPos As Long
Dim lFindLength As Long
Dim MakeSure As Boolean
GoTo Skip:
frmChat.Text1.SetFocus
Start:
frmChat.Text1.SetFocus
MakeSure = True
Skip:
lFoundPos = Rch.Find(":)", 0, , rtfNoHighlight)
While lFoundPos > 0
Rch.SelStart = lFoundPos
Rch.SelLength = 2
Rch.SelText = ""
Rch.OLEObjects.Add , , App.Path & "\smileys\1.bmp" 'Add the picture after it has deleted the string
frmChat.Text1.SetFocus
DoEvents
lFoundPos = Rch.Find(sFindString, lFoundPos + 2)
frmChat.Text1.SetFocus
Wend
If MakeSure = False Then GoTo Start
End Function

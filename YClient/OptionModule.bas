Attribute VB_Name = "Opts_Mod"
Public Username As String
Public Password As String
Public Room As String
Public U2D As String
Public U2D2 As Integer
Public RoomUser As String
Public data As String
Public dData() As String
Public sData() As String
Public PM(999) As New frmPM
Public OpenPMS As Integer
Public Const Code As String = """"
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Function PlayWav(Snd As String)
  Dim PlayIt
    PlayIt = sndPlaySound(Snd, 1)
End Function
Public Sub Pause(Seconds)
Dim Slayer            As Long
    Slayer = Timer
    Do
        DoEvents
    Loop Until Slayer + Seconds <= Timer
End Sub
Public Function SplitRoom(ByVal data As String, List As TreeView)
On Error Resume Next
Dim I As Integer
Dim Phrase1 As String
Dim Phrase2 As String
Dim RoomUser() As String
RoomUser = Split(data, "À€110À€")
For I = 1 To UBound(RoomUser)
If InStr(RoomUser(I), "À€") Then
Phrase1 = InStrRev(RoomUser(I), "À€")
Phrase2 = Mid(RoomUser(I), Phrase1 + 2)
RoomUser(I) = Phrase2
End If
List.Nodes.Add , , RoomUser(I), RoomUser(I), 2
Next
End Function
Public Function PhraseRoomData(ByVal data)
Dim Phrase1, Phrase2 As String
Phrase1 = InStr(data, "À€109À€")
Phrase2 = Mid(data, Phrase1 + 7)
Phrase1 = InStrRev(Phrase2, "À€110À€")
Phrase2 = Left(Phrase2, Phrase1 - 1)
PhraseRoomData = Phrase2
End Function
Public Function PhraseData(ByVal data, String1, String2 As String)
Dim Phrase1, Phrase2 As String
Phrase1 = InStr(data, String1)
Phrase2 = Mid(data, Phrase1 + Len(String1))
Phrase1 = InStr(Phrase2, String2)
Phrase2 = Left(Phrase2, Phrase1 - 1)
PhraseData = Phrase2
End Function
Public Sub SaveList(fileName As String, List As Listbox)
    On Error Resume Next
    Dim lngSave As Long
    
    If fileName$ = "" Then Exit Sub
    
    Open fileName$ For Output As #1
        For lngSave& = 0 To List.ListCount - 1
            Print #1, List.List(lngSave&)
        Next lngSave&
    Close #1
End Sub
Public Sub LoadList(fileName As String, List As Listbox)
  If fileName = "" Then Exit Sub
   Dim lstInput As String
    On Error Resume Next
    Open fileName$ For Input As #1
    While Not EOF(1)
        Input #1, lstInput$
    If lstInput$ = "" Then Exit Sub
        List.AddItem lstInput$
    Wend
    Close #1
End Sub

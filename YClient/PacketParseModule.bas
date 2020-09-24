Attribute VB_Name = "Parse_Mod"
Dim BuddyUser As String
Public RKey As String, RM_Space As String
Public Function InfoData(Sock As Winsock) As String
On Error Resume Next
Select Case Asc(Mid(data, 12, 1))
Case 87
GetString Sock
Case 85
LoginNow
UserStatus frmBuddy.TreeView1
Case 84
frmBuddy.Label1.Caption = "Sign In"
frmBuddy.Toolbar1.Visible = False
frmBuddy.TreeView1.Visible = False
frmBuddy.Label1.Visible = True
frmBuddy.Image1.Visible = True
frmBuddy.Label2.Visible = True
frmBuddy.Label3.Visible = True
frmBuddy.Label4.Visible = True
frmBuddy.Label1.FontUnderline = True
Case 1
UserStatus frmBuddy.TreeView1
If InStr(1, data, "¿Ä32¿Ä") Then
GetOfflines frmOfflines.List1, frmOfflines.List2, frmOfflines
End If
Case 2
If InStr(1, data, "¿Ä10¿Ä") Then
UserOffline frmBuddy.TreeView1
End If
Case 198
UserStatus frmBuddy.TreeView1
Case 6 'PM
GetPM
Case 152
If Mid(data, 13, 4) = "ˇˇˇˇ" Then
Select Case Mid(data, 21, 8)
Case Is = "114¿Ä16" & Chr(&HA)
frmChatLoad.Label1.Caption = "You are Already in this Room"
Case Is = "114¿Ä-33"
frmChatLoad.Label1.Caption = "This Room is Private"
Case Is = "114¿Ä-35"
frmChatLoad.Label1.Caption = "This Room is Full Please try Again Later"
End Select
End If

If InStr(1, data, "¿Ä128¿Ä") Then
If Not frmChat.List1.Nodes.Count = 0 Then
frmChat.List1.Nodes.Clear
End If
StartVoice
GetDesc
frmChat.Show
Unload frmChatLoad
Dim RoomList As String
RoomList = PhraseRoomData(data)
Call SplitRoom(RoomList, frmChat.List1)
End If

If InStr(1, data, "¿Ä109¿Ä") Then
UserJoin
End If
Case 155
If InStr(1, data, "¿Ä109¿Ä") Then
UserLeft
End If
Case 168
GetRoomText
Case 157
GetChatInvite
Case 75
MakeType
Case 15
BuddyUser = Split(data, "3¿Ä")(1)
BuddyUser = Split(BuddyUser, "¿Ä")(0)
If InStr(1, data, "1¿Ä") And InStr(1, data, "¿Ä3¿") Then
AskBuddy
End If
If InStr(1, data, "3¿Ä" & BuddyUser & "¿Ä14¿Ä") Then
DenyBuddy
End If
If InStr(1, data, "¿Ä10¿Ä") Then
Pause 1
UserStatus frmBuddy.TreeView1
End If
Case 131
NewBuddy
Case 132
frmBuddy.TreeView1.Nodes.Remove U2D2
Case 77
GetFile
Case 24
GetConf
End Select
Debug.Print Asc(Mid(data, 12, 1)) & " - " & data
End Function
Public Function SplitBuddy(Listbox As TreeView)
On Error Resume Next
If InStr(1, data, "7¿Ä") Then
Dim I As Integer
Dim Group As String
Dim FirstList As String
Dim List() As String
FirstList = Split(data, "7¿Ä")(1)
FirstList = Split(FirstList, "¿Ä")(0)
FirstList = Replace(FirstList, ":", ":,")
FirstList = Replace(FirstList, Chr(&HA), ",")
frmProfiles.FriendList.Clear
List = Split(FirstList, ",")
For I = 0 To UBound(List)
If List(I) = "" Then GoTo 1
If InStr(1, List(I), ":") Then
frmProfiles.FriendList.AddItem Split(List(I), ":")(0)
Group = Split(List(I), ":")(0)
Listbox.Nodes.Add , , Group, Group, 5
Else
Listbox.Nodes.Add Group, tvwChild, List(I), List(I), 1
End If
1
Next
For X = 1 To Listbox.Nodes.Count
Listbox.Nodes(X).Expanded = True
Next
End If
End Function
Public Function LoginNow() As String
frmBuddy.TreeView1.Nodes.Add , , "YahooBuddyMainList", "Friends for - " & Username, 3
SplitBuddy frmBuddy.TreeView1
frmBuddy.TreeView1.Visible = True
frmBuddy.Label1.Visible = False
frmBuddy.Label2.Visible = False
frmBuddy.Label3.Visible = False
frmBuddy.Label4.Visible = False
frmBuddy.Toolbar1.Visible = True
GetProfiles
End Function
Public Function GetString(Sock As Winsock) As String
    sData = Split(data, "¿Ä")
    Call getencrstrings(Username, Password, sData(3), Crypt(0), Crypt(1), 1)
    Sock.SendData Login(Username, Crypt(0), Crypt(1))
End Function
Public Function UserStatus(Listbox As TreeView) As String
Dim I As Integer
For I = 1 To Listbox.Nodes.Count
Dim User As String
Dim SelectC As String
User = Listbox.Nodes(I).Key

If InStr(1, data, "7¿Ä" & User & "¿Ä10¿Ä0¿Ä11¿Ä") Then
Listbox.Nodes(I).Bold = True
Listbox.Nodes(I).Image = 2
Listbox.Nodes(I).Text = Listbox.Nodes(I).Key
frmStatus.Show
frmStatus.Label1.Caption = " " & Listbox.Nodes(I).Key & ": Available"
frmStatus.Left = 0
frmStatus.Top = 0
End If

If InStr(1, data, "7¿Ä" & User & "¿Ä10¿Ä1") Then
Listbox.Nodes(I).Bold = True
Listbox.Nodes(I).Image = 7
Listbox.Nodes(I).Text = Listbox.Nodes(I).Key & " (" & "Be Right Back" & ")"
frmStatus.Show
frmStatus.Label1.Caption = " " & Listbox.Nodes(I).Key & ": Be Right Back"
frmStatus.Left = 0
frmStatus.Top = 0
End If
If InStr(1, data, "7¿Ä" & User & "¿Ä10¿Ä2") Then
Listbox.Nodes(I).Bold = True
Listbox.Nodes(I).Image = 7
Listbox.Nodes(I).Text = Listbox.Nodes(I).Key & " (" & "Busy" & ")"
frmStatus.Show
frmStatus.Label1.Caption = " " & Listbox.Nodes(I).Key & ": Busy"
frmStatus.Left = 0
frmStatus.Top = 0
End If
If InStr(1, data, "7¿Ä" & User & "¿Ä10¿Ä3") Then
Listbox.Nodes(I).Bold = True
Listbox.Nodes(I).Image = 7
Listbox.Nodes(I).Text = Listbox.Nodes(I).Key & " (" & "Not At Home" & ")"
frmStatus.Show
frmStatus.Label1.Caption = " " & Listbox.Nodes(I).Key & ": Not At Home"
frmStatus.Left = 0
frmStatus.Top = 0
End If
If InStr(1, data, "7¿Ä" & User & "¿Ä10¿Ä4") Then
Listbox.Nodes(I).Bold = True
Listbox.Nodes(I).Image = 7
Listbox.Nodes(I).Text = Listbox.Nodes(I).Key & " (" & "Not At My Desk" & ")"
frmStatus.Show
frmStatus.Label1.Caption = " " & Listbox.Nodes(I).Key & ": Not At My Desk"
frmStatus.Left = 0
frmStatus.Top = 0
End If
If InStr(1, data, "7¿Ä" & User & "¿Ä10¿Ä5") Then
Listbox.Nodes(I).Bold = True
Listbox.Nodes(I).Image = 7
Listbox.Nodes(I).Text = Listbox.Nodes(I).Key & " (" & "Not In The Office" & ")"
frmStatus.Show
frmStatus.Label1.Caption = " " & Listbox.Nodes(I).Key & " is now Not In The Office"
frmStatus.Left = 0
frmStatus.Top = 0
End If
If InStr(1, data, "7¿Ä" & User & "¿Ä10¿Ä6") Then
Listbox.Nodes(I).Bold = True
Listbox.Nodes(I).Image = 7
Listbox.Nodes(I).Text = Listbox.Nodes(I).Key & " (" & "On The Phone" & ")"
frmStatus.Show
frmStatus.Label1.Caption = " " & Listbox.Nodes(I).Key & ": On The Phone"
frmStatus.Left = 0
frmStatus.Top = 0
End If
If InStr(1, data, "7¿Ä" & User & "¿Ä10¿Ä7") Then
Listbox.Nodes(I).Bold = True
Listbox.Nodes(I).Image = 7
Listbox.Nodes(I).Text = Listbox.Nodes(I).Key & " (" & "On Vacation" & ")"
frmStatus.Show
frmStatus.Label1.Caption = " " & Listbox.Nodes(I).Key & ": On Vacation"
frmStatus.Left = 0
frmStatus.Top = 0
End If
If InStr(1, data, "7¿Ä" & User & "¿Ä10¿Ä8") Then
Listbox.Nodes(I).Bold = True
Listbox.Nodes(I).Image = 7
Listbox.Nodes(I).Text = Listbox.Nodes(I).Key & " (" & "Out To Lunch" & ")"
frmStatus.Show
frmStatus.Label1.Caption = " " & Listbox.Nodes(I).Key & ": Out To Lunch"
frmStatus.Left = 0
frmStatus.Top = 0
End If
If InStr(1, data, "7¿Ä" & User & "¿Ä10¿Ä9") Then
Listbox.Nodes(I).Bold = True
Listbox.Nodes(I).Image = 7
Listbox.Nodes(I).Text = Listbox.Nodes(I).Key & " (" & "Stepped Out" & ")"
frmStatus.Show
frmStatus.Label1.Caption = " " & Listbox.Nodes(I).Key & ": Stepped Out"
frmStatus.Left = 0
frmStatus.Top = 0
End If
If InStr(1, data, "7¿Ä" & User & "¿Ä10¿Ä99¿Ä19¿Ä") Then
Listbox.Nodes(I).Bold = True
Dim CustomMSG As String
CustomMSG = Split(data, User & "¿Ä10¿Ä99¿Ä19¿Ä")(1)
CustomMSG = Split(CustomMSG, "¿Ä")(0)
If InStr(1, data, CustomMSG & "¿Ä47¿Ä1") Then
Listbox.Nodes(I).Image = 7
Else
Listbox.Nodes(I).Image = 2
End If
Listbox.Nodes(I).Text = Listbox.Nodes(I).Key & " (" & CustomMSG & ")"
frmStatus.Show
frmStatus.Label1.Caption = " " & Listbox.Nodes(I).Key & ": " & CustomMSG
frmStatus.Left = 0
frmStatus.Top = 0
End If
If InStr(1, data, "7¿Ä" & User & "¿Ä10¿Ä999") Then
Listbox.Nodes(I).Bold = True
Listbox.Nodes(I).Image = 6
Listbox.Nodes(I).Text = Listbox.Nodes(I).Key & " (" & "Idle" & ")"
frmStatus.Label1.Caption = " " & Listbox.Nodes(I).Key & ": Idle"
frmStatus.Left = 0
frmStatus.Top = 0
End If
Next
End Function
Public Function GetProfiles() As String
Dim Dat As String
Dim Profils() As String
Dat = Split(data, "¿Ä89¿Ä")(1)
Dat = Split(Dat, "¿Ä")(0)
frmProfiles.ProfileList.Clear
Profils = Split(Dat, ",")
For X = 0 To UBound(Profils)
If Profils(X) = "" Then GoTo 1
frmProfiles.ProfileList.AddItem Profils(X)
1
Next
End Function
Public Function GetPM() As String
MakeFilter
On Error Resume Next
Dim FromWho As String
Dim Message As String
Dim ToWho As String
Dim N As Integer
N = InStr(data, "4¿Ä")
FromWho = Split(Mid(data, N + 3, Len(data)), "¿Ä")(0)
ToWho = Split(data, "5¿Ä")(1)
ToWho = Split(ToWho, "¿Ä")(0)
Message = Split(data, "¿Ä14¿Ä")(1)
Message = Split(Message, "¿Ä")(0)
If FromWho = "" Then GoTo 1
If Message = "" Then GoTo 1
For X = 0 To frmConfig.List2.ListCount
If LCase(FromWho) = LCase(frmConfig.List2.List(X)) Then
Exit Function
End If
Next
For X = 1 To UBound(PM)
If LCase(PM(X).Label4.Caption = LCase(FromWho)) Then
If LCase(Left(Message, 6)) = "<ding>" Then
ProcessDing PM(X).Text2, PM(X)
GoTo 3
End If
PM(X).Text2.SelBold = True
PM(X).Label8.Caption = X
PM(X).Text2.SelFontSize = 10
PM(X).Text2.SelFontName = "Arial"
PM(X).Text2.SelColor = vbBlue
PM(X).Text2.SelItalic = False
PM(X).Text2.SelUnderline = False
PM(X).Text2.SelText = FromWho & ": "
PM(X).Text2.SelBold = False
PM(X).Text2.SelColor = vbBlack
ProcessText Message, PM(X).Text2
PM(X).Text2.SelText = vbCrLf
PM(X).Text2.SelStart = Len(PM(X).Text2)
3:
PM(X).Caption = FromWho & " - " & ToWho
PM(X).StatusBar1.Panels(1).Text = "Last Message received at " & Time
Exit For
ElseIf PM(X).Visible = False Then
PM(X).Show: DoEvents
If LCase(Left(Message, 6)) = "<ding>" Then
ProcessDing PM(X).Text2, PM(X)
PM(X).Caption = FromWho & " - " & ToWho
PM(X).Label4.Caption = FromWho
PM(X).Label4.Visible = True
PM(X).Text3.Text = FromWho
PM(X).Text3.Visible = False
PM(X).Label2.Caption = ToWho
PM(X).Label2.Visible = True
PM(X).Combo1.Visible = False
GoTo 4
End If
If frmConfig.Check8.Value = 1 Then
PlayWav frmConfig.Text1
End If
PM(X).Label8.Caption = X
PM(X).Caption = FromWho & " - " & ToWho
PM(X).Label4.Caption = FromWho
PM(X).Label4.Visible = True
PM(X).Text3.Text = FromWho
PM(X).Text3.Visible = False
PM(X).Label2.Caption = ToWho
PM(X).Label2.Visible = True
PM(X).Combo1.Visible = False
PM(X).Text2.SelBold = True
PM(X).Text2.SelColor = vbBlue
PM(X).Text2.SelFontSize = 10
PM(X).Text2.SelFontName = "Arial"
PM(X).Text2.SelItalic = False
PM(X).Text2.SelUnderline = False
PM(X).Text2.SelText = FromWho & ": "
PM(X).Text2.SelBold = False
PM(X).Text2.SelColor = vbBlack
ProcessText Message, PM(X).Text2
PM(X).Text2.SelText = vbCrLf
4:
PM(X).Text2.SelStart = Len(PM(X).Text2)
PM(X).StatusBar1.Panels(1).Text = "Last Message received at " & Time
Exit For
End If
Next
1
End Function
Public Function UserOffline(Listbox As TreeView) As String
For I = 1 To Listbox.Nodes.Count
Dim User As String
User = Listbox.Nodes(I).Key
If InStr(1, data, "7¿Ä" & User & "¿Ä10¿Ä0") Then
Listbox.Nodes(I).Bold = False
Listbox.Nodes(I).Image = 1
Listbox.Nodes(I).Text = Listbox.Nodes(I).Key
frmStatus.Show
frmStatus.Label1.Caption = " " & Listbox.Nodes(I).Key & ": Offline"
frmStatus.Top = 0
frmStatus.Left = 0
End If
Next
End Function
Public Function StartVoice() As String
RKey = Split(data, "¿Ä130¿Ä")(1)
RKey = Split(RKey, "¿Ä")(0)
RM_Space = Split(data, "¿Ä129¿Ä")(1)
RM_Space = Split(RM_Space, "¿Ä")(0)
End Function
Public Function EnableVoice()
With frmChat.Voice1
    .leaveConference
    .HostName = "vc1.vip.scd.yahoo.com"
    .Username = RoomUser
    .appInfo = "mc(5, 6, 7, 1358)&u=" & RoomUser & "&ia=us"
    .confKey = RKey
    .ConfName = "ch/" & Room & "::" & RM_Space
    .createAndJoinConference
End With
End Function
Public Function UserLeft() As String
Dim NewUserOnRoom As String
NewUserOnRoom = Split(data, "¿Ä109¿Ä")(1)
NewUserOnRoom = Split(NewUserOnRoom, "¿Ä")(0)
For X = 1 To frmChat.List1.Nodes.Count
If frmChat.List1.Nodes(X) = NewUserOnRoom Then
frmChat.StatusBar1.Panels(2).Text = NewUserOnRoom & " left the Room"
frmChat.List1.Nodes.Remove X
frmChat.Text2.SelColor = vbRed
frmChat.Text2.SelFontName = "Arial"
frmChat.Text2.SelItalic = False
frmChat.Text2.SelBold = False
frmChat.Text2.SelUnderline = False
frmChat.Text2.SelFontSize = "10"
frmChat.Text2.SelText = NewUserOnRoom
frmChat.Text2.SelColor = vbBlack
frmChat.Text2.SelText = " left the Room" & vbCrLf
frmChat.Text2.SelStart = Len(frmChat.Text2)
GoTo ExSu
End If
Next
ExSu:
End Function
Public Function UserJoin() As String
On Error Resume Next
Dim NewUserOnRoom As String
NewUserOnRoom = Split(data, "¿Ä109¿Ä")(1)
NewUserOnRoom = Split(NewUserOnRoom, "¿Ä")(0)
For X = 1 To frmChat.List1.Nodes.Count
If frmChat.List1.Nodes(X) = NewUserOnRoom Then
frmChat.List1.Nodes.Remove X
End If
Next
frmChat.List1.Nodes.Add , , NewUserOnRoom, NewUserOnRoom, 2
frmChat.StatusBar1.Panels(2).Text = NewUserOnRoom & " joined the Room"
frmChat.Text2.SelColor = vbRed
frmChat.Text2.SelFontSize = "10"
frmChat.Text2.SelFontName = "Arial"
frmChat.Text2.SelItalic = False
frmChat.Text2.SelBold = False
frmChat.Text2.SelUnderline = False
frmChat.Text2.SelText = NewUserOnRoom
frmChat.Text2.SelColor = vbBlack
frmChat.Text2.SelText = " joined the Room" & vbCrLf
frmChat.Text2.SelStart = Len(frmChat.Text2)
End Function
Public Function GetDesc() As String
Dim Spons As String
Spons = Split(data, "¿Ä105¿Ä")(1)
Spons = Split(Spons, "¿Ä")(0)
frmChat.StatusBar1.Panels(1).Text = "You are In " & Code & Room & Code & " (" & Spons & ")"
End Function
Public Function GetRoomText() As String
On Error Resume Next
Dim UserText As String
Dim UserOnChat As String
UserOnChat = Split(data, "¿Ä109¿Ä")(1)
UserOnChat = Split(UserOnChat, "¿Ä")(0)
UserText = Split(data, "¿Ä117¿Ä")(1)
UserText = Split(UserText, "¿Ä")(0)
frmChat.Text2.SelStart = Len(frmChat.Text2)
frmChat.Text2.SelBold = True
frmChat.Text2.SelFontName = "Arial"
frmChat.Text2.SelItalic = False
frmChat.Text2.SelUnderline = False
frmChat.Text2.SelColor = &H40&
frmChat.Text2.SelFontSize = "10"
frmChat.Text2.SelText = UserOnChat & ": "
frmChat.Text2.SelColor = vbBlack
frmChat.Text2.SelBold = False
ProcessText UserText, frmChat.Text2
If frmConfig.Check7.Value = 1 Then
PlayWav frmConfig.Text2
End If
frmChat.Text2.SelText = vbCrLf
If frmConfig.Check7.Value = 1 Then
PlayWav frmConfig.Text2
End If
End Function
Public Function MakeType() As String
On Error GoTo Error1
Dim UserTyping As String
UserTyping = Split(data, "4¿Ä")(1)
UserTyping = Split(UserTyping, "¿Ä")(0)
For X = 0 To OpenPMS
If LCase(PM(X).Label4.Caption = LCase(UserTyping)) Then
PM(X).StatusBar1.Panels(1).Text = UserTyping & " is typing a Message"
End If
Next
Error1:
End Function
Public Function NewBuddy() As String
On Error GoTo 1
frmAddBuddy.Frame1.Visible = False
frmAddBuddy.Frame2.Visible = True
frmAddBuddy.Label5.Caption = frmAddBuddy.Text1 & " got added in the BuddyList"
For X = 1 To frmBuddy.TreeView1.Nodes.Count
Dim Group As String
If frmBuddy.TreeView1.Nodes(X).Key = frmAddBuddy.Combo1 Then
Group = frmAddBuddy.Combo1
frmBuddy.TreeView1.Nodes.Add Group, tvwChild, frmAddBuddy.Text1, frmAddBuddy.Text1, 1
Exit Function
End If
Next
Group = frmAddBuddy.Combo1
Dim ID As String
ID = frmAddBuddy.Text1
frmBuddy.TreeView1.Nodes.Add , , Group, Group, 4
frmBuddy.TreeView1.Nodes.Add Group, tvwChild, , ID, 1
For X = 1 To frmBuddy.TreeView1.Nodes.Count
If frmBuddy.TreeView1.Nodes(X).Image = 4 Then
frmBuddy.TreeView1.Nodes(X).Expanded = True
frmBuddy.TreeView1.Nodes(X).Image = 5
End If
Next
1
End Function
Public Function DenyBuddy() As String
On Error Resume Next
Dim DeclineMSG As String
DeclineMSG = Split(data, "14¿Ä")(1)
DeclineMSG = Split(DeclineMSG, "¿Ä")(0)
MsgBox BuddyUser & " has declined your Request: " & DeclineMSG, vbOKOnly, "Declined"
For X = 1 To frmBuddy.TreeView1.Nodes.Count
If frmBuddy.TreeView1.Nodes(X).Key = BuddyUser Then
frmBuddy.TreeView1.Nodes.Remove X
End If
Next
End Function
Public Function AskBuddy() As String
'15 - YMSG     -    ÇÅ1ﬂ1¿Äyaaahmaster22¿Ä3¿Äe_c_programmer¿Ä14¿Äyo¿Ä
On Error Resume Next
Dim From As String
Dim WhoTo As String
Dim MSG As String
From = Split(data, "¿Ä3¿Ä")(1)
From = Split(From, "¿Ä")(0)
WhoTo = Split(data, "1¿Ä")(1)
WhoTo = Split(WhoTo, "¿Ä")(0)
MSG = Split(data, "¿Ä14¿Ä")(1)
MSG = Split(MSG, "¿Ä")(0)
With frmBuddyRequest
.Show
.Label2.Caption = From
.Text1 = MSG
.Label1.Caption = From & " has added you in his Buddy List"
End With
End Function
Public Function GetFile() As String
Dim From As String
Dim fileName As String
Dim FileUrl As String
From = Split(data, "4¿Ä")(1)
From = Split(From, "¿Ä")(0)
fileName = Split(data, "7¿Ä")(1)
fileName = Split(fileName, "¿Ä")(0)
FileUrl = Split(data, "20¿Ä")(1)
FileUrl = Split(FileUrl, "¿Ä")(0)
Dim XBox As String
XBox = MsgBox("From: " & From & vbCrLf & "File: " & fileName & vbCrLf & vbCrLf & "Download File", vbYesNo, "File Request")
Select Case XBox
Case vbYes
frmDl.WebBrowser1.Navigate FileUrl
Case vbNo
frmDl.WebBrowser1.Navigate FileUrl
frmDl.WebBrowser1.Stop
End Select
End Function
Public Function GetConf() As String
'24 - YMSG     Ë    ÇÑÍ∂1¿Äyaaahmaster28¿Ä57¿Äe_c_programmer-vUeEnwWn4_gqLsK4X6v6nA--¿Ä50¿Äe_c_programmer¿Ä52¿Äyaaahmaster28¿Ä58¿ÄJoin My Conference...¿Ä13¿Ä256¿Ä234¿Äe_c_programmer-vUeEnwWn4_gqLsK4X6v6nA--¿Ä233¿ÄYe2BBB__Chv1zu_NizlaZzjQKbJAIV6I8-¿Ä97¿Ä1¿Ä
Dim From As String
Dim WhoTo As String
Dim Conf As String
Dim ConfMessage As String
WhoTo = Split(data, "1¿Ä")(1)
WhoTo = Split(WhoTo, "¿Ä")(0)
Conf = Split(data, "57¿Ä")(1)
Conf = Split(Conf, "¿Ä")(0)
From = Split(data, "50¿Ä")(1)
From = Split(From, "¿Ä")(0)
ConfMessage = Split(data, "58¿Ä")(1)
ConfMessage = Split(ConfMessage, "¿Ä")(0)
frmConfInvite.Show
frmConfInvite.Text2 = Conf
frmConfInvite.Text1 = From
frmConfInvite.Text3 = WhoTo
frmConfInvite.Text4 = ConfMessage
End Function
Public Function GetOfflines(UserBox As Listbox, OfflineBox As Listbox, OffWindow As Form) As String
On Error GoTo 2
OffWindow.Show
Dim Usr() As String, Off() As String
Usr = Split(data, "¿Ä4¿Ä"): Off = Split(data, "¿Ä14¿Ä")
For X = 1 To UBound(Usr)
If Usr(X) = "" Then GoTo 1:
Dim Usr2 As String, Off2 As String
Usr2 = Split(Usr(X), "¿Ä")(0)
Off2 = Split(Off(X), "¿Ä")(0)
UserBox.AddItem Usr2
OfflineBox.AddItem Off2
1:
Next
2:
End Function
Public Function GetChatInvite()
Dim From As String
Dim WhoTo As String
Dim Room As String
Dim MSG As String
From = Split(data, "119¿Ä")(1)
From = Split(From, "¿Ä")(0)
WhoTo = Split(data, "118¿Ä")(1)
WhoTo = Split(WhoTo, "¿Ä")(0)
Room = Split(data, "4¿Ä")(1)
Room = Split(Room, "¿Ä")(0)
MSG = Split(data, "¿Ä117¿Ä")(1)
MSG = Split(MSG, "¿Ä")(0)
frmChatInvite.Show
frmChatInvite.Text3 = WhoTo
frmChatInvite.Text1 = From
frmChatInvite.Text2 = Room
frmChatInvite.Text4 = MSG
End Function
Sub MakeFilter()
Dim strCuss As String, Spt() As String
If frmConfig.Check10.Value = 1 Then
For X = 0 To frmConfig.List3.ListCount
strCuss = strCuss & frmConfig.List3.List(X) & ","
Next
Spt = Split(strCuss, ",")
For I = 0 To UBound(Spt)
data = Replace(data, Spt(I), String(Len(Spt(I)), "*"), , , vbTextCompare)
Next
End If
End Sub

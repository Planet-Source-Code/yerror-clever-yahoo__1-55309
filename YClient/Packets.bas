Attribute VB_Name = "Packets"
'Packets made by Y-Error
'(C)Copyright 2004 Yah-Masters.com
'For Questions email me at: xxxxxx_anonymous_xxxxxx@yahoo.com
'-------------------------------------------------------------
Public Function YPager(whoto As String) As String
Packet = "7À€" & whoto & "À€10À€0À€11À€7D65DAC2À€17À€0À€13À€0À€"
YPager = Header("02", Packet)
Debug.Print YPager
End Function
Public Function AddMyFriend(from As String, whoto As String, Group As String, message As String) As String
Dim Packet As String
Packet = "1À€" & from & "À€7À€" & whoto & "À€14À€À€65À€" & Group & "sÀ€97À€1À€216À€À€"
AddMyFriend = Header("D0", Packet)
End Function
Public Function DeleteFriend(from As String, FriendToDelete As String, Group As String) As String
Dim Packet As String
Packet = "1À€" & from & "À€7À€" & FriendToDelete & "À€65À€" & Group & "À€"
DeleteFriend = Header("84", Packet)
End Function
Public Function ConfInvite(from As String, whoto As String, message As String, confrence As String) As String
Dim Packet As String
Packet = "1À€" & from & "À€50À€" & from & "À€57À€" & from & "À€57À€" & from & "À€57À€" & from & "À€57À€" & from & "À€57À€" & from & "À€57À€" & from & "À€57À€" & from & "À€57À€" & from & confrence & "À€58À€                                                              " _
& message & "À€97À€1À€52À€" & whoto & "À€13À€256À€"
ConfInvite = Header("18", Packet)
End Function
Public Function VoiceInvite(from As String, whoto As String) As String
Dim Packet As String
Packet = "1À€" & from & "À€5À€" & whoto & "À€57" & "À€À€13À€1À€"
VoiceInvite = Header("4A", Packet)
End Function
Public Function SendPM(from As String, whoto As String, message As String) As String
Dim Packet As String
Packet = "1À€" & from & "À€5À€" & whoto & "À€14À€" & message & "À€97À€1" _
& "À€63À€;0À€64À€0À€1002À€1À€206À€0À€15À€1086903880À€11À€-1820828541À€"
SendPM = Header("06", Packet)
End Function
Public Function Status(message As String, busy As Boolean) As String
Dim Packet As String
If busy = True Then
Packet = "10À€99À€19À€" & message & "À€47À€1À€187À€0À€"
Else
Packet = "10À€99À€19À€" & message & "À€47À€0À€187À€0À€"
End If
Status = Header("C6", Packet)
End Function
Public Function SendFile(from As String, whoto As String, file As String) As String
Dim Packet As String
Packet = "5À€" & whoto & "À€49À€FILEXFERÀ€1À€" & from & "À€13À€1À€27À€" & file & "À€28À€485À€20À€"
SendFile = Header("4D", Packet)
End Function
Public Function Deny(from As String, whoto As String, message As String) As String
Dim Packet As String
Packet = "1À€" & from & "À€7À€" & whoto & "À€14À€" & message & "À€"
Deny = Header("86", Packet)
End Function
Public Function JoinRoom(user As String) As String
'######## Note: this will make the Room Ready
Dim Packet As String
Packet = "109À€" & user & "À€1À€" & user & "À€6À" & "€abcdeÀ€98À€usÀ€" _
& "135À€ym6,0,0,1643À€"
JoinRoom = Header("96", Packet)
End Function
Public Function GotoRoom(user As String, room As String) As String
'######## Note this will Join the Room
Dim Packet As String
Packet = "1À€" & user & "À€104À€" & room & "À€12" & "9À€1600326535À€6" _
& "2À€2À€"
GotoRoom = Header("98", Packet)
End Function
Public Function LeaveRoom(user As String, room As String) As String
Dim Packet As String
Packet = "1À€" & user & "À€1005À€322" & "85272À€"
LeaveRoom = Header("A0", Packet)
End Function
Public Function ChatText(user As String, room As String, message As String) As String
Dim Packet As String
Packet = "1À€" & user & "À€104À€" & room & "À€117À€" _
& message & "À€124À€1À€"
ChatText = Header("A8", Packet)
End Function
Public Function CamStatus(Status As String) As String
Dim Packet As String
Packet = "10À€99À€19À€" & Status & "À€184À€" & "YSTATUS=1À€47" _
& "À€0À€187À€1À€"
CamStatus = Header("C6", Packet)
End Function
Public Function Unknown(user As String) As String
'I dont know this Packet i sniffed some
'Packets and saw this i dont know what it does
'But im added it here
Dim Packet As String
Packet = "1À€" & user & "À€212À€1À€192À€-650782246À€"
Unkown = Header("BD", Packet)
End Function
Public Function Boot(from As String, whoto As String) As String
Dim Packet As String
Dim Packet2 As String
Packet = "1À€" & from & "À€5À€" & whoto & "À€212À€1À€192À€-650782246À€"
Packet2 = "1À€" & from & "À€5À€" & whoto & "À€212À€1À€192À€-650782246À€"
Boot = Header("BD", Packet) & Header("BD", Packet)
End Function
Public Function Invisible() As String
'This will make you Invisible
Dim Packet As String
Packet = "13À€2À€"
Invisible = Header("C5", Packet)
End Function
Public Function Audibles(from As String, whoto As String, audible As String) As String
Dim Packet As String
If audible = "nosepick" Then audible = "hello.nosepick1"
If audible = "dude" Then audible = "hello.dude"
If audible = "seeyou" Then audible = "hello.seeyou"
If audible = "toothy" Then audible = "htf.toothy"
Packet = "1À€" & from & "À€5À€" & whoto & "À€230À€" & "base.us." & audible _
& "1À€231À€" & "Yo!" & "À€232À€0d29376f051ad14fb85f2fc9ff8b03b31" _
& "06d4689À€"
Audibles = Header("D0", Packet)
End Function


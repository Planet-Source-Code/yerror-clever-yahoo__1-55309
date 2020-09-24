Attribute VB_Name = "Packets_Mod"
Public Function JoinRoom(User As String) As String
Dim Packet As String
Packet = "109À€" & User & "À€1À€" & User & "À€6À" & "€abcdeÀ€98À€usÀ€" _
& "135À€ym6,0,0,1643À€"
JoinRoom = Header("96", Packet)
End Function
Public Function GoToRoom(User As String, Room As String) As String
Dim Packet As String
Packet = "1À€" & User & "À€104À€" & Room & "À€12" & "9À€1600326535À€6" _
& "2À€2À€"
GoToRoom = Header("98", Packet)
End Function
Public Function Typing(User As String, WhoTo As String) As String
Dim Packet As String
Packet = "5À€" & WhoTo & "À€4À€" & User & "À€14À€ À€13À€1À€49À€TYPINGÀ€"
Typing = Header("4B", Packet)
End Function
Public Function SendPM(From As String, WhoTo As String, Message As String) As String
Dim Packet As String
Packet = "1À€" & From & "À€5À€" & WhoTo & "À€14À€" & Message & "À€97À€1" _
& "À€63À€;0À€64À€0À€1002À€1À€206À€0À€15À€1086903880À€11À€-1820828541À€"
SendPM = Header("06", Packet)
End Function
Public Function ChatText(User As String, Room As String, Message As String) As String
Dim Packet As String
Packet = "1À€" & User & "À€104À€" & Room & "À€117À€" _
& Message & "À€124À€1À€"
ChatText = Header("A8", Packet)
End Function
Public Function AddFriend(ID As String, Buddy As String, Group As String, Message As String) As String
Dim Packet As String
Packet = "1À€" & ID & "À€7À€" & Bud & "À€14À€" & Message & "À€65À€" & Grp & "À€"
AddFriend = Header("83", Packet)
End Function
Public Function DeleteFriend(From As String, FriendToDelete As String, Group As String) As String
Dim Packet As String
Packet = "1À€" & From & "À€7À€" & FriendToDelete & "À€65À€" & Group & "À€"
DeleteFriend = Header("84", Packet)
End Function
Public Function ReFresh(From As String) As String
Dim Packet As String
Packet = "1À€" & From & "À€"
ReF = Header("55", Packet)
End Function
Public Function LeaveRoom(User As String) As String
Dim Packet As String
Packet = "1À€" & User & "À€1005À€35745352À€"
LeaveRoom = Header("A0", Packet)
End Function
Public Function DenyBudd(From As String, WhoTo As String, MSG As String) As String
Dim Packet As String
Packet = "1À€" & From & "À€7À€" & WhoTo & "À€14À€" & MSG & "À€"
DenyBudd = Header("86", Packet)
End Function
Public Function FollowUser(From As String, WhoTo As String) As String
Dim Packet As String
Packet = "109À€" & WhoTo & "À€1À€" & From & "À€62À€2À€"
FollowUser = Header("97", Packet)
End Function
Public Function YStatus(Message As String, busy As Boolean) As String
Dim Packet As String
If busy = True Then
Packet = "10À€99À€19À€" & Message & "À€47À€1À€187À€0À€"
Else
Packet = "10À€99À€19À€" & Message & "À€47À€0À€187À€0À€"
End If
YStatus = Header("C6", Packet)
End Function
Public Function InvI() As String
Dim Packet As String
Packet = "13À€2À€"
invisible = Header("C5", Packet)
End Function
Public Function AcceptConf(From As String, Conf As String, ConfName As String) As String
Dim Packet As String
Packet = "1À€" & From & "À€57À€" & Conf & "À€56À€" & ConfName & "À€"
AcceptConf = Header("1B", Packet)
End Function
Public Function AcceptConf2(From As String, Conf As String) As String
Dim Packet As String
Packet = "1À€" & From & "À€57À€" & Conf & "À€"
AcceptConf2 = Header("19", Packet)
End Function


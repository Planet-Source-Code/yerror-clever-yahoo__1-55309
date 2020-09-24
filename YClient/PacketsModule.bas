Attribute VB_Name = "Packets_Mod"
Public Function JoinRoom(User As String) As String
Dim Packet As String
Packet = "109¿Ä" & User & "¿Ä1¿Ä" & User & "¿Ä6¿" & "Äabcde¿Ä98¿Äus¿Ä" _
& "135¿Äym6,0,0,1643¿Ä"
JoinRoom = Header("96", Packet)
End Function
Public Function GoToRoom(User As String, Room As String) As String
Dim Packet As String
Packet = "1¿Ä" & User & "¿Ä104¿Ä" & Room & "¿Ä12" & "9¿Ä1600326535¿Ä6" _
& "2¿Ä2¿Ä"
GoToRoom = Header("98", Packet)
End Function
Public Function Typing(User As String, WhoTo As String) As String
Dim Packet As String
Packet = "5¿Ä" & WhoTo & "¿Ä4¿Ä" & User & "¿Ä14¿Ä ¿Ä13¿Ä1¿Ä49¿ÄTYPING¿Ä"
Typing = Header("4B", Packet)
End Function
Public Function SendPM(From As String, WhoTo As String, Message As String) As String
Dim Packet As String
Packet = "1¿Ä" & From & "¿Ä5¿Ä" & WhoTo & "¿Ä14¿Ä" & Message & "¿Ä97¿Ä1" _
& "¿Ä63¿Ä;0¿Ä64¿Ä0¿Ä1002¿Ä1¿Ä206¿Ä0¿Ä15¿Ä1086903880¿Ä11¿Ä-1820828541¿Ä"
SendPM = Header("06", Packet)
End Function
Public Function ChatText(User As String, Room As String, Message As String) As String
Dim Packet As String
Packet = "1¿Ä" & User & "¿Ä104¿Ä" & Room & "¿Ä117¿Ä" _
& Message & "¿Ä124¿Ä1¿Ä"
ChatText = Header("A8", Packet)
End Function
Public Function AddFriend(ID As String, Buddy As String, Group As String, Message As String) As String
Dim Packet As String
Packet = "1¿Ä" & ID & "¿Ä7¿Ä" & Bud & "¿Ä14¿Ä" & Message & "¿Ä65¿Ä" & Grp & "¿Ä"
AddFriend = Header("83", Packet)
End Function
Public Function DeleteFriend(From As String, FriendToDelete As String, Group As String) As String
Dim Packet As String
Packet = "1¿Ä" & From & "¿Ä7¿Ä" & FriendToDelete & "¿Ä65¿Ä" & Group & "¿Ä"
DeleteFriend = Header("84", Packet)
End Function
Public Function ReFresh(From As String) As String
Dim Packet As String
Packet = "1¿Ä" & From & "¿Ä"
ReF = Header("55", Packet)
End Function
Public Function LeaveRoom(User As String) As String
Dim Packet As String
Packet = "1¿Ä" & User & "¿Ä1005¿Ä35745352¿Ä"
LeaveRoom = Header("A0", Packet)
End Function
Public Function DenyBudd(From As String, WhoTo As String, MSG As String) As String
Dim Packet As String
Packet = "1¿Ä" & From & "¿Ä7¿Ä" & WhoTo & "¿Ä14¿Ä" & MSG & "¿Ä"
DenyBudd = Header("86", Packet)
End Function
Public Function FollowUser(From As String, WhoTo As String) As String
Dim Packet As String
Packet = "109¿Ä" & WhoTo & "¿Ä1¿Ä" & From & "¿Ä62¿Ä2¿Ä"
FollowUser = Header("97", Packet)
End Function
Public Function YStatus(Message As String, busy As Boolean) As String
Dim Packet As String
If busy = True Then
Packet = "10¿Ä99¿Ä19¿Ä" & Message & "¿Ä47¿Ä1¿Ä187¿Ä0¿Ä"
Else
Packet = "10¿Ä99¿Ä19¿Ä" & Message & "¿Ä47¿Ä0¿Ä187¿Ä0¿Ä"
End If
YStatus = Header("C6", Packet)
End Function
Public Function InvI() As String
Dim Packet As String
Packet = "13¿Ä2¿Ä"
invisible = Header("C5", Packet)
End Function
Public Function AcceptConf(From As String, Conf As String, ConfName As String) As String
Dim Packet As String
Packet = "1¿Ä" & From & "¿Ä57¿Ä" & Conf & "¿Ä56¿Ä" & ConfName & "¿Ä"
AcceptConf = Header("1B", Packet)
End Function
Public Function AcceptConf2(From As String, Conf As String) As String
Dim Packet As String
Packet = "1¿Ä" & From & "¿Ä57¿Ä" & Conf & "¿Ä"
AcceptConf2 = Header("19", Packet)
End Function


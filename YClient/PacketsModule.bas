Attribute VB_Name = "Packets_Mod"
Public Function JoinRoom(User As String) As String
Dim Packet As String
Packet = "109��" & User & "��1��" & User & "��6�" & "�abcde��98��us��" _
& "135��ym6,0,0,1643��"
JoinRoom = Header("96", Packet)
End Function
Public Function GoToRoom(User As String, Room As String) As String
Dim Packet As String
Packet = "1��" & User & "��104��" & Room & "��12" & "9��1600326535��6" _
& "2��2��"
GoToRoom = Header("98", Packet)
End Function
Public Function Typing(User As String, WhoTo As String) As String
Dim Packet As String
Packet = "5��" & WhoTo & "��4��" & User & "��14�� ��13��1��49��TYPING��"
Typing = Header("4B", Packet)
End Function
Public Function SendPM(From As String, WhoTo As String, Message As String) As String
Dim Packet As String
Packet = "1��" & From & "��5��" & WhoTo & "��14��" & Message & "��97��1" _
& "��63��;0��64��0��1002��1��206��0��15��1086903880��11��-1820828541��"
SendPM = Header("06", Packet)
End Function
Public Function ChatText(User As String, Room As String, Message As String) As String
Dim Packet As String
Packet = "1��" & User & "��104��" & Room & "��117��" _
& Message & "��124��1��"
ChatText = Header("A8", Packet)
End Function
Public Function AddFriend(ID As String, Buddy As String, Group As String, Message As String) As String
Dim Packet As String
Packet = "1��" & ID & "��7��" & Bud & "��14��" & Message & "��65��" & Grp & "��"
AddFriend = Header("83", Packet)
End Function
Public Function DeleteFriend(From As String, FriendToDelete As String, Group As String) As String
Dim Packet As String
Packet = "1��" & From & "��7��" & FriendToDelete & "��65��" & Group & "��"
DeleteFriend = Header("84", Packet)
End Function
Public Function ReFresh(From As String) As String
Dim Packet As String
Packet = "1��" & From & "��"
ReF = Header("55", Packet)
End Function
Public Function LeaveRoom(User As String) As String
Dim Packet As String
Packet = "1��" & User & "��1005��35745352��"
LeaveRoom = Header("A0", Packet)
End Function
Public Function DenyBudd(From As String, WhoTo As String, MSG As String) As String
Dim Packet As String
Packet = "1��" & From & "��7��" & WhoTo & "��14��" & MSG & "��"
DenyBudd = Header("86", Packet)
End Function
Public Function FollowUser(From As String, WhoTo As String) As String
Dim Packet As String
Packet = "109��" & WhoTo & "��1��" & From & "��62��2��"
FollowUser = Header("97", Packet)
End Function
Public Function YStatus(Message As String, busy As Boolean) As String
Dim Packet As String
If busy = True Then
Packet = "10��99��19��" & Message & "��47��1��187��0��"
Else
Packet = "10��99��19��" & Message & "��47��0��187��0��"
End If
YStatus = Header("C6", Packet)
End Function
Public Function InvI() As String
Dim Packet As String
Packet = "13��2��"
invisible = Header("C5", Packet)
End Function
Public Function AcceptConf(From As String, Conf As String, ConfName As String) As String
Dim Packet As String
Packet = "1��" & From & "��57��" & Conf & "��56��" & ConfName & "��"
AcceptConf = Header("1B", Packet)
End Function
Public Function AcceptConf2(From As String, Conf As String) As String
Dim Packet As String
Packet = "1��" & From & "��57��" & Conf & "��"
AcceptConf2 = Header("19", Packet)
End Function


Attribute VB_Name = "Packets"
'Packets made by Y-Error
'(C)Copyright 2004 Yah-Masters.com
'For Questions email me at: xxxxxx_anonymous_xxxxxx@yahoo.com
'-------------------------------------------------------------
Public Function YPager(whoto As String) As String
Packet = "7��" & whoto & "��10��0��11��7D65DAC2��17��0��13��0��"
YPager = Header("02", Packet)
Debug.Print YPager
End Function
Public Function AddMyFriend(from As String, whoto As String, Group As String, message As String) As String
Dim Packet As String
Packet = "1��" & from & "��7��" & whoto & "��14����65��" & Group & "s��97��1��216����"
AddMyFriend = Header("D0", Packet)
End Function
Public Function DeleteFriend(from As String, FriendToDelete As String, Group As String) As String
Dim Packet As String
Packet = "1��" & from & "��7��" & FriendToDelete & "��65��" & Group & "��"
DeleteFriend = Header("84", Packet)
End Function
Public Function ConfInvite(from As String, whoto As String, message As String, confrence As String) As String
Dim Packet As String
Packet = "1��" & from & "��50��" & from & "��57��" & from & "��57��" & from & "��57��" & from & "��57��" & from & "��57��" & from & "��57��" & from & "��57��" & from & "��57��" & from & confrence & "��58��                                                              " _
& message & "��97��1��52��" & whoto & "��13��256��"
ConfInvite = Header("18", Packet)
End Function
Public Function VoiceInvite(from As String, whoto As String) As String
Dim Packet As String
Packet = "1��" & from & "��5��" & whoto & "��57" & "����13��1��"
VoiceInvite = Header("4A", Packet)
End Function
Public Function SendPM(from As String, whoto As String, message As String) As String
Dim Packet As String
Packet = "1��" & from & "��5��" & whoto & "��14��" & message & "��97��1" _
& "��63��;0��64��0��1002��1��206��0��15��1086903880��11��-1820828541��"
SendPM = Header("06", Packet)
End Function
Public Function Status(message As String, busy As Boolean) As String
Dim Packet As String
If busy = True Then
Packet = "10��99��19��" & message & "��47��1��187��0��"
Else
Packet = "10��99��19��" & message & "��47��0��187��0��"
End If
Status = Header("C6", Packet)
End Function
Public Function SendFile(from As String, whoto As String, file As String) As String
Dim Packet As String
Packet = "5��" & whoto & "��49��FILEXFER��1��" & from & "��13��1��27��" & file & "��28��485��20��"
SendFile = Header("4D", Packet)
End Function
Public Function Deny(from As String, whoto As String, message As String) As String
Dim Packet As String
Packet = "1��" & from & "��7��" & whoto & "��14��" & message & "��"
Deny = Header("86", Packet)
End Function
Public Function JoinRoom(user As String) As String
'######## Note: this will make the Room Ready
Dim Packet As String
Packet = "109��" & user & "��1��" & user & "��6�" & "�abcde��98��us��" _
& "135��ym6,0,0,1643��"
JoinRoom = Header("96", Packet)
End Function
Public Function GotoRoom(user As String, room As String) As String
'######## Note this will Join the Room
Dim Packet As String
Packet = "1��" & user & "��104��" & room & "��12" & "9��1600326535��6" _
& "2��2��"
GotoRoom = Header("98", Packet)
End Function
Public Function LeaveRoom(user As String, room As String) As String
Dim Packet As String
Packet = "1��" & user & "��1005��322" & "85272��"
LeaveRoom = Header("A0", Packet)
End Function
Public Function ChatText(user As String, room As String, message As String) As String
Dim Packet As String
Packet = "1��" & user & "��104��" & room & "��117��" _
& message & "��124��1��"
ChatText = Header("A8", Packet)
End Function
Public Function CamStatus(Status As String) As String
Dim Packet As String
Packet = "10��99��19��" & Status & "��184��" & "YSTATUS=1��47" _
& "��0��187��1��"
CamStatus = Header("C6", Packet)
End Function
Public Function Unknown(user As String) As String
'I dont know this Packet i sniffed some
'Packets and saw this i dont know what it does
'But im added it here
Dim Packet As String
Packet = "1��" & user & "��212��1��192��-650782246��"
Unkown = Header("BD", Packet)
End Function
Public Function Boot(from As String, whoto As String) As String
Dim Packet As String
Dim Packet2 As String
Packet = "1��" & from & "��5��" & whoto & "��212��1��192��-650782246��"
Packet2 = "1��" & from & "��5��" & whoto & "��212��1��192��-650782246��"
Boot = Header("BD", Packet) & Header("BD", Packet)
End Function
Public Function Invisible() As String
'This will make you Invisible
Dim Packet As String
Packet = "13��2��"
Invisible = Header("C5", Packet)
End Function
Public Function Audibles(from As String, whoto As String, audible As String) As String
Dim Packet As String
If audible = "nosepick" Then audible = "hello.nosepick1"
If audible = "dude" Then audible = "hello.dude"
If audible = "seeyou" Then audible = "hello.seeyou"
If audible = "toothy" Then audible = "htf.toothy"
Packet = "1��" & from & "��5��" & whoto & "��230��" & "base.us." & audible _
& "1��231��" & "Yo!" & "��232��0d29376f051ad14fb85f2fc9ff8b03b31" _
& "06d4689��"
Audibles = Header("D0", Packet)
End Function


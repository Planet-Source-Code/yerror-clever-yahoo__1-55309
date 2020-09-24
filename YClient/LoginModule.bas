Attribute VB_Name = "Login_Mod"
Option Explicit
Public S(999) As Boolean
Public Crypt(1) As String
Public SessionKey As String, Packet As String
Private Declare Function YCrypt Lib "YCrypt.dll" (ByVal Username As String, ByVal Password As String, ByVal Seed As String, ByVal result_6 As String, ByVal result_96 As String, intt As Long) As Boolean
Public Function getencrstrings(name As String, Pass As String, Seed As String, Str1 As String, Str2 As String, mode As Long) As Boolean
    Dim Ts As String, Ts2 As String, N As Long
    On Error GoTo err
    Ts = String(80, vbNullChar)
    Ts2 = String(80, vbNullChar)
    getencrstrings = YCrypt(name, Pass, Seed, Ts, Ts2, mode)
    N = InStr(1, Ts, vbNullChar)
    Str1 = Left$(Ts, N - 1)
    N = InStr(1, Ts2, vbNullChar)
    Str2 = Left$(Ts2, N - 1)
    Exit Function
err:
    getencrstrings = False
End Function
Public Function Header(ByVal PacketType As String, ByVal Pck As String) As String
    Dim I As Integer
    Dim x As Integer
    x = 0
    I = Len(Pck)
    Do While I > 255
    I = I - 256
    x = x + 1
    Loop
Header = "YMSG" & Chr(0) & Chr(12) & String(2, 0) & Chr(x) & Chr(I) & Chr(0) & _
Chr("&H" & PacketType) & String(8, 0) & Pck
Debug.Print Header
End Function
Public Function Get_Key(ByVal Username As String) As String
Get_Key = "1À€" & Username & "À€"
Get_Key = Header(57, Get_Key)
End Function
Public Function Login(ByVal Username As String, ByVal Crypt1 As String, ByVal crypt2 As String) As String
Login = "6À€" & Crypt1 & "À€96À€" & crypt2 & "À€0À€" & Username & "À€2À€1À€1À€" & Username & _
"À€135À€5, 6, 0, 1347À€148À€300À€"
Login = Header(54, Login)
End Function

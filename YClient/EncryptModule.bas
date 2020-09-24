Attribute VB_Name = "EncryptModule"
Option Explicit
Public Crypt(1) As String
Public SessionKey As String, Packet As String
Private Declare Function YCrypt Lib "YCrypt.dll" (ByVal Username As String, ByVal Password As String, ByVal Seed As String, ByVal result_6 As String, ByVal result_96 As String, intt As Long) As Boolean
Public Function getencrstrings(name As String, Pass As String, Seed As String, Str1 As String, Str2 As String, mode As Long) As Boolean
    Dim Ts As String, Ts2 As String, n As Long
    On Error GoTo err
    Ts = String(80, vbNullChar)
    Ts2 = String(80, vbNullChar)
    getencrstrings = YCrypt(name, Pass, Seed, Ts, Ts2, mode)
    n = InStr(1, Ts, vbNullChar)
    Str1 = Left$(Ts, n - 1)
    n = InStr(1, Ts2, vbNullChar)
    Str2 = Left$(Ts2, n - 1)
    Exit Function
err:
    getencrstrings = False
End Function



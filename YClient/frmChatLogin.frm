VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmChatLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Chat Room List"
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   6975
   Icon            =   "frmChatLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   6975
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "frmChatLogin.frx":030A
      Left            =   1320
      List            =   "frmChatLogin.frx":0314
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   120
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   9
      Top             =   3720
      Width           =   2055
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   4920
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   -240
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChatLogin.frx":0329
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChatLogin.frx":043B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView treCats 
      Height          =   2895
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   5106
      _Version        =   393217
      HideSelection   =   0   'False
      Sorted          =   -1  'True
      Style           =   1
      ImageList       =   "ImageList1"
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin MSComctlLib.TreeView treRooms 
      Height          =   2895
      Left            =   3480
      TabIndex        =   2
      Top             =   720
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   5106
      _Version        =   393217
      Style           =   1
      ImageList       =   "ImageList1"
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin Project1.chameleonButton chameleonButton1 
      Height          =   315
      Left            =   3840
      TabIndex        =   3
      Top             =   3720
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Cancel"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14933984
      FCOL            =   0
   End
   Begin Project1.chameleonButton chameleonButton2 
      Height          =   315
      Left            =   2280
      TabIndex        =   4
      Top             =   3720
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Go to Room"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14933984
      FCOL            =   0
   End
   Begin Project1.chameleonButton chameleonButton3 
      Height          =   315
      Left            =   5400
      TabIndex        =   5
      Top             =   3720
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Refresh"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14933984
      FCOL            =   0
   End
   Begin MSWinsockLib.Winsock wskHTTP 
      Left            =   2760
      Top             =   4920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.TreeView treURooms 
      Height          =   2895
      Left            =   3480
      TabIndex        =   12
      Top             =   720
      Visible         =   0   'False
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   5106
      _Version        =   393217
      Style           =   1
      ImageList       =   "ImageList1"
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "User Rooms"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5160
      TabIndex        =   11
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Rooms"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3480
      TabIndex        =   8
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   "Enter chat room as:"
      Height          =   255
      Left            =   3480
      TabIndex        =   7
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Language:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmChatLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Timeout     As Integer
Dim HTML        As String
Dim HTTP_Server As String
Dim HTTP_Page   As String
Private Sub chameleonButton1_Click()
On Error Resume Next
Unload Me
End Sub
Public Function HTTP(Page As String, Optional Host As String = "127.0.0.1")
HTTP = "GET /" & Page & " HTTP/1.1" & vbCrLf & _
"Host: " & Host & vbCrLf & _
"User-Agent: Mozilla/5.0 (Windows) frmHTTP" & vbCrLf & _
"Accept: text/html,*/*" & vbCrLf & _
"Accept -Language: en -ca" & vbCrLf & vbCrLf
End Function

Public Function OpenURL(ByVal URL As String, Optional Port As Integer = 80)
If Timeout = 0 Then Timeout = 10
On Error GoTo Error
Dim i As Integer
If LCase(Left(URL, 7)) = "http://" Then URL = Mid(URL, 8, Len(URL) - 7)
i = InStr(URL, "/")
If i < 1 Then
    URL = URL & "/"
    i = Len(URL)
End If
HTTP_Server = Left(URL, i - 1)
HTTP_Page = Right(URL, Len(URL) - i)
HTML = ""
wskHTTP.Connect HTTP_Server, Port
Dim Sec As Long
Sec = Timer + Timeout
Do Until Timer > Sec
    DoEvents
    If wskHTTP.State = 0 Then GoTo Done
Loop
wskHTTP.Close
'HTML = "Time out"
Done:
OpenURL = HTML
Exit Function
Error:
OpenURL = "Error!"
End Function
Sub GetCats(Cats As String)
On Error Resume Next
treCats.Nodes.Clear
Dim Spt() As String, K As String
Spt = Split(Cats, "<category")
For i = 1 To UBound(Spt)
    K = "chatroom_" & GetVal(Spt(i), "id")
    Spt(i) = GetVal(Spt(i), "name")
    Spt(i) = Replace(Spt(i), "&amp;", "&", , , vbTextCompare)
    Spt(i) = Replace(Spt(i), "&apos;", "'", , , vbTextCompare)
    Spt(i) = Replace(Spt(i), "l&#xe4;", "ä")
    treCats.Nodes.Add , , Trim(K), Trim(Spt(i)), 1
Next
End Sub

Sub GetRooms(Rooms As String)
On Error Resume Next
Dim Spt() As String, Spt2() As String, r As String, K As String
Spt = Split(Rooms, "<room")
treURooms.Nodes.Clear
treRooms.Nodes.Clear
For i = 1 To UBound(Spt)
 If Left(Spt(i), 11) = " type=" & Chr(34) & "user" Then ' user room
    K = "(" & GetVal(Spt(i), "users") & ")"
    Spt(i) = GetVal(Spt(i), "name")
    Spt(i) = Replace(Spt(i), "&apos;", "'", , , vbTextCompare)
    Spt(i) = Replace(Spt(i), "&amp;", "&", , , vbTextCompare)
    treURooms.Nodes.Add , , Spt(i), Spt(i) & " " & K, 2
 Else ' yahoo room
    r = GetVal(Spt(i), "name")
    r = Replace(r, "&apos;", "'", , , vbTextCompare)
    r = Replace(r, "&amp;", "&", , , vbTextCompare)
    Spt2 = Split(Spt(i), "<lobby")
    For ii = 1 To UBound(Spt2)
     K = GetVal(Spt2(ii), "count")
     Dim NewRoom As String
     NewRoom = r & ":" & K & " (" & GetVal(Spt2(ii), "users") & ")"
     List1.AddItem Split(NewRoom, " (")(0)
     treRooms.Nodes.Add , , (r & ":" & K), NewRoom, 2
    Next
 End If
Next

End Sub

Function GetVal(ByVal StrAll As String, Str As String)
' get value like  'name2' out of '<name1="matt" name2="fred" name3="bob">'
' only works if string uses "s and =s
Dim i As Integer
On Error GoTo Error
i = InStr(1, StrAll, Str & "=" & Chr(34), vbTextCompare)
If i < 2 Then GoTo Error
StrAll = Mid(StrAll, i + Len(Str) + 2)
i = InStr(StrAll, Chr(34))
If i < 2 Then GoTo Error
StrAll = Left(StrAll, i - 1)
Error:
GetVal = StrAll
End Function




Sub LoadRooms(Cat As String)
If Cat = "" Then Exit Sub
treRooms.Nodes.Clear
treRooms.Nodes.Add , , , "Loading..."
Dim data As String
If Combo2.Text = "English" Then
data = "http://insider.msg.yahoo.com/ycontent/?" & Cat
Else
data = "http://insider.msg.yahoo.com/ycontent/?" & Cat & "&intl=de"
End If
DoEvents
data = OpenURL(data)
DoEvents
GetRooms data
End Sub

Private Sub cmdJoin_Click()
If txtRoom = "" Then Exit Sub
'code to join room here
'Hide
End Sub



Private Sub chameleonButton2_Click()
On Error Resume Next
RoomUser = Combo1.Text
Room = Text1
If Room = "" Then
MsgBox "Please Select a Room", vbOKOnly, "Error"
Exit Sub
End If
frmLogin.Socket.SendData JoinRoom(RoomUser)
Pause (0.02)
frmLogin.Socket.SendData GoToRoom(RoomUser, Room)
frmChatLoad.Show
wskHTTP.Close
Unload Me
End Sub

Private Sub chameleonButton3_Click()
On Error Resume Next
RoomStart
End Sub

Private Sub Combo2_Change()
treCats.Nodes.Clear
treRooms.Nodes.Clear
treURooms.Nodes.Clear
RoomStart
End Sub

Private Sub Combo2_Click()
RoomStart
End Sub

Private Sub Combo2_Scroll()
RoomStart
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
wskHTTP.Close
Unload Me
End Sub

Private Sub Timer1_Timer()

End Sub

Private Sub Label1_Click()
treRooms.Visible = False
treURooms.Visible = True
Label2.FontBold = False
Label1.FontBold = True
End Sub

Private Sub Label2_Click()
treRooms.Visible = True
treURooms.Visible = False
Label2.FontBold = True
Label1.FontBold = False
End Sub

Private Sub treCats_BeforeLabelEdit(Cancel As Integer)
Cancel = 1
End Sub

Private Sub treCats_Click()
On Error Resume Next
LoadRooms treCats.SelectedItem.Key
End Sub

Private Sub treRooms_BeforeLabelEdit(Cancel As Integer)
Cancel = 1
End Sub

Private Sub treRooms_Click()
On Error Resume Next
Text1 = Split(treRooms.SelectedItem, " (")(0)
End Sub

Private Sub treRooms_DblClick()
RoomUser = Combo1.Text
Room = Split(treRooms.SelectedItem, " (")(0)
If Room = "" Then
MsgBox "Please Select a Room", vbOKOnly, "Error"
Exit Sub
End If
frmLogin.Socket.SendData JoinRoom(RoomUser)
Pause 1
frmLogin.Socket.SendData GoToRoom(RoomUser, Room)
frmChatLoad.Show
Unload Me
End Sub


Private Sub treURooms_BeforeLabelEdit(Cancel As Integer)
Cancel = 1
End Sub

Private Sub treURooms_Click()
Text1 = Split(treURooms.SelectedItem, " (")(0)
End Sub

Private Sub treURooms_DblClick()
RoomUser = Combo1.Text
Room = Split(treURooms.SelectedItem, " (")(0)
If Room = "" Then
MsgBox "Please Select a Room", vbOKOnly, "Error"
Exit Sub
End If
frmLogin.Socket.SendData JoinRoom(RoomUser)
Pause 1
frmLogin.Socket.SendData GoToRoom(RoomUser, Room)
frmChatLoad.Show
Unload Me
End Sub
Sub CancelURL()
HTML = "Cancelled"
wskHTTP.Close
End Sub

Private Sub cmdTest_Click()
Dim Test As String
Test = OpenURL(txtURL)
Debug.Print Test
MsgBox Test
End Sub


Private Sub wskHTTP_Close()
wskHTTP.Close
End Sub

Private Sub wskHTTP_Connect()
wskHTTP.SendData HTTP(HTTP_Page, HTTP_Server)
End Sub

Private Sub wskHTTP_DataArrival(ByVal bytesTotal As Long)
Dim data As String
wskHTTP.GetData data
HTML = HTML & data
End Sub
Public Function RoomStart() As String
treCats.Nodes.Clear
Timeout = 10
Dim data As String
treCats.Nodes.Clear
treCats.Nodes.Add , , , "Loading...", 1
DoEvents
If Combo2.Text = "English" Then
wskHTTP.Close
data = OpenURL("http://insider.msg.yahoo.com/ycontent/?chatcat")
Else
wskHTTP.Close
data = OpenURL("http://insider.msg.yahoo.com/ycontent/?chatcat&intl=de")
End If
DoEvents
GetCats data
End Function



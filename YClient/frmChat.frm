VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{2B323CCC-50E3-11D3-9466-00A0C9700498}#1.0#0"; "yacscom.dll"
Object = "{7D1E9C3C-BD6A-11D3-87A8-009027A35D73}#1.0#0"; "yacsui.dll"
Begin VB.Form frmChat 
   ClientHeight    =   6720
   ClientLeft      =   165
   ClientTop       =   750
   ClientWidth     =   10215
   Icon            =   "frmChat.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6720
   ScaleWidth      =   10215
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.Slider YSlider1 
      Height          =   255
      Left            =   5040
      TabIndex        =   16
      Top             =   4800
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   450
      _Version        =   393216
      Min             =   1
      Max             =   100
      SelStart        =   100
      Value           =   100
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000C000&
      Caption         =   "Talk"
      Enabled         =   0   'False
      Height          =   300
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4830
      Width           =   855
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Hands-free"
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   4830
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Top             =   5640
      Width           =   9015
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   6465
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   3720
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   5280
      Width           =   750
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   5280
      Width           =   3495
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   6240
      Top             =   5400
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6600
      Top             =   5400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":045C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView List1 
      Height          =   3975
      Left            =   7800
      TabIndex        =   0
      Top             =   720
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   7011
      _Version        =   393217
      Sorted          =   -1  'True
      Style           =   1
      ImageList       =   "ImageList1"
      Appearance      =   1
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   1111
      ButtonWidth     =   1455
      ButtonHeight    =   1058
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Voice"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Chat"
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin RichTextLib.RichTextBox Text2 
      Height          =   3975
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   7011
      _Version        =   393217
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      MousePointer    =   99
      OLEDropMode     =   0
      TextRTF         =   $"frmChat.frx":056E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.chameleonButton chameleonButton3 
      Default         =   -1  'True
      Height          =   735
      Left            =   9240
      TabIndex        =   5
      Top             =   5640
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   "Send"
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
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   7200
      Top             =   5400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   34
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":05E5
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":1959
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":2CCD
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image Image2 
      Height          =   360
      Left            =   2280
      Picture         =   "frmChat.frx":35B5
      Top             =   4800
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   4560
      Picture         =   "frmChat.frx":37C5
      Top             =   4800
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Line Line1 
      X1              =   3600
      X2              =   3600
      Y1              =   5160
      Y2              =   4800
   End
   Begin VB.Label Label4 
      Height          =   255
      Left            =   6000
      TabIndex        =   15
      Top             =   4890
      Width           =   4095
   End
   Begin YACSUILibCtl.YVuMeter Meter1 
      Height          =   375
      Left            =   2760
      OleObjectBlob   =   "frmChat.frx":39D5
      TabIndex        =   14
      Top             =   4800
      Width           =   735
   End
   Begin YACSUILibCtl.YVuMeter Meter2 
      Height          =   375
      Left            =   3720
      OleObjectBlob   =   "frmChat.frx":39F9
      TabIndex        =   13
      Top             =   4800
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "&U"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5280
      TabIndex        =   8
      Top             =   5280
      Width           =   255
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "I"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4920
      TabIndex        =   7
      Top             =   5280
      Width           =   255
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4560
      TabIndex        =   6
      Top             =   5280
      Width           =   255
   End
   Begin YACSCOMLibCtl.YAcs Voice1 
      Left            =   120
      OleObjectBlob   =   "frmChat.frx":3A1D
      Top             =   6840
   End
   Begin VB.Menu menu1 
      Caption         =   "File"
      Begin VB.Menu chatopts 
         Caption         =   "Chat Options"
      End
   End
End
Attribute VB_Name = "frmChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chameleonButton3_Click()
If LCase(Left(Text1, 6)) = "/join " Then
Dim NewRoom As String
NewRoom = Split(Text1, "/join ")(1)
Command1.Enabled = False
Text1 = ""
Room = NewRoom
Voice1.leaveConference
frmLogin.Socket.SendData JoinRoom(RoomUser)
Pause 1
frmLogin.Socket.SendData GoToRoom(RoomUser, Room)
GoTo 3
End If
If Text1 = "" Then
MsgBox "Please enter a Message", vbOKOnly, "Error"
Exit Sub
End If
On Error Resume Next
Dim MSG As String
Dim FontF As String
Dim FontS As String
Dim FontC As String
Dim AllFont As String
Dim Font As String
Dim OtherFont As String
If Label3.BorderStyle = 1 Then
OtherFont = OtherFont & "<b>"
End If
If Label2.BorderStyle = 1 Then
OtherFont = OtherFont & "<i>"
End If
If Label1.BorderStyle = 1 Then
OtherFont = OtherFont & "<u>"
End If
MSG = Text1.Text
Text1 = ""
Text2.SelFontSize = 10
Text2.SelFontName = "Arial"
Text2.SelBold = True
Text2.SelItalic = False
Text2.SelUnderline = False
Text2.SelColor = vbBlue
Text2.SelStart = Len(Text2)
Text2.SelText = RoomUser & ": "
Text2.SelColor = vbBlack
Text2.SelBold = False
Font = "<font face=" & Code & Combo2 & Code & " size=" & Code & Combo3 & Code & ">"
ProcessText Font & OtherFont & MSG, Text2
Text2.SelText = vbCrLf
Text2.SelStart = Len(Text2)
frmLogin.Socket.SendData ChatText(RoomUser, Room, OtherFont & Font & MSG)
3:
End Sub

Private Sub chatopts_Click()
frmConfig.Show
frmConfig.Frame1.Visible = False
frmConfig.Frame2.Visible = True
frmConfig.Frame3.Visible = False
frmConfig.Frame4.Visible = False
frmConfig.List1.ListIndex = 3
End Sub

Private Sub Check1_Click()
If Check1.Value = 1 Then
Voice1.startTransmit
Else
Voice1.stopTransmit
End If
End Sub

Private Sub Combo2_Change()
Text1.FontName = Combo2
End Sub

Private Sub Combo3_Change()
Text1.FontSize = Combo3
End Sub

Private Sub Command1_Click()
Voice1.stopTransmit
End Sub

Private Sub Command1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
With frmChat.Voice1
.startTransmit
End With
End Sub

Private Sub Form_Load()
Dim I As Integer
For I = 0 To Screen.FontCount
Combo2.AddItem Screen.FontS(I)
Next
For I = 1 To 32
Combo3.AddItem I
Next
Combo3.Text = "10"
Combo2.Text = "Arial"
Me.Caption = Room
End Sub

Private Sub Form_Resize()
On Error Resume Next
If Me.Height < 4095 Then
Me.Height = 4096
GoTo 1
End If
If Me.Width < 5775 Then
Me.Width = 5776
GoTo 1
End If
On Error Resume Next
Image1.Top = Meter1.Top
Image2.Top = Image1.Top
Text2.Width = Me.Width - 2700
List1.Left = Text2.Width + 200
Text1.Width = Me.Width - 1300
chameleonButton3.Left = Text1.Width + 250
Text2.Height = Me.Height - 3500
List1.Height = Text2.Height
Frame1.Top = Text2.Height + 500
Check1.Top = Text2.Height + 850
Command1.Top = Check1.Top
Meter1.Top = Command1.Top - 40
Meter2.Top = Meter1.Top
YSlider1.Top = Meter1.Top + 45
Line1.Y2 = Meter2.Top
Line1.Y1 = Meter2.Top + 400
Combo2.Top = Check1.Top + 460
StatusBar1.Panels(1).Width = Me.ScaleWidth / 2 + 2000
StatusBar1.Panels(2).Width = Me.ScaleWidth / 2 - 2000
Combo3.Top = Combo2.Top
Label1.Top = Combo3.Top
Label2.Top = Combo3.Top
Label3.Top = Combo3.Top
Text1.Top = Combo3.Top + 350
Label4.Top = Meter1.Top + 60
Label4.Width = Me.Width - 5000
chameleonButton3.Top = Text1.Top
1
End Sub

Private Sub Form_Unload(Cancel As Integer)
List1.Nodes.Clear
Text2.Text = ""
Text1.Text = ""
Voice1.leaveConference
Voice1.stopSound 99
On Error Resume Next
frmLogin.Socket.SendData LeaveRoom(RoomUser)
End Sub

Private Sub Label1_Click()
If Label1.BorderStyle = 1 Then
Label1.BorderStyle = 0
Else
Label1.BorderStyle = 1
End If
End Sub

Private Sub Label2_Click()
If Label2.BorderStyle = 1 Then
Label2.BorderStyle = 0
Else
Label2.BorderStyle = 1
End If
End Sub

Private Sub Label3_Click()
If Label3.BorderStyle = 1 Then
Label3.BorderStyle = 0
Else
Label3.BorderStyle = 1
End If
End Sub

Private Sub List1_DblClick()
Dim U As Integer
For U = 0 To frmProfiles.ProfileList.ListCount - 1
PM(OpenPMS).Combo1.AddItem frmProfiles.ProfileList.List(U)
Next
PM(OpenPMS).Combo1.Text = frmProfiles.ProfileList.List(0)
PM(OpenPMS).Show
PM(OpenPMS).Label4.Caption = List1.SelectedItem.Key
PM(OpenPMS).Text3.Visible = False
PM(OpenPMS).Label8.Caption = OpenPMS
PM(OpenPMS).Label4.Visible = True
OpenPMS = OpenPMS + 1
End Sub

Private Sub Text2_Click()
ClickedURL Text2
End Sub

Private Sub Text2_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error Resume Next

Dim Str As String
Str = TextClicked(Text2)
If LCase(Left(Str, 7)) = "http://" Or LCase(Left(Str, 4)) = "www." Then
Dim I As Integer, N As String, E As Integer
I = Text2.Find("http://", , , rtfNoHighlight)
Text2.SelStart = I
E = Mid(Text2, I, Len(Text2))
Text2.SelBold = True
Text2.SelStart = E
Text2.SelBold = False
End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
Case "Voice"
If Button.Image = 1 Then
Button.Image = 2
EnableVoice
Else
If Button.Image = 2 Then
Button.Image = 1
Voice1.leaveConference
Check1.Enabled = False
Command1.Enabled = False
End If
End If
Case "Chat"
frmChatLogin.Show
End Select
End Sub

Private Sub Voice1_onAudioError(ByVal Code As Long, ByVal message As String)
If Code = 1 Then
MsgBox "You dont have Permissions to Join this Confrence", vbOKOnly, "Voice Error"
End If
End Sub

Private Sub Voice1_onConferenceNotReady()
Command1.Enabled = False
Check1.Enabled = False
End Sub

Private Sub Voice1_onConferenceReady()
Command1.Enabled = True
Check1.Enabled = True
End Sub

Private Sub Voice1_onInputGainChange(ByVal gain As Integer)
Meter1.maxValue = gain
End Sub

Private Sub Voice1_onInputLevelChange(ByVal level As Integer)
Dim Level2 As String
Level2 = Split(level, "-")(1)
If Level2 = "" Then GoTo 3
Meter1.Value = Level2
3:
End Sub

Private Sub Voice1_onLocalOffAir()
Image2.Visible = False
Image2.Visible = False
End Sub

Private Sub Voice1_onLocalOnAir()
Image2.Visible = True
Image2.Visible = True
End Sub

Private Sub Voice1_onOutputGainChange(ByVal gain As Integer)
Meter2.maxValue = gain
End Sub

Private Sub Voice1_onOutputLevelChange(ByVal level As Integer)
On Error Resume Next
Dim Level2 As String
Level2 = Split(level, "-")(1)
If Level2 = "" Then GoTo 3
Meter2.Value = Level2 * 3
3:
End Sub

Private Sub Voice1_onRemoteSourceOffAir(ByVal sourceId As Long, ByVal sourceName As String)
Image1.Visible = False
Image1.Visible = False
Label4.Caption = ""
Label4.Caption = ""
End Sub

Private Sub Voice1_onRemoteSourceOnAir(ByVal sourceId As Long, ByVal sourceName As String)
Image1.Visible = True
Image1.Visible = True
Label4.Caption = sourceName
End Sub

Private Sub Voice1_onSourceEntry(ByVal sourceId As Long, ByVal sourceName As String)
For x = 1 To List1.Nodes.Count
If List1.Nodes(x).Text = sourceName Then
List1.Nodes(x).Image = 1
End If
Next
End Sub

Private Sub Voice1_onSourceExit(ByVal sourceId As Long, ByVal sourceName As String)
For x = 1 To List1.Nodes.Count
If List1.Nodes(x).Text = sourceName Then
List1.Nodes(x).Image = 2
List1.Nodes(x).ForeColor = vbBlack
End If
Next
End Sub

Private Sub YSlider1_Change()
With Voice1
.outputGain = YSlider1.Value
    .inputAGC = YSlider1.Value
    .inputSource = YSlider1.Value
End With
End Sub


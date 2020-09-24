VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{AAC8DFAF-8A34-11D3-B327-000021C5C8A9}#1.0#0"; "Systray.ocx"
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   2520
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3735
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   3735
   StartUpPosition =   3  'Windows Default
   Begin SysTrayCtl.cSysTray cSysTray1 
      Left            =   3960
      Top             =   120
      _ExtentX        =   900
      _ExtentY        =   900
      InTray          =   -1  'True
      TrayIcon        =   "frmLogin.frx":030A
      TrayTip         =   "Clever Y!"
   End
   Begin MSWinsockLib.Winsock Socket 
      Left            =   0
      Top             =   2400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "User Login"
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3495
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmLogin.frx":0624
         Left            =   1080
         List            =   "frmLogin.frx":0649
         TabIndex        =   9
         Top             =   1080
         Width           =   2295
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Save Username && Password"
         Height          =   255
         Left            =   720
         TabIndex        =   7
         Top             =   1560
         Value           =   1  'Checked
         Width           =   2655
      End
      Begin Project1.chameleonButton chameleonButton1 
         Default         =   -1  'True
         Height          =   315
         Left            =   240
         TabIndex        =   5
         Top             =   1920
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Sign In"
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
      Begin VB.TextBox Pass 
         Height          =   285
         Left            =   1080
         TabIndex        =   4
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox User 
         Height          =   285
         Left            =   1080
         TabIndex        =   3
         Top             =   360
         Width           =   2295
      End
      Begin VB.Timer Timer2 
         Left            =   3840
         Top             =   600
      End
      Begin Project1.chameleonButton chameleonButton2 
         Height          =   315
         Left            =   1800
         TabIndex        =   6
         Top             =   1920
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
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Server:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Username:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Password:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chameleonButton1_Click()
On Error Resume Next
Me.Hide
frmBuddy.Show
Username = User.Text
Password = Pass.Text
Socket.Close
Socket.Connect "cs10.msg.dcn.yahoo.com", 5050
If Check1.Value = 1 Then
SaveSetting App.EXEName, "User", "User", User
SaveSetting App.EXEName, "Pass", "Pass", Pass
SaveSetting App.EXEName, "Server", "Server", Combo1.Text
End If
SaveSetting App.EXEName, "Save", "Save", Check1.Value
SaveSetting App.EXEName, "Login", "Login", Check2.Value
Me.Hide
End Sub

Private Sub cSysTray1_MouseDblClick(Button As Integer, ID As Long)
If Button = 1 Then
frmBuddy.Show
End If
End Sub

Private Sub cSysTray1_MouseDown(Button As Integer, ID As Long)
If Button = 2 Then
PopupMenu frmMenus.Menu3
End If
End Sub

Private Sub Form_Load()
GetSetts
Dim MSG As String
MSG = GetSetting(App.EXEName, "1", "1")
If MSG = "" Then GoTo 1
frmMenus.custom.Caption = MSG
1:
Combo1.Text = GetSetting(App.EXEName, "Server", "Server")
If Combo1.Text = "" Then
Combo1.Text = Combo1.List(0)
End If
OpenPMS = 1
On Error Resume Next
Check1.Value = GetSetting(App.EXEName, "Save", "Save")
Check2.Value = GetSetting(App.EXEName, "Login", "Login")
User = GetSetting(App.EXEName, "User", "User")
Pass = GetSetting(App.EXEName, "Pass", "Pass")
If User = "" Then User = "<Username>"
If Pass = "" Then
Pass = "<Password>"
Else
Pass.PasswordChar = "*"
End If
End Sub

Private Sub Pass_Change()
Pass.PasswordChar = "*"
End Sub

Private Sub Socket_Connect()
frmStatus.Label1.Caption = "Connecting to Yahoo!"
frmStatus.Show
frmStatus.Left = 0
frmStatus.Top = 0
frmBuddy.Label1.Caption = "Connecting to Yahoo!"
Socket.SendData Get_Key(Username)
End Sub

Private Sub Socket_DataArrival(ByVal bytesTotal As Long)
Socket.GetData data, vbString, bytesTotal
dData = Split(data, "YMSG" & Chr(0))
InfoData Socket
End Sub
Sub GetSetts()
On Error Resume Next
frmConfig.Check1.Value = GetSetting(App.EXEName, "OptCheck1", "OptCheck1")
frmConfig.Check2.Value = GetSetting(App.EXEName, "OptCheck2", "OptCheck2")
If frmConfig.Check2.Value = 0 Then frmConfig.Check2.Value = 1
frmConfig.Check3.Value = GetSetting(App.EXEName, "OptCheck3", "OptCheck3")
If frmConfig.Check3.Value = 0 Then frmConfig.Check3.Value = 1
frmConfig.Check4.Value = GetSetting(App.EXEName, "OptCheck4", "OptCheck4")
frmConfig.Check5.Value = GetSetting(App.EXEName, "OptCheck5", "OptCheck5")
If frmConfig.Check5.Value = 0 Then frmConfig.Check5.Value = 1
frmConfig.Check6.Value = GetSetting(App.EXEName, "OptCheck6", "OptCheck6")
frmConfig.Check7.Value = GetSetting(App.EXEName, "OptCheck7", "OptCheck7")
frmConfig.Check8.Value = GetSetting(App.EXEName, "OptCheck8", "OptCheck8")
If frmConfig.Check8.Value = 0 Then frmConfig.Check8.Value = 1
frmConfig.Check9.Value = GetSetting(App.EXEName, "OptCheck9", "OptCheck9")
frmConfig.Check10.Value = GetSetting(App.EXEName, "OptCheck10", "OptCheck10")
If frmConfig.Check10.Value = 0 Then frmConfig.Check10.Value = 1
frmConfig.Check11.Value = GetSetting(App.EXEName, "OptCheck11", "OptCheck11")
frmConfig.Check12.Value = GetSetting(App.EXEName, "OptCheck12", "OptCheck12")
If frmConfig.Check12.Value = 0 Then frmConfig.Check12.Value = 1
frmConfig.Text1 = GetSetting(App.EXEName, "Sound1", "Sound1")
frmConfig.Text2 = GetSetting(App.EXEName, "Sound2", "Sound2")
frmConfig.Text3 = GetSetting(App.EXEName, "Sound3", "Sound3")
If frmConfig.Text1 = "" Then frmConfig.Text1 = App.Path & "\Sounds\im.wav"
If frmConfig.Text2 = "" Then frmConfig.Text2 = App.Path & "\Sounds\chat.wav"
If frmConfig.Text3 = "" Then frmConfig.Text3 = App.Path & "\Sounds\buzz.wav"
LoadList App.Path & "\Ignored.txt", frmConfig.List2
LoadList App.Path & "\Filter.txt", frmConfig.List3
End Sub

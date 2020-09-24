VERSION 5.00
Begin VB.Form frmMenus 
   Caption         =   "Form1"
   ClientHeight    =   795
   ClientLeft      =   165
   ClientTop       =   750
   ClientWidth     =   4320
   Icon            =   "frmMenus.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   795
   ScaleWidth      =   4320
   StartUpPosition =   3  'Windows Default
   Begin VB.Image Image20 
      Height          =   270
      Left            =   2520
      Picture         =   "frmMenus.frx":030A
      Top             =   480
      Width           =   270
   End
   Begin VB.Image Image19 
      Height          =   270
      Left            =   2160
      Picture         =   "frmMenus.frx":1496
      Top             =   480
      Width           =   270
   End
   Begin VB.Image Image18 
      Height          =   270
      Left            =   1800
      Picture         =   "frmMenus.frx":1B37
      Top             =   480
      Width           =   270
   End
   Begin VB.Image Image17 
      Height          =   270
      Left            =   1440
      Picture         =   "frmMenus.frx":1F96
      Top             =   480
      Width           =   270
   End
   Begin VB.Image Image16 
      Height          =   270
      Left            =   1080
      Picture         =   "frmMenus.frx":228F
      Top             =   480
      Width           =   270
   End
   Begin VB.Image Image15 
      Height          =   270
      Left            =   720
      Picture         =   "frmMenus.frx":29A2
      Top             =   480
      Width           =   270
   End
   Begin VB.Image Image14 
      Height          =   270
      Left            =   360
      Picture         =   "frmMenus.frx":2BA0
      Top             =   480
      Width           =   270
   End
   Begin VB.Image Image13 
      Height          =   270
      Left            =   0
      Picture         =   "frmMenus.frx":330D
      Top             =   480
      Width           =   270
   End
   Begin VB.Image Image12 
      Height          =   270
      Left            =   3960
      Picture         =   "frmMenus.frx":3761
      Top             =   120
      Width           =   300
   End
   Begin VB.Image Image11 
      Height          =   270
      Left            =   3600
      Picture         =   "frmMenus.frx":4104
      Top             =   120
      Width           =   330
   End
   Begin VB.Image Image10 
      Height          =   270
      Left            =   3240
      Picture         =   "frmMenus.frx":4C28
      Top             =   120
      Width           =   270
   End
   Begin VB.Image Image9 
      Height          =   270
      Left            =   2880
      Picture         =   "frmMenus.frx":4E24
      Top             =   120
      Width           =   270
   End
   Begin VB.Image Image8 
      Height          =   270
      Left            =   2520
      Picture         =   "frmMenus.frx":53D1
      Top             =   120
      Width           =   330
   End
   Begin VB.Image Image7 
      Height          =   270
      Left            =   2160
      Picture         =   "frmMenus.frx":5D1E
      Top             =   120
      Width           =   270
   End
   Begin VB.Image Image6 
      Height          =   270
      Left            =   1800
      Picture         =   "frmMenus.frx":62B9
      Top             =   120
      Width           =   270
   End
   Begin VB.Image Image5 
      Height          =   270
      Left            =   1320
      Picture         =   "frmMenus.frx":67C3
      Top             =   120
      Width           =   510
   End
   Begin VB.Image Image4 
      Height          =   270
      Left            =   1080
      Picture         =   "frmMenus.frx":754E
      Top             =   120
      Width           =   270
   End
   Begin VB.Image Image3 
      Height          =   270
      Left            =   720
      Picture         =   "frmMenus.frx":7F95
      Top             =   120
      Width           =   270
   End
   Begin VB.Image Image2 
      Height          =   270
      Left            =   360
      Picture         =   "frmMenus.frx":815B
      Top             =   120
      Width           =   270
   End
   Begin VB.Image Image1 
      Height          =   270
      Left            =   0
      Picture         =   "frmMenus.frx":8784
      Top             =   120
      Width           =   270
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808080&
      Caption         =   "Label1"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Menu Menu1 
      Caption         =   "Buddy"
      Begin VB.Menu sendmsg 
         Caption         =   "Send Message"
      End
      Begin VB.Menu del 
         Caption         =   "Delete"
      End
      Begin VB.Menu joinuser 
         Caption         =   "Join User in Chat"
      End
      Begin VB.Menu showprofile 
         Caption         =   "View Profile"
      End
   End
   Begin VB.Menu Menu2 
      Caption         =   "Status"
      Begin VB.Menu available 
         Caption         =   "I'm Available"
      End
      Begin VB.Menu menuline1 
         Caption         =   "-"
      End
      Begin VB.Menu busy 
         Caption         =   "Busy"
      End
      Begin VB.Menu steppedout 
         Caption         =   "Stepped Out"
      End
      Begin VB.Menu brb 
         Caption         =   "Be Right Back"
      End
      Begin VB.Menu notatdesk 
         Caption         =   "Not at My Desk"
      End
      Begin VB.Menu onthephone 
         Caption         =   "On the Phone"
      End
      Begin VB.Menu menuline2 
         Caption         =   "-"
      End
      Begin VB.Menu cmdstatus 
         Caption         =   "Custom Status"
      End
      Begin VB.Menu custom 
         Caption         =   ""
      End
      Begin VB.Menu menuline3 
         Caption         =   "-"
      End
      Begin VB.Menu invisible 
         Caption         =   "Invisible"
      End
   End
   Begin VB.Menu Menu3 
      Caption         =   "CleverY"
      Begin VB.Menu showprog 
         Caption         =   "Show Clever Yahoo!"
      End
      Begin VB.Menu menulinelol 
         Caption         =   "-"
      End
      Begin VB.Menu changeuser 
         Caption         =   "Change User"
      End
      Begin VB.Menu dc 
         Caption         =   "Disconnect"
      End
      Begin VB.Menu menulinedontknow 
         Caption         =   "-"
      End
      Begin VB.Menu endtheprog 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmMenus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub available_Click()
frmLogin.Socket.SendData Header("04", "")
frmBuddy.Status1.Panels(1).Text = "I'm Available"
End Sub

Private Sub brb_Click()
frmLogin.Socket.SendData YStatus("Be Right Back", True)
frmBuddy.Status1.Panels(1).Text = "Be Right Back"
End Sub

Private Sub busy_Click()
frmLogin.Socket.SendData YStatus("Busy", True)
frmBuddy.Status1.Panels(1).Text = "Busy"
End Sub

Private Sub changeuser_Click()
Unload frmBuddy
frmLogin.Show
End Sub

Private Sub cmdstatus_Click()
frmCustom.Show
End Sub

Private Sub custom_Click()
frmLogin.Socket.SendData YStatus(Me.custom.Caption, False)
frmBuddy.Status1.Panels(1).Text = Me.custom.Caption
End Sub

Private Sub dc_Click()
frmLogin.Socket.Close
frmBuddy.Label1.Caption = "Sign In"
frmBuddy.Toolbar1.Visible = False
frmBuddy.TreeView1.Visible = False
frmBuddy.Label1.Visible = True
frmBuddy.Image1.Visible = True
frmBuddy.Label2.Visible = True
frmBuddy.Label3.Visible = True
frmBuddy.Label4.Visible = True
frmBuddy.Label1.FontUnderline = True
End Sub

Private Sub del_Click()
Dim Group As String
Group = Split(frmBuddy.TreeView1.SelectedItem.FullPath, "\")(0)
U2D = frmBuddy.TreeView1.SelectedItem.Key
U2D2 = frmBuddy.TreeView1.SelectedItem.Index
frmLogin.Socket.SendData DeleteFriend(Username, U2D, Group)
End Sub

Private Sub endtheprog_Click()
Dim Inp As String
Inp = MsgBox("You are Sure", vbYesNo, "Exit")
Select Case Inp
Case vbYes
frmLogin.cSysTray1.InTray = False
End
Case vbNo
frmBuddy.Show
End Select
End Sub

Private Sub Invisible_Click()
frmLogin.Socket.SendData InvI
frmBuddy.Status1.Panels(1).Text = "Invisible"
End Sub

Private Sub joinuser_Click()
Dim Usr As String
Usr = frmBuddy.TreeView1.SelectedItem.Key
frmLogin.Socket.SendData FollowUser(Username, Usr)
End Sub

Private Sub notatdesk_Click()
frmLogin.Socket.SendData YStatus("Not at My Desk", True)
frmBuddy.Status1.Panels(1).Text = "Not at My Desk"
End Sub

Private Sub onthephone_Click()
frmLogin.Socket.SendData YStatus("On the Phone", True)
frmBuddy.Status1.Panels(1).Text = "On the Phone"
End Sub

Private Sub sendmsg_Click()
Dim U As Integer
For U = 0 To frmProfiles.ProfileList.ListCount - 1
PM(OpenPMS + 1).Combo1.AddItem frmProfiles.ProfileList.List(U)
Next
PM(OpenPMS + 1).Combo1.Text = frmProfiles.ProfileList.List(0)
PM(OpenPMS + 1).Show
PM(OpenPMS + 1).Label4.Caption = frmBuddy.TreeView1.SelectedItem.Key
PM(OpenPMS + 1).Text3.Visible = False
PM(OpenPMS + 1).Label4.Visible = True
OpenPMS = OpenPMS + 1
End Sub

Private Sub showprofile_Click()
Shell "explorer http://profiles.yahoo.com/" & frmBuddy.TreeView1.SelectedItem.Key
End Sub

Private Sub showprog_Click()
frmBuddy.Show
End Sub

Private Sub steppedout_Click()
frmLogin.Socket.SendData YStatus("Stepped Out", True)
frmBuddy.Status1.Panels(1).Text = "Stepped Out"
End Sub

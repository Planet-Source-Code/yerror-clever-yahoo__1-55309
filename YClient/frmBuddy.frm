VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmBuddy 
   Caption         =   "Friends"
   ClientHeight    =   4395
   ClientLeft      =   165
   ClientTop       =   750
   ClientWidth     =   3300
   Icon            =   "frmBuddy.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   4395
   ScaleWidth      =   3300
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   4200
      Top             =   2280
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   3300
      _ExtentX        =   5821
      _ExtentY        =   1111
      ButtonWidth     =   1455
      ButtonHeight    =   1058
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Message"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Chat"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Webcam"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Add"
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   3480
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   34
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuddy.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuddy.frx":167E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuddy.frx":29F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuddy.frx":3D66
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   3960
      Top             =   720
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3600
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuddy.frx":50DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuddy.frx":562C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuddy.frx":573E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuddy.frx":5850
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuddy.frx":5962
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuddy.frx":5A74
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuddy.frx":5B88
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   3495
      Left            =   0
      TabIndex        =   0
      Top             =   630
      Visible         =   0   'False
      Width           =   3300
      _ExtentX        =   5821
      _ExtentY        =   6165
      _Version        =   393217
      Style           =   1
      ImageList       =   "ImageList1"
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.StatusBar Status1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   4140
      Width           =   3300
      _ExtentX        =   5821
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5900
            MinWidth        =   5900
            Text            =   "I'm Available"
            TextSave        =   "I'm Available"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Connecting to Yahoo!"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   840
      TabIndex        =   1
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Height          =   1815
      Left            =   0
      TabIndex        =   3
      Top             =   2280
      Width           =   3375
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   1800
      Width           =   3375
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   720
      Left            =   960
      Picture         =   "frmBuddy.frx":5C9C
      Top             =   720
      Width           =   1470
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Height          =   1815
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   3375
   End
   Begin VB.Menu menu1 
      Caption         =   "File"
      Begin VB.Menu ChangeUser 
         Caption         =   "Change User"
      End
      Begin VB.Menu Disconnect 
         Caption         =   "Disconnect"
         Shortcut        =   ^D
      End
      Begin VB.Menu menuline1 
         Caption         =   "-"
      End
      Begin VB.Menu editprofiles 
         Caption         =   "Edit Profiles"
      End
      Begin VB.Menu Configurations 
         Caption         =   "Preferences"
         Shortcut        =   ^P
      End
      Begin VB.Menu menuline2 
         Caption         =   "-"
      End
      Begin VB.Menu ex 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu menu2 
      Caption         =   "Info"
      Begin VB.Menu aboutinfo 
         Caption         =   "About Clever Y!"
      End
   End
End
Attribute VB_Name = "frmBuddy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub aboutinfo_Click()
frmAbout.Show
End Sub

Private Sub changeuser_Click()
Unload Me
frmLogin.Show
End Sub

Private Sub Configurations_Click()
frmConfig.Show
End Sub

Private Sub Disconnect_Click()
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

Private Sub editprofiles_Click()
Shell "Explorer http://edit.yahoo.com/config/eval_profile"
End Sub

Private Sub ex_Click()
Dim Inp As String
Inp = MsgBox("You are sure to Exit", vbYesNo, "Exit")
Select Case Inp
Case vbYes
End
End Select
End Sub

Private Sub Form_Load()
On Error GoTo err
Me.Width = GetSetting("Yah!Zone", "MeWidth", "MeWidth")
Me.Height = GetSetting("Yah!Zone", "MeHeight", "MeHeight")
err:
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Label1.ForeColor = vbBlack
End Sub

Private Sub Form_Resize()
SaveSetting "Yah!Zone", "MeWidth", "MeWidth", Me.Width
SaveSetting "Yah!Zone", "MeHeight", "MeHeight", Me.Height
On Error Resume Next
If Me.Width < 3420 Then
Me.Width = 3421
GoTo 1
End If
If Me.Height < 3315 Then
Me.Height = 3316
GoTo 1
End If
Label2.Width = Me.Width
Label3.Width = Me.Width
Label3.Height = Me.Height
Image1.Left = Me.ScaleWidth / 2 - 650
Label1.Left = Me.ScaleWidth / 2 - 800
Label4.Width = Me.Width
TreeView1.Top = Toolbar1.Height
TreeView1.Width = Me.Width - 100
Status1.Panels(1).Width = Me.Width - 100
TreeView1.Height = Me.Height - 1600
1
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Label1.ForeColor = vbBlack
End Sub

Private Sub Label1_Click()
If Label1.Caption = "Sign In" Then
Unload Me
frmLogin.Show
End If
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Label1.ForeColor = vbBlue
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Label1.ForeColor = vbBlack
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Label1.ForeColor = vbBlack
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Label1.ForeColor = vbBlack
End Sub

Private Sub refmenu_Click()
frmLogin.Socket.SendData ReF(Username)
End Sub

Private Sub Status1_PanelClick(ByVal Panel As MSComctlLib.Panel)
PopupMenu frmMenus.Menu2, , Status1.Left, Status1.Top
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
Case "Message"
Dim U As Integer
For U = 0 To frmProfiles.ProfileList.ListCount - 1
PM(OpenPMS).Combo1.AddItem frmProfiles.ProfileList.List(U)
Next
PM(OpenPMS).Combo1.Text = frmProfiles.ProfileList.List(0)
PM(OpenPMS).Label8.Caption = OpenPMS
PM(OpenPMS).Show
OpenPMS = OpenPMS + 1
Case "Chat"
frmChatLogin.Show
For x = 0 To frmProfiles.ProfileList.ListCount - 1
frmChatLogin.Combo1.AddItem frmProfiles.ProfileList.List(x)
Next
frmChatLogin.Combo1.Text = frmChatLogin.Combo1.List(0)
frmChatLogin.Combo2.Text = frmChatLogin.Combo2.List(0)
Case "Webcam"
frmWebCam.Show
Case "Add"
frmAddBuddy.Show
End Select
End Sub

Private Sub TreeView1_BeforeLabelEdit(Cancel As Integer)
Cancel = 1
End Sub


Private Sub TreeView1_DblClick()
If TreeView1.SelectedItem.Expanded = True Then
If TreeView1.SelectedItem.Image = 4 Then
TreeView1.SelectedItem.Image = 5
End If
End If

If TreeView1.SelectedItem.Expanded = False Then
If TreeView1.SelectedItem.Image = 5 Then
TreeView1.SelectedItem.Image = 4
End If
End If

If TreeView1.SelectedItem.Image = 1 Then
Dim U As Integer
For U = 0 To frmProfiles.ProfileList.ListCount - 1
PM(OpenPMS).Combo1.AddItem frmProfiles.ProfileList.List(U)
Next
PM(OpenPMS).Combo1.Text = frmProfiles.ProfileList.List(0)
PM(OpenPMS).Label8.Caption = OpenPMS
PM(OpenPMS).Show
PM(OpenPMS).Label4.Caption = TreeView1.SelectedItem.Key
PM(OpenPMS).Text3.Visible = False
PM(OpenPMS).Label4.Visible = True
OpenPMS = OpenPMS + 1
End If

If TreeView1.SelectedItem.Image = 2 Then
For U = 0 To frmProfiles.ProfileList.ListCount - 1
PM(OpenPMS).Combo1.AddItem frmProfiles.ProfileList.List(U)
Next
PM(OpenPMS).Combo1.Text = frmProfiles.ProfileList.List(0)
PM(OpenPMS).Label8.Caption = OpenPMS
PM(OpenPMS).Show
PM(OpenPMS).Label4.Caption = TreeView1.SelectedItem.Key
PM(OpenPMS).Text3.Visible = False
PM(OpenPMS).Label4.Visible = True
PM(OpenPMS).Tag = OpenPMS
OpenPMS = OpenPMS + 1
End If

If TreeView1.SelectedItem.Image = 6 Then
For U = 0 To frmProfiles.ProfileList.ListCount - 1
PM(OpenPMS).Combo1.AddItem frmProfiles.ProfileList.List(U)
Next
PM(OpenPMS).Combo1.Text = frmProfiles.ProfileList.List(0)
PM(OpenPMS).Label8.Caption = OpenPMS
PM(OpenPMS).Show
PM(OpenPMS).Label4.Caption = TreeView1.SelectedItem.Key
PM(OpenPMS).Text3.Visible = False
PM(OpenPMS).Label4.Visible = True
PM(OpenPMS).Tag = OpenPMS
OpenPMS = OpenPMS + 1
End If

If TreeView1.SelectedItem.Image = 7 Then
For U = 0 To frmProfiles.ProfileList.ListCount - 1
PM(OpenPMS).Combo1.AddItem frmProfiles.ProfileList.List(U)
Next
PM(OpenPMS).Combo1.Text = frmProfiles.ProfileList.List(0)
PM(OpenPMS).Label8.Caption = OpenPMS
PM(OpenPMS).Show
PM(OpenPMS).Label4.Caption = TreeView1.SelectedItem.Key
PM(OpenPMS).Text3.Visible = False
PM(OpenPMS).Label4.Visible = True
PM(OpenPMS).Tag = OpenPMS
OpenPMS = OpenPMS + 1
End If

End Sub

Private Sub TreeView1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If TreeView1.SelectedItem.Image = 1 Then
If Button = 2 Then
PopupMenu frmMenus.Menu1
End If
End If
If TreeView1.SelectedItem.Image = 2 Then
If Button = 2 Then
PopupMenu frmMenus.Menu1
End If
End If
End Sub


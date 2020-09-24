VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmPM 
   Caption         =   "Instant Message"
   ClientHeight    =   3330
   ClientLeft      =   165
   ClientTop       =   750
   ClientWidth     =   4575
   Icon            =   "frmPM.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3330
   ScaleWidth      =   4575
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   1111
      ButtonWidth     =   1455
      ButtonHeight    =   1058
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Add"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Webcam"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Ignore"
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4560
      Top             =   720
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
            Picture         =   "frmPM.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPM.frx":167E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPM.frx":29F2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   14
      Top             =   2400
      Width           =   3600
   End
   Begin RichTextLib.RichTextBox Text2 
      Height          =   855
      Left            =   120
      TabIndex        =   13
      Top             =   1080
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   1508
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      OLEDragMode     =   0
      TextRTF         =   $"frmPM.frx":3D66
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
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   2880
      TabIndex        =   4
      Top             =   720
      Width           =   1575
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   720
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   720
      Width           =   1575
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   2040
      Width           =   2175
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   3720
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   2040
      Width           =   750
   End
   Begin Project1.chameleonButton chameleonButton1 
      Default         =   -1  'True
      Height          =   615
      Left            =   3720
      TabIndex        =   3
      Top             =   2400
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1085
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
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   3075
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   8229
            MinWidth        =   8229
         EndProperty
      EndProperty
   End
   Begin VB.Label Label8 
      Height          =   255
      Left            =   0
      TabIndex        =   16
      Top             =   3120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   270
      Left            =   1200
      Picture         =   "frmPM.frx":3DDD
      Top             =   2040
      Width           =   270
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "From:"
      Height          =   255
      Left            =   0
      TabIndex        =   12
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "To:"
      Height          =   255
      Left            =   2520
      TabIndex        =   11
      Top             =   720
      Width           =   375
   End
   Begin VB.Label Label4 
      Height          =   255
      Left            =   2880
      TabIndex        =   10
      Top             =   720
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label2 
      Height          =   255
      Left            =   720
      TabIndex        =   9
      Top             =   720
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label5 
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
      Left            =   120
      TabIndex        =   8
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label Label6 
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
      Left            =   480
      TabIndex        =   7
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label Label7 
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
      Left            =   840
      TabIndex        =   6
      Top             =   2040
      Width           =   255
   End
   Begin VB.Menu Menu1 
      Caption         =   "Friend"
      Begin VB.Menu buz 
         Caption         =   "Buzz"
         Shortcut        =   ^G
      End
      Begin VB.Menu addfriend 
         Caption         =   "Add as Friend"
         Shortcut        =   ^A
      End
      Begin VB.Menu viewcam 
         Caption         =   "View Webcam"
      End
      Begin VB.Menu ignore 
         Caption         =   "Ignore"
      End
   End
End
Attribute VB_Name = "frmPM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Before(400) As Boolean
Dim Y(400) As Boolean

Private Sub addfriend_Click()
frmAddBuddy.Show
frmAddBuddy.Text1 = Label4.Caption
End Sub

Private Sub buz_Click()
If Combo1.Visible = True Then
frmLogin.Socket.SendData SendPM(Combo1, Label4.Caption, "<ding>")
Else
frmLogin.Socket.SendData SendPM(Label2.Caption, Label4.Caption, "<ding>")
End If
ProcessDing Text2, frmPM
End Sub

Private Sub chameleonButton1_Click()
Unload frmPMSmileys
If Text1 = "" Then
MsgBox "Please enter a Message", vbOKOnly, "Error"
Exit Sub
End If
If Text3.Visible = True Then
Label4.Caption = Text3
Text3.Visible = False
End If
If Label2.Visible = False Then
Label2.Caption = Combo1
Combo1.Visible = False
Label2.Visible = True
End If
For x = 0 To frmConfig.List2.ListCount
If LCase(Label4.Caption) = LCase(frmConfig.List2.List(x)) Then
MsgBox "The User is on the Ignore List", vbOKOnly, "Error"
Exit Sub
End If
Next
Text2.SelStart = Len(Text2)
Text2.SelBold = True
Text2.SelColor = vbBlack
Text2.SelFontSize = 10
Text2.SelFontName = "Arial"
Text2.SelUnderline = False
Text2.SelItalic = False
Text2.SelText = Label2.Caption & ": "
Text2.SelBold = False
Text2.SelColor = vbBlack
frmLogin.Socket.SendData SendPM(Label2.Caption, Label4.Caption, Text1)
Text2.SelStart = Len(Text2)
ProcessText Text1, Text2
Text1 = ""
Text2.SelText = vbCrLf
Text2.SelStart = Len(Text2)
Me.Caption = Label4.Caption & " - " & Label2.Caption
1
End Sub

Private Sub Combo1_Change()
Unload frmPMSmileys
End Sub

Private Sub Form_Click()
Unload frmPMSmileys
End Sub

Private Sub Form_GotFocus()
Unload frmPMSmileys
End Sub

Private Sub Form_Load()
On Error Resume Next
For x = 0 To Screen.FontCount
Combo2.AddItem Screen.FontS(x)
Next
Combo2.Text = "Arial"
For x = 1 To 32
Combo3.AddItem x
Next
Combo3.Text = "10"
End Sub



Private Sub Form_LostFocus()
Me.Tag = "Lost"
End Sub

Private Sub Form_Resize()
Image1.Top = Label6.Top
On Error Resume Next
If Me.Height < 3420 Then
Me.Height = 3421
GoTo 1
End If
If Me.Width < 4695 Then
Me.Width = 4696
GoTo 1
End If
Combo1.Width = Me.ScaleWidth / 3
Label2.Width = Combo1.Width
Label3.Left = Combo1.Width + 910
Label4.Width = Text3.Width
Label4.Left = Label3.Left + 400
Text3.Left = Label4.Left
Combo2.Width = Me.Width - 2600
Combo3.Left = Combo2.Width + 1600
Text2.Height = Me.Height - 3100
Combo2.Top = Text2.Height + 1100
Combo3.Top = Combo2.Top
Label5.Top = Combo2.Top
Label6.Top = Combo2.Top
Label7.Top = Combo2.Top
Text1.Top = Label7.Top + 380
chameleonButton1.Top = Text1.Top
Text2.Width = Me.Width - 350
Text1.Width = Me.Width - 1200
Text3.Width = Me.ScaleWidth / 3 + 100
Label4.Width = Text3.Width
chameleonButton1.Left = Text1.Width + 200
StatusBar1.Panels(1).Width = Me.Width
1
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
OpenPMS = OpenPMS - 1
Unload Me
End Sub

Private Sub ignore_Click()
Dim Inp As String
Inp = MsgBox("Ignore this User", vbYesNo, "Ignore")
Select Case Inp
Case vbYes
frmConfig.List2.AddItem Label4.Caption
End Select
End Sub

Private Sub Image1_Click()
frmPMSmileys.Show
frmPMSmileys.Label1.Caption = Me.Label8.Caption
frmPMSmileys.Tag = Me.Tag
frmPMSmileys.Left = Me.Left + 1250
frmPMSmileys.Top = Me.Top + 2040
End Sub

Private Sub Label1_Click()
Unload frmPMSmileys
End Sub

Private Sub Label2_Click()
Unload frmPMSmileys
End Sub

Private Sub Label3_Click()
Unload frmPMSmileys
End Sub

Private Sub Label4_Click()
Unload frmPMSmileys
End Sub

Private Sub Label5_Click()
Unload frmPMSmileys
End Sub

Private Sub Label6_Click()
Unload frmPMSmileys
End Sub

Private Sub Label7_Click()
Unload frmPMSmileys
End Sub

Private Sub StatusBar1_PanelClick(ByVal Panel As MSComctlLib.Panel)
Unload frmPMSmileys
End Sub

Private Sub Text1_Change()
On Error Resume Next
If Len(Text1) = "1" Then
If Combo1.Visible = False Then
frmLogin.Socket.SendData Typing(Label2.Caption, Label4.Caption)
End If
End If
End Sub

Private Sub Text1_Click()
Unload frmPMSmileys
End Sub

Private Sub Text3_Change()
Unload frmPMSmileys
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Unload frmPMSmileys
On Error Resume Next
Select Case Button.Key
Case "Add"
frmAddBuddy.Show
frmAddBuddy.Text1 = Label4.Caption
Case "Webcam"
frmWebCam.Show
frmWebCam.Text1 = Label4.Caption
Case "Ignore"
Dim Inp As String
Inp = MsgBox("Ignore this User", vbYesNo, "Ignore")
Select Case Inp
Case vbYes
frmConfig.List2.AddItem Label4.Caption
End Select
End Select
End Sub


Private Sub Text2_Click()
Unload frmPMSmileys
On Error Resume Next
ClickedURL Text2
End Sub

Private Sub viewcam_Click()
frmWebCam.Show
frmWebCam.Text1 = Label4.Caption
End Sub

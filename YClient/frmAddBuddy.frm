VERSION 5.00
Begin VB.Form frmAddBuddy 
   Caption         =   "Add Friend"
   ClientHeight    =   2490
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   3855
   Icon            =   "frmAddBuddy.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2490
   ScaleWidth      =   3855
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2535
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3855
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1560
         TabIndex        =   4
         Top             =   120
         Width           =   2175
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1560
         TabIndex        =   3
         Text            =   "Combo1"
         Top             =   480
         Width           =   2175
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   840
         Width           =   2175
      End
      Begin VB.TextBox Text2 
         Height          =   735
         Left            =   1560
         MaxLength       =   126
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   1200
         Width           =   2175
      End
      Begin Project1.chameleonButton chameleonButton1 
         Default         =   -1  'True
         Height          =   315
         Left            =   360
         TabIndex        =   5
         Top             =   2040
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Add"
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
         Left            =   2040
         TabIndex        =   6
         Top             =   2040
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
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Yahoo! User ID:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Group:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Your ID:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Message:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1200
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   2655
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   3855
      Begin Project1.chameleonButton chameleonButton3 
         Height          =   345
         Left            =   240
         TabIndex        =   12
         Top             =   2040
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   609
         BTYPE           =   3
         TX              =   "Finish"
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
      Begin VB.Label Label5 
         Height          =   975
         Left            =   120
         TabIndex        =   13
         Top             =   120
         Width           =   3615
      End
   End
End
Attribute VB_Name = "frmAddBuddy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chameleonButton1_Click()
frmLogin.Socket.SendData AddFriend(Combo2, Text1, Combo1, Text2)
chameleonButton1.Enabled = False
End Sub

Private Sub chameleonButton2_Click()
Unload Me
End Sub

Private Sub chameleonButton3_Click()
Unload Me
End Sub

Private Sub Form_Load()
For X = 0 To frmProfiles.FriendList.ListCount - 1
If frmProfiles.FriendList.List(X) = "" Then GoTo 1
Combo1.AddItem frmProfiles.FriendList.List(X)
1
Next
Combo1.Text = Combo1.List(0)
For X = 0 To frmProfiles.ProfileList.ListCount - 1
If frmProfiles.ProfileList.List(X) = "" Then GoTo 2
Combo2.AddItem frmProfiles.ProfileList.List(X)
2
Next
Combo2.Text = Combo2.List(0)
End Sub


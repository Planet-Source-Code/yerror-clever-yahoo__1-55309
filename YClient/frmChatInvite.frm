VERSION 5.00
Begin VB.Form frmChatInvite 
   Caption         =   "Chat Invite"
   ClientHeight    =   1935
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4575
   Icon            =   "frmChatInvite.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1935
   ScaleWidth      =   4575
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   480
      Width           =   3135
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   840
      Width           =   3135
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   3135
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   1200
      Width           =   3135
   End
   Begin Project1.chameleonButton chameleonButton1 
      Default         =   -1  'True
      Height          =   315
      Left            =   1320
      TabIndex        =   8
      Top             =   1560
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Join"
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
      Left            =   3000
      TabIndex        =   9
      Top             =   1560
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
      Caption         =   "From:"
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Room:"
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "To:"
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Message:"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   1200
      Width           =   1335
   End
End
Attribute VB_Name = "frmChatInvite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chameleonButton1_Click()
frmLogin.Socket.SendData JoinRoom(Text3)
Pause 1
frmLogin.Socket.SendData GoToRoom(Text3, Room)
frmChatLoad.Show
Unload Me
End Sub

Private Sub chameleonButton3_Click()
Unload Me
End Sub

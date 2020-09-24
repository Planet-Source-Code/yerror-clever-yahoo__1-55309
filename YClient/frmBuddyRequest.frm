VERSION 5.00
Begin VB.Form frmBuddyRequest 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   4575
   Icon            =   "frmBuddyRequest.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   4575
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   975
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   480
      Width           =   4335
   End
   Begin Project1.chameleonButton chameleonButton1 
      Height          =   315
      Left            =   1680
      TabIndex        =   3
      Top             =   1560
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Decline"
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
      Default         =   -1  'True
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Allow"
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
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Blah added you in his List"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "frmBuddyRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DeclineMSG As String

Private Sub chameleonButton1_Click()
DeclineMSG = InputBox("Enter your Message Here", "Deny", "No Thanks!")
frmLogin.Socket.SendData DenyBudd(Username, Label2.Caption, DeclineMSG)
Unload Me
End Sub

Private Sub chameleonButton3_Click()
Unload Me
End Sub


Private Sub Form_Load()
DeclineMSG = "No Thanks!"
End Sub

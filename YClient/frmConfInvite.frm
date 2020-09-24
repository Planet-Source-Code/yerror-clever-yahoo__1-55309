VERSION 5.00
Begin VB.Form frmConfInvite 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Conference Invite"
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   4560
   Icon            =   "frmConfInvite.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   1200
      Width           =   3135
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   120
      Width           =   3135
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   840
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   480
      Width           =   3135
   End
   Begin Project1.chameleonButton chameleonButton3 
      Height          =   315
      Left            =   1680
      TabIndex        =   6
      Top             =   1560
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Close"
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
   Begin Project1.chameleonButton chameleonButton1 
      Height          =   345
      Left            =   120
      TabIndex        =   7
      Top             =   2640
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   609
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
   Begin Project1.chameleonButton chameleonButton2 
      Default         =   -1  'True
      Height          =   315
      Left            =   120
      TabIndex        =   10
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
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Message:"
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "To:"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Conference:"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "From:"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   1335
   End
End
Attribute VB_Name = "frmConfInvite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chameleonButton2_Click()
frmLogin.Socket.SendData AcceptConf(Text3, Text2, Text1)
Pause 1
frmLogin.Socket.SendData AcceptConf2(Text3, Text2)
End Sub

Private Sub chameleonButton3_Click()
Unload Me
End Sub

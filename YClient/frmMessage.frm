VERSION 5.00
Begin VB.Form frmMessage 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1080
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3435
   Icon            =   "frmMessage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1080
   ScaleWidth      =   3435
   StartUpPosition =   3  'Windows Default
   Begin Project1.chameleonButton chameleonButton3 
      Default         =   -1  'True
      Height          =   315
      Left            =   1920
      TabIndex        =   1
      Top             =   720
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "OK"
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
   Begin VB.Image Image1 
      Height          =   360
      Left            =   0
      Picture         =   "frmMessage.frx":030A
      Top             =   120
      Width           =   360
   End
   Begin VB.Label Label1 
      Caption         =   "This message has been sent. If the receipent(s) are currently offline, they will see the message when they next login."
      Height          =   735
      Left            =   480
      TabIndex        =   0
      Top             =   0
      Width           =   3015
   End
End
Attribute VB_Name = "frmMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chameleonButton3_Click()
Unload Me
End Sub


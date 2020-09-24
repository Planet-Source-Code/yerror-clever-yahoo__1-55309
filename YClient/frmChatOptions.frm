VERSION 5.00
Begin VB.Form frmChatOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Chat Options"
   ClientHeight    =   1785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4080
   Icon            =   "frmChatOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   4080
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3855
      Begin VB.CheckBox Check4 
         Caption         =   "Filter Colors"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   3615
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Filter Bad Words"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   3615
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Use only Font 10 and Face Arial"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   3615
      End
      Begin Project1.chameleonButton chameleonButton1 
         Default         =   -1  'True
         Height          =   315
         Left            =   360
         TabIndex        =   4
         Top             =   1200
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Apply"
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
         TabIndex        =   5
         Top             =   1200
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
   End
End
Attribute VB_Name = "frmChatOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chameleonButton1_Click()
Me.Hide
End Sub

Private Sub chameleonButton2_Click()
Me.Hide
End Sub

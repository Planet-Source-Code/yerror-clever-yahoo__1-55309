VERSION 5.00
Begin VB.Form frmCustom 
   Caption         =   "Form1"
   ClientHeight    =   1830
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3375
   LinkTopic       =   "Form1"
   ScaleHeight     =   1830
   ScaleWidth      =   3375
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Show Busy Icon"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   120
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   240
      Width           =   3135
   End
   Begin Project1.chameleonButton chameleonButton1 
      Default         =   -1  'True
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   1440
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
   Begin Project1.chameleonButton chameleonButton2 
      Height          =   315
      Left            =   1800
      TabIndex        =   4
      Top             =   1440
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
      Caption         =   "256 Left"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   1095
   End
End
Attribute VB_Name = "frmCustom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chameleonButton1_Click()
If Check1.Value = 1 Then
frmLogin.Socket.SendData YStatus(Text1, True)
Else
frmLogin.Socket.SendData YStatus(Text1, False)
End If
frmMenus.custom.Caption = Text1
SaveSetting App.EXEName, "1", "1", frmMenus.custom.Caption
Hide
End Sub

Private Sub chameleonButton2_Click()
Hide
End Sub

Private Sub Text1_Change()
Dim i As Long
i = 255
i = i - Len(Text1)
Label1.Caption = i & " Left"
End Sub

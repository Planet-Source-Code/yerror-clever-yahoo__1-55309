VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmOfflines 
   Caption         =   "Offline Messages"
   ClientHeight    =   3165
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4560
   Icon            =   "frmOfflines.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3165
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox rchOff 
      Height          =   855
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   1508
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmOfflines.frx":030A
   End
   Begin VB.ListBox List2 
      Height          =   255
      Left            =   5400
      TabIndex        =   2
      Top             =   840
      Width           =   2775
   End
   Begin VB.ListBox List1 
      Height          =   1230
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4335
   End
   Begin Project1.chameleonButton chameleonButton1 
      Height          =   315
      Left            =   3000
      TabIndex        =   5
      Top             =   2760
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
   Begin Project1.chameleonButton chameleonButton3 
      Height          =   315
      Left            =   1440
      TabIndex        =   6
      Top             =   2760
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Reply"
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
      Caption         =   "Message:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "From:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   1335
   End
End
Attribute VB_Name = "frmOfflines"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chameleonButton1_Click()
List1.Clear
List2.Clear
Text1 = ""
Unload Me
End Sub

Private Sub chameleonButton3_Click()
Dim U As Integer
For U = 0 To frmProfiles.ProfileList.ListCount - 1
PM(OpenPMS).Combo1.AddItem frmProfiles.ProfileList.List(U)
Next
PM(OpenPMS).Combo1.Text = frmProfiles.ProfileList.List(0)
PM(OpenPMS).Show
PM(OpenPMS).Label4.Caption = List1.Text
PM(OpenPMS).Text3.Visible = False
PM(OpenPMS).Label4.Visible = True
OpenPMS = OpenPMS + 1
End Sub

Private Sub List1_Click()
On Error GoTo 2
rchOff.Text = ""
rchOff.SelBold = True
rchOff.SelColor = vbBlue
rchOff.SelFontSize = 10
rchOff.SelFontName = "Arial"
rchOff.SelUnderline = False
rchOff.SelItalic = False
rchOff.SelText = List1 & ": "
rchOff.SelColor = vbBlack
rchOff.SelBold = False
ProcessText List2.List(List1.ListIndex), rchOff
2:
End Sub

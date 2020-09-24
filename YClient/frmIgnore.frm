VERSION 5.00
Begin VB.Form frmIgnore 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ignore List"
   ClientHeight    =   2100
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4095
   Icon            =   "frmIgnore.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2100
   ScaleWidth      =   4095
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   1425
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3855
   End
   Begin Project1.chameleonButton chameleonButton1 
      Default         =   -1  'True
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   1680
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
      Left            =   1680
      TabIndex        =   2
      Top             =   1680
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Remove"
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
Attribute VB_Name = "frmIgnore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chameleonButton1_Click()
Dim Inp As String
Inp = InputBox("Enter the Name to Ignore", "Ignore", "")
List1.AddItem Inp
SaveList App.Path & "\Ignored.txt", List1
End Sub

Private Sub chameleonButton2_Click()
Dim S As String
S = List1
If S = "" Then GoTo 1
List1.RemoveItem List1.ListIndex
SaveList App.Path & "\Ignored.txt", List1
1:
End Sub

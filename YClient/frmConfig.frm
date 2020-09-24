VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmConfig 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Preferences"
   ClientHeight    =   3540
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5145
   Icon            =   "frmConfig.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   5145
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame6 
      Caption         =   "Word Filter"
      Height          =   3015
      Left            =   1920
      TabIndex        =   35
      Top             =   0
      Visible         =   0   'False
      Width           =   3135
      Begin VB.ListBox List3 
         Height          =   2205
         Left            =   120
         TabIndex        =   36
         Top             =   240
         Width           =   2895
      End
      Begin Project1.chameleonButton chameleonButton11 
         Height          =   315
         Left            =   120
         TabIndex        =   37
         Top             =   2520
         Width           =   1335
         _ExtentX        =   2355
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
      Begin Project1.chameleonButton chameleonButton12 
         Height          =   315
         Left            =   1680
         TabIndex        =   38
         Top             =   2520
         Width           =   1335
         _ExtentX        =   2355
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
   Begin VB.Frame Frame5 
      Caption         =   "Ignore List"
      Height          =   3015
      Left            =   1920
      TabIndex        =   31
      Top             =   0
      Visible         =   0   'False
      Width           =   3135
      Begin VB.ListBox List2 
         Height          =   2205
         Left            =   120
         TabIndex        =   32
         Top             =   240
         Width           =   2895
      End
      Begin Project1.chameleonButton chameleonButton9 
         Height          =   315
         Left            =   120
         TabIndex        =   33
         Top             =   2520
         Width           =   1335
         _ExtentX        =   2355
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
      Begin Project1.chameleonButton chameleonButton10 
         Height          =   315
         Left            =   1680
         TabIndex        =   34
         Top             =   2520
         Width           =   1335
         _ExtentX        =   2355
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
   Begin MSComDlg.CommonDialog CMD1 
      Left            =   120
      Top             =   3480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame4 
      Caption         =   "Sounds"
      Height          =   3015
      Left            =   1920
      TabIndex        =   15
      Top             =   0
      Visible         =   0   'False
      Width           =   3135
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   720
         TabIndex        =   21
         Top             =   2280
         Width           =   2295
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   720
         TabIndex        =   19
         Top             =   1320
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   720
         TabIndex        =   17
         Top             =   360
         Width           =   2295
      End
      Begin Project1.chameleonButton chameleonButton3 
         Height          =   255
         Left            =   720
         TabIndex        =   25
         Top             =   720
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         BTYPE           =   3
         TX              =   "Browse"
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
      Begin Project1.chameleonButton chameleonButton4 
         Height          =   255
         Left            =   1920
         TabIndex        =   26
         Top             =   720
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         BTYPE           =   3
         TX              =   "Play"
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
      Begin Project1.chameleonButton chameleonButton5 
         Height          =   255
         Left            =   720
         TabIndex        =   27
         Top             =   1680
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         BTYPE           =   3
         TX              =   "Browse"
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
      Begin Project1.chameleonButton chameleonButton6 
         Height          =   255
         Left            =   1920
         TabIndex        =   28
         Top             =   1680
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         BTYPE           =   3
         TX              =   "Play"
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
      Begin Project1.chameleonButton chameleonButton7 
         Height          =   255
         Left            =   720
         TabIndex        =   29
         Top             =   2640
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         BTYPE           =   3
         TX              =   "Browse"
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
      Begin Project1.chameleonButton chameleonButton8 
         Height          =   255
         Left            =   1920
         TabIndex        =   30
         Top             =   2640
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         BTYPE           =   3
         TX              =   "Play"
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
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Buzz:"
         Height          =   255
         Left            =   0
         TabIndex        =   20
         Top             =   2280
         Width           =   615
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Chat:"
         Height          =   255
         Left            =   0
         TabIndex        =   18
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "IM:"
         Height          =   255
         Left            =   0
         TabIndex        =   16
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "IM"
      Height          =   3015
      Left            =   1920
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   3135
      Begin VB.CheckBox Check12 
         Caption         =   "Sound on Buzz"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   1800
         Value           =   1  'Checked
         Width           =   2775
      End
      Begin VB.CheckBox Check11 
         Caption         =   "Filter Colors"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1080
         Width           =   2055
      End
      Begin VB.CheckBox Check10 
         Caption         =   "Filter Bad Words"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   720
         Value           =   1  'Checked
         Width           =   2535
      End
      Begin VB.CheckBox Check9 
         Caption         =   "Use only Font 10 and Face Arial"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   2775
      End
      Begin VB.CheckBox Check8 
         Caption         =   "Sound on new Message"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1440
         Value           =   1  'Checked
         Width           =   2775
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Chat"
      Height          =   3015
      Left            =   1920
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   3135
      Begin VB.CheckBox Check7 
         Caption         =   "Sound on new Message"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1440
         Width           =   2775
      End
      Begin VB.CheckBox Check6 
         Caption         =   "Use only Font 10 and Face Arial"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   2775
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Filter Bad Words"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Value           =   1  'Checked
         Width           =   2535
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Filter Colors"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1080
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "General"
      Height          =   3015
      Left            =   1920
      TabIndex        =   1
      Top             =   0
      Width           =   3135
      Begin VB.CheckBox Check3 
         Caption         =   "Start Voice when User Joins Rooms"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Value           =   1  'Checked
         Width           =   2895
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Enable PM's"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Value           =   1  'Checked
         Width           =   2535
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Startup when Windows Starts"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   2775
      End
   End
   Begin VB.ListBox List1 
      Height          =   3375
      ItemData        =   "frmConfig.frx":030A
      Left            =   120
      List            =   "frmConfig.frx":0320
      TabIndex        =   0
      Top             =   90
      Width           =   1695
   End
   Begin Project1.chameleonButton chameleonButton1 
      Default         =   -1  'True
      Height          =   315
      Left            =   1920
      TabIndex        =   23
      Top             =   3120
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
      Left            =   3600
      TabIndex        =   24
      Top             =   3120
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
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chameleonButton1_Click()
SaveSetting App.EXEName, "OptCheck1", "OptCheck1", Check1.Value
SaveSetting App.EXEName, "OptCheck2", "OptCheck2", Check2.Value
SaveSetting App.EXEName, "OptCheck3", "OptCheck3", Check3.Value
SaveSetting App.EXEName, "OptCheck4", "OptCheck4", Check4.Value
SaveSetting App.EXEName, "OptCheck5", "OptCheck5", Check5.Value
SaveSetting App.EXEName, "OptCheck6", "OptCheck6", Check6.Value
SaveSetting App.EXEName, "OptCheck7", "OptCheck7", Check7.Value
SaveSetting App.EXEName, "OptCheck8", "OptCheck8", Check8.Value
SaveSetting App.EXEName, "OptCheck9", "OptCheck9", Check9.Value
SaveSetting App.EXEName, "OptCheck10", "OptCheck10", Check10.Value
SaveSetting App.EXEName, "OptCheck11", "OptCheck11", Check11.Value
SaveSetting App.EXEName, "OptCheck12", "OptCheck12", Check12.Value
SaveSetting App.EXEName, "Sound1", "Sound1", Text1
SaveSetting App.EXEName, "Sound2", "Sound2", Text2
SaveSetting App.EXEName, "Sound3", "Sound3", Text3
Hide
End Sub

Private Sub chameleonButton10_Click()
Dim S As String
S = List2
If S = "" Then GoTo 1
List2.RemoveItem List2.ListIndex
SaveList App.Path & "\Ignored.txt", List2
1:
End Sub

Private Sub chameleonButton11_Click()
Dim Inp As String
Inp = InputBox("Enter a Word to Filter", "Filter", "")
List3.AddItem Inp
SaveList App.Path & "\Filter.txt", List3
End Sub

Private Sub chameleonButton12_Click()
Dim S As String
S = List3
If S = "" Then GoTo 1
List3.RemoveItem List3.ListIndex
SaveList App.Path & "\Filter.txt", List3
1:
End Sub

Private Sub chameleonButton2_Click()
Hide
End Sub

Private Sub chameleonButton3_Click()
CMD1.DialogTitle = "Open Sound File"
CMD1.Filter = "*.WAV | *.wav"
CMD1.ShowOpen
Text1 = CMD1.FileName
CMD1.FileName = ""
End Sub

Private Sub chameleonButton4_Click()
PlayWav Text1
End Sub

Private Sub chameleonButton5_Click()
CMD1.DialogTitle = "Open Sound File"
CMD1.Filter = "*.WAV | *.wav"
CMD1.ShowOpen
Text2 = CMD1.FileName
CMD1.FileName = ""
End Sub

Private Sub chameleonButton6_Click()
PlayWav Text2
End Sub

Private Sub chameleonButton7_Click()
CMD1.DialogTitle = "Open Sound File"
CMD1.Filter = "*.WAV | *.wav"
CMD1.ShowOpen
Text3 = CMD1.FileName
CMD1.FileName = ""
End Sub

Private Sub chameleonButton8_Click()
PlayWav Text3
End Sub

Private Sub chameleonButton9_Click()
Dim Inp As String
Inp = InputBox("Enter the Name to Ignore", "Ignore", "")
List2.AddItem Inp
SaveList App.Path & "\Ignored.txt", List2
End Sub

Private Sub Form_Unload(Cancel As Integer)
Cancel = 1
Hide
End Sub

Private Sub List1_Click()
If List1.Text = "General" Then
Frame1.Visible = True
Frame2.Visible = False
Frame3.Visible = False
Frame4.Visible = False
Frame5.Visible = False
Frame6.Visible = False
End If
If List1.Text = "Sounds" Then
Frame1.Visible = False
Frame2.Visible = False
Frame3.Visible = False
Frame4.Visible = True
Frame5.Visible = False
Frame6.Visible = False
End If
If List1.Text = "Chat" Then
Frame1.Visible = False
Frame2.Visible = True
Frame3.Visible = False
Frame4.Visible = False
Frame5.Visible = False
Frame6.Visible = False
End If
If List1.Text = "IM" Then
Frame1.Visible = False
Frame2.Visible = False
Frame3.Visible = True
Frame4.Visible = False
Frame5.Visible = False
End If
If List1.Text = "Ignored" Then
Frame1.Visible = False
Frame2.Visible = False
Frame3.Visible = False
Frame4.Visible = False
Frame5.Visible = True
Frame6.Visible = False
End If
If List1.Text = "Filter" Then
Frame1.Visible = False
Frame2.Visible = False
Frame3.Visible = False
Frame4.Visible = False
Frame5.Visible = False
Frame6.Visible = True
End If
End Sub

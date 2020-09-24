VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{9D392231-AE8E-11D4-8FD3-00D0B7730277}#1.0#0"; "ywcvwr.dll"
Begin VB.Form frmWebCam 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "WebCam"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   2475
   Icon            =   "frmWebCam.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   2475
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   2640
      Width           =   2475
      _ExtentX        =   4366
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4498
            MinWidth        =   4498
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Top             =   1920
      Width           =   2480
   End
   Begin YWCVWRLibCtl.WcViewer WcViewer1 
      Height          =   1935
      Left            =   0
      OleObjectBlob   =   "frmWebCam.frx":030A
      TabIndex        =   0
      Top             =   0
      Width           =   2535
   End
   Begin Project1.chameleonButton chameleonButton1 
      Default         =   -1  'True
      Height          =   315
      Left            =   80
      TabIndex        =   4
      Top             =   2280
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "View"
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
      Left            =   1320
      TabIndex        =   3
      Top             =   2280
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Pause"
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
Attribute VB_Name = "frmWebCam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chameleonButton1_Click()
WcViewer1.TargetName = Text1
WcViewer1.ISP = "1"
WcViewer1.Token = Username
WcViewer1.CountryCode = "US"
WcViewer1.AppName = 3
WcViewer1.Age = 18
WcViewer1.Receive
End Sub

Private Sub chameleonButton2_Click()
If chameleonButton2.Caption = "Pause" Then
WcViewer1.Abort
chameleonButton2.Caption = "Resume"
Else
WcViewer1.Receive
chameleonButton2.Caption = "Pause"
End If
End Sub

Private Sub WcViewer1_OnConnectionStatusChanged(ByVal iStatusCode As Long)
If iStatusCode = 1 Then StatusBar1.Panels(1).Text = "User Declined Request"
If iStatusCode = 2 Then StatusBar1.Panels(1).Text = "Connecting to Cam"
End Sub

Private Sub WcViewer1_OnReceivedImage(ByVal iImageLen As Long, ByVal tTimeStamp As Long)
StatusBar1.Panels(1).Text = "Last Image: " & Time
End Sub


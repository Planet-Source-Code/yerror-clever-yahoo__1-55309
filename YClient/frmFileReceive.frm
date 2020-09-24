VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmReceive 
   Caption         =   "Download File"
   ClientHeight    =   2460
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2460
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar P 
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   1080
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
   End
   Begin MSWinsockLib.Winsock WS 
      Left            =   360
      Top             =   4200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   1080
      Top             =   4320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Download to..."
   End
   Begin VB.CommandButton cmdDL 
      Caption         =   "Download"
      Height          =   285
      Left            =   720
      TabIndex        =   1
      Top             =   3000
      Width           =   975
   End
   Begin VB.TextBox txtAd 
      Height          =   285
      Left            =   1800
      TabIndex        =   0
      Top             =   2880
      Width           =   2175
   End
   Begin Project1.chameleonButton chameleonButton1 
      Height          =   345
      Left            =   3600
      TabIndex        =   7
      Top             =   120
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "Download"
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
      Height          =   345
      Left            =   3600
      TabIndex        =   12
      Top             =   600
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   609
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
   Begin VB.Line Line4 
      BorderWidth     =   3
      X1              =   4560
      X2              =   4560
      Y1              =   1320
      Y2              =   1080
   End
   Begin VB.Line Line3 
      BorderWidth     =   3
      X1              =   120
      X2              =   4560
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   120
      X2              =   120
      Y1              =   1080
      Y2              =   1320
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   120
      X2              =   4560
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label Label4 
      Height          =   255
      Left            =   1200
      TabIndex        =   11
      Top             =   480
      Width           =   2295
   End
   Begin VB.Label Label3 
      Height          =   255
      Left            =   1200
      TabIndex        =   10
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Filename:"
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "From:"
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblS 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   1200
      TabIndex        =   6
      Top             =   1800
      Width           =   3375
   End
   Begin VB.Label lblDLFrom 
      BackStyle       =   0  'Transparent
      Caption         =   "Download From:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label lblDL 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   960
      TabIndex        =   4
      Top             =   1440
      Width           =   3615
   End
   Begin VB.Label lblSave 
      BackStyle       =   0  'Transparent
      Caption         =   "Saving to:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label lblStat 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Specify a file"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   975
   End
End
Attribute VB_Name = "frmReceive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Here is my HTTP File Downloader

Dim File As String
Dim Ad As String
Dim X As Integer
Dim I As Integer
Dim File2 As String
Dim Rec As String
Dim Num As Integer
Dim FSize As String
Dim Temp2 As String
Dim Str As String

Private Sub cmdBrowse_Click()
    CD.Filter = "All Files (*.*)|*.*"   'sets the file types
    CD.ShowSave 'shows the save dialog
    txtS.Text = CD.FileName
End Sub

Private Sub chameleonButton1_Click()
    Ad = txtAd.Text 'here is the long messy and confusing process of
    'seperating the address, file path, and file name
    'make sure ur only using /'s or else this wont work
    If InStr(1, Ad, "http://") Then 'remover "http://" if it is present
        Ad = Right(Ad, Len(Ad) - 7)
    End If
    
    Do Until X = Len(Ad)    'scans for the first /
        DoEvents
        X = X + 1
        If Mid(Ad, X, 1) = "/" Then
            File = Mid(Ad, X, Len(Ad))  'gets the end...the file path
            Ad = Mid(Ad, 1, X - 1)  'gets the address
            Exit Do 'stops loop
        End If
    Loop
    File2 = File
    If InStr(2, File2, "/") Then    'this will go to the final / to get just the file name
        Do Until InStr(2, File2, "/") = False
            Do
                DoEvents
                I = I + 1
                If Mid(File2, I, 1) = "/" Then  'when it finds a /
                    File2 = Mid(File2, I, Len(File2))   'file2 will contain from the current / until the end..
                    Exit Do
                End If
            Loop
        Loop
    End If
    
    CD.Filter = "All Files (*.*)|*.*"
    CD.FileName = Right(File2, Len(File2) - 1)
    CD.ShowSave
    File2 = CD.FileName

    lblS.Caption = File2
    
    WS.Close    'closes winsock
    WS.Connect Ad, 80   'connects to the address on port 80
    lblStat.Caption = "Connecting to: " & Ad
End Sub


Private Sub chameleonButton2_Click()
Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    WS.Close    'closes winsock before u exit the app
End Sub

Private Sub WS_Connect()
Dim Header As String
    lblStat.Caption = "Connected, requesting " & File
    'this prepares our header
    Header = Header & "GET " & File & " HTTP/1.1" & vbCrLf
    Header = Header & "Host: " & WS.RemoteHostIP & vbCrLf
    Header = Header & "User-Agent: Nullific Downloader\1.0" & vbCrLf
    Header = Header & "Accept: */*" & vbCrLf
    WS.SendData Header & vbCrLf 'sends the header
End Sub

Private Sub WS_DataArrival(ByVal bytesTotal As Long)
Dim Data As String
    WS.GetData Data 'gets the data that is sent

    Num = 0
    
If InStr(1, Data, "HTTP") Then  'processes the servers response for the file size
    Str = Data
    Do
        'Num = Num + 1
        If Data = "" Then

        End If
        Data = Right(Data, Len(Data) - 1)   'front of header
        If Mid(Data, 1, 15) = "Content-Length:" Then    'when the front is...
                Do  'for the file size...
                    Num = Num + 1
                    If Mid(Data, Num, 2) = vbCrLf Then  'finds the vbcrlf, telling us that the line with the file size has ended
                        Temp2 = Mid(Data, 1, Num)   'isolates the line with the size
                        FSize = Mid(Temp2, 16, Len(Temp2))  'removes "Content-Length: " and leaves only the file size
                        P.Max = FSize
                        Exit Do
                    End If
                Loop
            Exit Do
        End If
    Loop
    
    Num = 0
    
    Do
        Num = Num + 1
            If Mid(Str, Num, 4) = (vbCrLf & vbCrLf) Then    'at the end of the header may be the beginning of the file, seperated by two vbcrlfs
                Str = Mid(Str, Num + 4, Len(Str))   'when they are found
                P.Value = P.Value + Len(Str)
                Rec = Len(Str)
                lblDL.Caption = Rec & "/" & FSize
                
                Open File2 For Binary As #2 'writes to the file
                    Put #2, , Str
                Close #2
             Exit Do
             End If
        Loop
    Else
    
    Open File2 For Binary As #2
        P.Value = P.Value + Len(Data)
        Rec = Int(Rec) + Len(Data)  'adds to how many bytes have been recieved
        lblDL.Caption = Rec & "/" & FSize
        Temp = (LOF(2) + 1)
        If Temp = 0 Then
            Put #2, , Data
            Else
            Put #2, Temp, Data
        End If
    Close #2
    
    If P.Value = P.Max Then
    I = 0
    End If
End If

End Sub

Private Sub WS_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    I = 0
End Sub

Private Sub WS_SendComplete()
I = 0
End Sub

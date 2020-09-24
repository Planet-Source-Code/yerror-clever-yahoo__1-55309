VERSION 5.00
Begin VB.Form frmStatus 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   225
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   17280
   LinkTopic       =   "Form1"
   ScaleHeight     =   225
   ScaleWidth      =   17280
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   4000
      Left            =   0
      Top             =   360
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   17295
   End
End
Attribute VB_Name = "frmStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
On Error Resume Next
Dim I As Integer
For X = 0 To 100
I = I + 5
Trans Me.hwnd, 254 - I
Pause 0.01
Next
Unload Me
End Sub

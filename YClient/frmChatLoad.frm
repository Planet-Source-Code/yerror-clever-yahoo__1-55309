VERSION 5.00
Begin VB.Form frmChatLoad 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   465
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5040
   Icon            =   "frmChatLoad.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   465
   ScaleWidth      =   5040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   0
      Picture         =   "frmChatLoad.frx":030A
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   1
      Top             =   0
      Width           =   480
   End
   Begin VB.Label Label1 
      Caption         =   "Loading Chat Room Please wait this can take a few Minutes"
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "frmChatLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Unload(Cancel As Integer)
If Me.Label1.Caption = "Loading Chat Room Please wait this can take a few Minutes" Then
If frmChat.Visible = False Then
Cancel = 1
Else
Cancel = 0
End If
End If

End Sub

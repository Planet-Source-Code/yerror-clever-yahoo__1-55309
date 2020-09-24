VERSION 5.00
Begin VB.Form frmProfiles 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   90
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1470
   LinkTopic       =   "Form1"
   ScaleHeight     =   90
   ScaleWidth      =   1470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.ListBox ProfileList 
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   240
      Width           =   1455
   End
   Begin VB.ListBox FriendList 
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1455
   End
End
Attribute VB_Name = "frmProfiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

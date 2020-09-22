VERSION 5.00
Begin VB.Form E_Chat 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Private Encrypted Network Conversation"
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9945
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "EncryptedChat.frx":0000
   LinkTopic       =   "TCPIP_Chat"
   MaxButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   9945
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   330
      Left            =   8940
      TabIndex        =   4
      Top             =   4110
      Width           =   900
   End
   Begin VB.Frame Frame1 
      Caption         =   "Conversation"
      ForeColor       =   &H000000C0&
      Height          =   3525
      Left            =   90
      TabIndex        =   5
      Top             =   495
      Width           =   9765
      Begin VB.ListBox lstChat 
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   3150
         ItemData        =   "EncryptedChat.frx":038A
         Left            =   120
         List            =   "EncryptedChat.frx":038C
         TabIndex        =   3
         Top             =   270
         Width           =   9540
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Send"
      Default         =   -1  'True
      Height          =   330
      Left            =   8940
      TabIndex        =   2
      Top             =   60
      Width           =   900
   End
   Begin VB.TextBox Text1 
      ForeColor       =   &H0000C000&
      Height          =   300
      Left            =   735
      TabIndex        =   1
      Top             =   75
      Width           =   8190
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "2001 Matthew Hall"
      ForeColor       =   &H00808080&
      Height          =   240
      Left            =   2610
      TabIndex        =   7
      Top             =   4185
      Width           =   1395
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   135
      Picture         =   "EncryptedChat.frx":038E
      Top             =   4185
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Say this:"
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   75
      TabIndex        =   0
      Top             =   120
      Width           =   660
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Interworks IntelliCrypt 7.0 "
      ForeColor       =   &H00404040&
      Height          =   270
      Left            =   105
      TabIndex        =   6
      Top             =   4170
      Width           =   2265
   End
End
Attribute VB_Name = "E_Chat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
SendChatLine Text1
If lstChat.ListCount > 100 Then lstChat.Clear: lstChat.AddItem "{Chat History Cleared}"
lstChat.AddItem "Me: " + Text1
lstChat.Selected(lstChat.ListCount - 1) = True
Text1 = ""
End Sub

Public Sub AddLine(TextToAdd As String)
If lstChat.ListCount > 100 Then lstChat.Clear: lstChat.AddItem "{Chat History Cleared}"
lstChat.AddItem TextToAdd
lstChat.Selected(lstChat.ListCount - 1) = True
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

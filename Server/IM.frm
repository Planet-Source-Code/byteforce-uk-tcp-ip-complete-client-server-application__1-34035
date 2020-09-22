VERSION 5.00
Begin VB.Form IM 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Instant Message"
   ClientHeight    =   2910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "IM.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2145
      TabIndex        =   7
      Top             =   3525
      Width           =   1110
   End
   Begin VB.CommandButton SendBack 
      Caption         =   "Send"
      Height          =   375
      Left            =   3345
      TabIndex        =   6
      Top             =   3525
      Width           =   1110
   End
   Begin VB.Frame fraReply 
      Caption         =   "Reply"
      ForeColor       =   &H00008000&
      Height          =   1155
      Left            =   180
      TabIndex        =   5
      Top             =   2325
      Visible         =   0   'False
      Width           =   4305
      Begin VB.TextBox Text2 
         ForeColor       =   &H00008000&
         Height          =   825
         Left            =   195
         MultiLine       =   -1  'True
         TabIndex        =   8
         Top             =   225
         Width           =   3915
      End
   End
   Begin VB.CommandButton Cancel2 
      Caption         =   "Dont Reply"
      Height          =   345
      Left            =   2145
      TabIndex        =   3
      Top             =   2445
      Width           =   1095
   End
   Begin VB.CommandButton Reply 
      Caption         =   "Reply"
      Default         =   -1  'True
      Height          =   345
      Left            =   3345
      TabIndex        =   2
      Top             =   2445
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Message"
      ForeColor       =   &H000040C0&
      Height          =   1530
      Left            =   180
      TabIndex        =   1
      Top             =   750
      Width           =   4290
      Begin VB.TextBox Text1 
         ForeColor       =   &H000080FF&
         Height          =   915
         Left            =   180
         Locked          =   -1  'True
         MaxLength       =   126
         MultiLine       =   -1  'True
         TabIndex        =   4
         Text            =   "IM.frx":0CCA
         Top             =   360
         Width           =   3915
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Instant Message Recieved"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   195
      Width           =   3870
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   150
      Picture         =   "IM.frx":0CDC
      Top             =   135
      Width           =   480
   End
End
Attribute VB_Name = "IM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cancel2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Reply.Default = True
PROBas.PlayWav App.Path + "\IM.wav"
End Sub

Private Sub Reply_Click()
Me.Height = 4305
fraReply.Visible = True
SendBack.Default = True
End Sub

Private Sub CancelButton_Click()
Unload Me
End Sub

Private Sub SendBack_Click()
Server.SendStrData Text2
Unload Me
End Sub

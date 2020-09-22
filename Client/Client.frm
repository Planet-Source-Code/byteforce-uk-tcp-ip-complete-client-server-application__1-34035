VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Client 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "General TCP\IP Client"
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4665
   FillColor       =   &H80000012&
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
   Icon            =   "Client.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   4665
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock Winsock 
      Left            =   1980
      Top             =   3300
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close Client"
      Height          =   360
      Left            =   3420
      TabIndex        =   8
      Top             =   3255
      Width           =   1125
   End
   Begin VB.Frame Frame1 
      Caption         =   "Client"
      Height          =   3105
      Left            =   120
      TabIndex        =   9
      Top             =   45
      Width           =   4425
      Begin VB.CommandButton cmdAddressBook 
         Caption         =   "Address Book"
         Height          =   690
         Left            =   3090
         MouseIcon       =   "Client.frx":0CCA
         MousePointer    =   99  'Custom
         Picture         =   "Client.frx":0FD4
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   285
         Width           =   1200
      End
      Begin VB.TextBox txtRemotePort 
         Height          =   285
         Left            =   1815
         TabIndex        =   2
         Text            =   "1029"
         Top             =   675
         Width           =   1200
      End
      Begin VB.Frame fraConnection 
         Caption         =   "Server 127.0.0.1:1029"
         ForeColor       =   &H00008000&
         Height          =   1425
         Left            =   135
         TabIndex        =   12
         Top             =   1545
         Visible         =   0   'False
         Width           =   4155
         Begin VB.TextBox txtArg 
            ForeColor       =   &H0000C000&
            Height          =   270
            Left            =   165
            TabIndex        =   6
            Top             =   990
            Width           =   2070
         End
         Begin VB.CommandButton cmdSend 
            Caption         =   "Send Command"
            Height          =   375
            Left            =   2280
            TabIndex        =   7
            Top             =   945
            Width           =   1650
         End
         Begin VB.ComboBox cboCommand 
            Appearance      =   0  'Flat
            ForeColor       =   &H000000C0&
            Height          =   315
            ItemData        =   "Client.frx":1C9E
            Left            =   165
            List            =   "Client.frx":1CC0
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   555
            Width           =   3750
         End
         Begin VB.Label Label3 
            Caption         =   "Choose a command to carry out on the server..."
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   165
            TabIndex        =   13
            Top             =   270
            Width           =   3705
         End
      End
      Begin VB.CommandButton cmdConnect 
         Caption         =   "Connect to server..."
         Default         =   -1  'True
         Height          =   360
         Left            =   195
         TabIndex        =   4
         Top             =   1050
         Width           =   4020
      End
      Begin VB.TextBox ServerHandle 
         Height          =   285
         Left            =   1815
         TabIndex        =   1
         Top             =   285
         Width           =   1230
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Server Port:"
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   1
         Left            =   150
         TabIndex        =   14
         Top             =   705
         Width           =   1590
      End
      Begin VB.Label lblNoConnect 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "There is no current connection"
         ForeColor       =   &H000000FF&
         Height          =   210
         Left            =   1095
         TabIndex        =   11
         Top             =   1665
         Width           =   2340
      End
      Begin VB.Label Label2 
         Caption         =   "Server IP \ Computer:"
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   0
         Left            =   150
         TabIndex        =   10
         Top             =   315
         Width           =   1650
      End
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "2001 Matthew Hall"
      ForeColor       =   &H00808080&
      Height          =   240
      Left            =   60
      TabIndex        =   15
      Top             =   3450
      Width           =   1395
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   60
      TabIndex        =   0
      Top             =   -240
      Width           =   660
   End
End
Attribute VB_Name = "Client"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Planet Source Code Submission
'=============================
'Author Name:   Matt Hall
'D.O.B:         29 DEC 84
'Mug Shot:      www.faceparty.com\cyberdude69b
'Contact:       matt_hall44@ hotmail.com
'
'Date Started   Unknown
'Date Finished: Unknown
'Total Hours:   10.5
'
'Project Title: General TCP\IP Client
'Version:       2.0
'Credits:       PROBas.bas by Unknown author
'               Some icons taken from the Microsoft(R) Windows(R)
'               Operating System.
'
'Description:   Allows application of TCP\IP network communication via two seperate
'               applications.. You can Query Current Windows Username, Execute EXE file,
'               Toggle Hide Server, Query IP Address (for DNS resolution), Send Instant
'               Message, Windows ShellExec and RSA Skipjack Encrypted Chat. It is very easy
'               to implement more features into the project, just take a quick delve into the
'               code and all will be clear!

'Assumes:       Microsoft Winsock Control [Included]
'               Microsoft Windows 9x\ME\NT\2000\XP Operating System
'
'Notes:         You can find settings for TCP IP Client\Server in XSERV.INI in the Server Folder
'
'               Got problems or queries about this submission? E-Post me at matt_hall44@hotmail.com
'
'               *********************************************************************************
'               Although every step is taken to prevent it;
'
'               I CANNOT BE HELD RESPONSIBLE FOR ANY DATA LOSS OR CORRUPTION TO YOUR COMPUTER,
'               CAUSED DIRECTLY OR INDIRECTLY BY USAGE OF THE FILES PROVIDED IN THIS ARCHIVE SET.
'               *********************************************************************************
Dim COK As Integer

Private Sub cmdAddressBook_Click()
Addresses.Show
End Sub

Private Sub cmdConnect_Click()
On Error GoTo 99
If Left(cmdConnect.Caption, 1) = "T" Then GoTo 34 'Caption = (T)erminate connection
If ServerHandle = "" Then Exit Sub
If txtRemotePort = "" Then Exit Sub
If IsNumeric(txtRemotePort) = False Then Exit Sub
If Winsock.State <> sckClosed Then Winsock.Close
Winsock.Connect ServerHandle, txtRemotePort
fraConnection.Caption = "Connecting to server " + ServerHandle + ":" & Winsock.RemotePort
cmdConnect.Caption = "Terminate connection"
COK = 0
Exit Sub

34 'Terminate code
Winsock.SendData "4"
For I = 1 To 10000
DoEvents
Next
Winsock.Close
Call Winsock_Close
Exit Sub

99 'Error connecting
MsgBox "Could not connect to server!", 16
cmdConnect.Caption = "Connect to server..."
End Sub

Private Sub cmdSend_Click()
'Send unformatted data
If cboCommand.ListIndex = 8 Then Winsock.SendData txtArg: Exit Sub

'Send private encrypted chat text
If cboCommand.ListIndex = 9 Then
If txtArg = "" Then txtArg = "{Opened Chat Session}"
E_Chat.Show: E_Chat.AddLine txtArg
SendChatLine txtArg
Exit Sub
End If

'Send nominal formatted data
Winsock.SendData cboCommand.ListIndex + 1 & txtArg
End Sub

Private Sub Command1_Click()
Client.Enabled = False
Command1.Enabled = False
On Error Resume Next
If Winsock.State <> 0 Then
Winsock.SendData "4"
For I = 1 To 10000
DoEvents
Next
End If
End
End Sub

Private Sub Form_Load()
On Error Resume Next
cboCommand.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
Command1_Click
End Sub

Private Sub Winsock_Close()
If COK = 1 Then Exit Sub
COK = 1
If E_Chat.Visible = True Then Unload E_Chat
MsgBox "The connection has been terminated", 48
cmdConnect.Caption = "Connect to server..."
fraConnection.Visible = False
End Sub

Private Sub Winsock_Connect()
fraConnection.Caption = "Connected to server " + ServerHandle + ":" & Winsock.RemotePort
fraConnection.Visible = True
End Sub

Private Sub Winsock_DataArrival(ByVal bytesTotal As Long)
Dim p As String
Winsock.GetData p
If Left(p, 1) = "C" Then 'Start Encrypted Chat Session
E_Chat.Show
E_Chat.AddLine PROBas.IntelliCrypt_DeCrypt(Right(p, Len(p) - 1))
Else
MsgBox p, 64, "Server TCP\IP Return Data"
End If
End Sub

Private Sub Winsock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
MsgBox Description, 16, "ActiveX Winsock Error"
Winsock.Close
Winsock_Close
End Sub

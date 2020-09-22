VERSION 5.00
Begin VB.Form Addresses 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TCP\IP IP Address Book"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7110
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Addresses.frx":0000
   LinkTopic       =   "TCPIPAddressBook"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   7110
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   360
      Left            =   5565
      TabIndex        =   3
      Top             =   675
      Width           =   1395
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "Import"
      Default         =   -1  'True
      Height          =   360
      Left            =   5565
      TabIndex        =   2
      Top             =   210
      Width           =   1395
   End
   Begin VB.Frame Frame2 
      Caption         =   "Add \ Remove Entries"
      Height          =   1770
      Left            =   120
      TabIndex        =   9
      Top             =   2520
      Width           =   5325
      Begin VB.CommandButton cmdKillEntry 
         Caption         =   " Remove  Selected"
         Height          =   1275
         Left            =   3780
         TabIndex        =   8
         Top             =   360
         Width           =   1365
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add Entry"
         Height          =   375
         Left            =   105
         TabIndex        =   7
         Top             =   1290
         Width           =   3480
      End
      Begin VB.TextBox txtNewServerPort 
         Height          =   285
         Left            =   1215
         TabIndex        =   6
         Top             =   975
         Width           =   2355
      End
      Begin VB.TextBox txtNewServerName 
         Height          =   285
         Left            =   1215
         TabIndex        =   5
         Top             =   630
         Width           =   2355
      End
      Begin VB.TextBox txtNewRecordName 
         Height          =   285
         Left            =   1215
         TabIndex        =   4
         Top             =   285
         Width           =   2355
      End
      Begin VB.Label Label2 
         Caption         =   "Server Port:"
         Height          =   240
         Index           =   2
         Left            =   120
         TabIndex        =   12
         Top             =   990
         Width           =   1050
      End
      Begin VB.Label Label2 
         Caption         =   "Server Name:"
         Height          =   240
         Index           =   1
         Left            =   120
         TabIndex        =   11
         Top             =   645
         Width           =   1050
      End
      Begin VB.Label Label2 
         Caption         =   "Record Name:"
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   300
         Width           =   1050
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Contents"
      Height          =   2280
      Left            =   120
      TabIndex        =   0
      Top             =   105
      Width           =   5325
      Begin VB.ListBox lstEntries 
         Appearance      =   0  'Flat
         ForeColor       =   &H0000C000&
         Height          =   1980
         ItemData        =   "Addresses.frx":0CCA
         Left            =   105
         List            =   "Addresses.frx":0CCC
         TabIndex        =   1
         Top             =   225
         Width           =   5085
      End
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "2001 Matthew Hall"
      ForeColor       =   &H00808080&
      Height          =   240
      Left            =   5640
      TabIndex        =   13
      Top             =   4155
      Width           =   1395
   End
End
Attribute VB_Name = "Addresses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Entry_IP(200) As Variant
Dim Entry_PORT(200) As Integer

Private Sub cmdAdd_Click()
'Check for foul input
If txtNewRecordName = "" Then Exit Sub
If txtNewServerName = "" Then Exit Sub
If txtNewServerPort = "" Then Exit Sub
If IsNumeric(txtNewServerPort) = False Then Exit Sub

'Add the record

'See what record is available
For i = 1 To 200 'Max = 200
    'See if record name exists for this entry
    recname = INIGetSetting("entry" & i, "Alias", App.Path + "\Addresses.ini")
    If Not recname = "" Then GoTo nxta 'entry in use; skip this enquiry
    
    ix = ix + 1 'array identifier
    Entry_IP(ix) = txtNewServerName
    INISaveSetting txtNewServerName, "entry" & i, "IP", App.Path + "\Addresses.ini"
    Entry_PORT(ix) = txtNewServerPort
    INISaveSetting txtNewServerPort, "entry" & i, "Port", App.Path + "\Addresses.ini"
    INISaveSetting txtNewRecordName, "entry" & i, "Alias", App.Path + "\Addresses.ini"
    
    lstEntries.AddItem txtNewRecordName
    MsgBox "Your entry was added to the address book", vbInformation: Exit For
    cmdImport.Enabled = True
    cmdImport.Default = True
    
nxta:
Next
If ix = 0 Then MsgBox "Address book full. Please delete one or more entries to add new ones.", 16: Exit Sub
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdImport_Click()
If lstEntries = "" Then Exit Sub
Client.txtRemotePort = Entry_PORT(lstEntries.ListIndex + 1)
Client.ServerHandle = Entry_IP(lstEntries.ListIndex + 1)
Unload Me
End Sub

Private Sub cmdKillEntry_Click()
If lstEntries = "" Then Exit Sub
'Remove the address from the file
INISaveSetting "", "entry" & lstEntries.ListIndex + 1, "Alias", App.Path + "\Addresses.ini"
INISaveSetting "", "entry" & lstEntries.ListIndex + 1, "IP", App.Path + "\Addresses.ini"
INISaveSetting "", "entry" & lstEntries.ListIndex + 1, "Port", App.Path + "\Addresses.ini"
lstEntries.Clear
Form_Load
MsgBox "The entry was deleted.", vbInformation
End Sub


Private Sub Form_Load()
cmdImport.Enabled = False
'Load the address book from file

ix = 0
For i = 1 To 200 'Max = 200
    'See if record name exists for this entry
    recname = INIGetSetting("entry" & i, "Alias", App.Path + "\Addresses.ini")
    If recname = "" Then GoTo nxt 'No entry; skip this enquiry
    
    ix = ix + 1 'array identifier
    Entry_IP(ix) = INIGetSetting("entry" & i, "IP", App.Path + "\Addresses.ini")
    Entry_PORT(ix) = INIGetSetting("entry" & i, "Port", App.Path + "\Addresses.ini")
    lstEntries.AddItem recname
    
nxt:
Next

If Not ix = 0 Then cmdImport.Enabled = True
End Sub

Private Sub lstEntries_DblClick()
If lstEntries = "" Then Exit Sub
cmdImport_Click
End Sub

Private Sub txtNewServerPort_Change()
cmdAdd.Default = True
End Sub

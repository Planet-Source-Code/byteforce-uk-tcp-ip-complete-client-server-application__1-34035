Attribute VB_Name = "PROBas"
'Attribute VB_Name = "OpenFile32"

'########################################
'#    C O L D  F U S I O N Z . N E T    #
'#We are the programmers of your future.#
'#          ---- PRO.BAS ----           #
'#(some code was already made but just  #
'#  put in for convenience while other  #
'#     code was made by us =])          #
'########################################
'-----------------------------------------
'This BAS Has Been Categorized So You Can
'Find Stuff Easily...!!!
'EXAMPLE: If your looking for a function
'dealing with forms then look under anything
'beginning with forms..Wanna encrypt something?
'Look under Encrypt..How about some shell functions
'like moving directories and files around look
'under shell... And so on.....
'-----------------------------------------
'DECOMMENT THE BELOW PHRASE IF USING CLSOPENSAVE
'Global CommonDialog As New clsOpenSave
Global Encrypt_CiperPercent%

Public Const REG_SZ = 1
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const CB_FINDSTRING = &H14C
Public Const CB_FINDSTRINGEXACT = &H158
Public Declare Function RegCloseKey Lib _
"advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegCreateKey Lib _
"advapi32" Alias "RegCreateKeyA" (ByVal _
hKey As Long, ByVal lpszSubKey As String, _
phkResult As Long) As Long
Public Declare Function RegSetValueEx Lib _
"advapi32" Alias "RegSetValueExA" (ByVal _
hKey As Long, ByVal lpszValueName As String, _
ByVal dwReserved As Long, ByVal fdwType As _
Long, lpbData As Any, ByVal cbData As Long) As Long

Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Declare Function GetDesktopWindow& Lib "user32" ()
Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Sub releaseCapture Lib "user32" Alias "ReleaseCapture" ()
Declare Function SendMessageByString& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String)
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function GetWindowsDirectory& Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long)
Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Declare Function GetWindow& Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long)
Declare Function GetClassName& Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long)
Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Const LB_FINDSTRING = &H18F
Public Const LB_FINDSTRINGEXACT = &H1A2
Public Const WM_LBUTTONDBLCLICK = &H203
Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Any, ByVal fuWinIni As Long) As Long
Declare Function CloseWindow Lib "user32" (ByVal hwnd As Long) As Long
Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Public Const CC_ANYCOLOR = &H100
Public Const CC_ENABLEHOOK = &H10
Public Const CC_ENABLETEMPLATE = &H20
Public Const CC_ENABLETEMPLATEHANDLE = &H40
Public Const CC_FULLOPEN = &H2
Public Const CC_INTERIORS = 128
Public Const CC_PREVENTFULLOPEN = &H4
Public Const CC_RGBINIT = &H1
Public Const CC_SHOWHELP = &H8
Public Const CC_SOLIDCOLOR = &H80
Public Const CC_WIDESTYLED = 64
Public Const CC_WIDE = 16

Public Const SW_ERASE = &H4
Public Const WM_RBUTTONUP = &H205
Public Const WM_MOUSEMOVE = &H200
Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
Public Const GW_HWNDPREV = 3
Public Const GW_MAX = 5
Public Const GW_OWNER = 4
Public Const SW_MAXIMIZE = 3
Public Const SW_MINIMIZE = 6
Public Const SW_HIDE = 0
Public Const SW_RESTORE = 9
Public Const SW_SHOW = 5
Public Const SW_SHOWDEFAULT = 10
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWMINNOACTIVE = 7
Public Const SW_SHOWNOACTIVATE = 4
Public Const SW_SHOWNORMAL = 1
Public Const WM_QUIT = &H12
Public Const WM_DESTROY = &H2
Public Const WM_DDE_FIRST = &H3E0
Public Const WM_DDE_TERMINATE = (WM_DDE_FIRST + 1)
Public Const WM_SHOWWINDOW = &H18
Public Const WM_SETFOCUS = &H7

Public Const WM_CLOSE = &H10
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_LBUTTONDBLCLK = &H203
Public Type CHOOSECOLOR
  lStructSize As Long
  hwndOwner As Long
  hInstance As Long
  rgbResult As Long
  lpCustColors As String
  flags As Long
  lCustData As Long
  lpfnHook As Long
  lpTemplateName As String
End Type
Public Declare Function ShowColour Lib _
"comdlg32.dll" Alias "ChooseColorA" _
(pChoosecolor As CHOOSECOLOR) As Long
Public Const WM_UNDO = &H304
Public Const WM_ACTIVATE = &H6
Public Const WM_SETTEXT = &HC
Public Const WM_CHAR = &H102
Public Const GW_CHILD = 5
Private Type OPENFILENAME
lStructSize As Long
hwndOwner As Long
hInstance As Long
lpstrFilter As String
lpstrCustomFilter As String
nMaxCustFilter As Long
nFilterIndex As Long
lpstrFile As String
nMaxFile As Long
lpstrFileTitle As String
nMaxFileTitle As Long
lpstrInitialDir As String
lpstrTitle As String
flags As Long
nFileOffset As Integer
nFileExtension As Integer
lpstrDefExt As String
lCustData As Long
lpfnHook As Long
lpTemplateName As String
End Type
Public Type filetype

Extension As String
ProperName As String
FullName As String
ContentType As String
IconPath As String
IconIndex As Integer
End Type
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
#If Win32 Then
Public Declare Function sndPlaySound& Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long)
#Else
Public Declare Function sndPlaySound% Lib "mmsystem.dll" (ByVallpszSoundName As String, ByVal uFlags As Integer)
#End If
Global Const SND_SYNC = &H0
Global Const SND_ASYNC = &H1
Global Const SND_NODEFAULT = &H2
Global Const SND_LOOP = &H8
Global Const SND_NOSTOP = &H10
Public CIAsystray As NOTIFYICONDATA
Public Type NOTIFYICONDATA
cbSize As Long
hwnd As Long
uId As Long
uFlags As Long
ucallbackMessage As Long
hIcon As Long
szTip As String * 64
End Type
Public Const SW_NORMAL = 1
Public Declare Function GetCursorPos Lib "user32" _
(lpPoint As PointAPI) As Long
Public Pnt As PointAPI
'These values MUST be public
Public OldX As Long
Public OldY As Long
Public NewX As Long
Public NewY As Long
Public Type PointAPI
X As Long
Y As Long
End Type
'This Const determines the total timeout value in minutes
Global Const GFM_STANDARD = 0
Global Const GFM_RAISED = 1
Global Const GFM_SUNKEN = 2
' Control Shadow Styles
Global Const GFM_BACKSHADOW = 1
Global Const GFM_DROPSHADOW = 2
' Color constants
Global Const BOX_WHITE& = &HFFFFFF
Global Const BOX_LIGHTGRAY& = &HC0C0C0
Global Const BOX_DARKGRAY& = &H808080
Global Const BOX_BLACK& = &H0&
Public Const SPI_SCREENSAVERRUNNING = 97
Public Type SHFILEOPSTRUCT
hwnd As Long
wFunc As Long
pFrom As String
pTo As String
fFlags As Integer
fAborted As Boolean
hNameMaps As Long
sProgress As String
End Type
Public Type BrowseInfo
hwndOwner As Long
pIDLRoot As Long
pszDisplayName As Long
lpszTitle As Long
ulFlags As Long
lpfnCallback As Long
lParam As Long
iImage As Long
End Type
Global FileDestination As String
Public Const BIF_RETURNONLYFSDIRS = 1
Public Const MAX_PATH = 260
Public Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Public Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Public Const FO_COPY = &H2
Public Const FO_DELETE = &H3
Public Const FO_MOVE = &H1
Public Const FO_RENAME = &H4
Public Const FOF_ALLOWUNDO = &H40
Public Const FOF_CONFIRMMOUSE = &H2
Public Const FOF_FILESONLY = &H80                  '  on *.*, do only files
Public Const FOF_MULTIDESTFILES = &H1
Public Const FOF_NOCONFIRMATION = &H10             '  Don't prompt the user.
Public Const FOF_NOCONFIRMMKDIR = &H200            '  don't confirm making any needed dirs
Public Const FOF_RENAMEONCOLLISION = &H8
Public Const FOF_SILENT = &H4                      '  don't create progress/report
Public Const FOF_SIMPLEPROGRESS = &H100            '  means don't show names of files
Public Const FOF_WANTMAPPINGHANDLE = &H20          '  Fill in SHFILEOPSTRUCT.hNameMappings
Public Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Const GW_HWNDFIRST = 0
Const GW_HWNDLAST = 1
Const GW_HWNDNEXT = 2
Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal Hbrush As Long) As Long
Type RECT
Left As Long
Top As Long
Right As Long
Bottom As Long
End Type
Global fillarea As RECT
Public Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Global Const EWX_SHUTDOWN = 1
Global Const EWX_REBOOT = 2
Declare Function GetWindowLong Lib "user32" _
Alias "GetWindowLongA" (ByVal hwnd As Long, _
ByVal nIndex As Long) As Long
Declare Function SetWindowLong Lib "user32" _
Alias "SetWindowLongA" (ByVal hwnd As Long, _
ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function FindWindowEx Lib "user32" _
Alias "FindWindowExA" (ByVal hWndParent As _
Long, ByVal hWndChildWindow As Long, ByVal _
lpClassName As String, ByVal lpsWindowName _
As String) As Long
Public Const GWL_STYLE = &HFFF0
Public Const TBSTYLE_FLAT = &H800

Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const ERROR_SUCCESS = 0&

'--------
' Add Frame1 to yur main form put this code in form_load
' Creates the line effect below yur menus
' Frame1.Width = Screen.Width + 100
' Frame1.Move -50, 0
'--------
'--------
' Text1.SelStart = Len(Text1.Text)
' Keep Text in textbox at the top
'---------
Public Sub FormsUnloadAll(sFormName As String)
'Sometimes vb doesnt unload all the forms in windows
'so put this in the unload routine of yur form
On Error Resume Next
Dim Form As Form
For Each Form In Forms
If Form.Name <> sFormName Then
Unload Form
Set Form = Nothing
End If
Next Form
End Sub



'----------
Function l_FileCopy(Src As String, dst As String) As Boolean
'Copys a file to one spot to the other
'---++UNTESTED CODE++---
l_FileCopy = False
Static buf$
Dim BTest!, FSize!
Dim Chunk%, F1%, F2%
Const BUFSIZE = 1024
If Dir(Src) = "" Then Exit Function
On Error GoTo FileCopyError
F1 = FreeFile
Open Src For Binary As F1
F2 = FreeFile
Open dst For Binary As F2
FSize = LOF(F1)
BTest = FSize - LOF(F2)
Do Until BTest = 0

If BTest < BUFSIZE Then
Chunk = BTest
Else
Chunk = BUFSIZE
End If
buf = String(Chunk, " ")
Get F1, , buf
Put F2, , buf
BTest = FSize - LOF(F2)
'FileCopy = False
timeout 1
Loop
'FileCopy = True
Close F1
Close F2
Exit Function
FileCopyError:
MsgBox "Copy Error!"
Close F1
Close F2
Exit Function
End Function
'------
Public Sub EnsureLoad(ByRef f As Variant, ByVal bUnloadInvisible As Boolean)
'USAGE: put this in yur form load keeps your
'program from having critical errors before its
'fully loaded up
' Trap and memory errors while loading a form or array element
' If an error occurs, ask the user to close a program before continuing
' Also unload any INVISIBLE forms if bUnloadInvisible is true
Dim bContinue As Boolean
Dim iFormCount As Integer
Dim I As Integer
bContinue = True
Do
On Error Resume Next
Load f
Select Case Err.Number
Case 0
bContinue = False
Case 7
MsgBox "&lt;Nexus 2000 Light needs more memory in order to continue running." + vbCrLf + vbCrLf + _
"Close some of your other applications then click OK to try to continue", vbCritical, "Memory Allocation Error"
Case Else
MsgBox "&lt;Nexus 2000 Light encountered an unexpected error while loading a window." + vbCrLf + vbCrLf + _
"Error Number #" + CStr(Err.Number) + vbCrLf + _
Err.Description + vbCrLf + vbCrLf + _
"Close some of your other applications then click OK to try to continue", vbCritical, "Memory Allocation Error"
End Select
If bContinue And bUnloadInvisible Then
' try to free (unload) forms that are not visible
iFormCount = Forms.Count - 1
For I = iFormCount To 0 Step -1
If Forms(I).Visible = False Then
On Error Resume Next
Unload Forms(I)
Set Forms(I) = Nothing
End If
Next I
End If
Loop Until bContinue = False
End Sub

Public Function BuildParseStr(vArray As Variant) As String
'This Function Takes Each Element in the Array Passed and Creates _
a Parseable String. Using the " ," as the Delimeter. Could be Changed to _
Accept any Character as the Delimeter.
Dim I As Integer, BldStr As String
If Not IsArray(vArray) Then 'If not an array then return zero length string.
BuildParseStr = ""
Exit Function
End If
For I = LBound(vArray) To UBound(vArray) 'Go thru each element in the array
If VarType(vArray(I)) <> vbString Then ' Make sure all element are string type
vArray(I) = CStr(vArray(I))                       ' If Not Convert them to strings.
End If
If I = UBound(vArray) Then                      'Keep from Appending last "," at the end of the final returned string
BldStr = BldStr & vArray(I)
Else
BldStr = BldStr & vArray(I) & ","   'Build the String on the Fly.
End If
Next I
BuildParseStr = BldStr          ' Return Parseable String.
End Function

Public Function GetFileTitle(FullFileName As String)
For I = 1 To Len(FullFileName)
d = Right(FullFileName, I)
If Left(d, 1) = "\" Then GetFileTitle = Right(FullFileName, I - 1): Exit Function
GetFileTitle = FullFileName
Next
End Function

Public Function OpenIt(frm As Form, ToOpen As String)
'USAGE: OPENIT "c:\windows\blah.exe"
'USAGE: OPENIT "http://www.coldfusionz.net"
'USAGE: OPENIT "mailto: magadass@usa.net"
ShellExecute frm.hwnd, "Open", ToOpen, "", &O0, SW_NORMAL
End Function
Public Function TextBoxLoad(TextFile As String, text As TextBox)
'Loads the file into the textbox
'LOAD "stuff.txt",text1
On Error Resume Next
Dim a$
Open TextFile For Input As #1
text.text = Input(LOF(1), #1)
Close #1
End Function
Public Sub GetThemeInf(destForm As Form)
Dim p As String
'GET all the values
'Line Colour
On Error Resume Next
LineCol = PROBas.INIGetSetting("theme", "titlelinecol", PATHFile + "\Users\" + CurrentUser + "\CurrentTheme.NexTheme")
destForm.P_Line.BorderColor = LineCol

'Line Width
px = PROBas.INIGetSetting("theme", "titlelinewidth", PATHFile + "\Users\" + CurrentUser + "\CurrentTheme.NexTheme")
destForm.P_Line.BorderWidth = px

'Font Colour
FontCol = PROBas.INIGetSetting("theme", "titlefontcol", PATHFile + "\Users\" + CurrentUser + "\CurrentTheme.NexTheme")
destForm.P_Text.ForeColor = FontCol

a:
'Window Picture
p = PROBas.INIGetSetting("theme", "windowtitle", PATHFile + "\Users\" + CurrentUser + "\CurrentTheme.NexTheme")
If p = "#" Then
destForm.P_BackPic.Picture = Main.DefaultImg.Picture
Else
    
    If FileCheck(p) = False Then
234
    PROBas.INISaveSetting "#", "theme", "windowtitle", PATHFile + "\Users\" + CurrentUser + "\CurrentTheme.NexTheme"
    GoTo a:
    Else
    On Error GoTo 234
    destForm.P_BackPic.Picture = LoadPicture(p)
    End If

End If

'Window Title Font Name
p = PROBas.INIGetSetting("theme", "titlefont", PATHFile + "\Users\" + CurrentUser + "\CurrentTheme.NexTheme")
destForm.P_Text.FontName = p

'Window Title Font Properties

p = PROBas.INIGetSetting("theme", "titlebold", PATHFile + "\Users\" + CurrentUser + "\CurrentTheme.NexTheme")
If p = "1" Then
destForm.P_Text.FontBold = True
Else
destForm.P_Text.FontBold = False
End If


p = PROBas.INIGetSetting("theme", "titleitalic", PATHFile + "\Users\" + CurrentUser + "\CurrentTheme.NexTheme")
If p = "1" Then
destForm.P_Text.FontItalic = True
Else
destForm.P_Text.FontItalic = False
End If




p = PROBas.INIGetSetting("theme", "titleunderline", PATHFile + "\Users\" + CurrentUser + "\CurrentTheme.NexTheme")
If p = "1" Then
destForm.P_Text.FontUnderline = True
Else
destForm.P_Text.FontUnderline = False
End If


destForm.P_Text.FontSize = PROBas.INIGetSetting("theme", "titlesize", PATHFile + "\Users\" + CurrentUser + "\CurrentTheme.NexTheme")


End Sub
Public Function FindContainerFolder(FileOrFolder As String)
Dim thedir As String
p = Len(FileOrFolder) + 1
Do
p = p - 1
If p = 0 Then MsgBox "Could not find containing directory! Please re-specify shell string.", 48:  Exit Function
thedir = Left(FileOrFolder, p - 1)
Loop Until Right(thedir, 1) = "\" Or Len(thedir) = 2
If thedir = "" Then MsgBox "Couldn't find directory!", 16, "Command Protocol 8.2": Exit Function
FindContainerFolder = thedir
End Function




Public Sub INISaveSetting(ByVal vntValue As Variant, ByVal sSection As String, ByVal sKey As String, SFilename As String)
' Will save an INI Setting to the specified Section and Key in the INI file
' secified by the full path name in sFileName

#If Win32 Then
Dim xRet          As Long
#Else
Dim xRet          As Integer
#End If
xRet = WritePrivateProfileString(sSection, sKey, CStr(vntValue), SFilename)
End Sub
Public Function INIGetSetting(ByVal sSection As String, ByVal sKey As String, SFilename As String) As Variant
' Will return an INI entry in the specified section at the specified key in the INI file
' specified by the full path name in sFilename

#If Win32 Then
Dim xRet          As Long
#Else
Dim xRet          As Integer
#End If
Dim sReturnStr    As String
Dim nStringLen    As Integer
nStringLen = 255
sReturnStr = String(nStringLen, Chr$(0))  ' Buffer String
xRet = GetPrivateProfileString(sSection, sKey, "", sReturnStr, nStringLen, SFilename)
INIGetSetting = Left(sReturnStr, xRet)
End Function
Public Sub INIDeleteSetting(SFilename As String, ByVal sSection As String, Optional vntKey As Variant)
' If vntKey is specified it this will delete the entry specified by vntKey, if not
' it will delete the entire section sepecified by sSection in the INI specefied by
' sFilename

#If Win32 Then
Dim xRet          As Long
#Else
Dim xRet          As Integer
#End If
' If key was provided just delete that key and value, if not delete the
' entire section
If IsMissing(vntKey) Then
xRet = WritePrivateProfileString(sSection, 0&, 0&, SFilename)
Else
xRet = WritePrivateProfileString(sSection, CStr(vntKey), 0&, SFilename)
End If
End Sub
'Public Function INIGetAllSettings(SFilename As String, ByVal sSection As String) As Variant
' Returns an variant array of all keys(0) and values(1) same as GetAllSettings
' This is  the complicated one.    It reads all of the Key Names into a temporary array
' then after teh array has been read it will crate another array.  The new array is
' 2 dimensional, the first dimension is the pair number.   The second dimension
' is 0 for the keyname, 1 for the value.

'#If Win32 Then
'Dim xRet          As Long
'#Else
'Dim xRet          As Integer
'#End If
'Dim sReturnStr    As String
'Dim nStringLen    As Integer
'Dim nEndOfKey     As Integer
'Dim nNumKeys      As Integer
'Dim arrValues()   As Variant
'nStringLen = 5000        ' Must be big enough to hold all keys
'sReturnStr = String(nStringLen, Chr$(0))
'nNumKeys = -1
'xRet = GetPrivateProfileString(sSection, 0&, "", sReturnStr, nStringLen, SFilename)
'' Parse the string, and add the elements to the array
'Do While (InStr(sReturnStr, Chr$(0)) > 1)
'' Get each key in the section
'nEndOfKey = InStr(sReturnStr, Chr$(0))
'nNumKeys = nNumKeys + 1
'ReDim Preserve arrValues(nNumKeys)
'arrValues(nNumKeys) = Left$(sReturnStr, nEndOfKey - 1)
'sReturnStr = Mid(sReturnStr, nEndOfKey + 1)
'Loop
'Debug.Print INIGetAllSettings
'If nNumKeys = -1 Then
' if no keys return an empty variant
'INIGetAllSettings = Empty
'Else
' Get the values for each key and return that, to maintain compliance with
' GetAllSettings
'ReDim arrFullArray(0 To nNumKeys, 0 To 1) As Variant
'For nNumKeys = LBound(arrValues) To UBound(arrValues)
'arrFullArray(nNumKeys, 0) = arrValues(nNumKeys)
'arrFullArray(nNumKeys, 1) = INIGetSetting(SFilename, sSection, arrValues(nNumKeys))
'Next nNumKeys
'INIGetAllSettings = arrFullArray
'End If
'End Function
Public Function App_Version() As String
'Returns version of your application as
'1.4 for example. text1.text = version
App_Version = VB.app.Major & "." & VB.app.Minor
End Function
Public Function FileCheck(Path$) As Boolean
'USAGE: If FileCheck("C:\windows\kewl.exe") then msgbox "it was found"
FileCheck = True 'Assume Success
On Error Resume Next
Dim Disregard As Long
Disregard = FileLen(Path)
If Err <> 0 Then FileCheck = False
End Function
Function Encrypt_Crypto(text$) As String 'this is not strong encryption
'USAGE: text1.text = crypto("stuff") returns the encrypter string
Dim bleh$
Dim X As Integer
For X = 1 To Len(text$)
bleh$ = bleh$ & Chr$(Asc(Mid$(text$, X, 1)) Xor 5)
Next
Encrypt_Crypto = bleh$
End Function
Public Sub SystemTrayDeleteIcon(Form As Form)
Call Shell_NotifyIcon(NIM_DELETE, CIAsystray)
End Sub
Public Function SystemTrayAddIcon(Form As Form)
CIAsystray.cbSize = Len(CIAsystray)
CIAsystray.hwnd = Form.hwnd
CIAsystray.uId = vbNull
CIAsystray.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
CIAsystray.ucallbackMessage = WM_MOUSEMOVE
CIAsystray.hIcon = Form.Icon
CIAsystray.szTip = Form.Caption & vbNullChar
Call Shell_NotifyIcon(NIM_ADD, CIAsystray)
'*Note The Lines Below Hide The Form
'*After The Icon Has Been Added If You
'*Want To Exclude That Option Just Erase It
app.TaskVisible = False
Form.Hide

'----PUT THIS IN YOUR FORM mousemove CODE
'Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'On Error Resume Next
'Static lngMsg As Long
'Static blnFlag As Boolean
'Dim result As Long
'lngMsg = X / Screen.TwipsPerPixelX
'If blnFlag = False Then
'blnFlag = True
'   Select Case lngMsg
'   Case WM_LBUTTONDBLCLICK 'Double Click
'   Me.Show
'   Case WM_RBUTTONUP 'Right Button
'   PopupMenu TRAYMNU
'result = SetForegroundWindow(Me.hwnd)
'End Select
'blnFlag = False
'End If
'End Sub
'----------------END CODE FOR FORM
End Function
Public Function UpdateProgress(pb As Control, ByVal Percent)
'Replacement for progress bar..looks nicer also
Dim Num$ 'use percent
If Not pb.AutoRedraw Then 'picture in memory ?
pb.AutoRedraw = -1 'no, make one
End If
pb.Cls 'clear picture in memory
pb.ScaleWidth = 100 'new sclaemodus
pb.DrawMode = 10 'not XOR Pen Modus
Num$ = Format$(Percent, "###") + "%"
pb.CurrentX = 50 - pb.TextWidth(Num$) / 2
pb.CurrentY = (pb.ScaleHeight - pb.TextHeight(Num$)) / 2
pb.Print Num$ 'print percent
pb.Line (0, 0)-(Percent, pb.ScaleHeight), , BF
pb.Refresh 'show differents
End Function
Public Function FormsOnTop(frmForm As Form, fOnTop As Boolean)
'USAGE: ONTOP ME,TRUE   -ONTOP MOST
'       ONTOP ME,FALSE  -NOT TOP MOST
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Dim lState As Long
Dim iLeft As Integer, iTop As Integer, iWidth As Integer, iHeight As Integer
With frmForm
iLeft = .Left / Screen.TwipsPerPixelX
iTop = .Top / Screen.TwipsPerPixelY
iWidth = .Width / Screen.TwipsPerPixelX
iHeight = .Height / Screen.TwipsPerPixelY
End With
If fOnTop Then
lState = HWND_TOPMOST
Else
lState = HWND_NOTOPMOST
End If
Call SetWindowPos(frmForm.hwnd, lState, iLeft, iTop, iWidth, iHeight, 0)
End Function
Public Function SetDesktopWindow(frmForm As Form)
'USAGE: ONTOP ME,TRUE   -ONTOP MOST
'       ONTOP ME,FALSE  -NOT TOP MOST
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Dim lState As Long
Dim iLeft As Integer, iTop As Integer, iWidth As Integer, iHeight As Integer
With frmForm
iLeft = .Left / Screen.TwipsPerPixelX
iTop = .Top / Screen.TwipsPerPixelY
iWidth = .Width / Screen.TwipsPerPixelX
iHeight = .Height / Screen.TwipsPerPixelY
End With
lState = 1
Call SetWindowPos(frmForm.hwnd, lState, iLeft, iTop, iWidth, iHeight, 0)
End Function

Public Function FormsMove(theform As Form)
'USAGE: FORMMOVE ME    -PUT IN MOUSEDOWN
releaseCapture
Call SendMessage(theform.hwnd, &HA1, 2, 0&)
End Function
Public Function FindListBox(ListBox As ListBox, text As String, Optional Mode As Byte) As Integer
'EXAMPLE:FINDLIST(LIST1,"STUFF",0)
'--MODE 0 EXACT  --STARTPOS: where to begin in list
'--MODE 1 NOT EXACT --TEXT: STRING TO SEARCH FOR
Dim found As Integer
If Mode = 0 Then
found = SendMessageByString(ListBox.hwnd, LB_FINDSTRING, -1, text)
Else
found = SendMessageByString(ListBox.hwnd, LB_FINDSTRINGEXACT, -1, text)
End If
FindListBox = found
End Function
Public Function FindComboBox(ComboBox As ComboBox, text As String, Optional Mode As Byte) As Integer
'EXAMPLE:FINDLIST(LIST1,"STUFF",0)
'--MODE 0 EXACT  --STARTPOS: where to begin in list
'--MODE 1 NOT EXACT --TEXT: STRING TO SEARCH FOR
Dim found As Integer
If Mode = 0 Then
found = SendMessageByString(ComboBox.hwnd, CB_FINDSTRING, -1, text)
Else
found = SendMessageByString(ComboBox.hwnd, CB_FINDSTRINGEXACT, -1, text)
End If
FindComboBox = found
End Function
Sub ListBoxLoad(File As String, ListBox As ListBox)
'USAGE: LOADLIST(LIST1,"STUFF.LST")
'THAT WILL LOAD THE CONTENTS OF STUFF.LST
On Error Resume Next
Dim free%, G$
free = FreeFile
If FileCheck(File) = False Then Exit Sub
ListBox.Clear
Open File For Input As #free
Do Until EOF(free)
Line Input #free, G$
ListBox.AddItem G$
Loop
Close free
End Sub
Public Sub ListBoxSave(File As String, List As ListBox)
On Error Resume Next
Dim free%
free = FreeFile
Dim SaveList As Long
Open File For Output As #free
If FileCheck(File) = False Then Exit Sub
For SaveList& = 0 To List.ListCount - 1
Print #free, List.List(SaveList&)
Next SaveList&
Close #free
Finish:
End Sub
Function T_Inschr(ByVal Strin As String, ByVal InsMe As String) As String
Dim inptxt$, lenth%, numspc%, nextchr$, newsent$
Let inptxt$ = Strin
Let lenth% = Len(inptxt$)
Do While numspc% <= lenth%
Let numspc% = numspc% + 1
Let nextchr$ = Mid$(inptxt$, numspc%, 1)
Let nextchr$ = nextchr$ + InsMe
Let newsent$ = newsent$ + nextchr$
Loop
T_Inschr = newsent$
End Function
Function T_Spaced(Strin As String) As String
'  x = t_spaced(text1)
Dim inptxt$, lenth%, numspc%, nextchr$, newsent$
Let inptxt$ = Strin
Let lenth% = Len(inptxt$)
Do While numspc% <= lenth%
Let numspc% = numspc% + 1
Let nextchr$ = Mid$(inptxt$, numspc%, 1)
Let nextchr$ = nextchr$ + " "
Let newsent$ = newsent$ + nextchr$
Loop
T_Spaced = newsent$
End Function
Sub Draw_Make3d(myForm As Form, MyCtl As Control)
'Place in Form_Paint; works best with Grey background
'Example:
'Make3d me,command1
myForm.ScaleMode = 3
myForm.CurrentX = MyCtl.Left - 1
myForm.CurrentY = MyCtl.Top + MyCtl.Height
myForm.Line -Step(0, -(MyCtl.Height + 1)), RGB(92, 92, 92)
myForm.Line -Step(MyCtl.Width + 1, 0), RGB(92, 92, 92)
myForm.Line -Step(0, MyCtl.Height + 1), RGB(255, 255, 255)
myForm.Line -Step(-(MyCtl.Width + 1), 0), RGB(255, 255, 255)
End Sub
Sub FormsCenter(frm As Form)
'CenterForm me
Dim X%, Y%
X = Screen.Width / 2 - frm.Width / 2
Y = Screen.Height / 2 - frm.Height / 2
frm.Move X, Y
End Sub
Function SetTextSpecial(ByVal Handle&, ByVal TextToSend$) As Long
Dim dum%
'THIS IS A SPECIAL SETTEXT THAT WILL SEND TEXT TO ANYTHING
'USE EXAMPLE: SETTEXT "HANDLE TO WINDOW", "TEXT TO SEND"
SetTextSpecial = SendMessageByString(Handle, WM_SETTEXT, 0, TextToSend$)
dum = SendMessageByNum(Handle, WM_CHAR, 13, 0) ' 13 == return
End Function
Public Function FindWindow(WindowNAME As String)
Dim Desktop%, window%
Desktop% = GetDesktopWindow
window% = FindChildByClass(Desktop%, WindowNAME)
FindWindow = window%
End Function
Function FindChildByClass(parentw, childhand)
'Took this from an aol bas
Dim firs%, firss%, room%
firs% = GetWindow(parentw, 5)
If UCase(Mid(getclass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo bone
firs% = GetWindow(parentw, GW_CHILD)
If UCase(Mid(getclass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo bone
While firs%
firss% = GetWindow(parentw, 5)
If UCase(Mid(getclass(firss%), 1, Len(childhand))) Like UCase(childhand) Then GoTo bone
firs% = GetWindow(firs%, 2)
If UCase(Mid(getclass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo bone
Wend
FindChildByClass = 0
bone:
room% = firs%
FindChildByClass = room%
End Function
Function getclass(child)
Dim buffer$, getclas%
buffer$ = String$(250, 0)
getclas% = GetClassName(child, buffer$, 250)
getclass = buffer$
End Function
Public Function File_ByteConversion(NumberOfBytes As Single) As String
On Error Resume Next
If NumberOfBytes < 1024 Then 'checks to see if its so small that it cant be converted into larger grouping
File_ByteConversion = NumberOfBytes & " Bytes"
End If
If NumberOfBytes > 1024 Then  'Checks to see if file is big enough to convert into KB
File_ByteConversion = Format(NumberOfBytes / 1024, "0.00") & " KB"
End If
If NumberOfBytes > 1024000 Then 'Checks to see if its big enough to convert into MB
File_ByteConversion = Format(NumberOfBytes / 1024000, "###,###,##0.00") & " MB"
End If
End Function
Public Function GetNetworkUserName() As String
  Dim lpBuff   As String * 25
  Dim retval   As Long

  retval = GetUserName(lpBuff, 25)
  ' trim off any trailing spaces found in the name
  GetNetworkUserName = Left(lpBuff, InStr(lpBuff, Chr(0)) - 1)

End Function

Public Function GetWorkstationName() As String
  Dim lpBuff   As String * 25
  Dim retval   As Long

  retval = GetComputerName(lpBuff, 25)
  ' trim off any trailing spaces found in the name
  GetWorkstationName = Left(lpBuff, InStr(lpBuff, Chr(0)) - 1)

End Function

Sub Draw_3d_Border_Around_Form(frmForm As Form)
'++++++++-------UNTESTED CODE-------++++++++
On Error Resume Next
Const cPi = 3.1415926
Dim intLineWidth As Integer
intLineWidth = 5
'save scale mode
Dim intSaveScaleMode As Integer
intSaveScaleMode = frmForm.ScaleMode
frmForm.ScaleMode = 3
Dim intScaleWidth As Integer
Dim intScaleHeight As Integer
intScaleWidth = frmForm.ScaleWidth
intScaleHeight = frmForm.ScaleHeight
'clear form
frmForm.Cls
'draw white lines
frmForm.Line (0, intScaleHeight)-(intLineWidth, 0), &HFFFFFF, BF
frmForm.Line (0, intLineWidth)-(intScaleWidth, 0), &HFFFFFF, BF
'draw grey lines
frmForm.Line (intScaleWidth, 0)-(intScaleWidth - intLineWidth, intScaleHeight), &H808080, BF
frmForm.Line (intScaleWidth, intScaleHeight - intLineWidth)-(0, intScaleHeight), &H808080, BF
'draw triangles(actually circles) at corners
Dim intCircleWidth As Integer
intCircleWidth = Sqr(intLineWidth * intLineWidth + intLineWidth * intLineWidth)
frmForm.FillStyle = 0
frmForm.FillColor = QBColor(15)
frmForm.Circle (intLineWidth, intScaleHeight - intLineWidth), intCircleWidth, QBColor(15), -3.1415926, -3.90953745777778 '-180 * cPi / 180, -224 * cPi / 180
frmForm.Circle (intScaleWidth - intLineWidth, intLineWidth), intCircleWidth, QBColor(15), -0.78539815, -1.5707963 ' -45 * cPi / 180, -90 * cPi / 180
'draw black frame
frmForm.Line (0, intScaleHeight)-(0, 0), 0
frmForm.Line (0, 0)-(intScaleWidth - 1, 0), 0
frmForm.Line (intScaleWidth - 1, 0)-(intScaleWidth - 1, intScaleHeight - 1), 0
frmForm.Line (0, intScaleHeight - 1)-(intScaleWidth - 1, intScaleHeight - 1), 0
'restore scale mode
frmForm.ScaleMode = intSaveScaleMode
End Sub


Public Function PlayWav(Optional File As String)
On Error Resume Next
Dim rc As String
'to stop it just do playwav by itself
If File Then rc = sndPlaySound("", SND_ASYNC)
If FileCheck(File) = False Then Exit Function
rc = sndPlaySound("FILE", SND_ASYNC)
End Function
Public Function MouseMoving() As Boolean
'NOTE:THIS HAS TO GO IN A TIMER...Probably 100 interval or less
'Determines if the user is using the mouse...If not you can assume
'the AFK time of the user for various programs
'----
'Returns true for movement and false for no movement
'----
Dim TimeExpired%
GetCursorPos Pnt
NewX = Pnt.X
NewY = Pnt.Y
If OldX - NewX = 0 Then
MouseMoving = False
Else
MouseMoving = True
TimeExpired = 0
End If
OldX = NewX
OldY = NewY
End Function
Public Function TextBoxSave(FilePath As String, text As TextBox)
On Error GoTo done:
Dim fno%
fno = FreeFile
Open FilePath For Output As #fno
If FileCheck(FilePath) = False Then Exit Function
Print #fno, text.text
Close #fno
done:
End Function
Sub Draw_ControlShadow(f As Form, c As Control, shadow_effect As Integer, shadow_width As Integer, shadow_color As Long)
'+++++++-----UNTESTED CODE-----++++++++
Dim shColor As Long
Dim shWidth As Integer
Dim oldWidth As Integer
Dim oldScale As Integer
shWidth = shadow_width
shColor = shadow_color
oldWidth = f.DrawWidth
oldScale = f.ScaleMode
f.ScaleMode = 3 'Pixels
f.DrawWidth = 1
Select Case shadow_effect
Case GFM_DROPSHADOW
f.Line (c.Left + shWidth, c.Top + shWidth)-Step(c.Width - 1, c.Height - 1), shColor, BF
Case GFM_BACKSHADOW
f.Line (c.Left - shWidth, c.Top - shWidth)-Step(c.Width - 1, c.Height - 1), shColor, BF
End Select
f.DrawWidth = oldWidth
f.ScaleMode = oldScale
End Sub
Public Sub CADDisable()
'DISABLE CONTROL ALT DEL
'To use this just put this in Form_Load
'DisableCAD
Dim pOld$, ret$
ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, True, pOld, 0)
End Sub
Public Sub CADEnable()
'ENABLE CONTROL ALT DEL
'To use this just put this in Form_Load or in a Command button
'EnableCAD
Dim pOld$, ret$
ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, False, pOld, 0)
End Sub
Public Function ShellRename(ParamArray vntFileName() As Variant) As Long
'To use this put this in a command button
'lets suppose you want to rename the Windows folder, so you do
'ShellRename "C:\Windows"
Dim I As Integer
Dim sFileNames As String
Dim Dick As String
Dim SHFileOp As SHFILEOPSTRUCT
For I = LBound(vntFileName) To UBound(vntFileName)
sFileNames = sFileNames & vntFileName(I) & vbNullChar
Next
sFileNames = sFileNames & vbNullChar
Dick = FileDestination
With SHFileOp
.wFunc = &H4
.pFrom = sFileNames
.fFlags = FOF_ALLOWUNDO
.pTo = Dick
End With
ShellRename = SHFileOperation(SHFileOp)
End Function
Public Function ShellCopy(ParamArray vntFileName() As Variant) As Long
'To use this put this in a command button
'lets suppose you want to copy the Windows folder, so you do
'ShellCopy "C:\Windows"
Dim I As Integer
Dim sFileNames As Variant
Dim Dick As String
Dim SHFileOp As SHFILEOPSTRUCT
For I = LBound(vntFileName) To UBound(vntFileName)
sFileNames = sFileNames & vntFileName(I) & vbNullChar
Next
sFileNames = sFileNames & vbNullChar
Dick = FileDestination
With SHFileOp
.wFunc = &H2
.pFrom = sFileNames
.fFlags = FOF_ALLOWUNDO
.pTo = Dick
End With
ShellCopy = SHFileOperation(SHFileOp)
End Function
Public Function ShellMove(ParamArray vntFileName() As Variant) As Long
'To use this put this in a command button
'lets suppose you want to Move the Windows folder, so you do
'ShellMove "C:\Windows"
Dim I As Integer
Dim sFileNames As Variant
Dim Dick As String
Dim SHFileOp As SHFILEOPSTRUCT
For I = LBound(vntFileName) To UBound(vntFileName)
sFileNames = sFileNames & vntFileName(I) & vbNullChar
Next
sFileNames = sFileNames & vbNullChar
Dick = FileDestination
With SHFileOp
.wFunc = &H1
.pFrom = sFileNames
.fFlags = FOF_ALLOWUNDO
.pTo = Dick
End With
ShellMove = SHFileOperation(SHFileOp)
End Function
Public Function ShellDelete(ParamArray vntFileName() As Variant) As Long
'To use this put this in a command button
'lets suppose you want to delete the Windows folder, so you do
'ShellDelete "C:\Windows"
Dim I As Integer
Dim sFileNames As String
Dim SHFileOp As SHFILEOPSTRUCT
For I = LBound(vntFileName) To UBound(vntFileName)
sFileNames = sFileNames & vntFileName(I) & vbNullChar
Next
sFileNames = sFileNames & vbNullChar
With SHFileOp
.wFunc = FO_DELETE
.pFrom = sFileNames
.fFlags = FOF_NOCONFIRMATION + FOF_SILENT
End With
ShellDelete = SHFileOperation(SHFileOp)
End Function
Sub Shell_GetRunningApplications(frm As Form, lst As ListBox)
'To use this put this in a Command button
'lets pretend your Listbox's name is List1
'getrunningapplications me, list1
lst.Clear
Dim lLgthChild As Long
Dim sNameChild As String
Dim lLgthOwner As Long
Dim sNameOwner As String
Dim lHwnd As Long
Dim lHwnd2 As Long
Const vbTextCompare = 1
lHwnd = GetWindow(frm.hwnd, GW_HWNDFIRST)
While lHwnd <> 0
lHwnd2 = GetWindow(lHwnd, GW_OWNER)
lLgthOwner = GetWindowTextLength(lHwnd2)
sNameOwner = String$(lLgthOwner + 1, Chr$(0))
lLgthOwner = GetWindowText(lHwnd2, sNameOwner, lLgthOwner + 1)
lLgthChild = GetWindowTextLength(lHwnd)
sNameChild = String$(lLgthChild + 1, Chr$(0))
lLgthChild = GetWindowText(lHwnd, sNameChild, lLgthChild + 1)
If lLgthChild <> 0 Then
sNameChild = Left$(sNameChild, InStr(1, sNameChild, Chr$(0), vbTextCompare) - 1)
sNameChild = Trim(sNameChild)
If FindListBox(lst, sNameChild, 0) > -1 Then GoTo noadd:
lst.AddItem sNameChild & " - [HWND: " & lHwnd & "]"
noadd:
End If
lHwnd = GetWindow(lHwnd, GW_HWNDNEXT)
DoEvents
Wend
End Sub
Public Function Draw_Gradient(frm As Form, Color1 As Long, Color2 As Long)
'EXAMPLE: Draw_gradient picture1,vbblue,vbblack
' OR draw_gradient me,vbred,vbblack
timeout 0.5
Dim r1, g1, b1, r2, g2, b2, boxStep, posY, I As Integer
Dim redStep, greenStep, BlueStep As Integer
' separate color1 to red,green and blue
r1 = Color1 Mod &H100
g1 = (Color1 \ &H100) Mod &H100
b1 = (Color1 \ &H10000) Mod &H100
' separate color2 to red,green and blue
r2 = Color2 Mod &H100
g2 = (Color2 \ &H100) Mod &H100
b2 = (Color2 \ &H10000) Mod &H100
' calculate box height
boxStep = frm.ScaleHeight / 63
posY = 0
If g1 > g2 Then
greenStep = 0
ElseIf g2 > g1 Then
greenStep = 1
Else
greenStep = 2
End If
If r1 > r2 Then
redStep = 0
ElseIf r2 > r1 Then
redStep = 1
Else
redStep = 2
End If
If b1 > b2 Then
BlueStep = 0
ElseIf b2 > b1 Then
BlueStep = 1
Else
BlueStep = 2
End If
For I = 1 To 63
frm.Line (0, posY)-(frm.ScaleWidth, posY + boxStep), RGB(r1, g1, b1), BF
If redStep = 1 Then
r1 = r1 + 4
If r1 > r2 Then
r1 = r2
End If
ElseIf redStep = 0 Then
r1 = r1 - 4
If r1 < r2 Then
r1 = r2
End If
End If
If greenStep = 1 Then
g1 = g1 + 4
If g1 > g2 Then
g1 = g2
End If
ElseIf greenStep = 0 Then
g1 = g1 - 4
If g1 < g2 Then
g1 = g2
End If
End If
If BlueStep = 1 Then
b1 = b1 + 4
If b1 > b2 Then
b1 = b2
End If
ElseIf BlueStep = 0 Then
b1 = b1 - 4
If b1 < b2 Then
b1 = b2
End If
End If
posY = posY + boxStep
Next I
End Function
Function FormsHinge(Mainfrm As Form, HingedFrm As Form, Method As Integer)
On Error Resume Next
'THIS CONNECTS A FORM TO ANOTHER FORM...
'IT STAYS WITH IT EVEN ON MOVING.COORDINATES
'ARE LISTED BELOW FOR METHOD
'Usage:
'---------------------
'There are 8 Methods
'1 = Top
'2 = Top Right
'3 = Right
'4 = Bottom Right
'5 = Bottom
'6 = Bottom Left
'7 = Left
'8 = Top Left
'--------------------
'Let's Suppose the Form you want to Hinge
'is Form2, and you want it at the Bottom Right
'Then Put this Code in a Timer
'With an Interval of "1"
'---------------------
'HingeForm Me, Form2, 4
'---------------------
Dim G As Boolean
HingedFrm.Visible = True
If HingedFrm.Visible = True Then
G = True
Else
G = False
End If
If G = True Then
Select Case Method
Case 1
HingedFrm.Left = Mainfrm.Left
HingedFrm.Top = Mainfrm.Top - HingedFrm.Height
Case 2
HingedFrm.Left = Mainfrm.Left + Mainfrm.Width
HingedFrm.Top = Mainfrm.Top - HingedFrm.Height
Case 3
HingedFrm.Left = Mainfrm.Left + Mainfrm.Width
HingedFrm.Top = Mainfrm.Top
Case 4
HingedFrm.Left = Mainfrm.Left + Mainfrm.Width
HingedFrm.Top = Mainfrm.Top + Mainfrm.Height
Case 5
HingedFrm.Left = Mainfrm.Left
HingedFrm.Top = Mainfrm.Top + Mainfrm.Height
Case 6
HingedFrm.Left = Mainfrm.Left - HingedFrm.Width
HingedFrm.Top = Mainfrm.Top + Mainfrm.Height
Case 7
HingedFrm.Left = Mainfrm.Left - HingedFrm.Width
HingedFrm.Top = Mainfrm.Top
Case 8
HingedFrm.Left = Mainfrm.Left - HingedFrm.Width
HingedFrm.Top = Mainfrm.Top - HingedFrm.Height
Case Else
End Select
ElseIf G = False Then
End If
End Function
Function PictureBox_LoadPic(PictureBoxName As PictureBox, PicDirectory As String)
On Error Resume Next
'Usage:
'----------------
'LoadPic picture1,"C:\Windows\Desktop\Blah.jpg"
'----------------
PictureBoxName.Picture = LoadPicture(PicDirectory)
End Function
Function PictureBox_TilePic(Mainfrm As Form, PictureToTile As PictureBox)
On Error Resume Next
'Usage:
'Put the Following Code in your Forms'
'Form_Load *Note: You might want to put it
'in Form_Resize Aswell
'Let's Suppose you name the Picturebox
'As Picture1, Then Do
'----------------
'Tilepic Me, Picture1
'----------------
Mainfrm.ScaleMode = 3
Mainfrm.AutoRedraw = True
PictureToTile.ScaleMode = 3

'Get dimensions
Dim FormHeight As Long
Dim FormWidth As Long
Dim PictureHeight As Long
Dim PictureWidth As Long
Dim X%, Y%
'Assign dimensions
FormHeight = Mainfrm.ScaleHeight
FormWidth = Mainfrm.ScaleWidth
PictureHeight = PictureToTile.ScaleHeight
PictureWidth = PictureToTile.ScaleWidth

'Tile bitmap
For Y = 0 To FormHeight Step PictureHeight
For X = 0 To FormWidth Step PictureWidth
Mainfrm.PaintPicture PictureToTile.Picture, X, Y
Next X
Next Y
PictureToTile.Visible = False
End Function
Function Form_TilePic(Mainfrm As Form, PictureToTile As Form)
On Error Resume Next
'Usage:
'Put the Following Code in your Forms'
'Form_Load *Note: You might want to put it
'in Form_Resize Aswell
'Let's Suppose you name the Picturebox
'As Picture1, Then Do
'----------------
'Tilepic Me, Picture1
'----------------
Mainfrm.ScaleMode = 3
Mainfrm.AutoRedraw = True
PictureToTile.ScaleMode = 3

'Get dimensions
Dim FormHeight As Long
Dim FormWidth As Long
Dim PictureHeight As Long
Dim PictureWidth As Long
Dim X%, Y%
'Assign dimensions
FormHeight = Mainfrm.ScaleHeight
FormWidth = Mainfrm.ScaleWidth
PictureHeight = PictureToTile.ScaleHeight
PictureWidth = PictureToTile.ScaleWidth

'Tile bitmap
For Y = 0 To FormHeight Step PictureHeight
For X = 0 To FormWidth Step PictureWidth
Mainfrm.PaintPicture PictureToTile.Picture, X, Y
Next X
Next Y
PictureToTile.Visible = False
End Function

Function PictureBox_MDITilePic(MDIMainfrm As Form, MDIPictureToTile As PictureBox)
On Error Resume Next
'Usage:
'Put the Following Code in your Forms'
'Form_Load *Note: You might want to put it
'in Form_Resize Aswell
'Let's Suppose you name the Picturebox
'As Picture1, Then Do
'----------------
'MDIFormTilePic Me, Picture1
'----------------

' Prepare form
MDIPictureToTile.AutoRedraw = True
MDIPictureToTile.Visible = False

' Get dimensions
Dim FormHeight As Long
Dim FormWidth As Long
Dim PictureHeight As Long
Dim PictureWidth As Long
Dim X%, Y%
' Assign dimensions
FormHeight = MDIMainfrm.Height
FormWidth = MDIMainfrm.Width
PictureHeight = MDIPictureToTile.ScaleX(MDIMainfrm.Picture.Height, 7, 1)
PictureWidth = MDIPictureToTile.ScaleY(MDIMainfrm.Picture.Width, 7, 1)

'Resize picturebox
MDIPictureToTile.Height = MDIMainfrm.Height

' Create a new tiled form of the bitmap
For Y = 0 To FormHeight Step PictureHeight
For X = 0 To FormWidth Step PictureWidth
MDIPictureToTile.PaintPicture MDIMainfrm.Picture, X, Y
Next X
Next Y

' Copy our new bitmap to the back of the MDIFrom
MDIMainfrm.Picture = MDIPictureToTile.Image
End Function

Public Function FileSize(FilePath As String) As String
'USAGE: Label1.Caption = FileSize("C:\Stuff.exe") 'would return something like 1.23 MB or 35 KB or 176 Bytes... Up To 999,999,999.99 MB
If FileCheck(FilePath) = False Then Exit Function
Dim a As Single
a = FileLen(FilePath)
FileSize = File_ByteConversion(a)
End Function
Public Function FileEXT(FileName As String, Optional ReturnPeriod As Boolean) As String
'RETURNS the filenames EXTENSION...
'Optional return of period or not
Dim a$, B$
If InStr(FileName, ".") = 0 Then Exit Function
If ReturnPeriod = True Then
a = Right(FileName, 4)
Else
a = Right(FileName, 3)
End If
FileEXT = a
End Function
Function FileGetName(FileNa As String) As String
'USAGE Label1.caption = FileGetName("C:\whatever\blah.exe") will return blah.exe
Dim FRes As String
Dim SLen As Integer
Dim lstpos%, I%, seppos%
SLen = Len(FileNa)
lstpos = 0
For I = 1 To SLen
seppos = InStr(I, FileNa, "\", 1)
If seppos > lstpos And (I + seppos) < SLen Then
lstpos = seppos
Else
Exit For
End If
I = I + seppos
Next I
FRes = Right(FileNa, SLen - seppos)
FileGetName = FRes
End Function

Function FileGetPath(FileNa As String) As String
'USAGE: Label1.caption = FileGetPath("C:\whatever\blah\hehe.exe") would return C:\whatever\blah\
Dim FRes As String
Dim SLen As Integer
SLen = Len(FileNa)
Dim lstpos%, I%, seppos%
lstpos = 0
For I = 1 To SLen
seppos = InStr(I, FileNa, "\", 1)
If seppos > lstpos And (I) < SLen Then
lstpos = seppos
Else
seppos = lstpos '+ 1 'i + 1
Exit For
End If
I = seppos 'i + (SepPos - 1)
Next I
FRes = Left(FileNa, seppos)
FileGetPath = FRes
End Function
Public Function FileShortName(Long_Path As String) As String
'USAGE: Label1.caption = File_ShortName("C:\Program Files\Icq\") ' Returns -  C:\PROGRA~1\ICQ\
Dim Short_Path As String
Dim Answer As Long
Short_Path = Space(250)
Answer = GetShortPathName(Long_Path, Short_Path, Len(Short_Path))
'Debug.Print Answer
If Answer > 0 Then
FileShortName = Left$(Short_Path, Answer)
End If
End Function
Public Sub TextBoxClearAll(frm As Form, ctl As Control)
'CLEARS ALL TEXTBOXES ON THE FORM
For Each ctl In frm
If TypeOf ctl Is TextBox Then
ctl.text = ""
End If
Next ctl
End Sub
Public Sub ClearAllControls(frmForm As Form)
'CLEARS ALL CONTROLS WITH TEXT INPUT OR INDEX INPUT CAPABILITYS ON THE FORM
Dim ctlControl As Object
On Error Resume Next
For Each ctlControl In frmForm.Controls
ctlControl.text = ""
ctlControl.ListIndex = -1
DoEvents
Next ctlControl
End Sub
Public Function FileDirCheck(ByVal sDirName As String) As Boolean
'RETURNS TRUE IF PATH EXISTS ELSE IT RETURNS FALSE
Dim sDir As String
On Error Resume Next
FileDirCheck = False
sDir = Dir$(sDirName, vbDirectory)
If (Len(sDir) > 0) And (Err = 0) Then
FileDirCheck = True
End If
End Function
Public Sub timeout(HowLong)


    Dim TheBeginning
    Dim NoFreeze As Integer
    TheBeginning = Timer


    Do
        If Timer - TheBeginning >= HowLong Then Exit Sub
        NoFreeze% = DoEvents()
    Loop

End Sub


Function Encrypt_Crypto2(text, types)
'to encrypt, example:
'encrypted$ = EncryptType("messagetoencrypt", 0)
'to decrypt, example:
'decrypted$ = EncryptType("decryptedmessage", 1)
'* First Paramete is the Message
'* Second Parameter is 0 for encrypt
'  or 1 for decrypt
'Now if this distorts yur text and all yur looking
'for is something so users cant change yur settings
'then use crypto one its not as strong but it works
'fine and I have never had a problem with it.
Dim God%, Current$, Process$
For God = 1 To Len(text)
If types = 0 Then
Current$ = Asc(Mid(text, God, 1)) - 1
Else
Current$ = Asc(Mid(text, God, 1)) + 1
End If
Process$ = Process$ & Chr(Current$)
Next God
Encrypt_Crypto2 = Process$
End Function
Function FreeProcess()
'Unfreezes a locked loop or subroutine
Dim Process%
Do: DoEvents
Process = Process + 1
If Process = 50 Then Exit Do
Loop
End Function
Function T_ReplaceText(text, charfind, charchange)
'Replaces text with other text
Dim Replace As Single
Dim thechar$, thechars$
For Replace = 1 To Len(text)
thechar = Mid(text, Replace, 1)
thechars = thechars & thechar
If thechar = charfind Then
thechars = Mid(thechars, 1, Len(thechars) - 1) + charchange
End If
Next Replace
T_ReplaceText = thechars
End Function

Public Function Encrypt_Cipher(PlainText As String, Secret As String) As String
'This is stronger and can enccrypt binary files
'unlike the other encrytps but is alot slower
Dim a, B, c As String
Dim pTb, cTb, cT As String
Dim I%, pseudoi%
For I = 1 To Len(PlainText)
pseudoi = I Mod Len(Secret)
If pseudoi = 0 Then pseudoi = 1
a = Mid(Secret, pseudoi, 1)
B = Mid(Secret, pseudoi + 1, 1)
c = Asc(a) Xor Asc(B)
pTb = Mid(PlainText, I, 1)
cTb = c Xor Asc(pTb)
cT = cT + Chr(cTb)
'Returns the progress it is in the string
DoEvents
Next I
Encrypt_Cipher = cT
End Function
Sub Draw_GradientTitleBar(Form As PictureBox, Color1 As Long, Color2 As Long, Optional text As String, Optional ForeColor As Long)
timeout 0.01
'DRAW GRADIENT LEFT TO RIGHT WITH OPTIONAL TEXT...
'USE THIS TO MAKE FAKE TITLE BARS EASILY
Dim X!, x2!, Y%, I%, red1%, green1%, blue1%, red2%, green2%, blue2%, pat1!, pat2!, pat3!, c1!, c2!, c3!
' find the length of the form and cut it into 80 pieces
x2 = Form.ScaleWidth / 80
Y% = Form.ScaleHeight
' separating red, green, and blue in each of the two colors
red1% = Color1 And 255
green1% = Color1 \ 256 And 255
blue1% = Color1 \ 65536 And 255
red2% = Color2 And 255
green2% = Color2 \ 256 And 255
blue2% = Color2 \ 65536 And 255
' cut the difference between the two colors into 100 pieces
pat1 = (red2% - red1%) / 80
pat2 = (green2% - green1%) / 80
pat3 = (blue2% - blue1%) / 80
' set the c variables at the starting colors
c1 = red1%
c2 = green1%
c3 = blue1%
' draw 80 different lines on the form
For I% = 1 To 80
Form.Line (X, 0)-(X + x2, Y%), RGB(c1, c2, c3), BF
X = X + x2 ' draw the next line one step up from the old step
c1 = c1 + pat1 ' make the c variable equal 2 it's next step
c2 = c2 + pat2
c3 = c3 + pat3
Next
Form.CurrentX = 0
Form.CurrentY = 0
Form.ForeColor = ForeColor
Form.Font.Size = 10
Form.Font.Name = "Arial"
Form.Print text
End Sub
Function App_ProgramAlreadyRunning() As Boolean
App_ProgramAlreadyRunning = False
If (app.PrevInstance = True) Then
App_ProgramAlreadyRunning = True
End If
End Function

Function Shell_ExitWindows(BootMode As Integer)
On Error Resume Next
'Usage:
'-------------
'BootModes are as Follows:
'1 = Shutdown Windows
'2 = Reboot Windows
'-------------
'To use this Put this code in a Command
'button or anywhere you want to use it
'-------------
'Exitwindows 1 '- To shutdown windows
'Exitwindows 2 '- To reboot windows
'-------------
Dim bootans As Integer
Dim bootvalue As Long
Select Case BootMode
Case 1 'Shutdown Windows
bootans = vbYes
If bootans = vbYes Then
bootvalue = ExitWindowsEx(EWX_SHUTDOWN, 0&)
End If
Case 2 ' Reboot Windows
bootans = MsgBox("Are you sure you want to reboot windows?", vbQuestion Or vbYesNo, "Reboot Windows")
If bootans = vbYes Then
bootvalue = ExitWindowsEx(EWX_REBOOT, 0&)
End If
Case Else
End Select
End Function

Public Function TextBoxInsert(Text1 As TextBox, TEXTSTRING)
Dim Num As Single
Dim LENGTH As Single
Dim STUFF$, a$, B$
STUFF = "": a = "": B = ""
Num = Text1.SelStart
STUFF = Text1.text
LENGTH = Len(STUFF)
a = Left(STUFF, Num)
B = Right(STUFF, LENGTH - Num)
STUFF = a & TEXTSTRING & B
Text1.text = STUFF
Text1.SelStart = Len(a) + Len(TEXTSTRING)
End Function
Public Function LineWrap(sString As String, iInterval As Integer) As String
Dim lPos As Long
Dim iPosCounter As Long
Dim lFinalLen As Long
Dim lBeginPos As Long
Dim lEndPos As Long
Dim iWordLen As Long
Dim iWordPos As Long
Dim dWrapThresh As Integer
lFinalLen = Len(sString)
Do Until lPos >= lFinalLen
If iPosCounter = iInterval Then 'ok, we hit the wrap point
iPosCounter = 0 'Reset the interval counter
'Get the beginning position of the current word
For lBeginPos = lPos To 0 Step -1
If Mid$(sString, lBeginPos, 1) = " " Then Exit For
Next lBeginPos
'Get the ending position of the current word
For lEndPos = lPos To lFinalLen
If Mid$(sString, lEndPos, 1) = " " Then Exit For
Next lEndPos
'Get the length of the current word
iWordLen = (lEndPos - 1) - (lBeginPos + 1)
'Find out at which character we are located in the word
iWordPos = lPos - (lBeginPos + 1)
'If we are over half way, then we move forward, otherwise we move
'     back
dWrapThresh = iWordLen / 2
If lEndPos > Len(sString) Then Exit Do
If iWordPos >= dWrapThresh Then 'Wrap at end of word
sString = Left$(sString, lEndPos) + vbCrLf + Right$(sString, lFinalLen - lEndPos)
Else 'Wrap at beginning of word
sString = Left$(sString, lBeginPos) + vbCrLf + Right$(sString, lFinalLen - lBeginPos)
End If
lFinalLen = Len(sString)
End If
iPosCounter = iPosCounter + 1
If lPos > 0 Then If Mid$(sString, lPos, 2) = vbCrLf Then iPosCounter = 0 'Reset if new line already
lPos = lPos + 1
Loop
LineWrap = sString
End Function

'Public Sub TBar97(TBar As Toolbar)
'Dim lTBarStyle As Long, lTBarHwnd As Long
'If Not TypeOf TBar Is Toolbar Then Exit Sub
'lTBarHwnd = FindWindowEx(TBar.hwnd, 0&, "ToolbarWindow32", vbNullString)
'lTBarStyle = GetWindowLong(lTBarHwnd, GWL_STYLE)
'lTBarStyle = lTBaSetWindowLongrStyle Or TBSTYLE_FLAT
'lTBarStyle = (lTBarHwnd, GWL_STYLE, lTBarStyle)
'TBar.Refresh
'End Sub
Public Function RegistryGet(hKey As Long, strPath As String, strValue As String)
'Usage:
'------------
'Call GetRegistryString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\", App.EXEName)
'------------
Dim r
Dim lValueType
Dim keyhand As Long
Dim datatype As Long
Dim lResult As Long
Dim strBuf As String
Dim lDataBufSize As Long
Dim intZeroPos As Integer
r = RegOpenKey(hKey, strPath, keyhand)
lResult = RegQueryValueEx(keyhand, strValue, 0&, lValueType, ByVal 0&, lDataBufSize)
If lValueType = REG_SZ Then
strBuf = String(lDataBufSize, " ")
lResult = RegQueryValueEx(keyhand, strValue, 0&, 0&, ByVal strBuf, lDataBufSize)
If lResult = ERROR_SUCCESS Then
intZeroPos = InStr(strBuf, Chr$(0))
If intZeroPos > 0 Then
RegistryGet = Left$(strBuf, intZeroPos - 1)
Else
RegistryGet = strBuf
End If
End If
End If
End Function
Public Sub RegistrySave(hKey As Long, strPath As String, strValue As String, strdata As String)
'Usage:
'----------
'Call SaveRegistryString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\", App.EXEName, App.EXEName)
'----------
Dim keyhand As Long
Dim r As Long
r = RegCreateKey(hKey, strPath, keyhand)
r = RegSetValueEx(keyhand, strValue, 0, REG_SZ, ByVal strdata, Len(strdata))
r = RegCloseKey(keyhand)
End Sub
Public Function IntelliCrypt_EnCrypt(InputString As String)
Dim z As String
z = EnHex(InputString)
If Not InputString = DeHex(z) Then
MsgBox "DeCrypt Error! Could not verfy the decrypted string! CODE: VERIFY_ERR", 48, "Interworks IntelliCrypt"
InputString = "VERIFY_ERR"
Exit Function
End If
IntelliCrypt_EnCrypt = z
End Function
Public Function IntelliCrypt_DeCrypt(DeString) As String
Dim w As String
w = DeString
IntelliCrypt_DeCrypt = DeHex(w)
End Function

Public Function BigShiftLeft(value1 As String, shifts As Integer) As String
Dim tempstr As String
Dim loopit As Integer, loopinner As Integer
Dim tempnum As Integer

    shifts = shifts Mod 32
    
    If shifts = 0 Then
        BigShiftLeft = value1
        Exit Function
    End If

    value1 = Right$(value1, 8)
    tempstr = String$(8 - Len(value1), "0") + value1
    value1 = ""

    ' Convert to binary
    For loopit = 1 To 8
        tempnum = Val("&H" + Mid$(tempstr, loopit, 1))
        For loopinner = 3 To 0 Step -1
            If tempnum And 2 ^ loopinner Then
                value1 = value1 + "1"
            Else
                value1 = value1 + "0"
            End If
        Next loopinner
    Next loopit
    
    For I = 1 To shifts
        For j = 1 To 32
            Mid(value1, j, 1) = Mid(value1, j + 1, 1)
        Next j
        If Not Mid(value1, 32, 1) = "0" Then Mid(value1, 32, 1) = "0"
    Next I
    tempstr = value1

    ' And convert back to hex
    value1 = ""
    For loopit = 0 To 7
        tempnum = 0
        For loopinner = 0 To 3
            If Val(Mid$(tempstr, 4 * loopit + loopinner + 1, 1)) Then
                tempnum = tempnum + 2 ^ (3 - loopinner)
            End If
        Next loopinner
        value1 = value1 + Hex$(tempnum)
    Next loopit

    BigShiftLeft = Right(value1, 8)
End Function
Public Function BigShiftRight(value1 As String, shifts As Integer) As String
Dim tempstr As String
Dim loopit As Integer, loopinner As Integer
Dim tempnum As Integer

    shifts = shifts Mod 32
    
    If shifts = 0 Then
        BigShiftRight = value1
        Exit Function
    End If

    value1 = Right$(value1, 8)
    tempstr = String$(8 - Len(value1), "0") + value1
    value1 = ""

    ' Convert to binary
    For loopit = 1 To 8
        tempnum = Val("&H" + Mid$(tempstr, loopit, 1))
        For loopinner = 3 To 0 Step -1
            If tempnum And 2 ^ loopinner Then
                value1 = value1 + "1"
            Else
                value1 = value1 + "0"
            End If
        Next loopinner
    Next loopit
    
    For I = 1 To shifts
        For j = 32 To 2 Step -1
            Mid(value1, j, 1) = Mid(value1, j - 1, 1)
        Next j
        If Not Mid(value1, 1, 1) = "0" Then Mid(value1, 1, 1) = "0"
    Next I
    tempstr = value1

    ' And convert back to hex
    value1 = ""
    For loopit = 0 To 7
        tempnum = 0
        For loopinner = 0 To 3
            If Val(Mid$(tempstr, 4 * loopit + loopinner + 1, 1)) Then
                tempnum = tempnum + 2 ^ (3 - loopinner)
            End If
        Next loopinner
        value1 = value1 + Hex$(tempnum)
    Next loopit

    BigShiftRight = Right(value1, 8)
End Function


Public Function RotLeft(ByVal value1 As String, ByVal rots As Integer) As String
Dim tempstr As String
Dim loopit As Integer, loopinner As Integer
Dim tempnum As Integer

    rots = rots Mod 32
    
    If rots = 0 Then
        RotLeft = value1
        Exit Function
    End If

    value1 = Right$(value1, 8)
    tempstr = String$(8 - Len(value1), "0") + value1
    value1 = ""

    ' Convert to binary
    For loopit = 1 To 8
        tempnum = Val("&H" + Mid$(tempstr, loopit, 1))
        For loopinner = 3 To 0 Step -1
            If tempnum And 2 ^ loopinner Then
                value1 = value1 + "1"
            Else
                value1 = value1 + "0"
            End If
        Next loopinner
    Next loopit
    tempstr = Mid$(value1, rots + 1) + Left$(value1, rots)

    ' And convert back to hex
    value1 = ""
    For loopit = 0 To 7
        tempnum = 0
        For loopinner = 0 To 3
            If Val(Mid$(tempstr, 4 * loopit + loopinner + 1, 1)) Then
                tempnum = tempnum + 2 ^ (3 - loopinner)
            End If
        Next loopinner
        value1 = value1 + Hex$(tempnum)
    Next loopit

    RotLeft = Right(value1, 8)
End Function
Public Function RotRight(ByVal value1 As String, ByVal rots As Integer) As String
Dim tempstr As String
Dim loopit As Integer, loopinner As Integer
Dim tempnum As Integer

    rots = rots Mod 32
    
    If rots = 0 Then
        RotRight = value1
        Exit Function
    End If

    value1 = Right$(value1, 8)
    tempstr = String$(8 - Len(value1), "0") + value1
    value1 = ""

    ' Convert to binary
    For loopit = 1 To 8
        tempnum = Val("&H" + Mid$(tempstr, loopit, 1))
        For loopinner = 3 To 0 Step -1
            If tempnum And 2 ^ loopinner Then
                value1 = value1 + "1"
            Else
                value1 = value1 + "0"
            End If
        Next loopinner
    Next loopit
    tempstr = Right$(value1, rots) + Mid$(value1, 1, Len(value1) - rots)

    ' And convert back to hex
    value1 = ""
    For loopit = 0 To 7
        tempnum = 0
        For loopinner = 0 To 3
            If Val(Mid$(tempstr, 4 * loopit + loopinner + 1, 1)) Then
                tempnum = tempnum + 2 ^ (3 - loopinner)
            End If
        Next loopinner
        value1 = value1 + Hex$(tempnum)
    Next loopit

    RotRight = Right(value1, 8)
End Function
Public Function BigAdd(ByVal value1 As String, ByVal value2 As String) As String
Dim valueans As String
Dim loopit As Integer, tempnum As Integer

    tempnum = Len(value1) - Len(value2)
    If tempnum < 0 Then
        value1 = Space$(Abs(tempnum)) + value1
    ElseIf tempnum > 0 Then
        value2 = Space$(Abs(tempnum)) + value2
    End If

    tempnum = 0
    For loopit = Len(value1) To 1 Step -1
        tempnum = tempnum + Val("&H" + Mid$(value1, loopit, 1)) + Val("&H" + Mid$(value2, loopit, 1))
        valueans = Hex$(tempnum Mod 16) + valueans
        tempnum = Int(tempnum / 16)
    Next loopit

    If tempnum <> 0 Then
        valueans = Hex$(tempnum) + valueans
    End If

    BigAdd = Right(valueans, 8)
End Function
Function BigAND(ByVal value1 As String, ByVal value2 As String) As String
Dim valueans As String
Dim loopit As Integer, tempnum As Integer

    tempnum = Len(value1) - Len(value2)
    If tempnum < 0 Then
        value2 = Mid$(value2, Abs(tempnum) + 1)
    ElseIf tempnum > 0 Then
        value1 = Mid$(value1, tempnum + 1)
    End If

    For loopit = 1 To Len(value1)
        valueans = valueans + Hex$(Val("&H" + Mid$(value1, loopit, 1)) And Val("&H" + Mid$(value2, loopit, 1)))
    Next loopit

    BigAND = valueans
End Function
Function BigMod32Add(ByVal value1 As String, ByVal value2 As String) As String
    BigMod32Add = Right$(BigAdd(value1, value2), 8)
End Function
Function BigNOT(ByVal value1 As String) As String
Dim valueans As String
Dim loopit As Integer

    value1 = Right$(value1, 8)
    value1 = String$(8 - Len(value1), "0") + value1
    For loopit = 1 To 8
        valueans = valueans + Hex$(15 Xor Val("&H" + Mid$(value1, loopit, 1)))
    Next loopit

    BigNOT = valueans
End Function
Function BigOR(ByVal value1 As String, ByVal value2 As String) As String
Dim valueans As String
Dim loopit As Integer, tempnum As Integer

    tempnum = Len(value1) - Len(value2)
    If tempnum < 0 Then
        valueans = Left$(value2, Abs(tempnum))
        value2 = Mid$(value2, Abs(tempnum) + 1)
    ElseIf tempnum > 0 Then
        valueans = Left$(value1, Abs(tempnum))
        value1 = Mid$(value1, tempnum + 1)
    End If

    For loopit = 1 To Len(value1)
        valueans = valueans + Hex$(Val("&H" + Mid$(value1, loopit, 1)) Or Val("&H" + Mid$(value2, loopit, 1)))
    Next loopit

    BigOR = valueans
End Function
Function BigXOR(ByVal value1 As String, ByVal value2 As String) As String
Dim valueans As String
Dim loopit As Integer, tempnum As Integer

    tempnum = Len(value1) - Len(value2)
    If tempnum < 0 Then
        valueans = Left$(value2, Abs(tempnum))
        value2 = Mid$(value2, Abs(tempnum) + 1)
    ElseIf tempnum > 0 Then
        valueans = Left$(value1, Abs(tempnum))
        value1 = Mid$(value1, tempnum + 1)
    End If

    For loopit = 1 To Len(value1)
        valueans = valueans + Hex$(Val("&H" + Mid$(value1, loopit, 1)) Xor Val("&H" + Mid$(value2, loopit, 1)))
    Next loopit

    BigXOR = Right(valueans, 8)
End Function
Public Function DeHex(InP As String) As String
For I = 1 To Len(InP) Step 2
    X = X & Chr(Val("&H" & Mid(InP, I, 2)))
Next I
DeHex = X
End Function
Public Function EnHex(X As String) As String
For I = 1 To Len(X)
    v = Hex(Asc(Mid(X, I, 1)))
    If Len(v) = 1 Then v = "0" & v
    InP = InP & v
Next I
EnHex = InP
End Function



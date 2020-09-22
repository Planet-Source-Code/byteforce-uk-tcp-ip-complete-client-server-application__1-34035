Attribute VB_Name = "EChatSubs"
Public Sub SendChatLine(TextToSend As String)
If TextToSend = "" Then MsgBox "Please type something to say in the box!", 16: Exit Sub
Server.WinSock.SendData "C" + PROBas.IntelliCrypt_EnCrypt(Server.WinSock.LocalIP + ": " + TextToSend)
End Sub



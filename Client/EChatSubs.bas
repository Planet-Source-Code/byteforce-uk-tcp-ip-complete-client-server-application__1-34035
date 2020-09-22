Attribute VB_Name = "EChatSubs"
Public Sub SendChatLine(TextToSend As String)
If TextToSend = "" Then MsgBox "Please type something to say in the box!", 16: Exit Sub
Client.Winsock.SendData "C" + PROBas.IntelliCrypt_EnCrypt(Client.Winsock.LocalIP + ": " + TextToSend)
End Sub




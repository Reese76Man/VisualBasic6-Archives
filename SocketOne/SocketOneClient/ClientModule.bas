Attribute VB_Name = "ClientModule"
Option Explicit

Public Sub SocketSend(OutData As String)
    'For Future Implimentation of Encryption
    frmMain.Wsck.SendData OutData
End Sub

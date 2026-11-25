Attribute VB_Name = "SocketOneServer"
Option Explicit
Public RS As ADODB.Recordset
Public SocketCount As Integer              'How many Sockets incremented so far?
Public SocketOpen(0 To 32768) As Boolean   'Is Socket Open or Closed?
Public SocketLogin(0 To 32768) As Long     'User's Database ID of Socket

Public Function UserStats(SocketID As Integer) As Boolean
    RS.MoveFirst
    RS.Find "UserID = " & Trim(Str(SocketLogin(SocketID))) & "", 0, adSearchForward, BookmarkEnum.adBookmarkFirst
    If Not RS.EOF Then UserStats = True Else UserStats = False
End Function

Public Function UserLogin(UName As String, UPass As String) As Long
   'UserLogin > 0     'Means that Login Found and Password Correct
   'UserLogin = 0     'Means that Login Found but Password Incorrect
    UserLogin = -1    'Means that Login not Found.
    
    RS.MoveFirst
    RS.Find "UserName = '" & Trim(UName) & "'", 0, adSearchForward, BookmarkEnum.adBookmarkFirst
        
    If Not RS.EOF Then
        UserLogin = Val(RS!UserID)
        If Trim(UCase(RS!UserPass)) <> Trim(UCase(UPass)) Then UserLogin = 0
    End If
End Function

Public Function Login2User(ByVal ULogin As String) As String
    'Extracts UserName from Login Packet
    Login2User = Trim(Left(ULogin, InStr(ULogin, Chr(0)) - 1))
End Function

Public Function Login2Pass(ByVal UPass As String) As String
    'Extracts User Password from Login Packet
    Login2Pass = Trim(Mid(UPass, (InStr(UPass, Chr(0)) + 1), Len(UPass) - InStr(UPass, Chr(0))))
End Function

Public Sub ServerSend(Index As Integer, Msg As String)
    'Routine to send a dynamic/custom message from server to client.
    ' which will be displayed in a message box.
    SocketSend Index, (ServerMessage & Trim(Msg))
End Sub

Public Sub ManualClose(Index As Integer)
Remove (frmMain.WSCK(Index).RemoteHostIP)
frmMain.WSCK(Index).Close
Unload frmMain.WSCK(Index)
SocketOpen(Index) = False
SocketLogin(Index) = 0
End Sub

Public Sub Remove(Sts As String)
Dim T As Integer
For T = 1 To frmMain.LSTConnects.ListCount
    If frmMain.LSTConnects.List(T - 1) = Sts Then frmMain.LSTConnects.RemoveItem (T - 1)
Next T
End Sub

Public Sub SocketSend(Index As Integer, OutData As String)
    'For future implimentation of Encryption
    frmMain.WSCK(Index).SendData OutData
    DoEvents
End Sub

Public Sub UserConnected(Index As Integer, IDex As Long)
    'IDex is the DATABASE UserID / Index is the User's Socket Index
    'Future routines when user connects to server
End Sub

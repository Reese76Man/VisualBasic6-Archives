Attribute VB_Name = "SocketOneModule"
'Client Codes (Codes issued by the Client Side)
Public Const LoginCode = "L:"     'Login Code:LoginName + NULL + Password

'Server Codes (Codes issued by the Server Side)
Public Const RequestLogin = "R:"  'Request for LoginName + NULL + Password
Public Const ServerMessage = "M:" 'Server is sending a dynamic message to Client
Public Const Failure = "D:"       'Disconnect Code to Client Side
Public Const Success = "S:"       'Sent to Client to show success of login.

Public Function CodeCheck(ByVal Sts As String)
    CodeCheck = UCase(Left(Sts, 2))
End Function

Public Function CodeLess(ByVal Sts As String)
    CodeLess = Right(Sts, Len(Sts) - 2)
End Function



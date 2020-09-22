Attribute VB_Name = "ModConnection"
Public Declare Function InternetGetConnectedStateEx Lib "wininet.dll" Alias "InternetGetConnectedStateExA" _
   (ByRef lpdwFlags As Long, _
   ByVal lpszConnectionName As String, _
   ByVal dwNameLen As Long, _
   ByVal dwReserved As Long _
   ) As Long

Public Enum EIGCInternetConnectionState
   INTERNET_CONNECTION_MODEM = &H1&
   INTERNET_CONNECTION_LAN = &H2&
   INTERNET_CONNECTION_PROXY = &H4&
   INTERNET_CONNECTION_OFFLINE = &H20&
   End Enum

Public Property Get InternetConnected( _
     Optional ByRef eConnectionInfo As EIGCInternetConnectionState, _
     Optional ByRef sConnectionName As String _
   ) As Boolean
Dim dwFlags As Long
Dim sNameBuf As String
Dim lR As Long
Dim iPos As Long
   sNameBuf = String$(513, 0)
   lR = InternetGetConnectedStateEx(dwFlags, sNameBuf, 512, 0&)
   eConnectionInfo = dwFlags
   iPos = InStr(sNameBuf, vbNullChar)
   If iPos > 0 Then
     sConnectionName = Left$(sNameBuf, iPos - 1)
   ElseIf Not sNameBuf = String$(513, 0) Then
     sConnectionName = sNameBuf
   End If
     InternetConnected = (lR = 1)
End Property




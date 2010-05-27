VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CJSONSocket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim mId As Long

Dim WithEvents theSocket As CSocket
Attribute theSocket.VB_VarHelpID = -1

Private Sub Class_Initialize()
'//
End Sub

Private Sub Class_Terminate()
'//
End Sub

Public Sub Accept(ByVal Id As Long)

    g_Debug "CJSONSocket.Accept(): Id=0x" & g_HexStr(Id), LEMON_LEVEL_PROC

    mId = Id
    Set theSocket = New CSocket
    theSocket.Accept Id

End Sub

Private Sub theSocket_OnClose()

    g_Debug "CJSONSocket.OnClose(): id=0x" & g_HexStr(mId), LEMON_LEVEL_PROC

End Sub

Private Sub theSocket_OnConnect()

    g_Debug "CJSONSocket.OnConnect(): id=0x" & g_HexStr(mId), LEMON_LEVEL_PROC

End Sub

Private Sub theSocket_OnConnectionRequest(ByVal requestID As Long)

    g_Debug "CJSONSocket.OnConnectionRequest(): id=0x" & g_HexStr(mId) & " requestID=0x" & g_HexStr(requestID), LEMON_LEVEL_PROC

End Sub

Private Sub theSocket_OnDataArrival(ByVal bytesTotal As Long)
Dim sz As String

    g_Debug "CJSONSocket.OnDataArrival(): id=0x" & g_HexStr(mId) & " bytesTotal=" & CStr(bytesTotal), LEMON_LEVEL_PROC

    theSocket.GetData sz

    If g_ConfigGet("listen_for_json") Then _
        uDo sz

    theSocket.SendData "HTTP/1.0 200 OK" & vbCrLf & Format$(Now(), "ddd, d mmm yyyy hh:mm:ss") & " GMT" & vbCrLf & _
                       "Content-Type: text/html" & vbCrLf & "Content-Length: 0" & vbCrLf

    theSocket.CloseSocket

End Sub

Private Sub theSocket_OnError(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

    g_Debug "CJSONSocket.OnError(): id=0x" & g_HexStr(mId) & " Description=" & Description, LEMON_LEVEL_PROC

End Sub

Private Sub theSocket_OnSendComplete()

    g_Debug "CJSONSocket.OnSendComplete(): id=0x" & g_HexStr(mId), LEMON_LEVEL_PROC

End Sub

Private Sub theSocket_OnSendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)

    g_Debug "CJSONSocket.OnSendProgress(): id=0x" & g_HexStr(mId) & " bytesSent=" & CStr(bytesSent), LEMON_LEVEL_PROC

End Sub

Private Sub uDo(ByVal HTTP As String)
Dim po As BJSONObject

    Set po = New BJSONObject
    If Not po.SetFromHTTP(HTTP) Then
        g_Debug "CJSONSocket.uDo(): failed to create JSON object", LEMON_LEVEL_CRITICAL
        Exit Sub

    End If

Dim pv As BJSONValue

    If Not po.Find("action", pv) Then
        g_Debug "CJSONSocket.uDo(): 'action' value missing", LEMON_LEVEL_CRITICAL
        Exit Sub

    End If

    If pv.AsLong <> 1 Then
        g_Debug "CJSONSocket.uDo(): incorrect 'action' value", LEMON_LEVEL_CRITICAL
        Exit Sub

    End If

Dim tt As String
Dim tx As String

    If po.Find("description", pv) Then _
        tx = pv.AsString()

    If po.Find("title", pv) Then _
        tt = pv.AsString()

    snShowMessageEx SNARL_CLASS_JSON, tt, tx

End Sub





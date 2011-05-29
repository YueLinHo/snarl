VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TRemoteConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

    '/*********************************************************************************************
    '/
    '/  File:           TRemoteConnection.cls
    '/
    '/  Description:    Represents an active SNP/TCP, SNP2.0/TCP or GNTP connection
    '/
    '/  © 2010 full phat products
    '/
    '/  This file may be used under the terms of the Simplified BSD Licence
    '/
    '*********************************************************************************************/

' /* FIXFIXFIX */
    ' /* SNP1.x support is currently broken */
' /* FIXFIXFIX */


Private Enum T_NETWORK_ACTIONS
    T_ACTION_NULL
    T_ACTION_REGISTER
    T_ACTION_REVOKE
    T_ACTION_ADD_CLASS
    T_ACTION_SHOW_NOTIFICATION
    ' /* SNP 1.1 */
    T_ACTION_HELLO
    T_ACTION_VERSION

End Enum

Private Type T_NETWORK_INFO
    Valid As Boolean
    SenderIP As String
    action As T_NETWORK_ACTIONS
    App As String
    Class As String
    Title As String             ' // used for application name in 'register' action
    Text As String
    Icon As String              ' // V40.10 (SNP1.1)
    DefaultAck As String        ' // V40.10 (SNP1.1)
    Timeout As Long
    Version As Single

End Type

Dim mId As Long
'Dim mContent As String
'Dim mIsGNTP As Boolean

Dim mType As Long               ' // actually the port number
Dim mGNTPRequest As String

Dim WithEvents theSocket As CSocket
Attribute theSocket.VB_VarHelpID = -1

Private Sub Class_Initialize()
'//
End Sub

Private Sub Class_Terminate()

    If Not (theSocket Is Nothing) Then _
        theSocket.CloseSocket

End Sub

Public Sub Accept(ByVal Id As Long, ByVal PortNumber As Long) '// IsGNTP As Boolean)

    g_Debug "TRemoteConnection.Accept(): Id=0x" & g_HexStr(Id) & " Type=" & CStr(PortNumber), LEMON_LEVEL_PROC

    mId = Id
    mType = PortNumber
'    mIsGNTP = IsGNTP

    Set theSocket = New CSocket
    theSocket.Accept Id

End Sub

Private Sub theSocket_OnClose()

    g_Debug "TRemoteConnection.OnClose(): id=0x" & g_HexStr(mId), LEMON_LEVEL_PROC

End Sub

Private Sub theSocket_OnConnect()

    g_Debug "TRemoteConnection.OnConnect(): id=0x" & g_HexStr(mId), LEMON_LEVEL_PROC

End Sub

Private Sub theSocket_OnConnectionRequest(ByVal requestID As Long)

    g_Debug "TRemoteConnection.OnConnectionRequest(): id=0x" & g_HexStr(mId) & " requestID=0x" & g_HexStr(requestID), LEMON_LEVEL_PROC

End Sub

Private Sub theSocket_OnDataArrival(ByVal bytesTotal As Long)

    If mType = GNTP_DEFAULT_PORT Then
        uDoGNTP bytesTotal

    ElseIf mType = SNP_DEFAULT_PORT Then
        uDoSNP bytesTotal

    ElseIf mType = JSON_DEFAULT_PORT Then
        uDoJSON bytesTotal

    End If

End Sub

Private Sub theSocket_OnError(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

    g_Debug "TRemoteConnection.OnError(): id=0x" & g_HexStr(mId) & " Description=" & Description, LEMON_LEVEL_PROC

End Sub

Private Sub theSocket_OnSendComplete()

    g_Debug "TRemoteConnection.OnSendComplete(): id=0x" & g_HexStr(mId), LEMON_LEVEL_PROC

End Sub

Private Sub theSocket_OnSendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)

    g_Debug "TRemoteConnection.OnSendProgress(): id=0x" & g_HexStr(mId) & " bytesSent=" & CStr(bytesSent), LEMON_LEVEL_PROC

End Sub

'Private Sub uProcessSNP(ByVal Request As String)
'
'    g_Debug "TRemoteConnection.uProcessSNP(): incoming notification from " & theSocket.RemoteHostIP & ":" & theSocket.RemotePort & "..."
'
'    If g_ConfigGet("listen_for_snarl") = "0" Then
'        g_Debug "TRemoteConnection.uProcessSNP(): ignored (user disabled incoming notifications)"
'        theSocket.SendData uCreateReply(SNP_ERROR_NOT_RUNNING)
'        Exit Sub
'
'    End If
'
'    If LCase$(g_SafeLeftStr(Request, 6)) = "snp://" Then
'        uTranslateSNP2 g_SafeRightStr(Request, Len(Request) - 6)
'
'    Else
'        uTranslateSNP Request
'
'    End If
'
'End Sub

Private Sub uTranslateSNP2(ByVal Request As String)

    ' /* syntax is: snp://<action>[?arg=value[&arg=value]] CRLF
    '    on entry the 'snp://' prefix will have been stripped */

Dim sCommand As String
Dim ppd As BPackedData
Dim i As Long

    Set ppd = New BPackedData                           ' // g_DoAction() expects this, even if it's empty

    ' /* find the command */

    i = InStr(Request, "?")
    If i = 0 Then
        ' /* entire line is command */
        sCommand = LCase$(Request)

    Else
        ' /* get the command */
        sCommand = LCase$(g_SafeLeftStr(Request, i - 1))

        ' /* get the args */
        Request = g_SafeRightStr(Request, Len(Request) - i)

        ' /* translate from SNP syntax into Win32 syntax */
        Request = Replace(Request, "&", "#?")
        Request = Replace(Request, "=", "::")

        ' /* create the BPackedData */
        ppd.SetTo g_DecodeURL(Request)

    End If

    ' /* pre-processing */

Dim sz As String

    Select Case sCommand

    Case "register", "reg"
        ' /* use default icon if none provided */
        If Not ppd.Exists("icon") Then _
            ppd.Add "icon", g_MakePath(App.Path) & "etc\icons\snp-no_icon.png"

    Case "notify"
        ' /* if encoded icon data exists, extract it now */
        If ppd.Exists("icon64") Then
            sz = uDecodeBase64(Replace$(ppd.ValueOf("icon64"), "%", "="))
            ppd.Add "icon", sz

        End If

    End Select

Dim hr As Long

    ' /* process the command */

    hr = g_DoAction(sCommand, Val(ppd.ValueOf("token")), ppd, , theSocket)
    If hr <> 0 Then
        ' /* success */
        uSNP2Reply SNARL_SUCCESS, IIf(hr <> -1, CStr(hr), "")

    Else
        ' /* failed */
        uSNP2Reply GetProp(ghWndMain, "last_error")

    End If

    ' /* post-processing */

    If (sCommand = "notify") And (sz <> "") Then _
        DeleteFile sz

End Sub

Private Sub uTranslateSNP(ByVal Request As String)
Dim pPacket As T_NETWORK_INFO
Dim sz As String
Dim pc As TAlert
Dim pa As TApp

    On Error GoTo er

    With New CConfFile
        If .SetFromText(Request, "#?") Then
            sz = .GetValueWithDefault("type")
            If sz = "SNP" Then
                pPacket.Version = Val(.GetValueWithDefault("version", "0"))
                If pPacket.Version >= 1 Then
                    ' /* required for all packet types */
                    pPacket.SenderIP = theSocket.RemoteHost
                    pPacket.App = .GetValueWithDefault("app")
                    If pPacket.App <> "" Then _
                        pPacket.App = pPacket.App & " on " & pPacket.SenderIP

                    sz = .GetValueWithDefault("action")

                    Select Case sz
                    Case "register"
                        pPacket.action = T_ACTION_REGISTER
                        pPacket.Valid = (pPacket.App <> "")

                    Case "unregister"
                        pPacket.action = T_ACTION_REVOKE
                        pPacket.Valid = (pPacket.App <> "")

                    Case "add_class"
                        pPacket.action = T_ACTION_ADD_CLASS
                        pPacket.Title = .GetValueWithDefault("title")
                        pPacket.Class = .GetValueWithDefault("class")
                        ' /* V40.10 (SNP1.1) */
                        pPacket.Icon = .GetValueWithDefault("icon")
                        pPacket.Valid = (pPacket.App <> "") And (pPacket.Class <> "")

                    Case "notification"
                        pPacket.action = T_ACTION_SHOW_NOTIFICATION
                        pPacket.Class = .GetValueWithDefault("class")
                        pPacket.Title = .GetValueWithDefault("title")
                        pPacket.Text = .GetValueWithDefault("text")
                        pPacket.Timeout = Val(.GetValueWithDefault("timeout", "0"))
                        ' /* V40.10 (SNP1.1) */
                        pPacket.Icon = .GetValueWithDefault("icon")
                        pPacket.DefaultAck = .GetValueWithDefault("default_ack")
                        pPacket.Valid = True

                    Case "hello"
                        ' /* introduced in SNP1.1 */
                        pPacket.action = T_ACTION_HELLO
                        pPacket.Valid = True

                    Case "version"
                        ' /* introduced in SNP1.1 */
                        pPacket.action = T_ACTION_VERSION
                        pPacket.Valid = True

                    Case Else
                        g_Debug "TRemoteConnection.uTranslateSNP(): unknown action '" & sz & "'", LEMON_LEVEL_CRITICAL

                    End Select

                Else
                    g_Debug "TRemoteConnection.uTranslateSNP(): unsupported version '" & CStr(pPacket.Version) & "'", LEMON_LEVEL_CRITICAL

                End If
            Else
                g_Debug "TRemoteConnection.uTranslateSNP(): bad packet type '" & sz & "'", LEMON_LEVEL_CRITICAL

            End If
        End If

    End With

    ' /* check here that the packet is valid (i.e. version, type, etc.) */

    If Not pPacket.Valid Then
        theSocket.SendData uCreateReply(SNP_ERROR_BAD_PACKET)
        theSocket.CloseSocket
        Exit Sub

    End If

    g_Debug "TRemoteConnection.uTranslateSNP(): packet is valid (ver=" & pPacket.Version & ")"

    ' /* do some general safety checking first */

    If (g_AppRoster Is Nothing) Or (g_NotificationRoster Is Nothing) Then
        g_Debug "TRemoteConnection.uTranslateSNP(): action '" & pPacket.action & "': app_roster or notification_roster not found", LEMON_LEVEL_CRITICAL
        theSocket.SendData uCreateReply(SNP_ERROR_FAILED)
        theSocket.CloseSocket
        Exit Sub

    End If

    ' /* figure out the packet action */

Dim dw As Long

    With pPacket

        Select Case .action

        Case T_ACTION_REGISTER
            dw = g_AppRoster.IndexOfPidAndName(-1, .App)
            If dw <> 0 Then
                ' /* already registered */
                g_Debug "TRemoteConnection.uTranslateSNP(): remote app '" & .App & "' is already registered", LEMON_LEVEL_CRITICAL
                theSocket.SendData uCreateReply(SNP_ERROR_ALREADY_REGISTERED)

            Else
                ' /* use the special "Add2()" method which allows us to register multiple apps using the same PID - in
                '    this case -1, which represents remote apps */
                g_AppRoster.OldAdd .App, 0, 0, g_MakePath(App.Path) & "etc\icons\remote_app.png", g_MakePath(App.Path) & "etc\icons\remote.png", -1
                g_Debug "TRemoteConnection.uTranslateSNP(): remote app '" & .App & "' registered"
                theSocket.SendData uCreateReply(SNP_SUCCESS)

            End If


        Case T_ACTION_REVOKE
            dw = g_AppRoster.IndexOfPidAndName(-1, .App)
            If dw = 0 Then
                ' /* not registered */
                g_Debug "TRemoteConnection.uTranslateSNP(): remote app '" & .App & "' is not registered", LEMON_LEVEL_CRITICAL
                theSocket.SendData uCreateReply(SNP_ERROR_NOT_REGISTERED)

            Else
                ' /* remove safely */
                g_AppRoster.Remove dw
                frmAbout.bUpdateAppList
                g_Debug "TRemoteConnection.uTranslateSNP(): remote app '" & .App & "' unregistered"
                theSocket.SendData uCreateReply(SNP_SUCCESS)

            End If


        Case T_ACTION_ADD_CLASS
            dw = g_AppRoster.IndexOfPidAndName(-1, .App)
            If dw = 0 Then
                ' /* not registered */
                g_Debug "TRemoteConnection.uTranslateSNP(): remote app '" & .App & "' is not registered", LEMON_LEVEL_CRITICAL
                theSocket.SendData uCreateReply(SNP_ERROR_NOT_REGISTERED)

            Else
                ' /* found the app */
                Set pa = g_AppRoster.AppAt(dw)
                If Not pa.FindAlert(.Class, pc) Then
                    ' /* okay to add the class */
                    If pa.AddAlert(.Class, .Title, pc) = M_OK Then _
                        pc.DefaultIcon = pPacket.Icon

                    frmAbout.bUpdateAppList
                    g_Debug "TRemoteConnection.uTranslateSNP(): class '" & .Class & "' add"
                    theSocket.SendData uCreateReply(SNP_SUCCESS)

                Else
                    ' /* class already exists */
                    g_Debug "TRemoteConnection.uTranslateSNP(): class '" & .Class & "' is already registered", LEMON_LEVEL_CRITICAL
                    theSocket.SendData uCreateReply(SNP_ERROR_CLASS_ALREADY_EXISTS)

                End If

            End If


        Case T_ACTION_SHOW_NOTIFICATION
            .Text = Replace$(.Text, "\n", vbCrLf)
            .Title = Replace$(.Title, "\n", vbCrLf)

            If pPacket.Icon = "" Then _
                pPacket.Icon = g_MakePath(App.Path) & "etc\icons\remote.png"

            dw = g_AppRoster.IndexOfPidAndName(-1, .App)
            If dw = 0 Then
                ' /* not registered: show anyway using the 'anonymous network notification' class */
                g_Debug "TRemoteConnection.uTranslateSNP(): remote app '" & .App & "' is not registered", LEMON_LEVEL_CRITICAL

'FIXFIXFIX
                ' /* we no longer include frmAbout.hWnd and WM_REMOTENOTIFY with this */
                g_PrivateNotify SNARL_CLASS_ANON_NET, .Title, .Text, .Timeout, .Icon, , .DefaultAck, , NF_REMOTE
'FIXFIXFIX
                g_Debug "TRemoteConnection.uTranslateSNP(): snShowMessageEx() returned 0x" & g_HexStr(dw)

            Else
                ' /* app found */
                Set pa = g_AppRoster.AppAt(dw)
                dw = pa.Show(.Class, .Title, .Text, .Timeout, pPacket.Icon) ', frmAbout.hWnd, WM_REMOTENOTIFY)
                g_Debug "TRemoteConnection.uTranslateSNP(): Show() returned 0x" & g_HexStr(dw)

            End If

            If dw > 0 Then
                theSocket.SendData uCreateReply(SNP_SUCCESS, CStr(dw))
'                frmAbout.AddRemoteNotification dw, Me

            Else
                theSocket.SendData uCreateReply(SNP_ERROR_FAILED)

            End If


        Case T_ACTION_HELLO
            ' /* send release back */
            theSocket.SendData uCreateReply(SNP_SUCCESS, App.Title & " " & App.Comments)

        Case T_ACTION_VERSION
            ' /* send version back */
            theSocket.SendData uCreateReply(SNP_SUCCESS, CStr(App.Major) & "." & CStr(App.Revision))

        Case Else
            ' /* unknown action */
            g_Debug "TRemoteConnection.uTranslateSNP(): action '" & .action & "' is unknown", LEMON_LEVEL_CRITICAL
            theSocket.SendData uCreateReply(SNP_ERROR_UNKNOWN_COMMAND)

        End Select

    End With

    Exit Sub

er:
    g_Debug "TRemoteConnection.uTranslateSNP(): " & err.Description, LEMON_LEVEL_CRITICAL
    theSocket.SendData uCreateReply(SNP_ERROR_FAILED)
    theSocket.CloseSocket

End Sub


'                If Not pa.FindAlert(.Class, pc) Then
'                    g_Debug "TRemoteConnection.uTranslate(): T_ACTION_SHOW_NOTIFICATION: class '" & .Class & "' is not registered", LEMON_LEVEL_WARNING
'                    If pa.CountAlerts = 0 Then
'                        g_Debug "TRemoteConnection.uTranslate(): T_ACTION_SHOW_NOTIFICATION: app '" & .App & "' has no classes!", LEMON_LEVEL_CRITICAL
'                        Exit Sub
'
'                    Else
'                        Set pc = pa.AlertAt(1)
'
'                    End If
'
'                End If
'
'                ' /* found the class (or using the catch-all) */
'
'                pc.Show .Title, .Text, .Timeout, g_MakePath(App.Path) & "etc\icons\remote.png"

Private Function uCreateReply(ByVal ReturnCode As Long, Optional ByVal Data As String) As String
Dim sz As String

    Select Case ReturnCode

    ' /* 0xx - Success code */

    Case SNP_SUCCESS
        sz = "OK"

    ' /* 1xx - Protocol errors */

    Case SNP_ERROR_FAILED
        sz = "Internal error"

    Case SNP_ERROR_UNKNOWN_COMMAND
        sz = "Unknown command"

    Case SNP_ERROR_TIMED_OUT
        sz = "Timed out"

    Case SNP_ERROR_IS_BUSY
        sz = "Server too busy"

    Case SNP_ERROR_BAD_PACKET
        sz = "Bad packet"


    ' /* 2xx - Server errors */

    Case SNP_ERROR_NOT_RUNNING
        sz = "Not running"

    Case SNP_ERROR_NOT_REGISTERED
        sz = "Application is not registered"

    Case SNP_ERROR_ALREADY_REGISTERED
        sz = "Application is already registered"

    Case SNP_ERROR_CLASS_ALREADY_EXISTS
        sz = "Class is already registered"


    ' /* 3xx - Notifications */

    Case SNP_NOTIFY_CANCELLED
        sz = "Notification cancelled"

    Case SNP_NOTIFY_TIMED_OUT
        sz = "Notification timed out"

    Case SNP_NOTIFY_ACK
        sz = "Notification acknowledged"

    Case SNP_NOTIFY_MENU
        sz = "Notification menu selected"

    Case SNP_NOTIFY_MIDDLE_BUTTON
        sz = "Notification middle button clicked"

    Case SNP_NOTIFY_CLOSED
        sz = "Notification closed"

    End Select

    uCreateReply = "SNP/1.1/" & CStr(ReturnCode) & "/" & sz & IIf(Data <> "", "/" & Data, "") & vbCrLf

End Function

Private Sub uSNP2Reply(ByVal StatusCode As Long, Optional ByVal Data As String)

    If (theSocket Is Nothing) Then _
        Exit Sub

Dim sz As String

    Select Case StatusCode

    Case SNARL_SUCCESS
        sz = "OK"

    Case SNARL_ERROR_FAILED
        sz = "Unspecified failure"

    Case SNARL_ERROR_UNKNOWN_COMMAND
        sz = "Unknown command"

    Case SNARL_ERROR_TIMED_OUT
        sz = "Timed out"

    Case SNARL_ERROR_ARG_MISSING
        sz = "Required argument missing"

    Case SNARL_ERROR_SYSTEM
        sz = "System error"

    Case SNARL_ERROR_NOT_RUNNING
        sz = "Not running"

    Case SNARL_ERROR_NOT_REGISTERED
        sz = "App is not registered"

    Case SNARL_ERROR_ALREADY_REGISTERED
        sz = "App is already registered"

    Case SNARL_ERROR_CLASS_ALREADY_EXISTS
        sz = "Class already exists"

    Case SNARL_ERROR_CLASS_BLOCKED
        sz = "Class blocked"

    Case SNARL_ERROR_CLASS_NOT_FOUND
        sz = "Class not found"

    Case SNARL_ERROR_NOTIFICATION_NOT_FOUND
        sz = "Notification not found"

    Case SNARL_ERROR_FLOODING
        sz = "Flooding detect"

    Case SNARL_ERROR_DO_NOT_DISTURB
        sz = "Do Not Disturb"

    Case SNARL_ERROR_COULD_NOT_DISPLAY
        sz = "Could not display"

    Case SNARL_ERROR_AUTH_FAILURE
        sz = "Authentication failure"

    Case Else
        sz = "Undefined error"

    End Select

    sz = "SNP/2.0/" & CStr(StatusCode) & "/" & sz & IIf(Data <> "", "/" & Data, "") & vbCrLf
    Debug.Print "uSNPReply(): replying with '" & sz & "'"
    theSocket.SendData sz

End Sub

Public Sub Notify(ByVal Notification As Long, ByVal Token As String)

    If Not (theSocket Is Nothing) Then _
        theSocket.SendData uCreateReply(Notification + 270, Token)

End Sub

Private Function uDecodeBase64(ByVal Data As String) As String
Dim bErr As Boolean
Dim sz As String
Dim i As Integer

    On Error Resume Next

    sz = Decode64(Data, bErr)
    uDecodeBase64 = g_GetSafeTempIconPath()

    i = FreeFile()

    err.Clear
    Open uDecodeBase64 For Binary Access Write As #i

    If err.Number = 0 Then
        Put #i, , sz
        Close #i

    End If

    Debug.Print "uDecodeBase64(): writing icon to '" & uDecodeBase64 & "'"

End Function

Private Sub uDoGNTP(ByVal bytesTotal As Long)
Dim ppd As BPackedData
Dim bKeepOpen As Boolean
Dim szReply As String
Dim sz As String

    On Error GoTo er

    theSocket.PeekData sz

    If g_SafeRightStr(sz, 4) = vbCrLf & vbCrLf Then
        theSocket.GetData sz
        gntp_Process sz, theSocket, szReply, bKeepOpen
        uOutput "TRemoteConnection.uDoGNTP(): response=" & Replace$(szReply, vbCrLf, "¶")
        theSocket.SendData szReply

        If Not bKeepOpen Then _
            theSocket.CloseSocket

    End If
    Exit Sub

er:
    g_Debug "TRemoteConnection.uDoGNTP(): " & err.Description, LEMON_LEVEL_CRITICAL

    ' /* drop the connection */
    theSocket.CloseSocket

End Sub

Private Sub uDoSNP(ByVal bytesTotal As Long)
Dim sz As String

    On Error GoTo er

    theSocket.PeekData sz

    If g_SafeRightStr(sz, 1) <> vbLf Then _
        Exit Sub

    g_Debug "TRemoteConnection.uDoSNP(): SNP packet received from " & theSocket.RemoteHostIP & ":" & theSocket.RemotePort & "..."
    theSocket.GetData sz

    If g_ConfigGet("listen_for_snarl") = "0" Then
        g_Debug "TRemoteConnection.uDoSNP(): ignored (user disabled incoming notifications)"
        theSocket.SendData uCreateReply(SNP_ERROR_NOT_RUNNING)
        theSocket.CloseSocket           ' // R2.4.2
        Exit Sub

    End If

    sz = g_SafeLeftStr(sz, Len(sz) - 1)        ' // strip the trailing CR

    If LCase$(g_SafeLeftStr(sz, 6)) = "snp://" Then
        g_Debug "TRemoteConnection.uDoSNP(): is SNP2 request"
        uTranslateSNP2 g_SafeRightStr(sz, Len(sz) - 6)

    Else
        g_Debug "TRemoteConnection.uDoSNP(): is SNP1 request"
        uTranslateSNP sz

    End If

'Dim i As Long
'
'    theSocket.GetData sz
'
'    i = InStr(sz, Chr$(13))
'    If i > 0 Then
'        g_Debug "TRemoteConnection.uDoSNP(): SNP packet received"
'        mContent = mContent & g_SafeLeftStr(sz, i - 1)
'        uProcessSNP mContent
'        mContent = ""
'
'    Else
'        mContent = mContent & sz
'
'    End If

    Exit Sub

er:
    g_Debug "TRemoteConnection.uDoSNP(): " & err.Description, LEMON_LEVEL_CRITICAL
    uSNP2Reply SNARL_ERROR_BAD_PACKET

    ' /* drop the connection */

    theSocket.CloseSocket

End Sub

Private Sub uOutput(ByVal Text As String)

#If GNTP_TEST = 1 Then
    Form1.output Text

#Else
    g_Debug Text

#End If

End Sub

Private Sub uDoJSON(ByVal bytesTotal As Long)
Dim szReply As String
Dim sz As String

    On Error GoTo er

    g_Debug "TRemoteConnection.uDoJSON(): id=0x" & g_HexStr(mId) & " bytesTotal=" & CStr(bytesTotal), LEMON_LEVEL_PROC

    theSocket.GetData sz

    If g_ConfigGet("listen_for_json") = "1" Then
        If uParseJSON(sz) Then
            szReply = "200 OK"

        Else
            szReply = "402 Payment Required"

        End If

    Else
        ' /* JSON listening is not enabled */
        g_Debug "TRemoteConnection.uDoJSON(): user has disabled JSON listening", LEMON_LEVEL_INFO
        szReply = "404 Not Found"

    End If

    theSocket.SendData "HTTP/1.0 " & szReply & vbCrLf & Format$(Now(), "ddd, d mmm yyyy hh:mm:ss") & " GMT" & vbCrLf & _
                       "Content-Type: text/html" & vbCrLf & "Content-Length: 0" & vbCrLf & vbCrLf

er:
    theSocket.CloseSocket

End Sub

Private Function uParseJSON(ByVal HTTP As String) As Boolean
Dim po As BJSONObject

    Set po = New BJSONObject
    If Not po.SetFromHTTP(HTTP) Then
        g_Debug "TRemoteConnection.uParseJSON(): failed to create JSON object", LEMON_LEVEL_CRITICAL
        Exit Function

    End If

Dim pv As BJSONValue

    If Not po.Find("action", pv) Then
        g_Debug "TRemoteConnection.uParseJSON(): 'action' value missing", LEMON_LEVEL_CRITICAL
        Exit Function

    End If

    If pv.AsLong <> 1 Then
        g_Debug "TRemoteConnection.uParseJSON(): incorrect 'action' value", LEMON_LEVEL_CRITICAL
        Exit Function

    End If

Dim tt As String
Dim tx As String

    If po.Find("description", pv) Then _
        tx = pv.AsString()

    If po.Find("title", pv) Then _
        tt = pv.AsString()

    g_Debug "TRemoteConnection.uParseJSON(): request #" & g_HexStr(mId) & " title=" & tt & " text=" & tx

    uParseJSON = (g_PrivateNotify(SNARL_CLASS_JSON, tt, tx) <> 0)

'    uDo = (sn41EZNotify(gSnarlToken, SNARL_CLASS_JSON, tt, tx) <> 0)

End Function



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
    '/  Description:    Represents an active SNP, SNP2, SNP3 or GNTP connection
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
Dim mDoingSNP3 As Boolean

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

    g_Debug "TRemoteConnection.Accept(): Id=0x" & g_HexStr(Id) & " Port=" & CStr(PortNumber), LEMON_LEVEL_PROC

    mId = Id
    mType = PortNumber
'    mIsGNTP = IsGNTP

    Set theSocket = New CSocket
    theSocket.Accept Id

End Sub

Private Sub theSocket_OnClose()

    g_Debug "TRemoteConnection.OnClose(): id=0x" & g_HexStr(mId), LEMON_LEVEL_PROC

'    Debug.Print "!!!!!!!!!!!!!!!!!!!!!!!!"
    Debug.Print "* connection dropped *"
'    Debug.Print "!!!!!!!!!!!!!!!!!!!!!!!!"

End Sub

Private Sub theSocket_OnConnect()

    g_Debug "TRemoteConnection.OnConnect(): id=0x" & g_HexStr(mId), LEMON_LEVEL_PROC

End Sub

Private Sub theSocket_OnConnectionRequest(ByVal requestID As Long)

    g_Debug "TRemoteConnection.OnConnectionRequest(): id=0x" & g_HexStr(mId) & " requestID=0x" & g_HexStr(requestID), LEMON_LEVEL_PROC

End Sub

Private Sub theSocket_OnDataArrival(ByVal bytesTotal As Long)

    Debug.Print "TRemoteConnection::OnDataArrival()"

    Select Case mType
    Case GNTP_DEFAULT_PORT
        uDoGNTP bytesTotal

    Case SNP_DEFAULT_PORT, MELON_DEFAULT_PORT
        If g_ConfigGet("listen_for_snarl") = "0" Then
            g_Debug "TRemoteConnection.OnDataArrival(): ignored (user disabled incoming notifications)"
            theSocket.SendData uSNPCreateReply(SNARL_ERROR_NOT_RUNNING)
            theSocket.CloseSocket           ' // R2.4.2

        Else
            ' /* process */
            uDoSNP

        End If

    Case JSON_DEFAULT_PORT
        uDoJSON bytesTotal

    End Select

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
'        theSocket.SendData uSNPCreateReply(SNARL_ERROR_NOT_RUNNING)
'        Exit Sub
'
'    End If
'
'    If LCase$(g_SafeLeftStr(Request, 6)) = "snp://" Then
'        uSNP2Translate g_SafeRightStr(Request, Len(Request) - 6)
'
'    Else
'        uSNPTranslate Request
'
'    End If
'
'End Sub

Private Sub uSNP2Translate(ByVal Request As String)

    ' /* syntax is: snp://<action>[?arg=value[&arg=value]] CRLF
    '    on entry the 'snp://' prefix will have been stripped */

'    Request = g_UTF8(Request)

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
        ppd.SetTo g_URLDecode(Request)

    End If

    ' /* pre-processing */

    Debug.Print ">>> '" & sCommand; "'"

Dim sz As String

    Select Case sCommand

    Case "register", "reg"
        ' /* use default icon if none provided */
'        If Not ppd.Exists("icon") Then _
            ppd.Add "icon", g_MakePath(App.Path) & "etc\icons\snp-no_icon.png"

    Case "notify"
        ' /* if encoded icon data exists, extract it now */
        If ppd.Exists("icon64") Then
            sz = uDecodeBase64(Replace$(ppd.ValueOf("icon64"), "%", "="))
            ppd.Add "icon", sz

        End If

    End Select

Dim hr As Long

    ' /* process the command (note the hard-coding to V42 of the API) */

    hr = g_DoAction(sCommand, Val(ppd.ValueOf("token")), ppd, SN_NF_REMOTE Or 42, theSocket)
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

Private Sub uSNPTranslate(ByVal Request As String)
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
                        g_Debug "TRemoteConnection.uSNPTranslate(): unknown action '" & sz & "'", LEMON_LEVEL_CRITICAL

                    End Select

                Else
                    g_Debug "TRemoteConnection.uSNPTranslate(): unsupported version '" & CStr(pPacket.Version) & "'", LEMON_LEVEL_CRITICAL

                End If
            Else
                g_Debug "TRemoteConnection.uSNPTranslate(): bad packet type '" & sz & "'", LEMON_LEVEL_CRITICAL

            End If
        End If

    End With

    ' /* check here that the packet is valid (i.e. version, type, etc.) */

    If Not pPacket.Valid Then
        theSocket.SendData uSNPCreateReply(SNARL_ERROR_BAD_PACKET)
        theSocket.CloseSocket
        Exit Sub

    End If

    g_Debug "TRemoteConnection.uSNPTranslate(): packet is valid (ver=" & pPacket.Version & ")"

    ' /* do some general safety checking first */

    If (g_AppRoster Is Nothing) Or (g_NotificationRoster Is Nothing) Then
        g_Debug "TRemoteConnection.uSNPTranslate(): action '" & pPacket.action & "': app_roster or notification_roster not found", LEMON_LEVEL_CRITICAL
        theSocket.SendData uSNPCreateReply(SNARL_ERROR_FAILED)
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
                g_Debug "TRemoteConnection.uSNPTranslate(): remote app '" & .App & "' is already registered", LEMON_LEVEL_CRITICAL
                theSocket.SendData uSNPCreateReply(SNARL_ERROR_ALREADY_REGISTERED)

            Else
                ' /* use the special "Add2()" method which allows us to register multiple apps using the same PID - in
                '    this case -1, which represents remote apps */
                g_AppRoster.OldAdd .App, 0, 0, g_MakePath(App.Path) & "etc\icons\remote_app.png", g_MakePath(App.Path) & "etc\icons\remote.png", -1
                g_Debug "TRemoteConnection.uSNPTranslate(): remote app '" & .App & "' registered"
                theSocket.SendData uSNPCreateReply(SNARL_SUCCESS)

            End If


        Case T_ACTION_REVOKE
            dw = g_AppRoster.IndexOfPidAndName(-1, .App)
            If dw = 0 Then
                ' /* not registered */
                g_Debug "TRemoteConnection.uSNPTranslate(): remote app '" & .App & "' is not registered", LEMON_LEVEL_CRITICAL
                theSocket.SendData uSNPCreateReply(SNARL_ERROR_NOT_REGISTERED)

            Else
                ' /* remove safely */
                g_AppRoster.Remove dw
                frmAbout.bUpdateAppList
                g_Debug "TRemoteConnection.uSNPTranslate(): remote app '" & .App & "' unregistered"
                theSocket.SendData uSNPCreateReply(SNARL_SUCCESS)

            End If


        Case T_ACTION_ADD_CLASS
            dw = g_AppRoster.IndexOfPidAndName(-1, .App)
            If dw = 0 Then
                ' /* not registered */
                g_Debug "TRemoteConnection.uSNPTranslate(): remote app '" & .App & "' is not registered", LEMON_LEVEL_CRITICAL
                theSocket.SendData uSNPCreateReply(SNARL_ERROR_NOT_REGISTERED)

            Else
                ' /* found the app */
                Set pa = g_AppRoster.AppAt(dw)
                If Not pa.FindAlert(.Class, pc) Then
                    ' /* okay to add the class */
                    If pa.AddAlert(.Class, .Title, pc) = M_OK Then _
                        pc.DefaultIcon = pPacket.Icon

                    frmAbout.bUpdateAppList
                    g_Debug "TRemoteConnection.uSNPTranslate(): class '" & .Class & "' add"
                    theSocket.SendData uSNPCreateReply(SNARL_SUCCESS)

                Else
                    ' /* class already exists */
                    g_Debug "TRemoteConnection.uSNPTranslate(): class '" & .Class & "' is already registered", LEMON_LEVEL_CRITICAL
                    theSocket.SendData uSNPCreateReply(SNARL_ERROR_CLASS_ALREADY_EXISTS)

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
                g_Debug "TRemoteConnection.uSNPTranslate(): remote app '" & .App & "' is not registered", LEMON_LEVEL_CRITICAL

'FIXFIXFIX
                ' /* we no longer include frmAbout.hWnd and WM_REMOTENOTIFY with this */
                g_PrivateNotify SNARL_CLASS_ANON_NET, .Title, .Text, .Timeout, .Icon, , .DefaultAck, , SN_NF_REMOTE
'FIXFIXFIX
                g_Debug "TRemoteConnection.uSNPTranslate(): snShowMessageEx() returned 0x" & g_HexStr(dw)

            Else
                ' /* app found */
                Set pa = g_AppRoster.AppAt(dw)
                dw = pa.Show(.Class, .Title, .Text, .Timeout, pPacket.Icon) ', frmAbout.hWnd, WM_REMOTENOTIFY)
                g_Debug "TRemoteConnection.uSNPTranslate(): Show() returned 0x" & g_HexStr(dw)

            End If

            If dw > 0 Then
                theSocket.SendData uSNPCreateReply(SNARL_SUCCESS, CStr(dw))
'                frmAbout.AddRemoteNotification dw, Me

            Else
                theSocket.SendData uSNPCreateReply(SNARL_ERROR_FAILED)

            End If


        Case T_ACTION_HELLO
            ' /* send release back */
            theSocket.SendData uSNPCreateReply(SNARL_SUCCESS, App.Title & " " & App.Comments)

        Case T_ACTION_VERSION
            ' /* send version back */
            theSocket.SendData uSNPCreateReply(SNARL_SUCCESS, CStr(App.Major) & "." & CStr(App.Revision))

        Case Else
            ' /* unknown action */
            g_Debug "TRemoteConnection.uSNPTranslate(): action '" & .action & "' is unknown", LEMON_LEVEL_CRITICAL
            theSocket.SendData uSNPCreateReply(SNARL_ERROR_UNKNOWN_COMMAND)

        End Select

    End With

    Exit Sub

er:
    g_Debug "TRemoteConnection.uSNPTranslate(): " & err.Description, LEMON_LEVEL_CRITICAL
    theSocket.SendData uSNPCreateReply(SNARL_ERROR_FAILED)
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

Private Function uSNPCreateReply(ByVal ReturnCode As Long, Optional ByVal Data As String) As String
Dim sz As String

'    Select Case ReturnCode
'
'    ' /* 0xx - Success code */
'
'    Case SNARL_SUCCESS
'        sz = "OK"
'
'    ' /* 1xx - Protocol errors */
'
'    Case SNARL_ERROR_FAILED
'        sz = "Internal error"
'
'    Case SNARL_ERROR_UNKNOWN_COMMAND
'        sz = "Unknown command"
'
'    Case SNARL_ERROR_TIMED_OUT
'        sz = "Timed out"
'
'    Case SNARL_ERROR_BUSY
'        sz = "Server too busy"
'
'    Case SNARL_ERROR_BAD_PACKET
'        sz = "Bad packet"
'
'    Case SNARL_ERROR_ACCESS_DENIED
'        sz = "Access denied"
'
'    ' /* 2xx - Server errors */
'
'    Case SNARL_ERROR_NOT_RUNNING
'        sz = "Not running"
'
'    Case SNARL_ERROR_NOT_REGISTERED
'        sz = "Application is not registered"
'
'    Case SNARL_ERROR_ALREADY_REGISTERED
'        sz = "Application is already registered"
'
'    Case SNARL_ERROR_CLASS_ALREADY_EXISTS
'        sz = "Class is already registered"
'
'
'    ' /* 3xx - Notifications */
'
'    Case SNP_NOTIFY_CANCELLED
'        sz = "Notification cancelled"
'
'    Case SNP_NOTIFY_TIMED_OUT
'        sz = "Notification timed out"
'
'    Case SNP_NOTIFY_ACK
'        sz = "Notification acknowledged"
'
'    Case SNP_NOTIFY_MENU
'        sz = "Notification menu selected"
'
'    Case SNP_NOTIFY_MIDDLE_BUTTON
'        sz = "Notification middle button clicked"
'
'    Case SNP_NOTIFY_CLOSED
'        sz = "Notification closed"
'
'    End Select

    uSNPCreateReply = "SNP/" & SNP_VERSION & "/" & CStr(ReturnCode) & "/" & snp3_StatusName(ReturnCode) & IIf(Data <> "", "/" & Data, "") & vbCrLf

End Function

Private Sub uSNP2Reply(ByVal StatusCode As Long, Optional ByVal Data As String)

    If (theSocket Is Nothing) Then _
        Exit Sub

Dim sz As String

'    Select Case StatusCode
'
'    Case SNARL_SUCCESS
'        sz = "OK"
'
'    Case SNARL_ERROR_FAILED
'        sz = "Unspecified failure"
'
'    Case SNARL_ERROR_UNKNOWN_COMMAND
'        sz = "Unknown command"
'
'    Case SNARL_ERROR_TIMED_OUT
'        sz = "Timed out"
'
'    Case SNARL_ERROR_ARG_MISSING
'        sz = "Required argument missing"
'
'    Case SNARL_ERROR_SYSTEM
'        sz = "System error"
'
'    Case SNARL_ERROR_NOT_RUNNING
'        sz = "Not running"
'
'    Case SNARL_ERROR_NOT_REGISTERED
'        sz = "App is not registered"
'
'    Case SNARL_ERROR_ALREADY_REGISTERED
'        sz = "App is already registered"
'
'    Case SNARL_ERROR_CLASS_ALREADY_EXISTS
'        sz = "Class already exists"
'
'    Case SNARL_ERROR_CLASS_BLOCKED
'        sz = "Class blocked"
'
'    Case SNARL_ERROR_CLASS_NOT_FOUND
'        sz = "Class not found"
'
'    Case SNARL_ERROR_NOTIFICATION_NOT_FOUND
'        sz = "Notification not found"
'
'    Case SNARL_ERROR_FLOODING
'        sz = "Flooding detect"
'
'    Case SNARL_ERROR_DO_NOT_DISTURB
'        sz = "Do Not Disturb"
'
'    Case SNARL_ERROR_COULD_NOT_DISPLAY
'        sz = "Could not display"
'
'    Case SNARL_ERROR_AUTH_FAILURE
'        sz = "Authentication failure"
'
'    Case Else
'        sz = "Undefined error"
'
'    End Select

    sz = "SNP/" & SNP_VERSION & "/" & CStr(StatusCode) & "/" & snp3_StatusName(StatusCode) & IIf(Data <> "", "/" & Data, "") & vbCrLf
    Debug.Print "uSNPReply(): replying with '" & sz & "'"
    theSocket.SendData sz

End Sub

Public Sub Notify(ByVal Notification As Long, ByVal Token As String)

    If Not (theSocket Is Nothing) Then _
        theSocket.SendData uSNPCreateReply(Notification + 270, Token)

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

    If g_SafeRightStr(sz, 1) = Chr$(0) Then
        ' /* is it a flash permission request? */
        If LCase$(g_TrimStr(sz)) = "<policy-file-request/>" Then _
            theSocket.SendData "<?xml version=""1.0""?><!DOCTYPE cross-domain-policy SYSTEM ""/xml/dtds/cross-domain-policy.dtd""><cross-domain-policy><site-control permitted-cross-domain-policies=""master-only""/><allow-access-from domain=""*"" to-ports=""*"" /></cross-domain-policy>" & Chr$(0)

    ElseIf g_SafeRightStr(sz, 4) = vbCrLf & vbCrLf Then
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

Private Sub uDoSNP()
Dim ppd As BPackedData
Dim sx As String
Dim sz As String

    On Error GoTo er

    theSocket.PeekData sz


    If g_SafeRightStr(sz, 1) = Chr$(0) Then
        ' /* is it a flash permission request? */
        If LCase$(g_TrimStr(sz)) = "<policy-file-request/>" Then
            theSocket.SendData "<?xml version=""1.0""?><!DOCTYPE cross-domain-policy SYSTEM ""/xml/dtds/cross-domain-policy.dtd""><cross-domain-policy><site-control permitted-cross-domain-policies=""master-only""/><allow-access-from domain=""*"" to-ports=""*"" /></cross-domain-policy>" & Chr$(0)
            theSocket.CloseSocket

        End If
    End If

    ' /* if it doesn't end with a least an LF, keep waiting */

    If g_SafeRightStr(sz, 1) <> vbLf Then _
        Exit Sub

    If mDoingSNP3 Then
        ' /* we're processing an incoming SNP3 packet... */
        Set ppd = New BPackedData
        ppd.SetTo g_SafeLeftStr(sz, Len(sz) - 2), vbCrLf, ""
        If ppd.Count > 0 Then
            uSNP3Parse ppd

        Else
            ' /* should never happen! */
            Debug.Print "SNP3: invalid packet"
            mDoingSNP3 = False
            theSocket.CloseSocket
            Exit Sub

        End If

        Exit Sub

    Else
        Set ppd = New BPackedData
        ppd.SetTo g_SafeLeftStr(sz, Len(sz) - 2), vbCrLf, ""
        If ppd.Count > 0 Then
            ppd.EntryAt 1, sx, ""
            If g_SafeLeftStr(sx, 4) = "SNP/" Then
                Debug.Print "SNP3+ packet tracking started..."
                mDoingSNP3 = True
                uSNP3Parse ppd
                Exit Sub

            End If
        End If
    End If

    g_Debug "TRemoteConnection.uDoSNP(): SNP packet received from " & theSocket.RemoteHostIP & ":" & theSocket.RemotePort & "..."
    theSocket.GetData sz

    ' /* remove CRs and LFs */

    sz = Replace$(sz, vbCr, "")
    sz = Replace$(sz, vbLf, "")

    If LCase$(g_SafeLeftStr(sz, 6)) = "snp://" Then
        g_Debug "TRemoteConnection.uDoSNP(): is SNP2 request"
        uSNP2Translate g_SafeRightStr(sz, Len(sz) - 6)

    Else
        g_Debug "TRemoteConnection.uDoSNP(): is SNP1 request"
        uSNPTranslate sz

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

Private Function uSNP3Translate(ByRef Request As BPackedData, ByRef Response As String) As Boolean
Dim bErr As Boolean
Dim szr As String

    ' /* return False if the socket should be closed */

    Debug.Print "----SNP3---"
    Debug.Print Request.AsString
    Debug.Print "----SNP3---"

Dim sz As String

    ' /* get the header */

    Request.EntryAt 1, sz, ""
    If g_SafeLeftStr(sz, 7) <> "SNP/3.0" Then
        Response = snp3_BuildResponse(SNARL_ERROR_UNSUPPORTED_VERSION, "Invalid version in header")
        Exit Function

    End If

Dim ppdHeader As BPackedData

    ' /* split header into component parts */

    Set ppdHeader = New BPackedData
    ppdHeader.SetTo sz, " ", ":"

Dim i As Long
Dim n As Long

    ' /* check for empty message */

    With Request
        If .Count > 2 Then
            For i = 2 To .Count - 1
                .EntryAt i, sz, ""
                If (g_SafeLeftStr(sz, 1) <> "#") Then _
                    n = n + 1

            Next i
        End If

    End With

    If n = 0 Then
        Response = snp3_BuildResponse(SNARL_ERROR_NO_ACTIONS_PROVIDED, "Must supply at least one action")
        Exit Function

    End If


Dim szExpectedKey As String

    ' /* get current password */

    szExpectedKey = g_GetPassword()

    ' /* decode header - currently SNP/3.0 [HASH:value.salt] [ENCRYPTION:value] */

Dim szProvidedKey As String
Dim szProvidedSalt As String
Dim szHashType As String

    If ppdHeader.Count > 1 Then
        ' /* get provided hashing algorithm, key and salt */
        ppdHeader.EntryAt 2, szHashType, szProvidedKey

        If Not uGetSalt(szProvidedKey, szProvidedKey, szProvidedSalt) Then
            ' /* key and/or salt invalid */
            Response = snp3_BuildResponse(SNARL_ERROR_UNSUPPORTED_HASHING)
            Exit Function

        End If

        Debug.Print "type=" & szHashType & " key=" & szProvidedKey & " salt=" & szProvidedSalt

    End If

    ' /* check invalid cases */

    If (szExpectedKey = "") And (szProvidedKey <> "") Then
        ' /* invalid case: no password set but one provided by sender (we don't tell the sender this, however) */
        Debug.Print "uSNP3Translate(): no system password set but request contained one"
        Response = snp3_BuildResponse(SNARL_ERROR_AUTH_FAILURE)
        Exit Function

    End If

    If (szExpectedKey <> "") And (szProvidedKey = "") Then
        ' /* invalid case: password set but none provided by sender (we don't tell the sender this, however) */
        Debug.Print "uSNP3Translate(): password required"
        Response = snp3_BuildResponse(SNARL_ERROR_AUTH_FAILURE)
        Exit Function

    End If

    ' /* valid combos */

    If (szExpectedKey <> "") And (szProvidedKey <> "") Then
        ' /* compare supplied password with ours */

        szExpectedKey = szExpectedKey & szProvidedSalt

        Select Case szHashType
        Case "MD5"
            szExpectedKey = MD5DigestStrToHexStr(szExpectedKey)

        Case "SHA1"
            szExpectedKey = SHA1DigestStrToHexStr(szExpectedKey)

        Case "SHA256"
            SHA256Init
            szExpectedKey = SHA256DigestStrToHexStr(szExpectedKey)

        Case "SHA384", "SHA512"
            Response = snp3_BuildResponse(SNARL_ERROR_UNSUPPORTED_HASHING, "")
            Exit Function

        Case Else
            ' /* unknown hashing algorithm */
            Response = snp3_BuildResponse(SNARL_ERROR_UNSUPPORTED_HASHING, "")
            Exit Function

        End Select

    End If


'Dim szEncType As String
'Dim szProvidedEncKey As String
'
'        If ppdHeader.Count > 1 Then
'            ' /* get encryption algorithm */
'            ppdHeader.EntryAt 2, szEncType, szProvidedEncKey
'            Debug.Print "CYPHER ALGORITHM=" & szEncType & " value=" & szProvidedEncKey
'
'            Select Case szEncType
'            Case "AES"
'                Response = snp3_BuildResponse(SNARL_ERROR_UNSUPPORTED_ENCRYPTION, "AES not currently implemented")
'                Exit Function
'
'            Case "DES"
'                Response = snp3_BuildResponse(SNARL_ERROR_UNSUPPORTED_ENCRYPTION, "DES not currently implemented")
'                Exit Function
'
'            Case "3DES"
'                Response = snp3_BuildResponse(SNARL_ERROR_UNSUPPORTED_ENCRYPTION, "Triple-DES not currently implemented")
'                Exit Function
'
'            Case Else
'                Response = snp3_BuildResponse(SNARL_ERROR_UNSUPPORTED_ENCRYPTION, "")
'                Exit Function
'
'            End Select
'
'        End If




    ' /* check key */

    Debug.Print "expected key: " & szExpectedKey

    If szExpectedKey <> szProvidedKey Then
        Response = snp3_BuildResponse(SNARL_ERROR_AUTH_FAILURE, "Authorization failure")
        Exit Function

    End If

    szr = ""

    ' /* parse actions */

    With Request
        For i = 2 To .Count - 1
            .EntryAt i, sz, ""

            If (g_SafeLeftStr(sz, 1) <> "#") Then
                n = g_DoV42Request(sz, GetCurrentProcessId(), theSocket, SN_NF_REMOTE Or SN_NF_IS_SNP3)
                If n < 0 Then
'                        bErr = True     ' // set overall response flag
                    n = Abs(n)      ' // convert to SNP code

                Else
                    n = 0           ' // convert to SNP code

                End If

                szr = szr & "result: " & uGetAction(sz) & " " & CStr(n) & " " & snp3_StatusName(n) & vbCrLf

            End If
        Next i
    End With

    Response = snp3_BuildResponse(IIf(bErr, SNARL_ERROR_FAILED, SNARL_SUCCESS), , szr)

    ' /* leave connected irrespective of individual action success/failure */

    uSNP3Translate = True

End Function

'Private Function snp3_BuildResponse(ByVal StatusCode As SNARL_STATUS_CODE, Optional ByVal Hint As String, Optional ByVal Actions As String) As String
'Dim sz As String
'
'    sz = "SNP/" & SNP_VERSION & " " & IIf(StatusCode = SNARL_SUCCESS, "OK", "FAILED") & vbCrLf
'
'    ' /* only if StatusCode indicates error */
'
'    If (StatusCode <> SNARL_SUCCESS) And (Actions = "") Then
'        sz = sz & "error-code: " & CStr(StatusCode) & vbCrLf & _
'                  "error-name: " & snp3_StatusName(StatusCode) & vbCrLf
'
'        If Hint <> "" Then _
'            sz = sz & "error-hint: " & Hint & vbCrLf
'
'    End If
'
'    If Actions <> "" Then _
'        sz = sz & Actions
'
'    ' /* standard headers */
'
'    sz = sz & "--" & vbCrLf
'    sz = sz & "x-timestamp: " & Format$(Now(), "d mmm yyyy hh:mm:ss") & vbCrLf
'    sz = sz & "x-daemon: " & "Snarl " & CStr(APP_VER) & "." & CStr(APP_SUB_VER) & IIf(APP_SUB_SUB_VER <> 0, "." & CStr(APP_SUB_SUB_VER), "") & vbCrLf
'    sz = sz & "x-host: " & LCase$(g_GetComputerName()) & vbCrLf
'    sz = sz & "END" & vbCrLf
'
'    snp3_BuildResponse = sz
'
'End Function
'
'Private Function snp3_StatusName(ByVal StatusCode As SNARL_STATUS_CODE) As String
'Dim sz As String
'
'    Select Case StatusCode
'
'    Case SNARL_SUCCESS
'        sz = "Ok"
'
'    Case SNARL_ERROR_FAILED                      '// miscellaneous failure
'        sz = "Failed"
'
'    Case SNARL_ERROR_UNKNOWN_COMMAND             '// specified command not recognised
'        sz = "BadCommand"
'
'    Case SNARL_ERROR_TIMED_OUT                   '// Snarl took too long to respond
'        sz = "TimedOut"
'
'    Case SNARL_ERROR_BAD_SOCKET                  '// invalid socket (or some other socket-related error)
'        sz = "BadSocket"
'
'    Case SNARL_ERROR_BAD_PACKET                  '// badly formed request
'        sz = "BadPacket"
'
'    Case SNARL_ERROR_INVALID_ARG                 '// R2.4B4: arg supplied was invalid
'        sz = "InvalidArg"
'
'    Case SNARL_ERROR_ARG_MISSING                 '// required argument missing
'        sz = "ArgMissing"
'
'    Case SNARL_ERROR_SYSTEM                      '// internal system error
'        sz = "InternalError"
'
'    Case SNARL_ERROR_ACCESS_DENIED               '// libsnarl only
'        sz = "AccessDenied"
'
'    Case SNARL_ERROR_UNSUPPORTED_VERSION
'        sz = "UnsupportedVersion"
'
'    Case SNARL_ERROR_NO_ACTIONS_PROVIDED
'        sz = "NothingToDo"
'
'    Case SNARL_ERROR_UNSUPPORTED_ENCRYPTION
'        sz = "UnsupportedEncryption"
'
'    Case SNARL_ERROR_UNSUPPORTED_HASHING
'        sz = "UnsupportedHashing"
'
'
'
'    Case SNARL_ERROR_NOT_RUNNING                 '// Snarl handling window not found
'        sz = "NotRunning"
'
'    Case SNARL_ERROR_NOT_REGISTERED
'        sz = "NotRegistered"
'
'    Case SNARL_ERROR_ALREADY_REGISTERED          '// not used yet; sn41RegisterApp() returns existing token
'        sz = "SpuriousRegister"
'
'    Case SNARL_ERROR_CLASS_ALREADY_EXISTS        '// not used yet
'        sz = "SpuriousClass"
'
'    Case SNARL_ERROR_CLASS_BLOCKED
'        sz = "ClassBlocked"
'
'    Case SNARL_ERROR_CLASS_NOT_FOUND
'        sz = "InvalidClass"
'
'    Case SNARL_ERROR_NOTIFICATION_NOT_FOUND
'        sz = "InvalidNotification"
'
'    Case SNARL_ERROR_FLOODING                    '// notification generated by same class within quantum
'        sz = "FloodingAlert"
'
'    Case SNARL_ERROR_DO_NOT_DISTURB              '// DnD mode is in effect was not logged as missed
'        sz = "DoNotDisturb"
'
'    Case SNARL_ERROR_COULD_NOT_DISPLAY           '// not enough space on-screen to display notification
'        sz = "DisplayFailed"
'
'    Case SNARL_ERROR_AUTH_FAILURE                '// password mismatch
'        sz = "AuthenticationFailure"
'
'    Case SNARL_ERROR_DISCARDED                   '// discarded for some reason, e.g. foreground app match
'        sz = "WasDiscarded"
'
'    Case SNARL_ERROR_NOT_SUBSCRIBED                 '// 2.4.2 DR3: subscriber not found
'        sz = "NotSubscribed"
'
'    Case SNARL_ERROR_ALREADY_SUBSCRIBED
'        sz = "AlreadySubscribed"
'
'
'    Case Else
'        sz = "(Unknown)"
'
'    End Select
'
'    snp3_StatusName = sz
'
'End Function

Private Sub uSNP3Parse(ByRef Data As BPackedData)
Dim sx As String
Dim sr As String

    ' /* get last entry */

    Data.EntryAt Data.Count, sx, ""

    If sx = "END" Then
        Debug.Print "SNP3 packet received!!!"
        theSocket.GetData sx                    ' // empty the socket buffer
        mDoingSNP3 = False

        ' /* if translation fails, we must close the socket immediately after responding */

        If Not uSNP3Translate(Data, sr) Then
            theSocket.SendData sr
            theSocket.CloseSocket

        Else
            theSocket.SendData sr

        End If

    End If

End Sub

Private Function uGetAction(ByVal Request As String) As String
Dim i As Long

    i = InStr(Request, "?")
    If i Then
        uGetAction = g_SafeLeftStr(Request, i - 1)

    Else
        uGetAction = Request

    End If

End Function

Private Function uGetSalt(ByVal SaltedKey As String, ByRef Key As String, ByRef Salt As String) As Boolean
Dim i As Long

    i = InStr(SaltedKey, ".")
    If i = 0 Then _
        Exit Function

    Key = g_SafeLeftStr(SaltedKey, i - 1)
    Salt = g_SafeRightStr(SaltedKey, Len(SaltedKey) - i)

    uGetSalt = ((Key <> "") And (Salt <> ""))

End Function



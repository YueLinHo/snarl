VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSnarlSocket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

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
    Action As T_NETWORK_ACTIONS
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
Dim mPacket As String

Dim WithEvents theSocket As CSocket
Attribute theSocket.VB_VarHelpID = -1

Private Sub Class_Initialize()
'//
End Sub

Private Sub Class_Terminate()
'//
End Sub

Public Sub Accept(ByVal Id As Long)

    g_Debug "CSnarlSocket.Accept(): Id=0x" & g_HexStr(Id), LEMON_LEVEL_PROC

    mId = Id
    Set theSocket = New CSocket
    theSocket.Accept Id

End Sub

Private Sub theSocket_OnClose()

    g_Debug "CSnarlSocket.OnClose(): id=0x" & g_HexStr(mId), LEMON_LEVEL_PROC

End Sub

Private Sub theSocket_OnConnect()

    g_Debug "CSnarlSocket.OnConnect(): id=0x" & g_HexStr(mId), LEMON_LEVEL_PROC

End Sub

Private Sub theSocket_OnConnectionRequest(ByVal requestID As Long)

    g_Debug "CSnarlSocket.OnConnectionRequest(): id=0x" & g_HexStr(mId) & " requestID=0x" & g_HexStr(requestID), LEMON_LEVEL_PROC

End Sub

Private Sub theSocket_OnDataArrival(ByVal bytesTotal As Long)
Dim sz As String
Dim i As Integer

    g_Debug "CSnarlSocket.OnDataArrival(): id=0x" & g_HexStr(mId) & " bytesTotal=" & CStr(bytesTotal), LEMON_LEVEL_PROC

    theSocket.GetData sz

'    i = InStr(sz, vbCrLf)

    i = InStr(sz, Chr$(13))
    If i > 0 Then
        g_Debug "CSnarlSocket.OnDataArrival(): end_of_packet received"
        mPacket = mPacket & g_SafeLeftStr(sz, i - 1)
        uTranslate mPacket
        mPacket = ""

    Else
        mPacket = mPacket & sz

    End If

End Sub

Private Sub theSocket_OnError(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

    g_Debug "CSnarlSocket.OnError(): id=0x" & g_HexStr(mId) & " Description=" & Description, LEMON_LEVEL_PROC

End Sub

Private Sub theSocket_OnSendComplete()

    g_Debug "CSnarlSocket.OnSendComplete(): id=0x" & g_HexStr(mId), LEMON_LEVEL_PROC

End Sub

Private Sub theSocket_OnSendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)

    g_Debug "CSnarlSocket.OnSendProgress(): id=0x" & g_HexStr(mId) & " bytesSent=" & CStr(bytesSent), LEMON_LEVEL_PROC

End Sub

Private Sub uTranslate(ByVal Request As String)

    On Error GoTo er

    g_Debug "CSnarlSocket.uTranslate(): incoming notification from " & theSocket.RemoteHostIP & ":" & theSocket.RemotePort & "..."

    If g_ConfigGet("listen_for_snarl") = "0" Then
        g_Debug "CSnarlSocket.uTranslate(): ignored (user disabled incoming notifications)"
        theSocket.SendData uCreateReply(SNP_ERROR_NOT_RUNNING)
        Exit Sub

    End If

Dim sz As String
Dim pPacket As T_NETWORK_INFO
Dim pc As TAlert
Dim pa As TApp

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
                        pPacket.Action = T_ACTION_REGISTER
                        pPacket.Valid = (pPacket.App <> "")

                    Case "unregister"
                        pPacket.Action = T_ACTION_REVOKE
                        pPacket.Valid = (pPacket.App <> "")

                    Case "add_class"
                        pPacket.Action = T_ACTION_ADD_CLASS
                        pPacket.Title = .GetValueWithDefault("title")
                        pPacket.Class = .GetValueWithDefault("class")
                        ' /* V40.10 (SNP1.1) */
                        pPacket.Icon = .GetValueWithDefault("icon")
                        pPacket.Valid = (pPacket.App <> "") And (pPacket.Class <> "")

                    Case "notification"
                        pPacket.Action = T_ACTION_SHOW_NOTIFICATION
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
                        pPacket.Action = T_ACTION_HELLO
                        pPacket.Valid = True

                    Case "version"
                        ' /* introduced in SNP1.1 */
                        pPacket.Action = T_ACTION_VERSION
                        pPacket.Valid = True

                    Case Else
                        g_Debug "CSnarlSocket.uTranslate(): unknown action '" & sz & "'", LEMON_LEVEL_CRITICAL

                    End Select

                Else
                    g_Debug "CSnarlSocket.uTranslate(): unsupported version '" & CStr(pPacket.Version) & "'", LEMON_LEVEL_CRITICAL

                End If
            Else
                g_Debug "CSnarlSocket.uTranslate(): bad packet type '" & sz & "'", LEMON_LEVEL_CRITICAL

            End If
        End If

    End With

    ' /* check here that the packet is valid (i.e. version, type, etc.) */

    If Not pPacket.Valid Then
        theSocket.SendData uCreateReply(SNP_ERROR_BAD_PACKET)
        Exit Sub

    End If

    g_Debug "CSnarlSocket.uTranslate(): packet is valid (ver=" & pPacket.Version & ")"

    ' /* do some general safety checking first */

    If (g_AppRoster Is Nothing) Or (g_NotificationRoster Is Nothing) Then
        g_Debug "CSnarlSocket.uTranslate(): action '" & pPacket.Action & "': app_roster or notification_roster not found", LEMON_LEVEL_CRITICAL
        theSocket.SendData uCreateReply(SNP_ERROR_FAILED)
        Exit Sub

    End If

    ' /* figure out the packet action */

Dim dw As Long

    With pPacket

        Select Case .Action

        Case T_ACTION_REGISTER
            dw = g_AppRoster.IndexOfPidAndName(-1, .App)
            If dw <> 0 Then
                ' /* already registered */
                g_Debug "CSnarlSocket.uTranslate(): remote app '" & .App & "' is already registered", LEMON_LEVEL_CRITICAL
                theSocket.SendData uCreateReply(SNP_ERROR_ALREADY_REGISTERED)

            Else
                ' /* use the special "Add2()" method which allows us to register multiple apps using the same PID - in
                '    this case -1, which represents remote apps */
                g_AppRoster.OldAdd .App, 0, 0, g_MakePath(App.Path) & "etc\icons\remote_app.png", g_MakePath(App.Path) & "etc\icons\remote.png", -1
                g_Debug "CSnarlSocket.uTranslate(): remote app '" & .App & "' registered"
                theSocket.SendData uCreateReply(SNP_SUCCESS)

            End If


        Case T_ACTION_REVOKE
            dw = g_AppRoster.IndexOfPidAndName(-1, .App)
            If dw = 0 Then
                ' /* not registered */
                g_Debug "CSnarlSocket.uTranslate(): remote app '" & .App & "' is not registered", LEMON_LEVEL_CRITICAL
                theSocket.SendData uCreateReply(SNP_ERROR_NOT_REGISTERED)

            Else
                ' /* remove safely */
                g_AppRoster.Remove dw
                frmAbout.bUpdateAppList
                g_Debug "CSnarlSocket.uTranslate(): remote app '" & .App & "' unregistered"
                theSocket.SendData uCreateReply(SNP_SUCCESS)

            End If


        Case T_ACTION_ADD_CLASS
            dw = g_AppRoster.IndexOfPidAndName(-1, .App)
            If dw = 0 Then
                ' /* not registered */
                g_Debug "CSnarlSocket.uTranslate(): remote app '" & .App & "' is not registered", LEMON_LEVEL_CRITICAL
                theSocket.SendData uCreateReply(SNP_ERROR_NOT_REGISTERED)

            Else
                ' /* found the app */
                Set pa = g_AppRoster.AppAt(dw)
                If Not pa.FindAlert(.Class, pc) Then
                    ' /* okay to add the class */
                    If pa.AddAlert(.Class, .Title, pc) = M_OK Then _
                        pc.DefaultIcon = pPacket.Icon

                    frmAbout.bUpdateAppList
                    g_Debug "CSnarlSocket.uTranslate(): class '" & .Class & "' add"
                    theSocket.SendData uCreateReply(SNP_SUCCESS)

                Else
                    ' /* class already exists */
                    g_Debug "CSnarlSocket.uTranslate(): class '" & .Class & "' is already registered", LEMON_LEVEL_CRITICAL
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
                g_Debug "CSnarlSocket.uTranslate(): remote app '" & .App & "' is not registered", LEMON_LEVEL_CRITICAL
                dw = snShowMessageEx(SNARL_CLASS_ANON_NET, .Title, .Text, .Timeout, pPacket.Icon, frmAbout.hWnd, WM_REMOTENOTIFY)
                g_Debug "CSnarlSocket.uTranslate(): snShowMessageEx() returned 0x" & g_HexStr(dw)

            Else
                ' /* app found */
                Set pa = g_AppRoster.AppAt(dw)
                dw = pa.Show(.Class, .Title, .Text, .Timeout, pPacket.Icon, frmAbout.hWnd, WM_REMOTENOTIFY)
                g_Debug "CSnarlSocket.uTranslate(): Show() returned 0x" & g_HexStr(dw)

            End If

            If dw > 0 Then
                snChangeAttribute dw, SNARL_ATTRIBUTE_ACK, pPacket.DefaultAck       ' // V40.17
                theSocket.SendData uCreateReply(SNP_SUCCESS, CStr(dw))
                frmAbout.AddRemoteNotification dw, Me

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
            g_Debug "CSnarlSocket.uTranslate(): action '" & .Action & "' is unknown", LEMON_LEVEL_CRITICAL
            theSocket.SendData uCreateReply(SNP_ERROR_UNKNOWN_COMMAND)

        End Select

    End With

    Exit Sub

er:
    g_Debug "CSnarlSocket.uTranslate(): " & Err.Description, LEMON_LEVEL_CRITICAL
    theSocket.SendData uCreateReply(SNP_ERROR_FAILED)

End Sub



'                If Not pa.FindAlert(.Class, pc) Then
'                    g_Debug "CSnarlSocket.uTranslate(): T_ACTION_SHOW_NOTIFICATION: class '" & .Class & "' is not registered", LEMON_LEVEL_WARNING
'                    If pa.CountAlerts = 0 Then
'                        g_Debug "CSnarlSocket.uTranslate(): T_ACTION_SHOW_NOTIFICATION: app '" & .App & "' has no classes!", LEMON_LEVEL_CRITICAL
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

Public Sub Notify(ByVal Notification As Long, ByVal Token As String)

    If Not (theSocket Is Nothing) Then _
        theSocket.SendData uCreateReply(Notification + 270, Token)

End Sub

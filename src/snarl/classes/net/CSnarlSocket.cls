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

    On Error GoTo er

    g_Debug "CSnarlSocket.OnDataArrival(): id=0x" & g_HexStr(mId) & " bytesTotal=" & CStr(bytesTotal), LEMON_LEVEL_PROC

Dim sz As String

    theSocket.GetData sz

Dim i As Integer

    i = InStr(sz, Chr$(13))
    If i > 0 Then
        g_Debug "CSnarlSocket.OnDataArrival(): end_of_packet received"
        mPacket = mPacket & g_SafeLeftStr(sz, i - 1)
        uTranslate mPacket
        mPacket = ""

    Else
        mPacket = mPacket & sz

    End If

    Exit Sub

er:
    g_Debug "CSnarlSocket.OnDataArrival(): " & err.Description, LEMON_LEVEL_CRITICAL
    theSocket.SendData uCreateSNP2Reply(SNARL41_ERROR_BAD_PACKET)

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

    g_Debug "CSnarlSocket.uTranslate(): incoming notification from " & theSocket.RemoteHostIP & ":" & theSocket.RemotePort & "..."

    If g_ConfigGet("listen_for_snarl") = "0" Then
        g_Debug "CSnarlSocket.uTranslate(): ignored (user disabled incoming notifications)"
        theSocket.SendData uCreateReply(SNP_ERROR_NOT_RUNNING)
        Exit Sub

    End If

    If LCase$(g_SafeLeftStr(Request, 6)) = "snp://" Then
        uTranslateSNP2 g_SafeRightStr(Request, Len(Request) - 6)

    Else
        uTranslateSNP Request
    
    End If

End Sub

Private Sub uTranslateSNP2(ByVal Request As String)

    ' /* syntax is: snp://<action>[?arg=value[&arg=value]] CRLF
    '    on entry the 'snp://' prefix will have been stripped */

Dim sCommand As String
Dim ppd As BPackedData
Dim i As Integer

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
            ppd.Add "icon", g_MakePath(App.Path) & "etc\icons\remote.png"

    Case "notify"
        ' /* if encoded icon data exists, extract it now */
        If ppd.Exists("icon64") Then
            sz = uDecodeBase64(Replace$(ppd.ValueOf("icon64"), "%", "="))
            ppd.Add "icon", sz

        End If

    End Select

Dim hr As Long

    ' /* process the command */

    hr = g_DoAction(sCommand, Val(ppd.ValueOf("token")), ppd, True, theSocket)
    If hr <> 0 Then
        ' /* success */
        theSocket.SendData uCreateSNP2Reply(SNARL41_SUCCESS, IIf(hr <> -1, CStr(hr), ""))

    Else
        ' /* failed */
        theSocket.SendData uCreateSNP2Reply(GetProp(ghWndMain, "last_error"))

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
                        g_Debug "CSnarlSocket.uTranslateSNP(): unknown action '" & sz & "'", LEMON_LEVEL_CRITICAL

                    End Select

                Else
                    g_Debug "CSnarlSocket.uTranslateSNP(): unsupported version '" & CStr(pPacket.Version) & "'", LEMON_LEVEL_CRITICAL

                End If
            Else
                g_Debug "CSnarlSocket.uTranslateSNP(): bad packet type '" & sz & "'", LEMON_LEVEL_CRITICAL

            End If
        End If

    End With

    ' /* check here that the packet is valid (i.e. version, type, etc.) */

    If Not pPacket.Valid Then
        theSocket.SendData uCreateReply(SNP_ERROR_BAD_PACKET)
        Exit Sub

    End If

    g_Debug "CSnarlSocket.uTranslateSNP(): packet is valid (ver=" & pPacket.Version & ")"

    ' /* do some general safety checking first */

    If (g_AppRoster Is Nothing) Or (g_NotificationRoster Is Nothing) Then
        g_Debug "CSnarlSocket.uTranslateSNP(): action '" & pPacket.action & "': app_roster or notification_roster not found", LEMON_LEVEL_CRITICAL
        theSocket.SendData uCreateReply(SNP_ERROR_FAILED)
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
                g_Debug "CSnarlSocket.uTranslateSNP(): remote app '" & .App & "' is already registered", LEMON_LEVEL_CRITICAL
                theSocket.SendData uCreateReply(SNP_ERROR_ALREADY_REGISTERED)

            Else
                ' /* use the special "Add2()" method which allows us to register multiple apps using the same PID - in
                '    this case -1, which represents remote apps */
                g_AppRoster.OldAdd .App, 0, 0, g_MakePath(App.Path) & "etc\icons\remote_app.png", g_MakePath(App.Path) & "etc\icons\remote.png", -1
                g_Debug "CSnarlSocket.uTranslateSNP(): remote app '" & .App & "' registered"
                theSocket.SendData uCreateReply(SNP_SUCCESS)

            End If


        Case T_ACTION_REVOKE
            dw = g_AppRoster.IndexOfPidAndName(-1, .App)
            If dw = 0 Then
                ' /* not registered */
                g_Debug "CSnarlSocket.uTranslateSNP(): remote app '" & .App & "' is not registered", LEMON_LEVEL_CRITICAL
                theSocket.SendData uCreateReply(SNP_ERROR_NOT_REGISTERED)

            Else
                ' /* remove safely */
                g_AppRoster.Remove dw
                frmAbout.bUpdateAppList
                g_Debug "CSnarlSocket.uTranslateSNP(): remote app '" & .App & "' unregistered"
                theSocket.SendData uCreateReply(SNP_SUCCESS)

            End If


        Case T_ACTION_ADD_CLASS
            dw = g_AppRoster.IndexOfPidAndName(-1, .App)
            If dw = 0 Then
                ' /* not registered */
                g_Debug "CSnarlSocket.uTranslateSNP(): remote app '" & .App & "' is not registered", LEMON_LEVEL_CRITICAL
                theSocket.SendData uCreateReply(SNP_ERROR_NOT_REGISTERED)

            Else
                ' /* found the app */
                Set pa = g_AppRoster.AppAt(dw)
                If Not pa.FindAlert(.Class, pc) Then
                    ' /* okay to add the class */
                    If pa.AddAlert(.Class, .Title, pc) = M_OK Then _
                        pc.DefaultIcon = pPacket.Icon

                    frmAbout.bUpdateAppList
                    g_Debug "CSnarlSocket.uTranslateSNP(): class '" & .Class & "' add"
                    theSocket.SendData uCreateReply(SNP_SUCCESS)

                Else
                    ' /* class already exists */
                    g_Debug "CSnarlSocket.uTranslateSNP(): class '" & .Class & "' is already registered", LEMON_LEVEL_CRITICAL
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
                g_Debug "CSnarlSocket.uTranslateSNP(): remote app '" & .App & "' is not registered", LEMON_LEVEL_CRITICAL
                dw = snShowMessageEx(SNARL_CLASS_ANON_NET, .Title, .Text, .Timeout, pPacket.Icon, frmAbout.hWnd, WM_REMOTENOTIFY)
                g_Debug "CSnarlSocket.uTranslateSNP(): snShowMessageEx() returned 0x" & g_HexStr(dw)

            Else
                ' /* app found */
                Set pa = g_AppRoster.AppAt(dw)
                dw = pa.Show(.Class, .Title, .Text, .Timeout, pPacket.Icon, frmAbout.hWnd, WM_REMOTENOTIFY)
                g_Debug "CSnarlSocket.uTranslateSNP(): Show() returned 0x" & g_HexStr(dw)

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
            g_Debug "CSnarlSocket.uTranslateSNP(): action '" & .action & "' is unknown", LEMON_LEVEL_CRITICAL
            theSocket.SendData uCreateReply(SNP_ERROR_UNKNOWN_COMMAND)

        End Select

    End With

    Exit Sub

er:
    g_Debug "CSnarlSocket.uTranslateSNP(): " & err.Description, LEMON_LEVEL_CRITICAL
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

Private Function uCreateSNP2Reply(ByVal Error As Long, Optional ByVal Data As String) As String
Dim sz As String

    Select Case Error

    Case SNARL41_SUCCESS
        sz = "OK"

    Case SNARL41_ERROR_FAILED
        sz = "Unspecified failure"

    Case SNARL41_ERROR_UNKNOWN_COMMAND
        sz = "Unknown command"

    Case SNARL41_ERROR_TIMED_OUT
        sz = "Timed out"

    Case SNARL41_ERROR_ARG_MISSING
        sz = "Required argument missing"

    Case SNARL41_ERROR_SYSTEM
        sz = "System error"

    Case SNARL41_ERROR_NOT_RUNNING
        sz = "Not running"

    Case SNARL41_ERROR_NOT_REGISTERED
        sz = "App is not registered"

    Case SNARL41_ERROR_ALREADY_REGISTERED
        sz = "App is already registered"

    Case SNARL41_ERROR_CLASS_ALREADY_EXISTS
        sz = "Class already exists"

    Case SNARL41_ERROR_CLASS_BLOCKED
        sz = "Class blocked"

    Case SNARL41_ERROR_CLASS_NOT_FOUND
        sz = "Class not found"

    Case SNARL41_ERROR_NOTIFICATION_NOT_FOUND
        sz = "Notification not found"

    Case SNARL41_ERROR_FLOODING
        sz = "Flooding detect"

    Case SNARL42_ERROR_DO_NOT_DISTURB
        sz = "Do Not Disturb"

    Case SNARL42_ERROR_COULD_NOT_DISPLAY
        sz = "Could not display"

    Case SNARL42_ERROR_AUTH_FAILURE
        sz = "Authentication failure"

    Case Else
        sz = "Undefined error"

    End Select

    uCreateSNP2Reply = "SNP/2.0/" & CStr(Error) & "/" & sz & IIf(Data <> "", "/" & Data, "") & vbCrLf

End Function

Public Sub Notify(ByVal Notification As Long, ByVal Token As String)

    If Not (theSocket Is Nothing) Then _
        theSocket.SendData uCreateReply(Notification + 270, Token)

End Sub

Private Function uDecodeBase64(ByVal Data As String) As String
Dim sz As String
Dim i As Integer

    On Error Resume Next

    sz = Decode64(Data)
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


'Private Sub uTranslateSNP2(ByVal Request As String)
'
'    ' /* SNP2.0 format is "snp://<action>[?arg=value[&arg=value]] CRLF" */
'
'Dim sCommand As String
'Dim ppd As BPackedData
'Dim i As Integer
'Dim hr As Long
'
'    ' /* on entry "snp://" prefix has been stripped */
'
'    Set ppd = New BPackedData
'
'    ' /* find the command */
'
'    i = InStr(Request, "?")
'    If i = 0 Then
'        ' /* entire line is command */
'        sCommand = Request
'
'    Else
'        ' /* get the command */
'        sCommand = g_SafeLeftStr(Request, i - 1)
'
'        Debug.Print Request
'
'        ' /* get the args */
'        Request = g_SafeRightStr(Request, Len(Request) - i)
'
'        ' /* translate from SNP syntax into Win32 syntax */
'        Request = Replace(Request, "&", "#?")
'        Request = Replace(Request, "=", "::")
'
'        ' /* create the BPackedData */
'        ppd.SetTo Request
'
'    End If
'
'    ' /* parse the command */
'
'    Select Case sCommand
'
'    Case "test"
'
'        ' /* only available when Snarl is running in debug mode */
'
'        If gDebugMode Then
'            sn41EZNotify gSnarlToken, "", "SNP2.0", "Test Message"
'            theSocket.SendData uCreateSNP2Reply(SNARL41_SUCCESS)
'
'        Else
'            theSocket.SendData uCreateSNP2Reply(SNARL41_ERROR_UNKNOWN_COMMAND)
'
'        End If
'
'
'    Case "hello"
'
'        ' /* just reply with SUCCESS but it means the sender will
'        '    be able to tell which version of SNP we support */
'
'        ' /* To-do: reply with an error message if Snarl isn't
'        '    accepting requests, or DND mode enabled? */
'
'        theSocket.SendData uCreateSNP2Reply(SNARL41_SUCCESS)
'
'
'    Case "register", "reg"
'
'        ' /* register the app as per args provided */
'
'        If Not ppd.Exists("icon") Then _
'            ppd.Add "icon", g_MakePath(App.Path) & "etc\icons\remote.png"
'
'        hr = g_AppRoster.Add41(ppd, True)
'        If hr = 0 Then
'            theSocket.SendData uCreateSNP2Reply(GetProp(ghWndMain, "last_error"))
'
'        Else
'            theSocket.SendData uCreateSNP2Reply(SNARL41_SUCCESS, CStr(hr))
'
'        End If
'
'
'    Case "unregister", "unreg"
'
'        ' /* unregister the app as per the args provided */
'
'        hr = g_AppRoster.Unregister(Val(ppd.ValueOf("token")), ppd.ValueOf("password"))
'        If hr = 0 Then
'            theSocket.SendData uCreateSNP2Reply(GetProp(ghWndMain, "last_error"))
'
'        Else
'            theSocket.SendData uCreateSNP2Reply(SNARL41_SUCCESS)
'
'        End If
'
'
'    Case "updateapp", "update_app"
'
'        ' /* modify the app as per the args provided */
'
'        hr = g_AppRoster.Update(Val(ppd.ValueOf("token")), ppd)
'        If hr = 0 Then
'            theSocket.SendData uCreateSNP2Reply(GetProp(ghWndMain, "last_error"))
'
'        Else
'            theSocket.SendData uCreateSNP2Reply(SNARL41_SUCCESS)
'
'        End If
'
'
'    ' /* for info: SNARL41_SET_CALLBACK is not required as all SNP2.0 applications inherently
'    '    receive feedback to their local port */
'
'Dim pa As TApp
'
'    Case "addclass"
'
'        Debug.Print "XAC: " & g_SafeFindApp(Val(ppd.ValueOf("token")), ppd.ValueOf("password"), pa)
'
'        hr = sn41AddClass(Val(ppd.ValueOf("token")), ppd.ValueOf("id"), ppd.ValueOf("name"))
'        If hr = 0 Then
'            theSocket.SendData uCreateSNP2Reply(GetProp(ghWndMain, "last_error"))
'
'        Else
'            theSocket.SendData uCreateSNP2Reply(SNARL41_SUCCESS)
'
'        End If
'
'
'
'    Case "remclass"
'
'        Debug.Print "XRC: " & g_SafeFindApp(Val(ppd.ValueOf("token")), ppd.ValueOf("password"), pa)
'
'        hr = sn41RemoveClass(Val(ppd.ValueOf("token")), ppd.ValueOf("id"))
'        If hr = 0 Then
'            theSocket.SendData uCreateSNP2Reply(GetProp(ghWndMain, "last_error"))
'
'        Else
'            theSocket.SendData uCreateSNP2Reply(SNARL41_SUCCESS)
'
'        End If
'
'
'
'    Case "killclasses"
'        sn41RemoveAllClasses Val(ppd.ValueOf("token"))
'        theSocket.SendData uCreateSNP2Reply(SNARL41_SUCCESS)
'
'
'    Case "notify"
'
'        ' /* create a notification based on the args provided */
'
'        ppd.Add "remote", "1"       ' // add this to indicate sender is remote (or using IP, at least)
'
'Dim sz As String
'
'        If ppd.Exists("icon64") Then
'            sz = uDecodeBase64(Replace$(ppd.ValueOf("icon64"), "%", "="))
'            ppd.Add "icon", sz
'
'        End If
'
'        hr = g_DoNotify(Val(ppd.ValueOf("token")), ppd, theSocket)
'        If hr = 0 Then
'            theSocket.SendData uCreateSNP2Reply(GetProp(ghWndMain, "last_error"))
'
'        Else
'            theSocket.SendData uCreateSNP2Reply(SNARL41_SUCCESS, CStr(hr))
'
'        End If
'
'        If sz <> "" Then _
'            DeleteFile sz
'
'
'    Case "update"
'
'        ' /* update a notification based on the args provided */
'
'        hr = g_DoUpdate(Val(ppd.ValueOf("token")), ppd)
'        If hr = 0 Then
'            theSocket.SendData uCreateSNP2Reply(GetProp(ghWndMain, "last_error"))
'
'        Else
'            theSocket.SendData uCreateSNP2Reply(SNARL41_SUCCESS)
'
'        End If
'
'
'    Case "hide"
'        hr = sn41Hide(Val(ppd.ValueOf("token")))
'        If hr = 0 Then
'            theSocket.SendData uCreateSNP2Reply(GetProp(ghWndMain, "last_error"))
'
'        Else
'            theSocket.SendData uCreateSNP2Reply(SNARL41_SUCCESS)
'
'        End If
'
'
'    Case "isvisible"
'        hr = sn41IsVisible(Val(ppd.ValueOf("token")))
'        If hr = 0 Then
'            theSocket.SendData uCreateSNP2Reply(GetProp(ghWndMain, "last_error"))
'
'        Else
'            theSocket.SendData uCreateSNP2Reply(SNARL41_SUCCESS)
'
'        End If
'
'
'    Case "lasterror"
'        theSocket.SendData uCreateSNP2Reply(GetProp(ghWndMain, "last_error"))
'
'
'    Case "version"
'        theSocket.SendData uCreateSNP2Reply(SNARL41_SUCCESS, CStr(GetProp(ghWndMain, "_version")))
'
''    SNARL42_ADD_ACTION
''    SNARL42_CLEAR_ACTIONS
''    SNARL42_SHOW_REQUEST
'
'    Case Else
'        g_Debug "CSnarlSocket(): uTranslateSNP2(): unknown command '" & sCommand & "'", LEMON_LEVEL_WARNING
'        theSocket.SendData uCreateSNP2Reply(SNARL41_ERROR_UNKNOWN_COMMAND)
'
'    End Select
'
'End Sub



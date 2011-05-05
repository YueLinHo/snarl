Attribute VB_Name = "mGNTPSupport"
Option Explicit

    '/*********************************************************************************************
    '/
    '/  File:           mGNTPSupport.bas
    '/
    '/  Description:    GNTP support routines
    '/
    '/  � 2011 full phat products
    '/
    '/  This file may be used under the terms of the Simplified BSD Licence
    '/
    '*********************************************************************************************/

Public Enum E_GNTP_CODES
    INVALID_REQUEST = 300           '// The request contained an unsupported directive, invalid headers or values, or was otherwise malformed
    UNKNOWN_PROTOCOL = 301          '// The request was not a GNTP request
    UNKNOWN_PROTOCOL_VERSION = 302  '// The request specified an unknown or unsupported GNTP version
    REQUIRED_HEADER_MISSING = 303   '// The request was missing required information
    NOT_AUTHORIZED = 400            '// The request supplied a missing or wrong password/key or was otherwise not authorized
    UNKNOWN_APPLICATION = 401       '// Application is not registered to send notifications
    UNKNOWN_NOTIFICATION = 402      '// Notification type is not registered by the application
    INTERNAL_SERVER_ERROR = 500     '// An internal server error occurred while processing the request

End Enum


Private Type T_NOTIFY_TYPE
    Name As String
    DisplayName As String
    Enabled As Boolean
    Icon As String

End Type

Private Type T_REG
    Token As Long
    AppName As String
    NotificationType() As T_NOTIFY_TYPE
    Count As Long
    Signature As String
    IconPath As String

End Type

Dim mRegistration As T_REG

Private Type T_GNTP_NOTIFICATION
    AppName As String
    Name As String
    Title As String
    Id As String
    Text As String
    Sticky As Boolean
    Priority As Long
    Icon As String
    CoalesceID As String
    CallbackContext As String
    CallbackContextType As String
    CallbackContextTarget As String

End Type

Dim mSection() As String
Dim mDirective As String
Dim mCustomHeaders As String

' /*********************************************************************************************
'   gntp_Process() -- Master GNTP request handler
'
'   Inputs
'       Request - unabridged request content
'       Sender - sending socket object (only used for notifications)
'
'   Outputs
'       Response - GNTP response that should be sent back to the source socket
'
'   Return Value
'       True if the request was processed successfully, false otherwise.  Note that, as per
'       GNTP specifications, the resulting Response should be sent to the socket and the
'       socket closed irrespective of success or failure.
'
' *********************************************************************************************/

Public Function gntp_Process(ByVal Request As String, ByRef Response As String, ByRef Sender As CSocket) As Boolean

    ' /* return a GNTP error code here */

    On Error GoTo er

    mCustomHeaders = ""

    uOutput ""
    uOutput "gntp_Process(): " & Replace$(Request, vbCrLf, "�")

    ' /* split into sections */

    mSection = Split(Request, vbCrLf & vbCrLf)
    uOutput "gntp_Process(): sections=" & UBound(mSection) - 1

    If UBound(mSection) < 1 Then
        uOutput "gntp_Process(): failed: no sections"
        Response = uCreateResponse(INVALID_REQUEST)
        Exit Function

    End If

    ' /* parse section 1 - must be the info block */

    mDirective = ""

    If Not uParse(0, Response) Then
        uOutput "gntp_Process(): failed: invalid info block"
        Response = uCreateResponse(INVALID_REQUEST)
        Exit Function

    End If

    uOutput "gntp_Process(): info block okay, directive=" & mDirective

    Select Case mDirective
    Case "REGISTER"
        gntp_Process = uDoRegistration(Response)

    Case "NOTIFY"
        gntp_Process = uDoNotification(Response, Sender)

    Case Else
        uOutput "gntp_Process(): unsupported directive '" & mDirective & "'"
        Response = uCreateResponse(INVALID_REQUEST)

    End Select

''    ' /* was it successful? */
''
''    If gntp_Process <> 0 Then
''        ' /* error response headers */
''         mResponse = "Error-Code: " & CStr(gntp_Process) & vbCrLf & _
''                     "Error-Description: " & uErrStr(gntp_Process) & IIf(mExtError <> "", "(" & mExtError & ")", "") & vbCrLf
''
''    End If

    uOutput "gntp_Process(): done"

'    gntp_Process = True

    Exit Function

er:
    Debug.Print "gntp_Process(): panic: " & Err.Description

End Function

'Private Function uErrStr(ByVal Code As Long) As String
'
'    Select Case Code
'    Case 300
'        'The request contained an unsupported directive, invalid headers or values, or was otherwise malformed
'        uErrStr = "Invalid request"
'
'    Case 301
'        'The request was not a GNTP request
'        uErrStr = "Unknown protocol"
'
'    Case 302
'        'The request specified an unknown or unsupported GNTP version
'        uErrStr = "Unknown protocol version"
'
'    Case 303
'        'The request was missing required information
'        uErrStr = "Required header missing"
'
'    Case 400
'        'The request supplied a missing or wrong password/key or was otherwise not authorized
'        uErrStr = "Not authorized"
'
'    Case 401
'        'Application is not registered to send notifications
'        uErrStr = "Unknown application"
'
'    Case 402
'        'Notification type is not registered by the application
'        uErrStr = "Unknown notification"
'
'    Case 500
'        'An internal server error occurred while processing the request
'        uErrStr = "Internal server error"
'
'    End Select
'
'End Function

Private Function uBool(ByVal str As String) As Boolean

    Select Case LCase$(str)
    Case "yes", "true"
        uBool = True

    End Select

End Function

Private Function uDoRegistration(ByRef Response As String) As Boolean

    mCustomHeaders = "Response-Action: REGISTER" & vbCrLf

Dim pp As BPackedData

    Set pp = New BPackedData
    If Not pp.SetTo(mSection(0), vbCrLf, ": ") Then
        uOutput "uDoRegistration(): bad data"
        Response = uCreateResponse(INVALID_REQUEST)
        Exit Function

    End If

    ' /* required items */

Dim szAppName As String
Dim szAppSig As String

    If Not pp.Exists("Application-Name") Then
        uOutput "uDoRegistration(): missing app name"
        Response = uCreateResponse(REQUIRED_HEADER_MISSING)
        Exit Function

    Else
        szAppName = pp.ValueOf("Application-Name")
        szAppSig = "application/x-gntp-" & Replace$(szAppName, " ", "_")

    End If

Dim dwCount As Long

    If Not pp.Exists("Notifications-Count") Then
        uOutput "uDoRegistration(): missing count"
        Response = uCreateResponse(REQUIRED_HEADER_MISSING)
        Exit Function

    Else
        dwCount = g_SafeLong(pp.ValueOf("Notifications-Count"))

    End If

    ' /* special Snarl feature: zero notifications means unregister */

    If dwCount = 0 Then
        uOutput "uDoRegistration(): app requested an unregister"
        Response = uCreateResponse(0)

#If GNTP_TEST Then
        snarl_unregister szAppSig

#Else
        g_DoAction "unreg", 0, g_newBPackedData("app-sig::" & szAppSig)

#End If
        Exit Function

    End If


Dim px As T_REG

    LSet mRegistration = px

    With mRegistration
        .AppName = szAppName
        .Signature = szAppSig
        .Count = dwCount
        ReDim .NotificationType(.Count)

        If pp.Exists("Application-Icon") Then _
            .IconPath = pp.ValueOf("Application-Icon")

    End With

    ' /* otherwise must have the right number of sections */

    If UBound(mSection) < mRegistration.Count Then
        uOutput "uDoRegistration(): not enough sections"
        Response = uCreateResponse(INVALID_REQUEST)
        Exit Function

    End If

    ' /* sections 1 to pr.Count should be notification types */

Dim i As Long

    For i = 1 To mRegistration.Count
        ' /* if adding any of the notification types fails, we fail */
        If Not uAddNotificationType(mSection(i), i - 1, Response) Then _
            Exit Function

    Next i

    ' /* sections pr.Count to end should be resource identifiers */

    uOutput "uDoRegistration(): parsing resource identifiers " & CStr(mRegistration.Count) & " to " & CStr(UBound(mSection) - 1)

    For i = mRegistration.Count To UBound(mSection) - 1
        uParse i, ""

    Next i

    uOutput "uDoRegistration(): registering with Snarl..."

    ' /* register here */

    With mRegistration

#If GNTP_TEST Then
        .Token = snarl_register(.Signature, .AppName, .IconPath)

#Else
        .Token = g_DoAction("register", 0, _
                                         g_newBPackedData("app-sig::" & .Signature & _
                                                          "#?title::" & .AppName & _
                                                          "#?icon::" & .IconPath))
#End If

    End With

    ' /* if registration fails, quit now */

    If mRegistration.Token < 1 Then
        uOutput "uDoRegistration(): registration failed (" & CStr(Abs(mRegistration.Token)) & ")"
        Response = uCreateResponse(INTERNAL_SERVER_ERROR)
        Exit Function

    End If

    ' /* add notification types as classes */

Dim hr As Long

    With mRegistration
        For i = 0 To .Count - 1

#If GNTP_TEST Then
            hr = snDoRequest("addclass?app-sig=" & .Signature & _
                             "&id=" & .NotificationType(i).Name & _
                             "&name=" & .NotificationType(i).DisplayName & _
                             "&enabled=" & IIf(.NotificationType(i).Enabled, "1", "0") & _
                             "&icon=" & .NotificationType(i).Icon)

            uOutput "uDoRegistration(): addclass '" & .NotificationType(i).Name & "' returned " & CStr(hr)

#Else

            g_DoAction "addclass", 0, _
                       g_newBPackedData("app-sig::" & mRegistration.Signature & _
                                        "#?id::" & .NotificationType(i).Name & _
                                        "#?name::" & .NotificationType(i).DisplayName & _
                                        "#?enabled::" & IIf(.NotificationType(i).Enabled, "1", "0") & _
                                        "#?icon::" & .NotificationType(i).Icon)
#End If

        Next i

    End With

    ' /* done */

    Response = uCreateResponse(0)
    uDoRegistration = True

End Function

Private Function uAddNotificationType(ByVal str As String, ByVal Index As Long, ByRef Response As String) As Boolean
Dim pp As BPackedData

    Set pp = New BPackedData
    If Not pp.SetTo(str, vbCrLf, ": ") Then
        uOutput "uAddNotificationType(): invalid data for notification type #" & CStr(Index)
        Response = uCreateResponse(INVALID_REQUEST)
        Exit Function

    End If

    'For each notification being registered:
    '
    'Notification-Name: <string>
    'Required - The name (type) of the notification being registered

    If Not pp.Exists("Notification-Name") Then
        uOutput "uAddNotificationType(): missing required arg Notification-Name for notification type #" & CStr(Index)
        Response = uCreateResponse(REQUIRED_HEADER_MISSING)
        Exit Function

    End If

    'Notification-Display-Name: <string>
    'Optional - The name of the notification that is displayed to the user (defaults to the same value as Notification-Name)
    '
    'Notification-Enabled: <boolean>
    'Optional - Indicates if the notification should be enabled by default (defaults to False)
    '
    'Notification-Icon: <url> | <uniqueid>
    'Optional - The default icon to use for notifications of this type
    '
    'Each notification being registered should be seperated by a blank line, including the first notification.

Dim sx() As String

    With mRegistration.NotificationType(Index)
        .Name = pp.ValueOf("Notification-Name")
        .DisplayName = pp.ValueOf("Notification-Display-Name")
        .Enabled = uBool(pp.ValueOf("Notification-Enabled"))
        .Icon = pp.ValueOf("Notification-Icon")

        ' /* if the icon is actually an identifier, map it to the locally saved copy */

        If g_SafeLeftStr(.Icon, 19) = "x-growl-resource://" Then
            sx = Split(.Icon, "://")
            .Icon = g_GetTempPath() & "gntp-res-" & sx(1) & ".png"

        End If

        ' /* as per GNTP specification */

        If .DisplayName = "" Then _
            .DisplayName = .Name

        uOutput "uAddNotificationType(): got notification type " & CStr(Index) & " (" & .Name & "�" & .DisplayName & "�" & IIf(.Enabled, "Enabled", "Disabled") & "�" & .Icon & ")"

    End With

    uAddNotificationType = True

End Function

Private Function uParse(ByVal SectionIndex As Long, ByRef Response As String) As Boolean
Dim s() As String

    ' /* parses the first line of the given section and returns
    '    the appropriate result */

    s = Split(mSection(SectionIndex), vbCrLf)
    If UBound(s) < 1 Then
        uOutput "uParse(): failed: invalid section"
        Response = uCreateResponse(INVALID_REQUEST)
        Exit Function

    End If

    uOutput "uParse(): section header='" & s(0) & "'"

    ' /* identify section type from the first line */

Dim x() As String

    If g_SafeLeftStr(s(0), 4) = "GNTP" Then
        ' /* information line first */
        uParse = uParseInfoLine(s(0), Response)

    ElseIf g_SafeLeftStr(s(0), 12) = "Identifier: " Then
        ' /* resource identifier */
        x = Split(s(0), ": ")
        uSaveBinary SectionIndex + 1, x(1)
        uParse = True

    Else
        ' /* other headers... */

    End If

End Function

Private Function uParseInfoLine(ByVal str As String, ByRef Response As String) As Boolean
Dim s() As String

    s = Split(str, " ")
    If UBound(s) < 2 Then
        uOutput "uParseInfoLine(): not enough params"
        Response = uCreateResponse(INVALID_REQUEST)
        Exit Function           ' // not enough params

    End If

Dim v() As String

    v = Split(s(0), "/")
    If UBound(v) <> 1 Then
        uOutput "uParseInfoLine(): not GNTP"
        Response = uCreateResponse(UNKNOWN_PROTOCOL)
        Exit Function           ' // not GNTP
    
    End If

    If v(0) <> "GNTP" Then
        uOutput "uParseInfoLine(): not GNTP"
        Response = uCreateResponse(UNKNOWN_PROTOCOL)
        Exit Function           ' // not GNTP
    
    End If

    If v(1) <> "1.0" Then
        uOutput "uParseInfoLine(): not 1.0"
        Response = uCreateResponse(UNKNOWN_PROTOCOL_VERSION)
        Exit Function

    End If

    mDirective = ""

    Select Case s(1)
    Case "REGISTER", "NOTIFY"
        mDirective = s(1)

    Case Else
        uOutput "uParseInfoLine(): unsupported directive"
        Response = uCreateResponse(INVALID_REQUEST)
        Exit Function           ' // bad directive

    End Select

    Select Case s(2)
    Case "NONE"
    
    Case Else
        uOutput "uParseInfoLine(): unsupported encryption (" & s(2) & ")"
        Response = uCreateResponse(INVALID_REQUEST)
        Exit Function           ' // unsupported encryption

    End Select

    uParseInfoLine = True

End Function

Private Sub uSaveBinary(ByVal SectionIndex As Long, ByVal Identifier As String)
Dim i As Integer

    On Error Resume Next

    Identifier = "gntp-res-" & Identifier
    uOutput "uSaveBinary(): writing binary content to " & g_GetTempPath() & Identifier & ".png"

    i = FreeFile()
    Open g_GetTempPath() & Identifier & ".png" For Binary Access Write As #i
    Put #i, , mSection(SectionIndex)
    Close #i

End Sub

Private Function uDoNotification(ByRef Response As String, ByRef Sender As CSocket) As Boolean

    mCustomHeaders = "Response-Action: NOTIFY"

Dim pp As BPackedData

    ' /* convert the mime-style section into packed data */

    Set pp = New BPackedData
    If Not pp.SetTo(mSection(0), vbCrLf, ": ") Then
        uOutput "uDoNotification(): bad data"
        Response = uCreateResponse(INVALID_REQUEST)
        Exit Function

    End If

    ' /* required items */

    If (Not pp.Exists("Application-Name")) Or (Not pp.Exists("Notification-Name")) Or (Not pp.Exists("Notification-Title")) Then
        uOutput "uDoNotification(): missing required arg"
        Response = uCreateResponse(REQUIRED_HEADER_MISSING)
        Exit Function

    End If

'Dim pn As T_GNTP_NOTIFICATION
'
'    With pn
'        'Notification-Callback-Context: <string>
'        'Optional - Any data (will be passed back in the callback unmodified)
'        .CallbackContext = pp.ValueOf("Notification-Callback-Context")
'        'Notification-Callback-Context-Type: <string>
'        'Optional, but Required if 'Notification-Callback-Context' is passed - The type of data being passed in
'        'Notification-Callback-Context (will be passed back in the callback unmodified). This does not need to be of any pre-defined
'        'type, it is only a convenience to the sending application.
'        .CallbackContextType = pp.ValueOf("Notification-Callback-Context-Type")
'
'    End With

    ' /* build the Snarl packet */

Dim px As BPackedData

    Set px = New BPackedData
    With px
        'Application-Name: <string>
        'Required - The name of the application that sending the notification (must match a previously registered application)
        .Add "app-sig", "application/x-gntp-" & Replace$(pp.ValueOf("Application-Name"), " ", "_")

        'Notification-Name: <string>
        'Required - The name (type) of the notification (must match a previously registered notification name registered by the
        'application specified in Application-Name)
        .Add "id", pp.ValueOf("Notification-Name")

        'Notification-Title: <string>
        'Required - The notification's title
        .Add "title", g_toUnicodeUTF8(Replace$(pp.ValueOf("Notification-Title"), Chr$(13), vbCrLf))
        
        'Notification-Text: <string>
        'Optional - The notification's text. (defaults to "")
        .Add "text", g_toUnicodeUTF8(Replace$(pp.ValueOf("Notification-Text"), Chr$(13), vbCrLf))

        'Notification-Sticky: <boolean>
        'Optional - Indicates if the notification should remain displayed until dismissed by the user. (default to False)
        ' /* sticky == zero duration under Snarl */
        If uBool(pp.ValueOf("Notification-Sticky")) Then _
            .Add "timeout", "0"

        'Notification-Priority: <int>
        'Optional - A higher number indicates a higher priority. This is a display hint for the receiver which may be ignored. (valid
        'values are between -2 and 2, defaults to 0)
        .Add "priority", g_SafeLong(pp.ValueOf("Notification-Priority"))

        'Notification-Coalescing-ID: <string>
        'Optional - If present, should contain the value of the Notification-ID header of a previously-sent notification. This serves
        'as a hint to the notification system that this notification should replace/update the matching previous notification. The
        'notification system may ignore this hint.
        If pp.Exists("Notification-Coalescing-ID") Then _
            .Add "update-uid", pp.ValueOf("Notification-Coalescing-ID")

        'Notification-ID: <string>
        'Optional - A unique ID for the notification. If used, this should be unique for every request, even if the notification is
        'replacing a current notification (see Notification-Coalescing-ID)
        If pp.Exists("Notification-ID") Then _
            .Add "uid", pp.ValueOf("Notification-ID")

        'Notification-Callback-Target: <string>
        'Optional - An alternate target for callbacks from this notification. If passed, the standard behavior of performing the
        'callback over the original socket will be ignored and the callback data will be passed to this target instead. See the 'Url
        'Callbacks' section for more information.
        If pp.ValueOf("Notification-Callback-Target") <> "" Then _
            .Add "callback", pp.ValueOf("Notification-Callback-Target")


        ' /* sort out the icon */

Dim sx() As String
Dim sz As String

        'Notification-Icon: <url> | <uniqueid>
        'Optional - The icon to display with the notification.
        sz = pp.ValueOf("Notification-Icon")
        If g_SafeLeftStr(sz, 19) = "x-growl-resource://" Then
            sx = Split(sz, "://")
            .Add "icon", g_GetTempPath() & "gntp-res-" & sx(1) & ".png"

        ElseIf sz <> "" Then
            .Add "icon", sz

        End If

    End With

    ' /* do the notification */

Dim hr As Long

#If GNTP_TEST = 1 Then

    sz = px.AsString()
    sz = Replace$(sz, "::", "=")
    sz = Replace$(sz, "#?", "&")
    hr = snDoRequest("notify?" & sz)

    If hr > 0 Then
        Response = uCreateResponse(0)
        uDoNotification = True
        
    End If

#Else
    hr = g_DoAction("notify", 0, px, IIf(Sender.LocalIP <> "127.0.0.1", NF_REMOTE, 0) Or App.Major, Sender)
    If hr = 0 Then
        ' /* failed */

        uOutput "uDoNotification(): <notify> failed: " & CStr(g_QuickLastError())
        Select Case g_QuickLastError()
        Case SNARL_ERROR_AUTH_FAILURE
            Response = uCreateResponse(NOT_AUTHORIZED)

        Case SNARL_ERROR_NOT_REGISTERED
            Response = uCreateResponse(UNKNOWN_APPLICATION)

        Case Else
            Response = uCreateResponse(INTERNAL_SERVER_ERROR)

        End Select

    Else
        Response = uCreateResponse(0)
        uDoNotification = True

    End If


#End If

End Function

Private Function uCreateResponse(ByVal ResponseCode As E_GNTP_CODES) As String
Dim sz As String

    sz = "GNTP/1.0 "

    Select Case ResponseCode
    Case 0
        sz = sz & "-OK"

    Case Else
        sz = sz & "-ERROR"

    End Select

    sz = sz & " NONE" & vbCrLf

    If mCustomHeaders <> "" Then _
        sz = sz & mCustomHeaders & vbCrLf

    ' /* generic headers */

    sz = sz & "Origin-Machine-Name: " & get_host_name() & vbCrLf

#If GNTP_TEST = 1 Then
    sz = sz & "Origin-Software-Name: " & App.ProductName & vbCrLf
    sz = sz & "Origin-Software-Version: " & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf

#Else
    sz = sz & "Origin-Software-Name: Snarl" & vbCrLf
    sz = sz & "Origin-Software-Version: " & CStr(APP_VER) & "." & CStr(APP_SUB_VER) & " (" & CStr(App.Major) & "." & CStr(App.Revision) & ")" & vbCrLf

#End If
    sz = sz & "Origin-Platform-Name: Windows" & vbCrLf
    sz = sz & "Origin-Platform-Version: " & g_GetOSVersionString() & vbCrLf

    sz = sz & "X-Timestamp: " & Format$(Now(), "MM/DD/YYYY HH:MM:SS AMPM")

    uCreateResponse = sz & vbCrLf & vbCrLf

End Function

Private Sub uOutput(ByVal Text As String)

#If GNTP_TEST = 1 Then
    Form1.Output Text

#Else
    g_Debug Text

#End If

End Sub

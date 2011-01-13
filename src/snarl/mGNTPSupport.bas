Attribute VB_Name = "mGNTPSupport"
Option Explicit

    '/*********************************************************************************************
    '/
    '/  File:           mGNTPSupport.bas
    '/
    '/  Description:    GNTP support routines
    '/
    '/  © 2011 full phat products
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
Dim mResponse As String
Dim mExtError As String

Public Function gntp_Process(ByVal Request As String, ByRef ReplySocket As CSocket, ByRef Response As String) As E_GNTP_CODES

    ' /* return a GNTP error code here */

    On Error GoTo er

    Debug.Print "length: " & Len(Request)
    mResponse = ""

    ' /* split into sections */

    mSection = Split(Request, vbCrLf & vbCrLf)
    Debug.Print "sections: " & UBound(mSection)

    If UBound(mSection) > 0 Then

        ' /* parse section 1 - must be the info block */

        mDirective = ""

        gntp_Process = uParse(0)
        If gntp_Process = 0 Then
            Debug.Print "directive: " & mDirective

            Select Case mDirective
            Case "REGISTER"
                gntp_Process = uDoRegistration()
                If gntp_Process = 0 Then _
                    mResponse = "Response-Action: REGISTER" & vbCrLf

            Case "NOTIFY"
                gntp_Process = uDoNotification(ReplySocket)
                If gntp_Process = 0 Then _
                    mResponse = "Response-Action: NOTIFY" & vbCrLf

            Case Else
                Debug.Print "gntp_process(): unsupported directive '" & mDirective & "'"
                gntp_Process = INVALID_REQUEST

            End Select

        Else
            ' /* return value already set */
            Debug.Print "gntp_process(): invalid info block"

        End If

    Else
        Debug.Print "gntp_process(): invalid number of sections"
        gntp_Process = INVALID_REQUEST

    End If

    ' /* was it successful? */

    If gntp_Process <> 0 Then
        ' /* error response headers */
         mResponse = "Error-Code: " & CStr(gntp_Process) & vbCrLf & _
                     "Error-Description: " & uErrStr(gntp_Process) & IIf(mExtError <> "", "(" & mExtError & ")", "") & vbCrLf

    End If

    Exit Function

er:
    Debug.Print err.Description

End Function

Private Function uErrStr(ByVal Code As Long) As String

    Select Case Code
    Case 300
        'The request contained an unsupported directive, invalid headers or values, or was otherwise malformed
        uErrStr = "Invalid request"

    Case 301
        'The request was not a GNTP request
        uErrStr = "Unknown protocol"

    Case 302
        'The request specified an unknown or unsupported GNTP version
        uErrStr = "Unknown protocol version"

    Case 303
        'The request was missing required information
        uErrStr = "Required header missing"

    Case 400
        'The request supplied a missing or wrong password/key or was otherwise not authorized
        uErrStr = "Not authorized"

    Case 401
        'Application is not registered to send notifications
        uErrStr = "Unknown application"

    Case 402
        'Notification type is not registered by the application
        uErrStr = "Unknown notification"

    Case 500
        'An internal server error occurred while processing the request
        uErrStr = "Internal server error"

    End Select

End Function

Private Function uBool(ByVal str As String) As Boolean

    Select Case LCase$(str)
    Case "yes", "true"
        uBool = True

    End Select

End Function

Private Function uDoRegistration() As E_GNTP_CODES
Dim pp As BPackedData

    Set pp = New BPackedData
    If Not pp.SetTo(mSection(0), vbCrLf, ": ") Then
        Debug.Print "uDoRegistration: bad data"
        uDoRegistration = INVALID_REQUEST
        Exit Function

    End If

    ' /* required items */

    If Not pp.Exists("Application-Name") Then
        Debug.Print "uDoRegistration: missing app name"
        uDoRegistration = REQUIRED_HEADER_MISSING
        Exit Function

    End If

    If Not pp.Exists("Notifications-Count") Then
        Debug.Print "uDoRegistration: missing count"
        uDoRegistration = REQUIRED_HEADER_MISSING
        Exit Function

    End If

Dim px As T_REG

    LSet mRegistration = px

    With mRegistration
        .AppName = pp.ValueOf("Application-Name")
        .Count = g_SafeLong(pp.ValueOf("Notifications-Count"))
        ReDim .NotificationType(.Count)
        .Signature = "application/x-gntp-" & Replace$(.AppName, " ", "_")

    End With

    ' /* special Snarl feature: zero notifications means unregister */

    If mRegistration.Count = 0 Then
        Debug.Print "uDoRegistration: app requested unregister"
        g_DoAction "unreg", 0, g_newBPackedData("app-sig::" & mRegistration.Signature)
        Exit Function

    End If

    ' /* otherwise must have the right number of sections */

    If UBound(mSection) < mRegistration.Count Then
        Debug.Print "uDoRegistration: not enough sections"
        uDoRegistration = INVALID_REQUEST
        Exit Function

    End If

    ' /* sections 1 to pr.Count should be notification types */

Dim i As Long

    For i = 1 To mRegistration.Count
        uDoRegistration = uAddNotificationType(mSection(i), i - 1)
        If uDoRegistration <> 0 Then
            Debug.Print "uDoRegistration: bad notification type " & i
            Exit Function

        End If
    Next i

    ' /* sections pr.Count to end should be resource identifiers */

    Debug.Print "uDoRegistration: parsing resource identifiers " & CStr(mRegistration.Count) & " to " & CStr(UBound(mSection))

    For i = mRegistration.Count To UBound(mSection)
        uParse i

    Next i

    ' /* register here */

    mRegistration.Token = g_DoAction("register", 0, _
                                     g_newBPackedData("app-sig::" & mRegistration.Signature & _
                                                      "#?title::" & mRegistration.AppName & _
                                                      "#?icon::" & g_MakePath(App.Path) & "etc\icons\growl_app.png"))

    If mRegistration.Token = 0 Then
        Debug.Print "uDoRegistration: registration failed (" & CStr(g_QuickLastError()) & ")"
        uDoRegistration = INTERNAL_SERVER_ERROR
        Exit Function

    End If

    With mRegistration
        For i = 0 To .Count - 1
            ' /* check for error here? */
            Debug.Print "reg->" & g_DoAction("addclass", 0, _
                                             g_newBPackedData("app-sig::" & mRegistration.Signature & _
                                                              "#?id::" & .NotificationType(i).Name & _
                                                              "#?name::" & .NotificationType(i).DisplayName & _
                                                              "#?enabled::" & IIf(.NotificationType(i).Enabled, "1", "0") & _
                                                              "#?icon::" & .NotificationType(i).Icon))

        Next i

    End With

End Function

Private Function uAddNotificationType(ByVal str As String, ByVal Index As Long) As E_GNTP_CODES
Dim pp As BPackedData

    Set pp = New BPackedData
    If Not pp.SetTo(str, vbCrLf, ": ") Then
        Debug.Print "uAddNotificationType(): bad data"
        uAddNotificationType = INVALID_REQUEST
        Exit Function

    End If

    'For each notification being registered:
    '
    'Notification-Name: <string>
    'Required - The name (type) of the notification being registered

    If Not pp.Exists("Notification-Name") Then
        Debug.Print "missing required arg"
        uAddNotificationType = REQUIRED_HEADER_MISSING
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

        ' /* sort out the icon */

        If g_SafeLeftStr(.Icon, 19) = "x-growl-resource://" Then
            sx = Split(.Icon, "://")
            .Icon = g_GetTempPath() & "gntp-res-" & sx(1) & ".png"

        End If

        ' /* as per GNTP specification */

        If .DisplayName = "" Then _
            .DisplayName = .Name

    End With

End Function

Private Function uParse(ByVal SectionIndex As Long) As E_GNTP_CODES
Dim s() As String

    ' /* parses the first line of the given section and returns
    '    the appropriate result */

    s = Split(mSection(SectionIndex), vbCrLf)
    If UBound(s) < 1 Then
        Debug.Print "uParse(): bad data"
        uParse = INVALID_REQUEST
        Exit Function

    End If

'    Debug.Print "section length: " & UBound(s) & " line1: " & s(0)

    ' /* identify section type from the first line */

Dim x() As String

    If g_SafeLeftStr(s(0), 4) = "GNTP" Then
        ' /* information line first */
        uParse = uParseInfoLine(s(0))

    ElseIf g_SafeLeftStr(s(0), 12) = "Identifier: " Then
        ' /* resource identifier */
        x = Split(s(0), ": ")
        uSaveBinary SectionIndex + 1, x(1)

    Else
        ' /* other headers... */

    End If

End Function

Private Function uParseInfoLine(ByVal str As String) As E_GNTP_CODES
Dim s() As String

    s = Split(str, " ")
    If UBound(s) < 2 Then
        Debug.Print "uParseInfoLine(): not enough parms"
        uParseInfoLine = INVALID_REQUEST
        Exit Function           ' // not enough params

    End If

Dim v() As String

    v = Split(s(0), "/")
    If UBound(v) <> 1 Then
        Debug.Print "uParseInfoLine(): not GNTP"
        uParseInfoLine = UNKNOWN_PROTOCOL
        Exit Function           ' // not GNTP
    
    End If

    If v(0) <> "GNTP" Then
        Debug.Print "uParseInfoLine(): not GNTP"
        uParseInfoLine = UNKNOWN_PROTOCOL
        Exit Function           ' // not GNTP
    
    End If

    If v(1) <> "1.0" Then
        Debug.Print "uParseInfoLine(): not 1.0"
        uParseInfoLine = UNKNOWN_PROTOCOL_VERSION
        Exit Function

    End If

    mDirective = ""

    Select Case s(1)
    Case "REGISTER", "NOTIFY"
        mDirective = s(1)

    Case Else
        Debug.Print "uParseInfoLine(): unsupported directive"
        uParseInfoLine = INVALID_REQUEST
        Exit Function           ' // bad directive

    End Select

    Select Case s(2)
    Case "NONE"
    
    Case Else
        Debug.Print "uParseInfoLine(): unsupported encryption (" & s(2) & ")"
        uParseInfoLine = INVALID_REQUEST
        Exit Function           ' // unsupported encryption

    End Select

End Function

Private Sub uSaveBinary(ByVal SectionIndex As Long, ByVal Identifier As String)
Dim i As Integer

    On Error Resume Next

    Identifier = "gntp-res-" & Identifier

    Debug.Print "uSaveBinary(): writing binary to %TEMP%" & Identifier & ".png"

    i = FreeFile()
    Open g_GetTempPath() & Identifier & ".png" For Binary Access Write As #i
    Put #i, , mSection(SectionIndex)
    Close #i

End Sub

Private Function uDoNotification(ByRef ReplySocket As CSocket) As E_GNTP_CODES
Dim pp As BPackedData

    ' /* convert the mime-style section into packed data */

    Set pp = New BPackedData
    If Not pp.SetTo(mSection(0), vbCrLf, ": ") Then
        Debug.Print "uDoNotification: bad data"
        uDoNotification = INVALID_REQUEST
        Exit Function

    End If

'Dim i As Long
'Dim s1 As String
'Dim s2 As String
'
'    Form1.List1.AddItem "##########"
'
'    For i = 1 To pp.Count
'        pp.EntryAt i, s1, s2
'        Form1.List1.AddItem s1 & "->" & s2
'
'    Next i

    ' /* required items */

    If (Not pp.Exists("Application-Name")) Or (Not pp.Exists("Notification-Name")) Or (Not pp.Exists("Notification-Title")) Then
        Debug.Print "uDoNotification: missing required arg"
        uDoNotification = REQUIRED_HEADER_MISSING
        Exit Function

    End If

Dim pn As T_GNTP_NOTIFICATION

    With pn
        'Notification-Icon: <url> | <uniqueid>
        'Optional - The icon to display with the notification.
        .Icon = pp.ValueOf("Notification-Icon")
        'Notification-Coalescing-ID: <string>
        'Optional - If present, should contain the value of the Notification-ID header of a previously-sent notification. This serves
        'as a hint to the notification system that this notification should replace/update the matching previous notification. The
        'notification system may ignore this hint.
        .CoalesceID = pp.ValueOf("Notification-Coalescing-ID")
        'Notification-Callback-Context: <string>
        'Optional - Any data (will be passed back in the callback unmodified)
        .CallbackContext = pp.ValueOf("Notification-Callback-Context")
        'Notification-Callback-Context-Type: <string>
        'Optional, but Required if 'Notification-Callback-Context' is passed - The type of data being passed in
        'Notification-Callback-Context (will be passed back in the callback unmodified). This does not need to be of any pre-defined
        'type, it is only a convenience to the sending application.
        .CallbackContextType = pp.ValueOf("Notification-Callback-Context-Type")

    End With

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
        .Add "title", pp.ValueOf("Notification-Title")
        
        'Notification-Text: <string>
        'Optional - The notification's text. (defaults to "")
        .Add "text", pp.ValueOf("Notification-Text")

        'Notification-Sticky: <boolean>
        'Optional - Indicates if the notification should remain displayed until dismissed by the user. (default to False)
        ' /* sticky == zero duration under Snarl */
        If uBool(pp.ValueOf("Notification-Sticky")) Then _
            .Add "timeout", "0"

        'Notification-Priority: <int>
        'Optional - A higher number indicates a higher priority. This is a display hint for the receiver which may be ignored. (valid
        'values are between -2 and 2, defaults to 0)
        ' /* translate priority into a Snarl-friendly one */
        Select Case g_SafeLong(pp.ValueOf("Notification-Priority"))
        Case 1, 2
            .Add "priority", "1"

        Case -1, -2
            .Add "priority", "-1"

        End Select

        'Notification-Callback-Target: <string>
        'Optional - An alternate target for callbacks from this notification. If passed, the standard behavior of performing the
        'callback over the original socket will be ignored and the callback data will be passed to this target instead. See the 'Url
        'Callbacks' section for more information.
        If pp.ValueOf("Notification-Callback-Target") <> "" Then _
            .Add "callback", pp.ValueOf("Notification-Callback-Target")

        'Notification-ID: <string>
        'Optional - A unique ID for the notification. If used, this should be unique for every request, even if the notification is
        'replacing a current notification (see Notification-Coalescing-ID)
        If pp.ValueOf("Notification-ID") <> "" Then _
            .Add "uid", pp.ValueOf("Notification-ID")

    End With

    ' /* do the notification */

    If g_DoAction("notify", 0, px, True, ReplySocket) = 0 Then
        Debug.Print "uDoNotification(): <notify> failed: " & CStr(g_QuickLastError())
        Select Case g_QuickLastError()
        Case SNARL_ERROR_AUTH_FAILURE
            uDoNotification = NOT_AUTHORIZED

        Case SNARL_ERROR_NOT_REGISTERED
            uDoNotification = UNKNOWN_APPLICATION

        Case Else
            uDoNotification = INTERNAL_SERVER_ERROR

        End Select

    End If

End Function

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
    ID As String
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

Public Function gntp_Process(ByVal Request As String, ByRef ReplySocket As CSocket) As Long

    ' /* return a GNTP error code here */

    On Error GoTo er

    Debug.Print "length: " & Len(Request)

    ' /* split into sections */

    mSection = Split(Request, vbCrLf & vbCrLf)
    Debug.Print "sections: " & UBound(mSection)

    If UBound(mSection) < 1 Then
        Debug.Print "gntp_process(): invalid number of sections"
        gntp_Process = 700
        Exit Function

    End If

    ' /* parse section 1 - must be the info block */

    mDirective = ""

    If Not uParse(0) Then
        Debug.Print "gntp_process(): invalid info block"
        gntp_Process = 700
        Exit Function

    End If

    Debug.Print "directive: " & mDirective

    Select Case mDirective
    Case "REGISTER"
        gntp_Process = uGetRegistrationInfo()

    Case "NOTIFY"
        gntp_Process = uGetNotificationInfo(ReplySocket)

    Case Else
        Debug.Print "gntp_process(): unsupported directive '" & mDirective & "'"
        gntp_Process = 303

    End Select

    Exit Function

er:
    Debug.Print err.Description

End Function

Private Function uBool(ByVal str As String) As Boolean

    Select Case LCase$(str)
    Case "yes", "true"
        uBool = True

    End Select

End Function

Private Function uGetRegistrationInfo() As Long
Dim pp As BPackedData

    Set pp = New BPackedData
    If Not pp.SetTo(mSection(0), vbCrLf, ": ") Then
        Debug.Print "bad data"
        uGetRegistrationInfo = 303
        Exit Function

    End If

    ' /* required items */

    If Not pp.Exists("Application-Name") Then
        Debug.Print "missing app name"
        uGetRegistrationInfo = 303
        Exit Function

    End If

    If Not pp.Exists("Notifications-Count") Then
        Debug.Print "missing count"
        uGetRegistrationInfo = 303
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
        Debug.Print "app requested unregister"
        g_DoAction "unreg", 0, g_newBPackedData("app-sig::" & mRegistration.Signature)
        Exit Function

    End If


    If UBound(mSection) < mRegistration.Count Then
        Debug.Print "not enough sections"
        uGetRegistrationInfo = 303
        Exit Function

    End If

    ' /* sections 1 to pr.Count should be notification types */

Dim i As Long

    For i = 1 To mRegistration.Count
        If Not uAddNotificationType(mSection(i), i - 1) Then
            ' /* replycode will have been set */
            Debug.Print "bad notification type " & i
            uGetRegistrationInfo = 303
            Exit Function

        End If
    Next i


    ' /* sections pr.Count to end should be resource identifiers */

    Debug.Print "parsing resource identifiers " & CStr(mRegistration.Count) & " to " & CStr(UBound(mSection))

    For i = mRegistration.Count To UBound(mSection)
        uParse i

    Next i

    ' /* register here */

    mRegistration.Token = g_DoAction("register", 0, _
                                     g_newBPackedData("app-sig::" & mRegistration.Signature & "#?title::" & mRegistration.AppName))

    If mRegistration.Token = 0 Then
        Debug.Print "registration failed"
        uGetRegistrationInfo = 700
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
                                                              "#?icon::" & .NotificationType(i).Icon)) & " == " & snGetLastError()

        Next i

    End With

End Function

Private Function uAddNotificationType(ByVal str As String, ByVal Index As Long) As Boolean
Dim pp As BPackedData

    Set pp = New BPackedData
    If Not pp.SetTo(str, vbCrLf, ": ") Then _
        Exit Function

    'For each notification being registered:
    '
    'Notification-Name: <string>
    'Required - The name (type) of the notification being registered
    '
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

    If Not pp.Exists("Notification-Name") Then
        Debug.Print "invalid notification type section"
        Exit Function

    End If

Dim sx() As String

    With mRegistration.NotificationType(Index)
        .Name = pp.ValueOf("Notification-Name")
        .DisplayName = pp.ValueOf("Notification-Display-Name")
        .Enabled = uBool(pp.ValueOf("Notification-Enabled"))
        .Icon = pp.ValueOf("Notification-Icon")

        ' /* sort out the icon */

        If g_SafeLeftStr(.Icon, 19) = "x-growl-resource://" Then
            sx = Split(.Icon, "://")
            .Icon = g_MakePath(App.Path) & sx(1) & ".png"
        
        End If


        If .DisplayName = "" Then _
            .DisplayName = .Name

'        Form1.List1.AddItem "Type " & CStr(Index) & ": " & .Name & " " & .DisplayName & " Enabled=" & .Enabled & " icon=" & .Icon

    End With

    uAddNotificationType = True

End Function

Private Function uParse(ByVal SectionIndex As Long) As Boolean
Dim s() As String
Dim x() As String

    s = Split(mSection(SectionIndex), vbCrLf)
    If UBound(s) < 1 Then _
        Exit Function

'    Debug.Print "section length: " & UBound(s) & " line1: " & s(0)

    ' /* identify section type from the first line */

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

    uParse = True

End Function

Private Function uParseInfoLine(ByVal str As String) As Boolean
Dim s() As String

    s = Split(str, " ")
    If UBound(s) < 2 Then _
        Exit Function           ' // not enough params

    If s(0) <> "GNTP/1.0" Then _
        Exit Function           ' // not GNTP

    mDirective = ""

    Select Case s(1)
    Case "REGISTER", "NOTIFY"
        mDirective = s(1)

    Case Else
        Debug.Print "unsupported directive"
        Exit Function           ' // bad directive

    End Select

    Select Case s(2)
    Case "NONE"
    
    Case Else
        Debug.Print "unsupported encryption"
        Exit Function           ' // unsupported encryption

    End Select

    uParseInfoLine = True

End Function

Private Sub uSaveBinary(ByVal SectionIndex As Long, ByVal Identifier As String)
Dim i As Integer

    Debug.Print "writing binary to " & Identifier & ".png"

    i = FreeFile()
    Open g_MakePath(App.Path) & Identifier & ".png" For Binary Access Write As #i

    Put #i, , mSection(SectionIndex)

    Close #i

End Sub

Private Function uGetNotificationInfo(ByRef ReplySocket As CSocket) As Long
Dim pp As BPackedData

'GNTP/1.0 NOTIFY NONE
'application -Name: Test application
'Notification -Name: Notify 6
'Notification -Title: Sticky
'Notification-Text: And without icon
'Notification-Sticky: True
'Data-Internal-Notification-ID: 5
'uReply(): replying with 'GNTP/1.0 -OK NONE
'X-Timestamp: 01/07/2011 07:51:43 PM'

    Set pp = New BPackedData
    If Not pp.SetTo(mSection(0), vbCrLf, ": ") Then
        Debug.Print "bad data"
        uGetNotificationInfo = 303
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
        Debug.Print "missing required arg"
        uGetNotificationInfo = 303
        Exit Function

    End If

Dim pn As T_GNTP_NOTIFICATION

    With pn
        'Application-Name: <string>
        'Required - The name of the application that sending the notification (must match a previously registered application)
        'Notification-Name: <string>
        'Required - The name (type) of the notification (must match a previously registered notification name registered by the
        'application specified in Application-Name)
        'Notification-Title: <string>
        'Required - The notification's title
        'Notification-ID: <string>
        'Optional - A unique ID for the notification. If used, this should be unique for every request, even if the notification is
        'replacing a current notification (see Notification-Coalescing-ID)
        .ID = pp.ValueOf("Notification-ID")
        'Notification-Text: <string>
        'Optional - The notification's text. (defaults to "")
        'Notification-Sticky: <boolean>
        'Optional - Indicates if the notification should remain displayed until dismissed by the user. (default to False)
        'Notification-Priority: <int>
        'Optional - A higher number indicates a higher priority. This is a display hint for the receiver which may be ignored. (valid
        'values are between -2 and 2, defaults to 0)
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
        'Notification-Callback-Target: <string>
        'Optional - An alternate target for callbacks from this notification. If passed, the standard behavior of performing the
        'callback over the original socket will be ignored and the callback data will be passed to this target instead. See the 'Url
        'Callbacks' section for more information.

    End With

    ' /* build the Snarl packet */

Dim px As BPackedData

    Set px = New BPackedData
    With px
        .Add "app-sig", "application/x-gntp-" & Replace$(pp.ValueOf("Application-Name"), " ", "_")
        .Add "id", pp.ValueOf("Notification-Name")
        .Add "title", pp.ValueOf("Notification-Title")
        .Add "text", pp.ValueOf("Notification-Text")

        ' /* sticky == zero duration */

        If uBool(pp.ValueOf("Notification-Sticky")) Then _
            .Add "timeout", "0"

        ' /* translate priority into a Snarl-friendly one */

        Select Case g_SafeLong(pp.ValueOf("Notification-Priority"))
        Case 1, 2
            .Add "priority", "1"

        Case -1, -2
            .Add "priority", "-1"

        End Select

        If pp.ValueOf("Notification-Callback-Target") <> "" Then _
            .Add "callback", pp.ValueOf("Notification-Callback-Target")

    End With

    g_DoAction "notify", 0, px, True, ReplySocket

End Function






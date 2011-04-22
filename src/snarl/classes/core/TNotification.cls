VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TNotification"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'/*********************************************************************************************
'/
'/  File:           TNotification.cls
'/
'/  Description:    Represents a notification within the notification roster
'/
'/  © 2009 full phat products
'/
'/  This file may be used under the terms of the Simplified BSD Licence
'/
'*********************************************************************************************/

'// "http://www.google.com/ig/images/weather/chance_of_rain.gif"

'    Window As CSnarlWindow
'    Column As Integer           ' // used to help determine placement of other notifications
'    StyleName As String         ' // R2.31: expanded
'    SchemeName As String        ' // R2.31: expanded
'    Info As T_NOTIFICATION_INFO ' // R2.4 DR7: full originating information

Dim mWindow As CSnarlWindow
Dim mColumn As Integer          ' // used to help determine placement of other notifications
'Dim mStyleName As String        ' // R2.31: expanded
'Dim mSchemeName As String       ' // R2.31: expanded

Dim mInfo As T_NOTIFICATION_INFO
Dim mTempIcon As String
Dim mArgs As BPackedData

Public Enum E_NOTIFICATION_STATUS
    NS_ACTIVE = 1
    NS_WAITING = 2
    NS_MISSED = 3

End Enum

Dim mState As E_NOTIFICATION_STATUS
Dim mSeen As Boolean

Dim WithEvents theIconRequest As BWebRequest
Attribute theIconRequest.VB_VarHelpID = -1

Implements BTagItem

Public Function Init(ByVal URL As String, ByRef Args As BPackedData) As Boolean

    On Error Resume Next

    err.Clear
    Set theIconRequest = New BWebRequest
    If err.Number <> 0 Then
        g_Debug "TNotification.Init(): couldn't create BWebRequest", LEMON_LEVEL_CRITICAL
        Exit Function

    End If

    mTempIcon = g_GetSafeTempIconPath() & "." & g_GetExtension(URL)
    If mTempIcon = "" Then
        g_Debug "TNotification.Init(): couldn't create local temporary file", LEMON_LEVEL_CRITICAL
        Exit Function

    End If

    Set mArgs = Args

    g_Debug "TNotification.Init(): saving as '" & mTempIcon & "'"
    Init = theIconRequest.GetFile(URL, mTempIcon)

End Function

Friend Sub SetTo(ByVal State As E_NOTIFICATION_STATUS, ByRef Info As T_NOTIFICATION_INFO, Optional ByRef Window As CSnarlWindow, Optional ByVal Column As Integer, Optional ByRef Frame As BRect)   ', Optional ByVal StyleName As String, Optional ByVal SchemeName As String)

    LSet mInfo = Info
    uFixupContent

    Set mWindow = Window
    mColumn = Column
'    mStyleName = StyleName
'    mSchemeName = SchemeName

    mState = State

    If Not (mWindow Is Nothing) Then
        mWindow.Window.MoveTo Frame.Left, Frame.Top
        mWindow.Show

    End If

End Sub

Private Function BTagItem_Name() As String

    BTagItem_Name = CStr(mInfo.Token)

End Function

Private Function BTagItem_Value() As String

    BTagItem_Value = mInfo.ClassObj.App.Signature & mInfo.CustomUID

End Function

Private Sub theIconRequest_Completed(ByVal WasSuccessful As Boolean)
Dim sz As String

    Debug.Print "TNotification.Completed(): " & WasSuccessful

    If Not WasSuccessful Then
        ' /* didn't download for some reason so use the class default icon, if one exists */
        If Not (mInfo.ClassObj Is Nothing) Then
            sz = mInfo.ClassObj.DefaultIcon

        Else
            ' /* to-do: allow user to specify an icon? */

        End If

    Else
        sz = mTempIcon

    End If

    mInfo.IconPath = sz

    If Not (g_NotificationRoster Is Nothing) Then _
        g_NotificationRoster.Add mInfo, mArgs

End Sub

Public Function Window() As CSnarlWindow

    Set Window = mWindow

End Function

Public Function Column() As Integer

    Column = mColumn

End Function

Friend Function Info() As T_NOTIFICATION_INFO

    LSet Info = mInfo

End Function

Public Function StyleName() As String

    StyleName = mInfo.StyleName

End Function

Public Function StyleNameLCase() As String

    StyleNameLCase = LCase$(mInfo.StyleName)

End Function

'Private Sub uFixupContent(ByRef Info As T_NOTIFICATION_INFO)
'
'    If (Info.ClassObj Is Nothing) Then
'        Debug.Print "########### uFixupContent missing class object ###########"
'        Exit Sub
'
'    End If
'
'    With Info
'        .Title = Replace$(.Title, "\n", vbCrLf)
'        .Text = Replace$(.Text, "\n", vbCrLf)
'
'    End With
'
'Dim ppn As BPackedData
'
'    ' /* new */
'    Set ppn = New BPackedData
'
'    With ppn
'        .Add "id", Info.ClassObj.Name
'        .Add "title", Info.Title
'        .Add "text", Info.Text
'        .Add "timeout", CStr(Info.Timeout)
'        .Add "icon", Info.IconPath
'        .Add "priority", CStr(Info.Priority)
'        .Add "callback", Info.DefaultAck
'        .Add "value", Info.Value
'
''        If (Info.Flags And SNARL41_NOTIFICATION_ALLOWS_MERGE) Then _
'            .Add "merge", "1"
'
'    End With
'
'Dim ppd As BPackedData
'Dim szn As String
'Dim szv As String
'
'    ' /* add in all other content */
'
'    Set ppd = New BPackedData
'    With ppd
'        If .SetTo(Info.OriginalContent) Then
'            .Rewind
'            Do While .GetNextItem(szn, szv)
'                If Not ppn.Exists(LCase$(szn)) Then _
'                    ppn.Add szn, szv
'
'            Loop
'        End If
'    End With
'
'    Info.OriginalContent = ppn.AsString()
'
'End Sub

Public Sub Unset()

    Set mWindow = Nothing

End Sub

Public Sub Remove()

    If mState = NS_ACTIVE Then _
        mWindow.Hide

End Sub

Public Sub Quit()

    If Not (mWindow Is Nothing) Then _
        mWindow.Remove

End Sub

Public Function AppSigAndUID() As String

    AppSigAndUID = mInfo.ClassObj.App.Signature & mInfo.CustomUID

End Function

Public Function State() As E_NOTIFICATION_STATUS

    State = mState

End Function

Public Function GetExtent(ByVal Column As Integer, ByVal RightToLeft As Boolean, ByVal x As Long) As Long

    GetExtent = x

    ' /* in the correct column? */

    If mColumn <> Column Then _
        Exit Function

    If RightToLeft Then
        If mWindow.Frame.Left < x Then _
            GetExtent = mWindow.Frame.Left

    Else
        If mWindow.Frame.Right > x Then _
            GetExtent = mWindow.Frame.Right

    End If

End Function

Private Sub uFixupContent()

    g_KludgeNotificationInfo mInfo

'    ' /* translates current T_NOTIFICATION_INFO content into a BPackedString
'    '    and then stores that back into T_NOTIFICATION_INFO->OriginalContent
'    '    this is so V42 styles can access the entire request */
'
'    If (mInfo.ClassObj Is Nothing) Then
'        g_Debug "TNotification.uFixupContent(): missing ClassObj", LEMON_LEVEL_CRITICAL
'        Exit Sub
'
'    End If
'
'Dim ppd As BPackedData
'
'    Set ppd = New BPackedData
'
'    With mInfo
'        .Title = Replace$(.Title, "\n", vbCrLf)
'        .Text = Replace$(.Text, "\n", vbCrLf)
'
'        ppd.Add "id", .ClassObj.Name
'        ppd.Add "title", .Title
'        ppd.Add "text", .Text
'        ppd.Add "timeout", CStr(.Timeout)
'        ppd.Add "icon", .IconPath
'        ppd.Add "priority", CStr(.Priority)
'        ppd.Add "callback", .DefaultAck
'        ppd.Add "value", .Value
'
''        If (Info.Flags And SNARL41_NOTIFICATION_ALLOWS_MERGE) Then _
'            .Add "merge", "1"
'
'    End With
'
'Dim ppx As BPackedData
'Dim szn As String
'Dim szv As String
'
'    ' /* add in all other custom content */
'
''    If Info.OriginalContent <> "" Then _
'        MsgBox Info.OriginalContent
'
'    Set ppx = New BPackedData
'    With ppx
'        If .SetTo(Info.OriginalContent) Then
'            .Rewind
'            Do While .GetNextItem(szn, szv)
'                If Not ppd.Exists(LCase$(szn)) Then _
'                    ppd.Add szn, szv
'
'            Loop
'        End If
'    End With
'
'    mInfo.OriginalContent = ppd.AsString()

End Sub

Public Function SetTitle(ByVal Title As String) As Boolean

    mInfo.Title = Title
    uFixupContent
    SetTitle = True

    Select Case mState
    Case NS_ACTIVE
        mWindow.Update mInfo

    Case NS_MISSED

    Case NS_WAITING

    End Select

End Function

Public Function SetText(ByVal Text As String) As Boolean

    mInfo.Text = Text
    uFixupContent
    SetText = True

    Select Case mState
    Case NS_ACTIVE
        mWindow.Update mInfo

    Case NS_MISSED

    Case NS_WAITING

    End Select

End Function

Public Function SetIcon(ByVal Icon As String) As Boolean

    mInfo.IconPath = Icon
    uFixupContent
    SetIcon = True

    Select Case mState
    Case NS_ACTIVE
        mWindow.Update mInfo

    Case NS_MISSED

    Case NS_WAITING

    End Select

End Function

Public Function SetTimeout(ByVal Timeout As Long) As Boolean

    Select Case mState
    Case NS_ACTIVE
        mWindow.SetTimeout Timeout
        SetTimeout = True

    Case NS_MISSED

    Case NS_WAITING

    End Select

End Function

Public Function SetCallback(ByVal Callback As String) As Boolean

    Select Case mState
    Case NS_ACTIVE
        mWindow.SetAck Callback
        SetCallback = True

    Case NS_MISSED

    Case NS_WAITING
        mInfo.DefaultAck = Callback

    End Select

End Function

Public Function SetMenu(ByVal Content As String) As Boolean

    Select Case mState
    Case NS_ACTIVE
        mWindow.SetMenu Content
        SetMenu = True

    Case NS_MISSED

    Case NS_WAITING

    End Select

End Function

Public Function ClearActions() As Boolean

    ClearActions = True
    If mInfo.Actions.CountItems = 0 Then _
        Exit Function

    mInfo.Actions.MakeEmpty

    Select Case mState
    Case NS_ACTIVE
        mWindow.RethinkActions

    Case NS_MISSED
        ' /* need to update the missed notifications panel, if it's open */

    Case NS_WAITING

    End Select

End Function

Public Function AddAction(ByVal Label As String, ByVal Command As String) As Boolean

    With mInfo.Actions
        If .IndexOf(Label) Then
            .Update Label, Command

        Else
            .Add new_BTagItem(Label, Command)

        End If

    End With

    Select Case mState
    Case NS_ACTIVE
        mWindow.RethinkActions

    Case NS_MISSED
        ' /* need to update the missed notifications panel, if it's open */
        g_NotificationRoster.UpdateMissedPanel

    Case NS_WAITING

    End Select

    AddAction = True

End Function

Public Function UpdateOrMerge(ByRef Args As BPackedData, ByVal ShouldMerge As Boolean) As Boolean

    uUpdateContent Args, ShouldMerge

    Select Case mState
    Case NS_ACTIVE
        uFixupContent
        UpdateOrMerge = mWindow.Update(mInfo)


    Case NS_MISSED
        ' /* refresh the missed panel if it's open */
        g_NotificationRoster.UpdateMissedPanel
        UpdateOrMerge = True


    Case NS_WAITING
        UpdateOrMerge = True

    End Select

End Function

Private Sub uUpdateContent(ByRef Args As BPackedData, Optional ByVal ShouldMerge As Boolean = False)
Dim sz As String

    ' /* update or merge provided information with existing
    '    content - ShouldMerge controls which */

    With mInfo

        ' /* only change something if the entry _exists_ in Args */

        If Args.Exists("title") Then
            sz = uUnescape(Args.ValueOf("title"), .Title)
            .Title = g_TrimLines(sz)

        End If

        sz = ""
        If Args.Exists("text") Then
            If ShouldMerge Then _
                sz = mInfo.Text & vbCrLf

            sz = sz & uUnescape(Args.ValueOf("text"), .Text)
            .Text = g_TrimLines(sz)

        End If

        If Args.Exists("icon") Then
            .IconPath = Args.ValueOf("icon")

        ElseIf Args.Exists("icon-base64") Then
            ' /* R2.4b4 */
            .IconPath = g_GetBase64Icon(Args.ValueOf("icon-base64"))

        End If

        If Args.Exists("timeout") Then _
            .Timeout = g_SafeLong(Args.ValueOf("timeout"))          ' // g_SafeLong() returns 0 on error - not good...

        ' /* following added in R2.4 Beta 4 */

        If Args.Exists("callback") Then _
            .DefaultAck = Args.ValueOf("callback")

        If Args.Exists("reply-to") Then _
            .hWndReply = g_SafeLong(Args.ValueOf("reply-to"))

        If Args.Exists("reply") Then _
            .uReplyMsg = g_SafeLong(Args.ValueOf("reply"))

        ' /* clearly this can only be set if referencing the notification by its token */

        If Args.Exists("uid") Then _
            .CustomUID = Args.ValueOf("uid")

        ' /* still TBC: Priority, Actions, Value */

        .LastUpdated = Now()


        .OriginalContent = Args.AsString()

    End With

End Sub

Public Function CanMerge(ByRef Class As TAlert, ByVal Title As String) As Boolean

    If Not (mWindow Is Nothing) Then _
        CanMerge = mWindow.IsMergeCandidate(Class, Title)

End Function

Public Sub MarkAsSeen()

    mSeen = True

End Sub

Public Function WasSeen() As Boolean

    WasSeen = mSeen

End Function

Private Function uUnescape(ByVal Text As String, ByVal Replacement As String) As String
Dim sz As String

    sz = Text
    sz = Replace$(sz, "\+", Replacement)
    sz = Replace$(sz, "\user", g_GetUserName())
    sz = Replace$(sz, "\host", g_GetComputerName())
    sz = Replace$(sz, "\info", "Snarl " & App.Comments & " (V" & CStr(App.Major) & "." & CStr(App.Revision) & ")")
    sz = Replace$(sz, "\n", vbCrLf)

    uUnescape = sz

End Function

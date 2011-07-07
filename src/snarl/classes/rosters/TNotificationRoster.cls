VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TNotificationRoster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

    '/*********************************************************************************************
    '/
    '/  File:           TNotificationRoster.cls
    '/
    '/  Description:    Manages on-screen, missed and delayed notifications
    '/
    '/  © 2009 full phat products
    '/
    '/  This file may be used under the terms of the Simplified BSD Licence
    '/
    '*********************************************************************************************/

Const OBJ_DATE = "7-Apr-2011"

Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Dim mToken As Long
Dim mRefs As Long
Dim mMissedPanel As TMissedNotificationsPanel
Dim mActualMissed As Long                       ' // number that appeared since DnD enabled

    ' /* lists - a notification should only ever be in one of them */
Dim mActive As BTagList
Dim mWait As BTagList
Dim mMissedX As BTagList

Dim mCurrentLowPriority As TNotification

Enum E_STYLE_FORWARD_FLAGS
    SFF_WHEN_BUSY = 1
    SFF_WHEN_AWAY = 2
    SFF_WHEN_NOT_AWAY_OR_BUSY = 4

End Enum

Implements MLibrary
Implements MVersionInfo

Private Sub Class_Initialize()

    Set mActive = new_BTagList()
    Set mWait = new_BTagList()
    Set mMissedX = new_BTagList()

End Sub

Private Function MLibrary_Close() As melon.M_RESULT

    g_Debug "TNotificationRoster.Close()", LEMON_LEVEL_PROC

    mRefs = mRefs - 1
    If mRefs > 0 Then
        ' /* still open */
        Exit Function

    End If

    ' /* run tidyup code */

Dim i As Long

    With mActive
        If .CountItems Then
            For i = .CountItems To 1 Step -1
                uAsNotification(.TagAt(i)).Quit

            Next i

            .MakeEmpty

        End If

    End With

    ' /* remove temporary icons */

Dim sz As String

    sz = String$(MAX_PATH + 1, 0)
    GetTempPath MAX_PATH, sz
    sz = g_TrimStr(sz)
    If sz <> "" Then
        sz = g_MakePath(sz)
        i = 1
        Do While g_Exists(sz & "snarl-icon" & CStr(i))
            delete_file sz & "snarl-icon" & CStr(i)
            i = i + 1

        Loop

    End If

    g_Debug "TNotificationRoster.Close(): done"

End Function

Private Sub MLibrary_Initialize()
    ' /* not currently used */
End Sub

Private Function MLibrary_Magic() As Long
    MLibrary_Magic = &H23232323
End Function

Private Function MLibrary_Open() As melon.M_RESULT

    mRefs = mRefs + 1
    If mRefs = 1 Then
        ' /* run startup code */
        mToken = (&HC0 + (Rnd * 80)) And &HFFFFFFFC

    End If

End Function

Private Function MLibrary_OpenCount() As Long
    MLibrary_OpenCount = mRefs
End Function

Private Sub MLibrary_Uninitialize()
    ' /* not currently used */
End Sub

Private Property Get MVersionInfo_Date() As String
    MVersionInfo_Date = OBJ_DATE
End Property

Private Property Get MVersionInfo_Name() As String
    MVersionInfo_Name = "notification.roster"
End Property

Private Property Get MVersionInfo_Revision() As Long
    MVersionInfo_Revision = App.Revision
End Property

Private Property Get MVersionInfo_Version() As Long
    MVersionInfo_Version = App.Major
End Property

Friend Function Add(ByRef Info As T_NOTIFICATION_INFO, ByRef Args As BPackedData) As Long

    ' /* adds a notification to the roster: should only be called by:
    '       TAlert.ShowNotification()
    '       g_DoSchemePreview()
    '       TNotification (when download complete)
    '       Missed Notifications panel (when "Display" button invoked by user)
    ' */

    gSetLastError SNARL_ERROR_FAILED
    If g_IsQuitting Then _
        Exit Function               ' // shutting down - can't do anything

    ' /* log only? */

    If g_ConfigGet("log_only") = "1" Then
        g_WriteToLog Info.Title, Info.Text
        Add = -1                    ' // V40.7 - changed from M_OK to -1 (indicates success but no token available)
        gSetLastError SNARL_SUCCESS
        Exit Function               ' // TO-DO: determine if we should return 0 and set last_error instead

    End If

    ' /* for anything else, the class must be valid */

    If (Info.ClassObj Is Nothing) Then
        g_Debug "TNotificationRoster.Add(): Class is null for '" & Info.Title & "/" & Info.Text & "'", LEMON_LEVEL_CRITICAL
        Exit Function

    End If

    ' /* R2.31: log last time a notification was generated */

    gLastNotification = Now()

    ' /* R2.31: 'duplicates' is now a simple 'stop spam' toggle - what we do with the potential
    '    spam is for future releases; for now we just discard it */

Dim lQuantum As Long

    lQuantum = g_SafeLong(g_ConfigGet("flood_quantum"))
    If lQuantum < 100 Then _
        lQuantum = 750

    If (Info.ClassObj.GetValue("block_spam") = "1") And ((GetTickCount() - Info.ClassObj.LastNotificationGeneratedTick) < lQuantum) Then
        Debug.Print "!!! SPAM !!! " & CStr(lQuantum)

        ' /* R2.31 - for now, block and return a failure code.  For the future, it would be lovely to
        '    allow more functionality (e.g. update notification, append to notification, send to missed log, etc. */

        gSetLastError SNARL_ERROR_FLOODING
        Exit Function

    End If

    ' /* R2.4.2: erase callback if url detection enabled */

    If g_ConfigGet("no_callback_urls") = "1" Then
        If g_IsURL(Info.DefaultAck) Then _
            Info.DefaultAck = ""

    End If

    ' /* only record the tick of the last _generated_ notification */

    Info.ClassObj.LastNotificationGeneratedTick = GetTickCount()

    gSetLastError SNARL_SUCCESS

    ' /* R2.3: set datestamp */

    Info.DateStamp = Now()
    
    ' /* R2.4: overflow trimming */

    Info.Title = g_TrimLines(Info.Title)
    Info.Text = g_TrimLines(Info.Text)

    ' /* R2.4: generate the "OriginalContent" content */

    ' /* ===============================================
    '   Note: this is temporary until TNotification
    '   can take care of window creation (see notes below)
    '    =============================================== */

    g_KludgeNotificationInfo Info

    ' /* R2.4 DR7 - can include actions in request */

    If (Info.Actions Is Nothing) Then _
        Set Info.Actions = new_BTagList()                   ' // create list even if there are no actions defined

Dim sn As String
Dim sv As String

    If Not (Args Is Nothing) Then
        With Args
            .Rewind
            Do While .GetNextItem(sn, sv)
                If sn = "action" Then _
                    uGetAction Info.Actions, sv

            Loop

        End With
    End If

    ' /* 40.7 - icon can be an http:// or https:// url - if it is we initiate an
    '    async download of the icon and put the notification onto a stack while
    '    we wait for the download to complete */

Dim pn As TNotification

    If (LCase$(g_SafeLeftStr(Info.IconPath, 7)) = "http://") Or (LCase$(g_SafeLeftStr(Info.IconPath, 8)) = "https://") Then
        g_Debug "TNotificationRoster.Add(): remote icon specified, starting async download..."

        ' /* to-do: allow user to suppress this by specifying a maximum download wait
        '    of zero seconds - in which case use the class default icon */

        Set pn = New TNotification
        If pn.Init(Info.IconPath, Args) Then

            ' /* allocate a token now */

            Info.Token = mToken
            pn.SetTo NS_WAITING, Info

            ' /* add to the queue */

            mWait.Add pn
            Add = mToken                                '// return the token that was allocated
            mToken = mToken + 1
            Exit Function

        Else
            g_Debug "TNotificationRoster.Add(): remote icon specified but couldn't start icon download", LEMON_LEVEL_WARNING

        End If

    End If


    ' /* R2.3: send to any selected non-window styles */

Dim pCopyInfo As T_NOTIFICATION_INFO
Dim pt As BTagItem
Dim fc As Long              ' // needs to be an enum before we forget what the flags mean!

    If Not (Info.ClassObj.ForwardList Is Nothing) Then
        With Info.ClassObj.ForwardList
            .Rewind
            Do While .GetNextTag(pt) = B_OK
                fc = Val(pt.Value)

                Debug.Print "TNotificationRoster.Add(): is_dnd=" & g_IsDND() & " is_away=" & g_IsAway() & "cond=" & g_HexStr(fc)

                If (g_IsDND()) And ((fc And SFF_WHEN_BUSY) <> 0) Then
                    ' /* shadow only when busy */
                    fc = -1

                ElseIf (g_IsAway()) And ((fc And SFF_WHEN_AWAY) <> 0) Then
                    ' /* shadow only when away */
                    fc = -1

                ElseIf (fc And SFF_WHEN_NOT_AWAY_OR_BUSY) <> 0 Then
                    ' /* shadow when present and not busy */
                    fc = -1

                Else
                    ' /* no shadowing */
                    fc = 0

                End If

                If fc <> 0 Then
                    g_Debug "TNotificationRoster.Add(): shadowing to '" & pt.Name & "'..."
                    LSet pCopyInfo = Info
                    pCopyInfo.StyleName = style_GetStyleName(pt.Name)
                    pCopyInfo.SchemeName = style_GetSchemeName(pt.Name)
                    uCopyTo pCopyInfo ', Info.ClassObj

                End If
            Loop
        End With
    End If



    ' /* R2.4.2 DR3 - send to subscribers */

    g_SubsRoster.Broadcast Info


    If (Info.Priority < 0) And (Info.Timeout = 0) Then
        ' /* R2.4: low-priority notifications cannot be sticky */
        Info.Timeout = Val(g_ConfigGet("default_duration"))

    ElseIf Info.Timeout = -1 Then
        ' /* V41: if timeout is -1 then use global default timeout */
        Info.Timeout = Val(g_ConfigGet("default_duration"))

    End If


    ' /*
    '   R2.4 DR8: presence management, as follows:
    '
    '   Make Sticky takes priority over Do Not Disturb Mode, which in turn takes priority over Away Mode
    '
    '   End result is:
    '                       Low         Normal      High
    '                       ----------------------------------
    '   Sticky              Sticky      Sticky      Sticky
    '   Busy (DND=true)     Discarded   <user>      Displayed
    '   Away                Discarded   <user>      Missed
    '
    '   Items marked <user> are user configurable.
    '
    ' */

'    Debug.Print "++ current presence: " & g_HexStr(g_GetPresence())

    ' /* sticky snarls take precendence, then DND mode setting, then Away mode setting */

Dim nPMode As E_PRESENCE_ACTIONS

    If g_ConfigGet("sticky_snarls") = "1" Then
        g_Debug "TNotificationRoster.Add(): Sticky mode enabled"
        ' /* R2.4 DR8: *all* notifications are sticky */
        Info.Timeout = 0

    ElseIf g_IsDND() Then
        ' /* Do Not Disturb (Busy) mode */
        g_Debug "TNotificationRoster.Add(): Busy (aka DND) mode enabled: priority=" & CStr(Info.Priority)
        If Info.Priority < 0 Then
            ' /* low priority = ignored */
            gSetLastError SNARL_ERROR_DO_NOT_DISTURB
            Exit Function

        ElseIf Info.Priority = 0 Then
            ' /* determine what to do with normal priority notifications */
            nPMode = Val(Info.ClassObj.GetValue("override-busy")) - 1
            If nPMode = PA_DO_DEFAULT Then _
                nPMode = Val(g_ConfigGet("busy_mode"))

        End If

    ElseIf g_IsAway() Then
        ' /* Away mode */
        g_Debug "TNotificationRoster.Add(): Away mode enabled: priority=" & CStr(Info.Priority)
        If Info.Priority < 0 Then
            ' /* low priority = ignored */
            gSetLastError SNARL_ERROR_CLASS_BLOCKED
            Exit Function

        ElseIf Info.Priority = 0 Then
            ' /* determine what to do with normal priority notifications */
            nPMode = Val(Info.ClassObj.GetValue("override-away")) - 1
            Debug.Print "### away class setting is: " & nPMode
            If nPMode = PA_DO_DEFAULT Then _
                nPMode = Val(g_ConfigGet("away_mode"))

            Debug.Print "### away setting is: " & nPMode

        Else
            ' /* high priority = logged as missed irrespective of class setting */
            uAddToMissedList Info
            Add = Info.Token

        End If

    Else
        nPMode = Val(Info.ClassObj.GetValue("override-normal")) - 1

    End If

    g_Debug "TNotificationRoster.Add(): Presence action=" & CStr(nPMode)

    If nPMode <> PA_DO_DEFAULT Then
        ' /* we have something specific to do with the notification */

        Select Case nPMode

        Case PA_LOG_AS_MISSED
            ' /* put in missed list (unless class setting says otherwise) and return the token */
'            If Info.ClassObj.GetValue("log_during_dnd") = "1" Then
                uAddToMissedList Info
                Add = Info.Token
                Exit Function

'            End If

        Case PA_MAKE_STICKY
            Info.Timeout = 0

        Case PA_DO_NOTHING
            ' /* bin it */
            gSetLastError SNARL_ERROR_CLASS_BLOCKED
            Exit Function

        Case PA_DISPLAY_URGENT
            Info.Priority = 1

        Case PA_DISPLAY_NORMAL
            ' /* allow it through */

        Case PA_FORWARD
            Info.StyleName = style_GetStyleName(g_ConfigGet(IIf(g_IsDND, "busy_style", "away_style")))
            Info.SchemeName = style_GetSchemeName(g_ConfigGet(IIf(g_IsDND, "busy_style", "away_style")))

        End Select

    End If


    ' /* =============================================================================
    '
    '      if we're here we need to create an on-screen instance of the notification
    '
    '    ============================================================================= */


    ' /* R2.4: merging */

    If (Info.IntFlags And NF_MERGE) Then
        ' /* is there an existing notification we can merge with? */
        If uFindMergeCandidate(Info.ClassObj, Info.Title, pn) Then
            ' /* update content */
            pn.UpdateOrMerge Args, True
            gSetLastError SNARL_WAS_MERGED                      ' // set LastError to indicate the merge
            Add = pn.Info.Token                                 ' // return existing token
            Exit Function

        End If

    End If

    ' /* no style name?  that's okay, use the default */

    If Info.StyleName = "" Then
        Info.StyleName = style_GetStyleName(g_ConfigGet("default_style"))
        Info.SchemeName = style_GetSchemeName(g_ConfigGet("default_style"))

    End If

Dim hr As M_RESULT
Dim ar As BRect
Dim p As Long

    ' /* get the appropriate instance for the notification before we do anything else
    '    we must pass a valid style instance to CSnarlWindow, even if the style
    '    roster is invalid */

Dim psi As IStyleInstance
Dim pStyle As TStyle

    ' /* get the style from the roster here - this is guaranteed to return a usable TStyle */

    uSafeGetStyle Info.StyleName & "/" & Info.SchemeName, pStyle, psi

    ' /* get the path to the style - we need this for style-provided icons and sounds */

'    Info.IconPath = uTranslateIconPath(Info.IconPath, pStyle.Path)

    ' /* if S_STYLE_IS_WINDOWLESS is set, just pass it to the handler */
    ' /* TO-DO: do we still need to record it? what about our return value? */

    If (pStyle.Flags And S_STYLE_IS_WINDOWLESS) Then
        g_Debug "TNotificationRoster.Add(): style is handling notification display"
'        psi.UpdateContent pStyle.CreateNotificationInfo(Info.Title, Info.Text, Info.IconPath, Info.SchemeName, IIf(Info.Priority > 0, S_NOTIFICATION_IS_PRIORITY, 0))

        ' /* uCopyTo() will return TRUE if the style asked for the notification
        '    to be passed back for on-screen display (pass-through) */

        If uCopyTo(Info) Then
            ' /* style is asking for the notification to be displayed on-screen */
            Info.StyleName = style_GetStyleName(g_ConfigGet("default_style"))
            Info.SchemeName = style_GetSchemeName(g_ConfigGet("default_style"))
            Add = Me.Add(Info, Args)

        Else
            ' /* no need to pass through */

            Add = -1                ' // TEMPORARY while we figure out what best to return here

        End If

        Exit Function

    End If

    ' /* is it a single instance style?  if so, we need to check to see if we already have a
    '    a notification of this style on-screen */

    If (pStyle.Flags And S_STYLE_SINGLE_INSTANCE) Then
        If uFindbyStyleName(Info.StyleName, pn) Then
            ' /* notification found, so decide what to do next... */
            If (pStyle.Flags And S_STYLE_DONT_QUEUE) Then
                ' /* don't queue notifications - so remove the existing one and create new */
                pn.Remove

            Else
                ' /* TO DO: add to queue */
                Exit Function

            End If
        End If
    End If

    ' /* if the token hasn't been pre-allocated set it here */

    If Info.Token = 0 Then
        Info.Token = mToken
        mToken = mToken + 1

    End If

'Dim dwStyleFlags As S_STYLE_FLAGS
'
'    dwStyleFlags = S_STYLE_MULTI_INSTANCE
'
'    If Not (pStyle Is Nothing) Then _
'        dwStyleFlags = pStyle.Flags

'    Info.StyleToUse = style_GetSchemeName(Info.StyleToUse)

    ' /* create a new instance of the requested style - if this fails for any reason,
    '    we simply create an instance of the internal style */




    ' /* R2.4 DR8: can use a style which allows blocking the notification - if the style
    '    allows the notification, the default style will be used to display it */

'Dim pn As notification_info
'
'    If (pStyle.Flags And S_STYLE_WANTS_VETO) Then
'
'        psi.UpdateContent pn
'        If pn.Flags = -1 Then
'            Add = 0
'            Exit Function
'
'        Else
'            ' /* re-create the style
'            uSafeGetStyle g_ConfigGet("default_style"), pStyle, psi
'
'        End If
'
'    End If

    ' /* if S_STYLE_WANT_APP_NAME is set, prefix title with the sending app name */

    If (pStyle.Flags And S_STYLE_WANT_APP_NAME) Then
        g_Debug "TNotificationRoster.Add(): style wants application name"
        Info.Title = Info.ClassObj.App.Name & "#?" & Info.Title

    End If

    ' /* only one low priority notification can be on screen at once... */

    If (Info.Priority < 0) And (Not (mCurrentLowPriority Is Nothing)) Then
        With mCurrentLowPriority
            Me.Hide .Info.Token, .Info.CustomUID, .Info.ClassObj.App.Signature, .Info.ClassObj.App.Password

        End With

        Set mCurrentLowPriority = Nothing

    End If



    ' /* create the on-screen instance */


    ' /* =======================================================
    '
    '   For future development: it would be better if the
    '   TNotification object actually created the
    '   CSnarlWindow itself, or at least did the V42
    '   kludging for us
    '
    '    ======================================================= */

Dim psw As CSnarlWindow

'    MsgBox Info.OriginalContent

    Set psw = New CSnarlWindow
    hr = psw.Create(Info, psi, pStyle.Flags, pStyle.Path)

    If hr <> M_OK Then
        ' /* TO-DO: Create() should probably return a more useful set of errors rather than just M_OK */
        g_Debug "TNotificationRoster.Add(): couldn't create window", LEMON_LEVEL_CRITICAL
        Add = 0
        Exit Function

    End If

Dim i As Long

    ' /* figure out its location on screen */

    If Not uGetNotificationPosition(psw, Info.Position, ar, i) Then
        g_Debug "TNotificationRoster.Add(): couldn't display notification", LEMON_LEVEL_CRITICAL
        gSetLastError SNARL_ERROR_COULD_NOT_DISPLAY
        
        ' /* V42: couldn't find suitable position - mostly likely no room left on-screen */

        If Info.Priority >= 0 Then
            ' /* normal and high priority notifications are added to missed list */
            uAddToMissedList Info
            Add = Info.Token

        Else
            Add = 0

        End If

        ' /* zap the notification window */

        psw.Remove
        Set psw = Nothing
        Exit Function

    End If

    g_Debug "TNotificationRoster.Add(): placing at " & ar.Left & "," & ar.Top & "-" & ar.Right & "," & ar.Bottom


    ' /* add it to our list */

    Set pn = New TNotification
    pn.SetTo NS_ACTIVE, Info, psw, i, ar
    mActive.Add pn

    ' /* return value */

    Add = Info.Token

    ' /* R2.4: if this was a low priority notification, remember it */

    If Info.Priority < 0 Then _
        Set mCurrentLowPriority = pn


    ' /* R2.4 DR8: last thing to do - play the sound */

Dim szSound As String

    If g_ConfigGet("enable_sounds") = "1" Then
        szSound = trim(Info.SndFile)
        If g_ConfigGet("use_style_sounds") = "1" Then
            If (szSound = "") Or ((szSound <> "") And (g_ConfigGet("prefer_style_sounds") = "1")) Then _
                szSound = ""        ' // get_style_sound()

        End If

        ' /* no specific or style-provided sound? */

        If szSound = "" Then _
            szSound = IIf(Info.Priority > 0, g_ConfigGet("default_priority_sound"), g_ConfigGet("default_notification_sound"))

        ' /* play what we have - if we have something */

        g_QuickPlaySound szSound

    End If

End Function

Private Sub uGetAction(ByRef ListToAddTo As BTagList, ByVal PackedAction As String)
Dim i As Long

    ' /* only interested in _first_ comma */

    i = InStr(PackedAction, ",")
    If i = 0 Then _
        Exit Sub

Dim szLbl As String
Dim szCmd As String

    szLbl = g_SafeLeftStr(PackedAction, i - 1)
    szCmd = g_SafeRightStr(PackedAction, Len(PackedAction) - i)

    ' /* must have a label and command */

    If (szLbl = "") Or (szCmd = "") Then _
        Exit Sub

    ListToAddTo.Add new_BTagItem(szLbl, szCmd)

End Sub

Public Function Remove(ByVal Id As Long) As Boolean

    Remove = uRemove(mActive.IndexOf(Id))

End Function

Private Function uRemove(ByVal Index As Long) As Boolean

    If (Index < 1) Or (Index > mActive.CountItems) Then
        g_Debug "TNotificationRoster.uRemove(): bad index #" & CStr(Index), LEMON_LEVEL_CRITICAL
        Exit Function

    End If

    uAsNotification(mActive.TagAt(Index)).Unset
    mActive.Remove Index
    uRemove = True

End Function

Public Function Find(ByVal Token As Long, ByVal UID As String, ByVal AppSig As String, ByVal Password As String, ByRef Notification As TNotification) As Boolean

    ' /* searches all notification lists for the specified notification (token takes priority)
    '    and returns True if it's found, and sets Notification to it */

Dim pn As TNotification
Static i As Long

    If Token <> 0 Then
        ' /* search on token */
        i = mActive.IndexOf(CStr(Token))
        If i Then
            Set pn = mActive.TagAt(i)
            If pn.Info.ClassObj.App.Password = Password Then
                ' /* password ok */
                Set Notification = pn
                gSetLastError SNARL_SUCCESS
                Find = True

            Else
                ' /* password failure */
                gSetLastError SNARL_ERROR_AUTH_FAILURE

            End If

        Else
            ' /* search wait list */
            i = mWait.IndexOf(CStr(Token))
            If i Then
                Set pn = mWait.TagAt(i)
                If pn.Info.ClassObj.App.Password = Password Then
                    ' /* password ok */
                    Set Notification = pn
                    gSetLastError SNARL_SUCCESS
                    Find = True

                Else
                    ' /* password failure */
                    gSetLastError SNARL_ERROR_AUTH_FAILURE

                End If

            Else
                ' /* search missed list */
                i = mMissedX.IndexOf(CStr(Token))
                If i Then
                    Set pn = mMissedX.TagAt(i)
                    If pn.Info.ClassObj.App.Password = Password Then
                        ' /* password ok */
                        Set Notification = pn
                        gSetLastError SNARL_SUCCESS
                        Find = True

                    Else
                        ' /* password failure */
                        gSetLastError SNARL_ERROR_AUTH_FAILURE

                    End If

                Else
                    gSetLastError SNARL_ERROR_NOTIFICATION_NOT_FOUND

                End If
            End If
        End If

    ElseIf UID <> "" Then
        ' /* search on UID/App-Sig pair */
        UID = AppSig & UID

        With mActive
            .Rewind
            Do While .GetNextTag(pn) = B_OK
                Debug.Print "looking for '" & UID & "' found '" & pn.AppSigAndUID & "'"
                If pn.AppSigAndUID = UID Then
                    If pn.Info.ClassObj.App.Password = Password Then
                        ' /* password ok */
                        Set Notification = pn
                        gSetLastError SNARL_SUCCESS
                        Find = True

                    Else
                        ' /* password failure */
                        gSetLastError SNARL_ERROR_AUTH_FAILURE

                    End If

                    Exit Function

                End If
            Loop

        End With

        With mWait
            .Rewind
            Do While .GetNextTag(pn) = B_OK
                If pn.AppSigAndUID = UID Then
                    If pn.Info.ClassObj.App.Password = Password Then
                        ' /* password ok */
                        Set Notification = pn
                        gSetLastError SNARL_SUCCESS
                        Find = True

                    Else
                        ' /* password failure */
                        gSetLastError SNARL_ERROR_AUTH_FAILURE

                    End If

                    Exit Function

                End If
            Loop

        End With

        ' /* search missed list */

        With mMissedX
            .Rewind
            Do While .GetNextTag(pn) = B_OK
                If pn.AppSigAndUID = UID Then
                    If pn.Info.ClassObj.App.Password = Password Then
                        ' /* password ok */
                        Set Notification = pn
                        gSetLastError SNARL_SUCCESS
                        Find = True

                    Else
                        ' /* password failure */
                        gSetLastError SNARL_ERROR_AUTH_FAILURE

                    End If

                    Exit Function

                End If
            Loop

        End With

        gSetLastError SNARL_ERROR_NOTIFICATION_NOT_FOUND

    End If

End Function

Public Function Hide(ByVal Token As Long, ByVal UID As String, ByVal AppSig As String, ByVal Password As String) As Boolean
Dim pn As TNotification

    ' /* uFind() will set LastError for us */

    If Find(Token, UID, AppSig, Password, pn) Then

        Select Case pn.State
        Case NS_ACTIVE
'            MsgBox "removing"
            pn.Remove
            Hide = True

        Case NS_MISSED
            ' /* remove from missed list */
            mMissedX.Remove mMissedX.IndexOf(CStr(pn.Info.Token))
            UpdateMissedPanel

        Case NS_WAITING
            ' /* remove from wait list */

        End Select

    End If

End Function

Public Function IsVisible(ByVal Token As Long, ByVal UID As String, ByVal AppSig As String, ByVal Password As String) As Boolean

    ' /* uFind sets lasterror for us */
    IsVisible = Find(Token, UID, AppSig, Password, Nothing)

End Function

Public Function WasMissed(ByVal Token As Long, ByVal UID As String, ByVal AppSig As String, ByVal Password As String) As Long
Dim pn As TNotification

    If Find(Token, UID, AppSig, Password, pn) Then _
        WasMissed = CLng(pn.State = NS_MISSED)

End Function

Private Function uIsFreeSpace(ByRef aRect As BRect) As Boolean
Dim pn As TNotification

    ' /* return True if passed rect doesn't overlap an existing notification */

    uIsFreeSpace = True

    With mActive
        .Rewind
        Do While .GetNextTag(pn) = B_OK
            If pn.Window.Frame.Intersects(aRect) Then
                uIsFreeSpace = False
                Exit Function

            End If

        Loop
    End With

''    g_Debug "uIsFreeSpace(): testing " & aRect.Left & "," & aRect.Top & "-" & aRect.Right & "," & aRect.Bottom
''                g_Debug "blocked by item " & CStr(i) & " at " & .Left & "," & .Top & ";" & .Right & "," & .Bottom

End Function

Private Function uFindExtent(ByVal Column As Integer, ByVal RightToLeft As Boolean) As Long
Dim pn As TNotification

    uFindExtent = IIf(RightToLeft, 8192, 0)

    With mActive
        .Rewind
        Do While .GetNextTag(pn) = B_OK
            uFindExtent = pn.GetExtent(Column, RightToLeft, uFindExtent)

        Loop

    End With

    g_Debug "TNotificationRoster.uFindExtent(): returning " & uFindExtent

End Function

Public Sub CloseMultiple(ByVal AppToken As Long)
Dim n As Integer
Dim i As Long

    If AppToken = 0 Then
        ' /* close them all */
        g_Debug "TNotificationRoster.CloseMultiple(): closing all..."

        If mActive.CountItems Then
            ' /* fade all notifications out together */
            For n = 255 To 0 Step -(Val(g_ConfigGet("step_size")) * 8)
                For i = 1 To mActive.CountItems
                    uAsNotification(mActive.TagAt(i)).Window.bSetAlpha n

                Next i

            Next n

            ' /* make sure they're at zero alpha */

            For i = 1 To mActive.CountItems
                uAsNotification(mActive.TagAt(i)).Window.bSetAlpha 0

            Next i
        
            ' /* now zap them */

            For i = mActive.CountItems To 1 Step -1
                uAsNotification(mActive.TagAt(i)).Quit

            Next i

            mActive.MakeEmpty

        End If

    Else
        ' /* only close snarls with same creating app */

        g_Debug "CloseMultiple(): closing all Snarls from " & g_HexStr(AppToken) & "..."

        If mActive.CountItems Then
            ' /* fade all related notifications out together */
            For n = 255 To 0 Step -(Val(g_ConfigGet("step_size")) * 8)
                For i = 1 To mActive.CountItems
                    With uAsNotification(mActive.TagAt(i))
                        If .Info.ClassObj.App.Token = AppToken Then _
                            .Window.bSetAlpha n

                    End With

                Next i
            Next n

            ' /* make sure they're at zero alpha */

            For i = 1 To mActive.CountItems
                With uAsNotification(mActive.TagAt(i))
                    If .Info.ClassObj.App.Token = AppToken Then _
                        .Window.bSetAlpha 0

                End With
            Next i
            
            For i = mActive.CountItems To 1 Step -1
                If uAsNotification(mActive.TagAt(i)).Info.ClassObj.App.Token = AppToken Then
                    uAsNotification(mActive.TagAt(i)).Quit
                    uRemove i

                End If
            Next i
        End If
    End If

End Sub

Public Function SetAttribute(ByVal Id As Long, ByVal WhatToChange As SNARL_ATTRIBUTES, ByVal Value As String) As M_RESULT
Dim pn As TNotification

    ' /* must return a melon error for arcane legacy reasons */

    If Not Find(Id, "", "", "", pn) Then
        g_Debug "TNotificationRoster.SetAttribute(): notification " & CStr(Id) & " not found", LEMON_LEVEL_CRITICAL
        SetAttribute = M_NOT_FOUND
        Exit Function

    End If

    ' /* */

    SetAttribute = M_OK

    Select Case WhatToChange

    Case SNARL_ATTRIBUTE_TITLE
        pn.SetTitle Value

    Case SNARL_ATTRIBUTE_TEXT
        pn.SetText Value

    Case SNARL_ATTRIBUTE_ICON
        pn.SetIcon Value

    Case SNARL_ATTRIBUTE_TIMEOUT
        pn.SetTimeout Val(Value)

    Case SNARL_ATTRIBUTE_SOUND
        g_Debug "TNotificationRoster.SetAttribute(): SNARL_ATTRIBUTE_SOUND cannot be changed", LEMON_LEVEL_WARNING
        SetAttribute = M_NOT_IMPLEMENTED

    Case SNARL_ATTRIBUTE_ACK
        pn.SetCallback Value

    Case SNARL_ATTRIBUTE_MENU
        g_Debug "TNotificationRoster.SetAttribute(): SNARL_ATTRIBUTE_MENU is now deprecated", LEMON_LEVEL_WARNING
        pn.SetMenu Value

    Case Else
        g_Debug "TNotificationRoster.SetAttribute(): WhatToChange(" & CStr(WhatToChange) & ") is unknown", LEMON_LEVEL_CRITICAL
        SetAttribute = M_INVALID_ARGS

    End Select

End Function

Public Sub ShowMissedPanel()

    If (mMissedPanel Is Nothing) Then _
        Set mMissedPanel = New TMissedNotificationsPanel

    mMissedPanel.Go

End Sub

Public Function ActualMissedCount() As Long

    ActualMissedCount = mActualMissed

End Function

Public Sub ClearMissedList()

    ' /* clear the missed list and reset counters */

    mMissedX.MakeEmpty
    ResetMissedList

End Sub

Public Sub ResetMissedList()

    ' /* this tags the existing missed notifications as seen and resets
    '    the manual missed count tracker - it does NOT clear the list */

    mActualMissed = 0

    ' /* mark existing notifications as seen */

Dim pItem As TNotification

    With mMissedX
        .Rewind
        Do While .GetNextTag(pItem) = B_OK
            pItem.MarkAsSeen

        Loop

    End With

    ' /* tell the tray icon */

    frmAbout.bMissedNotificationsChanged

End Sub

Public Function HitTest(ByVal x As Long, ByVal y As Long) As Long

    If mActive.CountItems = 0 Then _
        Exit Function

Dim pn As TNotification
Dim px As BPoint
Dim i As Long

    Set px = new_BPoint(x, y)

    With mActive
        For i = 1 To .CountItems
            If uAsNotification(.TagAt(i)).Window.Frame.Contains(px) Then
                HitTest = i
                Exit Function

            End If
        Next i

    End With

End Function

Public Function NotificationAt(ByVal Index As Long) As CSnarlWindow

    If (Index > 0) And (Index <= mActive.CountItems) Then _
        Set NotificationAt = uAsNotification(mActive.TagAt(Index)).Window

End Function

Private Function uGetClassFlagsStr(ByRef Class As TAlert, ByVal ForcePriority As Boolean) As String
'Dim dw As E_ALERT_CLASS_FLAGS
'
'    If Not (Class Is Nothing) Then
'        dw = Class.GetFlags()
'
'        If (dw And E_CLASS_PRIORITY_NOTIFICATION) Then _
'            uGetClassFlagsStr = uGetClassFlagsStr & "p"
'
'        If (dw And E_CLASS_NOTIFICATION_ONLY) Then _
'            uGetClassFlagsStr = uGetClassFlagsStr & "n"
'
'    End If
'
'    If ForcePriority Then _
'        uGetClassFlagsStr = uGetClassFlagsStr & "p"

End Function

Private Function uFindbyStyleName(ByVal Style As String, ByRef Notification As TNotification) As Boolean

    ' /* only searches active list */

    If (mActive.CountItems = 0) Or (Style = "") Then _
        Exit Function

    Style = LCase$(Style)

Dim pn As TNotification

    With mActive
        .Rewind
        Do While .GetNextTag(pn) = B_OK
            If pn.StyleNameLCase() = Style Then
                Set Notification = pn
                uFindbyStyleName = True
                Exit Function

            End If
        Loop

    End With

End Function

Private Function uGetNotificationPosition(ByRef Window As CSnarlWindow, ByVal Position As E_START_POSITIONS, ByRef Frame As BRect, ByRef Column As Long) As Boolean

    If (Window Is Nothing) Then _
        Exit Function

Dim pv As mfxView

    Set pv = Window.GetView()
    If (pv Is Nothing) Then _
        Exit Function

    Column = 1

Dim ps As SIZEL

    With pv
        ps.cx = .Width
        ps.cy = .Height

    End With

    ' /* TODO: needs to take into account styles which veto the dropshadow */

    ps.cx = ps.cx + (2 * Val(g_ConfigGet("margin_spacing")))
    ps.cy = ps.cy + (2 * Val(g_ConfigGet("margin_spacing")))

Dim rcwa As RECT
Const MARGIN = 4

    g_GetWorkArea rcwa
    g_RectInset rcwa, MARGIN, MARGIN

    ' /* determine starting position */

    Set Frame = new_BRect(0, 0, ps.cx - 1, ps.cy - 1)

    If Position = E_START_DEFAULT_POS Then _
        Position = Val(g_ConfigGet("default_position"))

    Select Case Position

    Case E_START_TOP_LEFT
        Frame.OffsetBy rcwa.Left, rcwa.Top

    Case E_START_TOP_RIGHT
        Frame.OffsetBy rcwa.Right - ps.cx, rcwa.Top

    Case E_START_BOTTOM_LEFT
        Frame.OffsetBy rcwa.Left, rcwa.Bottom - ps.cy

    Case Else
        ' /* bottom right */
        Frame.OffsetBy rcwa.Right - ps.cx, rcwa.Bottom - ps.cy

    End Select

    ' /* loop looking for free space */

Dim p As Long
Const STEP_SIZE = 8

    Do While Not uIsFreeSpace(Frame)

        If (Position = E_START_TOP_LEFT) Or (Position = E_START_TOP_RIGHT) Then
            ' /* notifications run from top to bottom */

            Frame.OffsetBy 0, STEP_SIZE
            If Frame.Bottom > rcwa.Bottom Then
                ' /* notification would extend past bottom of work area so move to top of next column... */

                Frame.Top = rcwa.Top
                Frame.Bottom = Frame.Top + ps.cy

                If Position = E_START_TOP_RIGHT Then
                    ' /* ...which is to the left */
                    p = uFindExtent(Column, True) - STEP_SIZE
                    Frame.Right = p
                    Frame.Left = Frame.Right - ps.cx

                    If Frame.Left < rcwa.Left Then
                        ' /* going off-screen horizontally */
                        g_Debug "uGetNotificationPosition(): out of screen space", LEMON_LEVEL_CRITICAL
                        Exit Function

                    End If

                Else
                    ' /* ...which is to the right */
                    p = uFindExtent(Column, False) + STEP_SIZE
                    Frame.Left = p
                    Frame.Right = Frame.Left + ps.cx

                    If Frame.Right > rcwa.Right Then
                        ' /* going off-screen horizontally */
                        g_Debug "uGetNotificationPosition(): out of screen space", LEMON_LEVEL_CRITICAL
                        Exit Function

                    End If

                End If

                Column = Column + 1     ' // column count always increments

            End If

        Else
            ' /* notifications run from bottom to top */

            Frame.OffsetBy 0, -STEP_SIZE
            If Frame.Top < rcwa.Top Then
                ' /* notification would extend past top of work area so move to bottom of next column... */

                Frame.Top = rcwa.Bottom - ps.cy
                Frame.Bottom = Frame.Top + ps.cy

                If Position = E_START_BOTTOM_RIGHT Then
                    ' /* ...which is to the left */
                    p = uFindExtent(Column, True) - STEP_SIZE
                    Frame.Right = p
                    Frame.Left = Frame.Right - ps.cx

                    If Frame.Left < rcwa.Left Then
                        ' /* going off-screen horizontally */
                        g_Debug "uGetNotificationPosition(): out of screen space", LEMON_LEVEL_CRITICAL
                        Exit Function

                    End If

                Else
                    ' /* ...which is to the right */
                    p = uFindExtent(Column, False) + STEP_SIZE
                    Frame.Left = p
                    Frame.Right = Frame.Left + ps.cx

                    If Frame.Right > rcwa.Right Then
                        ' /* going off-screen horizontally */
                        g_Debug "uGetNotificationPosition(): out of screen space", LEMON_LEVEL_CRITICAL
                        Exit Function

                    End If

                End If

                Column = Column + 1

            End If

        End If
    Loop

    ' /* for 2.1+ style can veto position */

Dim pt As POINTAPI

    pt.x = Frame.Left
    pt.y = Frame.Top

    Window.AdjustPosition pt.x, pt.y

    Frame.Left = pt.x
    Frame.Top = pt.y

    uGetNotificationPosition = True

End Function

Private Sub uSafeGetStyle(ByVal StyleToUse As String, ByRef Style As TStyle, ByRef Instance As IStyleInstance)

    ' /* try to create a new TStyle from the provided Style name */

    If Not (g_StyleRoster Is Nothing) Then _
        g_StyleRoster.Find style_GetStyleName(StyleToUse), Style

    ' /* didn't work? use a default one */

    If (Style Is Nothing) Then _
        Set Style = New TStyle

    ' /* create a new instance from the provided Style */

    Set Instance = Style.CreateInstance()

    ' /* didn't work? use an internal style instance */

    If (Instance Is Nothing) Then _
        Set Instance = New TInternalStyle

End Sub

Private Function uCopyTo(ByRef Info As T_NOTIFICATION_INFO) As Boolean

    ' /* get the appropriate instance for the notification before we do anything else
    '    we must pass a valid style instance to CSnarlWindow, even if the style
    '    roster is invalid */

Dim psi As IStyleInstance
Dim pStyle As TStyle

    ' /* get the style from the roster here - this is guaranteed to return a usable TStyle */

    uSafeGetStyle Info.StyleName & "/" & Info.SchemeName, pStyle, psi

    ' /* this is just a safety check */

    If (pStyle.Flags And S_STYLE_IS_WINDOWLESS) = 0 Then
        Debug.Print "TNotificationRoster.uCopyTo(): '" & pStyle.Name & "' is not a windowless style"
        Exit Function

    End If

    ' /* get the path to the style - we need this for style-provided icons and sounds */

    Info.IconPath = g_TranslateIconPath(Info.IconPath, pStyle.Path)

Dim sTitle As String
Dim sText As String

    If (pStyle.Flags And S_STYLE_V42_CONTENT) Then

        ' /* title becomes sender name, text is set to the original packed string */

        sTitle = Info.ClassObj.App.Name
        sText = Info.OriginalContent

        ' /* R2.4 Beta 4 - if we should include the hostname, add it here */

        If g_ConfigGet("include_host_name_when_forwarding") = "1" Then _
            sTitle = sTitle & " on " & g_GetComputerName()

    Else

        ' /* pre-V42 model: <title> is either title or app/title pair, <text> is text */

        sText = Info.Text

        ' /* if S_STYLE_WANT_APP_NAME is set, prefix title with the sending app name */

        If (pStyle.Flags And S_STYLE_WANT_APP_NAME) Then
            g_Debug "TNotificationRoster.uCopyTo(): style wants application name"
            sTitle = Info.ClassObj.App.Name

            ' /* R2.4 Beta 4 - if app name is requested, make it "<app> on <computer>" */

            If g_ConfigGet("include_host_name_when_forwarding") = "1" Then _
                sTitle = sTitle & " on " & g_GetComputerName()

            sTitle = sTitle & "#?" & Info.Title

        Else
            ' /* R2.4 Beta 4 - make the title component "<title> (from <computer>)" */

            If g_ConfigGet("include_host_name_when_forwarding") = "1" Then _
                sTitle = Info.Title & " (from " & g_GetComputerName() & ")"

        End If

    End If

Dim pContent As notification_info

    g_Debug "TNotificationRoster.uCopyTo(): sending to '" & pStyle.Name & "'"

    LSet pContent = pStyle.CreateNotificationInfo(sTitle, sText, Info.IconPath, _
                                               Info.SchemeName, _
                                               IIf(Info.Priority > 0, S_NOTIFICATION_IS_PRIORITY, 0))

    psi.UpdateContent pContent

    ' /* R2.4 Beta 4: if the style set notificaton_info->flags to -1 we should pass the notification
    '    back around for on-screen display */

    uCopyTo = (pContent.Flags = -1)

End Function

Private Function uFindMergeCandidate(ByRef Class As TAlert, ByVal Title As String, ByRef Notification As TNotification) As Boolean

    ' /* only active (visible) notifications can be considered here */

    If (Class Is Nothing) Or (mActive.CountItems = 0) Then _
        Exit Function

Dim pn As TNotification

    With mActive
        .Rewind
        Do While .GetNextTag(pn) = B_OK
            If pn.CanMerge(Class, Title) Then
                Set Notification = pn
                uFindMergeCandidate = True
                Exit Function

            End If
        
        Loop

    End With

End Function

Public Function AddAction(ByVal Token As Long, ByRef Data As BPackedData) As Long
Dim pn As TNotification

    ' /* uFind sets lasterror for us */

    If Not Find(Token, Data.ValueOf("uid"), Data.ValueOf("app-sig"), Data.ValueOf("password"), pn) Then _
        Exit Function

    Debug.Print "*** AddAction: found " & Token & " - state = " & CStr(pn.State)

    ' /* param checking... */

    If (Data Is Nothing) Then
        gSetLastError SNARL_ERROR_FAILED
        Exit Function

    End If

    If (Data.ValueOf("label") = "") Or (Data.ValueOf("cmd") = "") Then
        g_Debug "TNotificationRoster.AddAction(): missing label and/or command", LEMON_LEVEL_CRITICAL
        gSetLastError SNARL_ERROR_ARG_MISSING
        Exit Function

    End If

    pn.AddAction Data.ValueOf("label"), Data.ValueOf("cmd")
    AddAction = -1

End Function

Public Function ClearActions(ByVal Token As Long, ByRef Data As BPackedData) As Long
Dim pn As TNotification

    ' /* uFind sets lasterror for us */

    If Find(Token, Data.ValueOf("uid"), Data.ValueOf("app-sig"), Data.ValueOf("password"), pn) Then
        pn.ClearActions
        ClearActions = -1

    End If

End Function

Private Sub uAddToMissedList(ByRef Info As T_NOTIFICATION_INFO)

    If Info.Token = 0 Then
        ' /* allocate a token now */
        mToken = mToken + 1
        Info.Token = mToken

    End If

Dim pNewList As BTagList

    ' /* create new list */

    Set pNewList = new_BTagList()

    ' /* create the new node and add it to the head */

Dim pn As TNotification

    Set pn = New TNotification
    pn.SetTo NS_MISSED, Info
    pNewList.Add pn

    With mMissedX
        .Rewind
        Do While .GetNextTag(pn) = B_OK
            pNewList.Add pn

        Loop

    End With

    ' /* assign our new list as the missed list */

    Set mMissedX = pNewList
    mActualMissed = mActualMissed + 1

    ' /* if the panel is open, update it */

    If Not (mMissedPanel Is Nothing) Then _
        mMissedPanel.UpdateList

    ' /* tell the tray icon */

    frmAbout.bMissedNotificationsChanged

End Sub

Public Function Update(ByVal Token As Long, ByRef Args As BPackedData) As Boolean
Dim pn As TNotification

    If Find(Token, Args.ValueOf("uid"), Args.ValueOf("app-sig"), Args.ValueOf("password"), pn) Then _
        Update = pn.UpdateOrMerge(Args, False)

End Function

Public Function Merge(ByVal Token As Long, ByRef Args As BPackedData) As Boolean
Dim pn As TNotification

    If Find(Token, Args.ValueOf("uid"), Args.ValueOf("app-sig"), Args.ValueOf("password"), pn) Then _
        Merge = pn.UpdateOrMerge(Args, True)

End Function

Public Sub CheckMissed()

    If mActualMissed < 1 Then _
        Exit Sub

Static iToken As Long

    If iToken Then _
        Me.Hide iToken, "", "", gSnarlPassword

    iToken = g_PrivateNotify(SNARL_CLASS_GENERAL, "While you were away...", _
                             "You missed " & CStr(mActualMissed) & " notification" & IIf(mActualMissed > 1, "s", ""), _
                             -1, _
                             g_MakePath(App.Path) & "etc\icons\snarl.png", _
                             , _
                             "!snarl show_missed_panel")

End Sub

Private Function uAsNotification(ByRef Obj As BTagItem) As TNotification

    Set uAsNotification = Obj

End Function

Public Function MissedList() As BTagList

    Set MissedList = mMissedX

End Function

Public Sub UpdateMissedPanel()

    If Not (mMissedPanel Is Nothing) Then _
        mMissedPanel.UpdateList

End Sub

Public Sub CloseMostRecent()
Dim pn As TNotification

    If mActive.CountItems > 0 Then
        Set pn = mActive.TagAt(mActive.CountItems)
        pn.Remove

    End If

End Sub

Public Function ActiveList() As BTagList

    Set ActiveList = mActive

End Function

Public Function WaitList() As BTagList

    Set WaitList = mWait

End Function
